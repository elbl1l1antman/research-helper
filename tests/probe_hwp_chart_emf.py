"""Local PowerPoint-to-Hancom integration probe using synthetic chart data."""

import argparse
import ctypes as C
from ctypes import wintypes as W
import json
from pathlib import Path
import shutil
import sys
import time

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from report_automation_engine.hwp_chart_probe import verify_chart_document
from report_automation_engine.hwp_com_writer import (
    close_hwp, create_hwp_object, insert_text, leave_hwp_table, new_report,
    open_document, run_action, save_as_hwpx, set_parameter_item, set_visible,
)


def make_deck(path):
    from pptx import Presentation
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.util import Inches

    deck = Presentation()
    for kind, categories, values in (
        (XL_CHART_TYPE.COLUMN_CLUSTERED, ('긴 보기 이름\n줄바꿈 예시', '보기 B', '보기 C'), (63.3, 25.0, 11.7)),
        (XL_CHART_TYPE.DOUGHNUT, ('있음', '없음'), (63.3, 36.7)),
    ):
        slide = deck.slides.add_slide(deck.slide_layouts[6])
        data = CategoryChartData()
        data.categories = categories
        data.add_series('전체', values)
        shape = slide.shapes.add_chart(kind, Inches(1), Inches(1), Inches(5.7), Inches(2.5), data)
        shape.chart.has_title = False
        shape.chart.plots[0].has_data_labels = True
        if kind == XL_CHART_TYPE.COLUMN_CLUSTERED:
            shape.chart.plots[0].gap_width = 140
    deck.save(path)


def restrict_clipboard_to_emf():
    # Duplicate the handle before EmptyClipboard releases PowerPoint's copy.
    import win32clipboard as clip
    import win32con
    import win32gui

    user, gdi = C.windll.user32, C.windll.gdi32
    user.GetClipboardData.argtypes, user.GetClipboardData.restype = (W.UINT,), W.HANDLE
    user.SetClipboardData.argtypes, user.SetClipboardData.restype = (W.UINT, W.HANDLE), W.HANDLE
    gdi.CopyEnhMetaFileW.argtypes, gdi.CopyEnhMetaFileW.restype = (W.HANDLE, W.LPCWSTR), W.HANDLE
    gdi.DeleteEnhMetaFile.argtypes, gdi.DeleteEnhMetaFile.restype = (W.HANDLE,), W.BOOL
    owner = win32gui.CreateWindow('STATIC', '', 0, 0, 0, 0, 0, 0, 0, 0, None)
    fmt = win32con.CF_ENHMETAFILE
    try:
        for attempt in range(3):
            try:
                clip.OpenClipboard(owner)
                break
            except Exception:
                if attempt == 2:
                    raise
                time.sleep(0.2)
        try:
            source = user.GetClipboardData(fmt)
            duplicate = gdi.CopyEnhMetaFileW(source, None) if source else None
            if not duplicate:
                raise RuntimeError('PowerPoint did not provide CF_ENHMETAFILE')
            try:
                clip.EmptyClipboard()
                if not user.SetClipboardData(fmt, duplicate):
                    raise C.WinError()
            except Exception:
                gdi.DeleteEnhMetaFile(duplicate)
                raise
        finally:
            clip.CloseClipboard()
    finally:
        win32gui.DestroyWindow(owner)


def add_container(hwp, title, prefix, report, paste=None):
    action = hwp.HAction
    sec = hwp.HParameterSet.HSecDef
    if not action.GetDefault('PageSetup', sec.HSet):
        raise RuntimeError('PageSetup unavailable')
    page = sec.PageDef
    body = (page.PaperHeight if page.Landscape else page.PaperWidth) - page.LeftMargin - page.RightMargin - page.GutterLen
    table = hwp.HParameterSet.HTableCreation
    action.GetDefault('TableCreate', table.HSet)
    table.Rows, table.Cols = 1, 1
    table.WidthType, table.HeightType = 2, 1
    table.WidthValue = body - 800
    table.CreateItemArray('ColWidth', 1)
    margins = table.TableProperties.CellMarginLeft + table.TableProperties.CellMarginRight
    table.ColWidth.SetItem(0, body - 800 - margins)
    if not action.Execute('TableCreate', table.HSet):
        raise RuntimeError('TableCreate failed')
    control = hwp.ParentCtrl
    if not run_action(hwp, 'TableCellBlock', report, 'probe'):
        raise RuntimeError('TableCellBlock failed')
    border = hwp.HParameterSet.HCellBorderFill
    action.GetDefault('CellBorderFill', border.HSet)
    for target in (border, border.SelCellsBorderFill):
        for side in ('Left', 'Right', 'Top', 'Bottom'):
            set_parameter_item(target, 'BorderType' + side, hwp.HwpLineType('None'))
    if not action.Execute('CellBorderFill', border.HSet):
        raise RuntimeError('CellBorderFill failed')
    run_action(hwp, 'Cancel', report, 'probe')
    cell = hwp.GetPosBySet()

    def select(target):
        run_action(hwp, 'Cancel', report, 'probe')
        if not hwp.SetPosBySet(target.GetAnchorPos(0)):
            raise RuntimeError('Could not reach object anchor')
        hwp.FindCtrl()

    select(control)
    for step in ('ShapeObjCaption', 'MoveListBegin', 'MoveSelListEnd', 'Delete'):
        if not run_action(hwp, step, report, 'probe'):
            raise RuntimeError(f'{step} failed')
    insert_text(hwp, f'<그림 {prefix}-', report)
    number = hwp.CreateSet('AutoNum')
    number.SetItem('NumType', 3)
    if hwp.InsertCtrl('atno', number).CtrlID != 'atno':
        raise RuntimeError('Picture autonumber failed')
    insert_text(hwp, '> ' + title, report)
    select(control)
    props = hwp.HParameterSet.HShapeObject
    if not action.GetDefault('TablePropertyDialog', props.HSet):
        raise RuntimeError('TablePropertyDialog unavailable')
    props.ShapeCaption.Side = hwp.SideType('Top')
    props.ShapeCaption.CapFullSize = 1
    props.TreatAsChar, props.AffectsLine, props.NumberingType = 1, 1, 1
    if not action.Execute('TablePropertyDialog', props.HSet):
        raise RuntimeError('Figure container settings failed')
    if paste:
        run_action(hwp, 'Cancel', report, 'probe')
        if not hwp.SetPosBySet(cell) or hwp.ParentCtrl.CtrlID != 'tbl':
            raise RuntimeError('Could not return to chart cell')
        picture = paste()
        picture_props = picture.Properties
        picture_props.SetItem('NumberingType', 0)
        picture_props.SetItem('TreatAsChar', 1)
        width, height = (int(picture_props.Item(key)) for key in ('Width', 'Height'))
        if min(width, height) <= 0:
            raise RuntimeError('Invalid pasted picture geometry')
        available = body - 1000 - margins
        if width > available:
            picture_props.SetItem('Width', available)
            picture_props.SetItem('Height', round(height * available / width))
        picture.Properties = picture_props
    leave_hwp_table(hwp, control, report)
    run_action(hwp, 'BreakPara', report, 'probe')


def run(output, method):
    from win32com.client import DispatchEx

    output.mkdir(parents=True, exist_ok=True)
    template = output / 'probe_template.hwpx'
    pptx, hwpx = output / 'chart_probe.pptx', output / 'chart_probe.hwpx'
    if any(path.exists() for path in (template, pptx, hwpx)):
        raise FileExistsError('Use a fresh output directory for each probe')
    shutil.copyfile(Path(__file__).resolve().parents[1] / 'templates/report_template_basic.hwpx', template)
    make_deck(pptx)
    report = new_report(pptx, pptx, template, hwpx, method == 'special')
    checkpoint = output / 'probe_progress.json'
    print('starting Hancom', flush=True)
    hwp = create_hwp_object(report, checkpoint, dispatch_mode='dispatch_ex')
    app = presentation = None
    try:
        print('opening template copy', flush=True)
        set_visible(hwp, method == 'special', report, checkpoint)
        open_document(hwp, template, report, checkpoint)
        run_action(hwp, 'MoveDocEnd', report, 'probe')
        run_action(hwp, 'BreakPara', report, 'probe')
        if method != 'empty':
            print('opening PowerPoint', flush=True)
            app = DispatchEx('PowerPoint.Application')
            presentation = app.Presentations.Open(str(pptx), -1, 0, 0)
        for index, title in enumerate(('세로 막대 검증', '도넛 검증', '수동 보정 대기'), 1):
            print(f'container {index}', flush=True)

            def paste():
                presentation.Slides(index).Shapes(1).Copy()
                print('PowerPoint copied chart', flush=True)
                if method == 'emf-only':
                    restrict_clipboard_to_emf()
                action = 'PasteSpecial' if method == 'special' else 'Paste'
                if method == 'special':
                    print('Manual probe: choose 그림 (메타파일) in Hancom 골라 붙이기.', flush=True)
                anchor = hwp.GetPosBySet()
                if not run_action(hwp, action, report, 'probe'):
                    raise RuntimeError(f'{action} failed')
                run_action(hwp, 'Cancel', report, 'probe')
                if not hwp.SetPosBySet(anchor):
                    raise RuntimeError('Could not return to pasted object anchor')
                hwp.FindCtrl()
                picture = hwp.CurSelectedCtrl
                if picture is None or picture.CtrlID.strip() != 'gso':
                    raise RuntimeError('Pasted object is not a picture control')
                return picture

            add_container(hwp, title, '3', report, paste if presentation is not None and index < 3 else None)
        save_as_hwpx(hwp, hwpx, report)
        try:
            figures = verify_chart_document(hwpx, prefix='3', expected=(method != 'empty', method != 'empty', False))
            result = {'status': 'ready', 'figures': figures}
        except ValueError as error:
            from zipfile import ZipFile

            with ZipFile(hwpx) as archive:
                media = [name for name in archive.namelist() if name.startswith('BinData/')]
            result = {'status': 'blocked', 'reason': str(error), 'saved_media': media}
            raise
        finally:
            (output / 'verification.json').write_text(json.dumps(
                {'method': method, 'manual_selection': method == 'special', **result},
                ensure_ascii=False, indent=2), encoding='utf-8')
        print(hwpx)
    finally:
        if presentation is not None:
            presentation.Close()
        if app is not None and app.Presentations.Count == 0:
            app.Quit()
        close_hwp(hwp, report)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--out-dir', type=Path, required=True)
    parser.add_argument('--method', choices=('empty', 'emf-only', 'special'), default='emf-only',
                        help='special opens a modal dialog and requires a human selection')
    args = parser.parse_args()
    run(args.out_dir.resolve(), args.method)
