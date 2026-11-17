"""Render editable survey charts with PowerPoint and place only EMFs in Hancom."""

from collections import Counter
from contextlib import contextmanager
from hashlib import sha256
from io import BytesIO
import json
from pathlib import Path
import re
from tempfile import mkdtemp, TemporaryDirectory

from pptx import Presentation

try:
    from .hwp_chart_probe import _is_emf, verify_chart_document
    from .survey_chart_style import REVIEW_CHART_PREFIX, survey_chart_plan, validate_local_deck, write_chart_review
except ImportError:
    from hwp_chart_probe import _is_emf, verify_chart_document
    from survey_chart_style import REVIEW_CHART_PREFIX, survey_chart_plan, validate_local_deck, write_chart_review


@contextmanager
def _powerpoint_deck(deck):
    from win32com.client import DispatchEx

    app = DispatchEx('PowerPoint.Application')
    presentation = app.Presentations.Open(str(deck), -1, 0, 0)
    try:
        yield presentation
    finally:
        try:
            presentation.Close()
        except Exception:
            pass
        # PowerPoint can return the user's existing application even with DispatchEx.


def _export_emf(shape, path):
    try:
        shape.Export(str(path), 5)  # ppShapeFormatEMF
        data = path.read_bytes()
        if not _is_emf(data):
            raise RuntimeError('PowerPoint에서 EMF를 내보내지 못했습니다.')
        return sha256(data).hexdigest()
    except Exception:
        path.unlink(missing_ok=True)
        raise


def export_chart_emfs(plan, directory, style_path):
    """Map only ready charts to EMF files; review/held charts get empty slots."""
    paths, errors = {}, {}
    ready = [item for item in plan if item['status'] == 'ready']
    if not ready:
        return paths, errors
    deck = write_chart_review({}, directory / 'survey_charts.pptx', style_path, plan=ready)
    try:
        with _powerpoint_deck(deck) as presentation:
            if presentation.Slides.Count != len(ready):
                raise RuntimeError('차트 검토 슬라이드 수가 계획과 다릅니다.')
            for index, item in enumerate(ready, 1):
                key = item['key']
                try:
                    shapes = [shape for shape in presentation.Slides(index).Shapes if shape.HasChart]
                    if len(shapes) != 1:
                        raise RuntimeError('슬라이드에 차트 객체가 하나여야 합니다.')
                    path = directory / f'survey_chart_{index}.emf'
                    _export_emf(shapes[0], path)
                    paths[key] = path
                except Exception as exc:
                    errors[key] = str(exc)
    except Exception as exc:
        for item in ready:
            if item['key'] not in paths:
                errors[item['key']] = f'PowerPoint 차트 준비 실패: {exc}'
    return paths, errors


def _verify_reviewed_series(chart, item):
    if chart.ChartData.IsLinked or chart.SeriesCollection().Count != 1:
        raise ValueError('외부 링크 없는 데이터 계열 한 개를 유지하세요.')
    series = chart.SeriesCollection(1)
    values = tuple(series.Values)
    labels = lambda items: tuple(''.join(str(label).split()) for label in items)
    if (any(isinstance(value, bool) or not isinstance(value, (int, float)) for value in values)
            or values != tuple(item['values'])):
        raise ValueError('차트 수치 또는 순서가 집계표와 다릅니다.')
    if labels(series.XValues) != labels(item['labels']):
        raise ValueError('차트 보기명 또는 순서가 집계표와 다릅니다. 줄바꿈과 공백만 바꿀 수 있습니다.')


def export_reviewed_charts(package, deck, output_dir):
    deck = Path(deck).resolve()
    if deck.suffix.lower() != '.pptx':
        raise ValueError('편집한 차트 검토 PPTX(.pptx)를 선택하세요.')
    payload = deck.read_bytes()
    pptx = Presentation(BytesIO(payload))
    validate_local_deck(pptx)
    plan = {item['key']: item for item in survey_chart_plan(package) if item['status'] in {'ready', 'review'}}
    entries = [(index, shape.name) for index, slide in enumerate(pptx.slides, 1)
               for shape in slide.shapes if shape.has_chart or shape.name.startswith(REVIEW_CHART_PREFIX)]
    if not any(name.startswith(REVIEW_CHART_PREFIX) for _, name in entries):
        raise ValueError('문항 식별자가 없습니다. 런처에서 차트 검토 PPTX를 새로 만들고 개체 이름을 유지하세요.')
    repeats = Counter(name for _, name in entries)
    slides = Counter(index for index, _ in entries)
    directory = Path(mkdtemp(prefix='survey_chart_emf_', dir=Path(output_dir).resolve()))
    results, pending = [], []
    for index, name in entries:
        key = name[len(REVIEW_CHART_PREFIX):] if name.startswith(REVIEW_CHART_PREFIX) else ''
        item = plan.get(key)
        result = dict(slide=index, table_key=key, title=item['title'] if item else '', status='rejected')
        results.append(result)
        if repeats[name] > 1:
            result['reason'] = '문항 식별자가 중복되었습니다.'
        elif slides[index] != 1:
            result['reason'] = '검토 슬라이드에 차트 객체가 하나여야 합니다.'
        elif not item or not re.fullmatch(r'[A-Za-z0-9_-]+', key):
            result['reason'] = '이 집계표의 유효한 차트 문항 식별자가 아닙니다.'
        else:
            pending.append((index, name, item, result))

    if pending:
        try:
            # Open the exact bytes inspected above, even when the source deck is open in PowerPoint.
            with TemporaryDirectory() as work:
                snapshot = Path(work) / 'reviewed.pptx'
                snapshot.write_bytes(payload)
                with _powerpoint_deck(snapshot) as presentation:
                    for index, name, item, result in pending:
                        try:
                            shape = presentation.Slides(index).Shapes(name)
                            if not shape.HasChart:
                                raise ValueError('편집 가능한 차트 객체가 아닙니다.')
                            _verify_reviewed_series(shape.Chart, item)
                            path = directory / f'chart_{index:03d}_{item["key"]}.emf'
                            result.update(status='exported', emf=path.name, sha256=_export_emf(shape, path))
                        except Exception as exc:
                            result['reason'] = str(exc)
        except Exception as exc:
            for _, _, _, result in pending:
                if result['status'] == 'rejected' and 'reason' not in result:
                    result['reason'] = f'PowerPoint 차트 준비 실패: {exc}'

    report = directory / 'emf_export_report.json'
    exported = sum(item['status'] == 'exported' for item in results)
    report.write_text(json.dumps(dict(presentation=str(deck), exported_count=exported,
                                      rejected_count=len(results) - exported, charts=results),
                                 ensure_ascii=False, indent=2), encoding='utf-8')
    return report


def verify_saved_figures(path, prefix, results):
    from zipfile import ZipFile

    figures = verify_chart_document(path, prefix=prefix,
                                    expected=tuple(item['status'] == 'inserted' for item in results),
                                    titles=tuple(item['title'] for item in results))
    with ZipFile(path) as archive:
        for result, figure in zip(results, figures):
            if figure['emf'] and sha256(archive.read(figure['emf'])).hexdigest() != result['sha256']:
                raise ValueError(f"{result['table_key']}: 저장된 EMF가 PowerPoint 원본과 다릅니다.")
    return figures


def insert_figure(hwp, title, prefix, emf, report, body_width):
    # Imported at call time because the writer calls this module after validation.
    if __package__:
        from .hwp_com_writer import insert_text, leave_hwp_table, run_action, same_hwp_control, set_parameter_item
    else:
        from hwp_com_writer import insert_text, leave_hwp_table, run_action, same_hwp_control, set_parameter_item
    action = hwp.HAction

    def checked(step):
        if not run_action(hwp, step, report, 'chart'):
            raise RuntimeError(step)

    paragraph = hwp.HParameterSet.HParaShape
    if not action.GetDefault('ParagraphShape', paragraph.HSet):
        raise RuntimeError('차트 문단 서식을 읽지 못했습니다.')
    paragraph.LineSpacingType, paragraph.LineSpacing = 0, 100
    if not action.Execute('ParagraphShape', paragraph.HSet):
        raise RuntimeError('차트 문단 줄간격을 적용하지 못했습니다.')
    table = hwp.HParameterSet.HTableCreation
    action.GetDefault('TableCreate', table.HSet)
    table.Rows, table.Cols = 1, 1
    table.WidthType, table.HeightType = 2, 1
    table.WidthValue = body_width - 800
    table.CreateItemArray('ColWidth', 1)
    margins = table.TableProperties.CellMarginLeft + table.TableProperties.CellMarginRight
    inner_width = body_width - 1000 - margins
    if inner_width <= 0:
        raise RuntimeError('차트 셀의 가용 폭이 없습니다.')
    table.ColWidth.SetItem(0, body_width - 800 - margins)
    if not action.Execute('TableCreate', table.HSet):
        raise RuntimeError('TableCreate')
    control = hwp.ParentCtrl
    checked('TableCellBlock')
    border = hwp.HParameterSet.HCellBorderFill
    action.GetDefault('CellBorderFill', border.HSet)
    for target in (border, border.SelCellsBorderFill):
        for side in ('Left', 'Right', 'Top', 'Bottom'):
            set_parameter_item(target, 'BorderType' + side, hwp.HwpLineType('None'))
    if not action.Execute('CellBorderFill', border.HSet):
        raise RuntimeError('CellBorderFill')
    run_action(hwp, 'Cancel', report, 'chart')
    cell = hwp.GetPosBySet()

    def select_table():
        run_action(hwp, 'Cancel', report, 'chart')
        if not hwp.SetPosBySet(control.GetAnchorPos(0)):
            raise RuntimeError('차트 표 앵커를 찾지 못했습니다.')
        hwp.FindCtrl()
        if not same_hwp_control(hwp.CurSelectedCtrl, control):
            raise RuntimeError('차트 표를 선택하지 못했습니다.')

    select_table()
    for step in ('ShapeObjCaption', 'MoveListBegin', 'MoveSelListEnd', 'Delete'):
        checked(step)
    insert_text(hwp, f'<그림 {prefix}-', report)
    number = hwp.CreateSet('AutoNum')
    number.SetItem('NumType', 3)
    if hwp.InsertCtrl('atno', number).CtrlID != 'atno':
        raise RuntimeError('그림 자동 번호를 넣지 못했습니다.')
    insert_text(hwp, '> ' + title, report)
    select_table()
    props = hwp.HParameterSet.HShapeObject
    if not action.GetDefault('TablePropertyDialog', props.HSet):
        raise RuntimeError('TablePropertyDialog')
    props.ShapeCaption.Side = hwp.SideType('Top')
    props.ShapeCaption.CapFullSize = 1
    props.TreatAsChar, props.AffectsLine, props.NumberingType = 1, 1, 1
    if not action.Execute('TablePropertyDialog', props.HSet):
        raise RuntimeError('차트 표를 글자처럼 취급하는 설정을 적용하지 못했습니다.')
    if emf:
        run_action(hwp, 'Cancel', report, 'chart')
        if not hwp.SetPosBySet(cell) or not same_hwp_control(hwp.ParentCtrl, control):
            raise RuntimeError('차트 표의 셀로 돌아오지 못했습니다.')
        # InsertPicture preserves EMF; the interactive PasteSpecial route converts it to WMF.
        picture = hwp.InsertPicture(str(emf), True, 0)
        if not picture or picture.CtrlID.strip() != 'gso':
            raise RuntimeError('EMF 그림을 삽입하지 못했습니다.')
        picture_props = picture.Properties
        picture_props.SetItem('NumberingType', 0)
        picture_props.SetItem('TreatAsChar', 1)
        width, height = (int(picture_props.Item(key)) for key in ('Width', 'Height'))
        if min(width, height) <= 0:
            raise RuntimeError('삽입한 차트의 크기를 확인하지 못했습니다.')
        section = hwp.HParameterSet.HSecDef
        if not action.GetDefault('PageSetup', section.HSet):
            raise RuntimeError('차트의 본문 높이를 확인하지 못했습니다.')
        page = section.PageDef
        inner_height = (page.PaperWidth if page.Landscape else page.PaperHeight) - page.TopMargin - page.BottomMargin - page.HeaderLen - page.FooterLen - 4500
        if inner_height <= 0:
            raise RuntimeError('차트의 본문 가용 높이가 없습니다.')
        scale = min(1, inner_width / width, inner_height / height)
        if scale < 1:
            picture_props.SetItem('Width', max(1, round(width * scale)))
            picture_props.SetItem('Height', max(1, round(height * scale)))
        picture.Properties = picture_props
    leave_hwp_table(hwp, control, report)
    checked('BreakPara')
