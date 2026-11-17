"""Extract native PowerPoint chart templates and build a compatible report starter."""

import argparse
from collections import Counter
import csv
from io import BytesIO
import json
from pathlib import Path
from tempfile import TemporaryDirectory, mkdtemp
from zipfile import BadZipFile, ZIP_DEFLATED, ZipFile

from lxml import etree
from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.enum.chart import XL_TICK_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.exc import PackageNotFoundError
from pptx.oxml.ns import qn
from pptx.util import Pt

if __package__:
    from .hwp_survey_charts import _powerpoint_deck
    from .survey_chart_style import FONT, STYLE_NAMES, create_chart_style, register_chart_style, validate_local_deck, _prototypes
else:
    from hwp_survey_charts import _powerpoint_deck
    from survey_chart_style import FONT, STYLE_NAMES, create_chart_style, register_chart_style, validate_local_deck, _prototypes


def _chart_shapes(shapes):
    for shape in shapes:
        if shape.has_chart:
            yield shape
        elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _chart_shapes(shape.shapes)


def _com_charts(shapes):
    for shape in shapes:
        if shape.HasChart:
            yield shape
        elif shape.Type == 6:  # msoGroup
            yield from _com_charts(shape.GroupItems)


def _local_name(el):
    return el.tag.rsplit('}', 1)[-1]


def clean_style_package(payload, chart_xml=None):
    """Remove Office document identity while keeping theme colors and fonts."""
    cleaned = BytesIO()
    with ZipFile(BytesIO(payload)) as source, ZipFile(cleaned, 'w', ZIP_DEFLATED) as target:
        previews = set()
        if chart_xml is not None and '_rels/.rels' in source.namelist():
            previews = {rel.get('Target', '').lstrip('/') for rel in etree.fromstring(source.read('_rels/.rels'))
                        if rel.get('Type', '').endswith('/image')}
        for name in source.namelist():
            lower = name.lower()
            if lower.startswith('docprops/thumbnail.'):
                continue
            data = source.read(name)
            if lower.endswith('.xlsx'):
                data = clean_style_package(data)
            elif name == 'chart/chart.xml' and chart_xml is not None:
                data = chart_xml
            elif name in previews:
                # The gallery preview can reveal the old series even without its workbook.
                picture = Image.new('RGB', (48, 48), 'white')
                draw = ImageDraw.Draw(picture)
                for x, y in ((7, 27), (19, 17), (31, 22)):
                    draw.rectangle((x, y, x + 8, 40), fill='#3979A6')
                preview = BytesIO()
                picture.save(preview, format=Image.registered_extensions().get(Path(name).suffix.lower(), 'BMP'))
                data = preview.getvalue()
            elif lower.startswith('docprops/') and lower.endswith('.xml'):
                root = etree.fromstring(data)
                root.attrib.clear()
                for node in list(root):
                    if lower != 'docprops/app.xml' or _local_name(node) not in ('Application', 'AppVersion'):
                        root.remove(node)
                data = etree.tostring(root, encoding='UTF-8')
            elif name in ('_rels/.rels', '[Content_Types].xml'):
                root = etree.fromstring(data)
                for node in list(root):
                    part = (node.get('Target') or node.get('PartName') or '').lstrip('/').lower()
                    if part.startswith('docprops/thumbnail.'):
                        root.remove(node)
                data = etree.tostring(root, encoding='UTF-8')
            elif '/theme' in lower and lower.endswith('.xml'):
                root = etree.fromstring(data)
                for node in root.iter():
                    if _local_name(node) in ('theme', 'themeOverride', 'clrScheme', 'fontScheme', 'fmtScheme') and 'name' in node.attrib:
                        node.set('name', 'Chart Theme')
                data = etree.tostring(root, encoding='UTF-8')
            target.writestr(name, data)
    return cleaned.getvalue()


def _role(chart):
    if len(chart.plots) == len(chart.series) == 1:
        return next((name for name, chart_type in STYLE_NAMES.items() if chart.chart_type == chart_type), '')
    return ''


def _describe_chart(slide_index, shape):
    chart = shape.chart
    root = chart._chartSpace
    layout = {}
    for tag in ('gapWidth', 'overlap', 'holeSize', 'firstSliceAng', 'barDir', 'grouping',
                'min', 'max', 'majorUnit', 'minorUnit', 'legendPos', 'dLblPos'):
        values = [el.get('val') for el in root.iter(qn('c:' + tag))]
        if values:
            layout[tag] = values
    plots = [_local_name(plot._element) for plot in chart.plots]
    try:
        kind = str(chart.chart_type)
        role = _role(chart)
    except (ValueError, NotImplementedError):
        kind, role = ', '.join(plots), ''
    fonts = sorted({el.get('typeface') for tag in ('latin', 'ea', 'cs')
                    for el in root.iter(qn('a:' + tag)) if el.get('typeface')})
    sizes = sorted({int(el.get('sz')) / 100 for tag in ('defRPr', 'rPr', 'endParaRPr')
                    for el in root.iter(qn('a:' + tag)) if (el.get('sz') or '').isdigit()})
    return dict(slide=slide_index, shape_id=shape.shape_id, shape_name=f'차트 {shape.shape_id}', chart_type=kind,
                plot_types=plots, series_count=len(chart.series), prototype=role, fonts=fonts, font_sizes_pt=sizes,
                number_formats=sorted({el.get('formatCode') for el in root.iter(qn('c:numFmt')) if el.get('formatCode')}),
                colors=sorted({el.get('val') for tag in ('srgbClr', 'schemeClr')
                               for el in root.iter(qn('a:' + tag)) if el.get('val')}),
                size_pt=[round(shape.width.pt, 2), round(shape.height.pt, 2)], layout=layout,
                theme_override=any(rel.reltype.endswith('/themeOverride') for rel in chart.part.rels.values()),
                linked=any(rel.is_external for rel in chart.part.rels.values()))


def _clean_crtx(payload):
    """Office removes data by design; verify that before saving a reusable template."""
    with ZipFile(BytesIO(payload)) as source:
        names = source.namelist()
        if any(name.lower().endswith(('.xlsx', '.xls', '.xlsm', '.bin')) for name in names):
            raise ValueError('CRTX에 내장 데이터 또는 실행 파일이 포함되어 있습니다.')
        for name in names:
            if name.endswith('.rels'):
                rels = etree.fromstring(source.read(name))
                if any(rel.get('TargetMode') == 'External' for rel in rels):
                    raise ValueError('외부 연결이 남은 CRTX는 사용할 수 없습니다.')
        chart_name = 'chart/chart.xml'
        root = etree.fromstring(source.read(chart_name))
        if root.tag != qn('c:chartSpace'):
            raise ValueError('올바른 CRTX 차트 서식이 아닙니다.')
        data_tags = {qn('c:' + tag) for tag in ('externalData', 'numCache', 'strCache', 'multiLvlStrCache', 'numLit', 'strLit', 'f', 'v')}
        if any(el.tag in data_tags for el in root.iter()):
            raise ValueError('CRTX에 원본 차트 데이터가 남아 있습니다.')
        for el in root.iter(qn('a:t')):
            el.text = ''
        return clean_style_package(payload, etree.tostring(root, encoding='UTF-8')), root


def _remove_linked_chart(slide, shape):
    chart_ref = next(shape._element.iter(qn('c:chart')))
    rid = chart_ref.get(qn('r:id'))
    shape._element.getparent().remove(shape._element)
    return slide.part, rid


def extract_chart_styles(source, output_dir):
    source = Path(source).resolve(strict=True)
    if source.suffix.lower() != '.pptx':
        raise ValueError('기존 그래프틀은 .pptx 파일로 선택하세요.')
    prs = Presentation(source)
    charts, removed = [], set()
    for index, slide in enumerate(prs.slides, 1):
        for shape in list(_chart_shapes(slide.shapes)):
            result = _describe_chart(index, shape)
            charts.append(result)
            if result['linked']:
                result.update(status='skipped', reason='외부 연결 차트는 링크를 열지 않고 제외했습니다.')
                removed.add(_remove_linked_chart(slide, shape))
    if not charts:
        raise ValueError('편집 가능한 차트 객체를 찾지 못했습니다. 이미지로 만든 그래프는 서식으로 추출할 수 없습니다.')
    for part, rid in removed:
        part.drop_rel(rid)
    directory = Path(mkdtemp(prefix='chart_style_library_', dir=Path(output_dir).resolve()))
    pending = [row for row in charts if 'status' not in row]
    if pending:
        try:
            validate_local_deck(prs)
            with TemporaryDirectory() as work:
                scratch = Path(work) / 'charts.pptx'
                download = Path(work) / 'style.crtx'
                prs.save(scratch)
                with _powerpoint_deck(scratch) as presentation:
                    for index in dict.fromkeys(row['slide'] for row in pending):
                        by_id = {shape.Id: shape for shape in _com_charts(presentation.Slides(index).Shapes)}
                        for row in (item for item in pending if item['slide'] == index):
                            name = f"slide_{index:03d}_shape_{row['shape_id']:03d}.crtx"
                            download.unlink(missing_ok=True)
                            try:
                                by_id[row['shape_id']].Chart.SaveChartTemplate(str(download))
                                validated, _ = _clean_crtx(download.read_bytes())
                                (directory / name).write_bytes(validated)
                                row.update(status='exported', file=name)
                            except Exception as exc:
                                row.update(status='failed', reason=str(exc))
        except Exception as exc:
            for row in pending:
                if 'status' not in row:
                    row.update(status='failed', reason=f'PowerPoint 준비 실패: {exc}')
    counts = Counter(row['status'] for row in charts)
    report = directory / 'chart_style_inventory.json'
    report.write_text(json.dumps(dict(chart_count=len(charts),
                                      exported_count=counts['exported'], skipped_count=counts['skipped'],
                                      failed_count=counts['failed'], charts=charts),
                                 ensure_ascii=False, indent=2), encoding='utf-8')
    columns = ('slide', 'shape_id', 'shape_name', 'chart_type', 'series_count', 'prototype', 'fonts',
               'font_sizes_pt', 'size_pt', 'layout', 'status', 'file', 'reason')
    with (directory / 'chart_style_inventory.csv').open('w', encoding='utf-8-sig', newline='') as stream:
        table = csv.DictWriter(stream, columns)
        table.writeheader()
        for row in charts:
            fields = {key: row.get(key, '') for key in columns}
            for key, value in fields.items():
                if isinstance(value, str) and value.lstrip().startswith(('=', '+', '-', '@')):
                    fields[key] = "'" + value
            table.writerow(fields)
    return report


def _template_role(root):
    plot_area = root.find('.//' + qn('c:plotArea'))
    plots = [] if plot_area is None else [p for p in plot_area if _local_name(p).endswith('Chart')]
    if len(plots) == 1 and len(plots[0].findall(qn('c:ser'))) == 1:
        plot = plots[0]
        if plot.tag == qn('c:doughnutChart'):
            return 'RA_CHART_DOUGHNUT'
        direction, grouping = plot.find(qn('c:barDir')), plot.find(qn('c:grouping'))
        if (plot.tag == qn('c:barChart') and direction is not None and direction.get('val') == 'col'
                and grouping is not None and grouping.get('val') == 'clustered'):
            return 'RA_CHART_COLUMN'
    raise ValueError('기본 PPTX에는 단일 계열 세로 묶은 막대 또는 도넛 CRTX를 선택하세요.')


def _ensure_category_labels(chart):
    axis = chart.category_axis
    if axis.visible and axis.tick_label_position != XL_TICK_LABEL_POSITION.NONE:
        return False
    # Original slides can place labels outside the chart; reports export only the chart.
    axis.visible = True
    axis.tick_label_position = XL_TICK_LABEL_POSITION.LOW
    font = axis.tick_labels.font
    if font.name is None:
        font.name = chart.font.name or FONT
    if font.size is None:
        font.size = Pt(12)
    layout = chart._chartSpace.find('.//' + qn('c:plotArea')).find(qn('c:layout'))
    if layout is not None:
        manual = layout.find(qn('c:manualLayout'))
        if manual is not None:
            layout.remove(manual)
    return True


def make_report_chart_style(source, output, style_path=None):
    source, output = Path(source), Path(output)
    if source.suffix.lower() != '.crtx' or output.suffix.lower() != '.pptx':
        raise ValueError('입력은 CRTX, 결과는 PPTX로 지정하세요.')
    if source.resolve() == output.resolve():
        raise ValueError('원본 CRTX는 결과 파일로 덮어쓸 수 없습니다.')
    if style_path and Path(style_path).resolve() == output.resolve():
        raise ValueError('기존 기본 PPTX에 덮어쓰지 말고 사본을 저장하세요.')
    payload, root = _clean_crtx(source.read_bytes())
    role = _template_role(root)
    with TemporaryDirectory() as work:
        scratch = Path(work) / 'base.pptx'
        imported = Path(work) / 'style.crtx'
        result = Path(work) / 'result.pptx'
        imported.write_bytes(payload)
        if style_path:
            register_chart_style(style_path, scratch)
        else:
            create_chart_style(scratch)
        with _powerpoint_deck(scratch) as presentation:
            for slide in presentation.Slides:
                shape = next((s for s in slide.Shapes if s.HasChart and s.Name == role), None)
                if shape is not None:
                    chart = shape.Chart
                    before = tuple(chart.SeriesCollection(1).Values)
                    chart.ApplyChartTemplate(str(imported))
                    if chart.SeriesCollection().Count != 1 or tuple(chart.SeriesCollection(1).Values) != before:
                        raise ValueError('서식 적용 중 예제 데이터가 바뀌었습니다.')
                    chart.HasTitle = False
                    presentation.SaveAs(str(result), 24)  # ppSaveAsOpenXMLPresentation
                    break
        if role == 'RA_CHART_COLUMN':
            deck = Presentation(result)
            if _ensure_category_labels(_prototypes(deck)[role].chart):
                deck.save(result)
        result.write_bytes(clean_style_package(result.read_bytes()))
        return register_chart_style(result, output)


def run_cli():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('action', choices=('extract', 'make-default'))
    parser.add_argument('--input', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--style', type=Path)
    args = parser.parse_args()
    try:
        print(extract_chart_styles(args.input, args.output) if args.action == 'extract' else
              make_report_chart_style(args.input, args.output, args.style))
    except (BadZipFile, PackageNotFoundError, KeyError, OSError, ValueError) as exc:
        parser.exit(1, f'{exc}\n')


if __name__ == '__main__':
    run_cli()
