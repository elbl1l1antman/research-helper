"""Common PowerPoint chart styles and editable survey review exports."""

import argparse
from collections import defaultdict
from copy import deepcopy
from io import BytesIO
import json
import math
from pathlib import Path
import re
from tempfile import NamedTemporaryFile
from zipfile import BadZipFile

from openpyxl.utils.cell import coordinate_to_tuple
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_DATA_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.exc import PackageNotFoundError
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.parts.chart import ChartPart
from pptx.util import Inches, Pt


STYLE_NAMES = {'RA_CHART_COLUMN': XL_CHART_TYPE.COLUMN_CLUSTERED, 'RA_CHART_DOUGHNUT': XL_CHART_TYPE.DOUGHNUT}
PROFILE_TITLES = {'응답자특성', '응답자특성표', '응답자구성', '응답자구성표', '표본특성표', '표본구성표'}
FONT = '맑은 고딕'


def _text(slide, text, x, y, width, height, size):
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(width), Inches(height))
    shape.text_frame.word_wrap = True
    run = shape.text_frame.paragraphs[0].add_run()
    run.text = text
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = size >= 20
    run.font.color.rgb = RGBColor.from_string('1F303A')
    return shape


def _data(categories, values, fmt='0.0'):
    data = CategoryChartData(number_format=fmt)
    data.categories = categories
    data.add_series('전체', values)
    return data


def _starter():
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    for name, kind in STYLE_NAMES.items():
        column = kind == XL_CHART_TYPE.COLUMN_CLUSTERED
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        _text(slide, '세로 막대 · 차트 서식' if column else '도넛 · 차트 서식', 0.7, 0.36, 11.9, 0.62, 24)
        samples = ('보기 A', '보기 B', '길이가 긴 보기\n줄바꿈 예시') if column else ('있음', '없음')
        values = (25, 63.3, 11.7) if column else (36.7, 63.3)
        shape = slide.shapes.add_chart(kind, Inches(0.8), Inches(1.25), Inches(11.7), Inches(5.45), _data(samples, values, '0.0"%"'))
        shape.name = name
        chart = shape.chart
        chart.has_title = False
        chart.font.name, chart.font.size = FONT, Pt(13)
        chart.has_legend = not column
        plot = chart.plots[0]
        plot.has_data_labels = True
        labels = plot.data_labels
        labels.show_value = True
        labels.number_format = '0.0"%"'
        labels.font.name, labels.font.size = FONT, Pt(13)
        if column:
            plot.gap_width = 125
            labels.position = XL_DATA_LABEL_POSITION.OUTSIDE_END
            chart.category_axis.tick_labels.font.name = FONT
            chart.value_axis.tick_labels.font.name = FONT
        else:
            chart.legend.position = XL_LEGEND_POSITION.RIGHT
            chart.legend.font.name = FONT
        for index, point in enumerate(chart.series[0].points):
            point.format.fill.solid()
            point.format.fill.fore_color.rgb = RGBColor.from_string('C65B39' if index == 1 else '357B99')
            if not column:
                point.format.line.color.rgb = RGBColor(255, 255, 255)
        _text(slide, '보기 A/첫 조각: 기본색     보기 B/두 번째 조각: 최대값 강조색', 0.72, 6.87, 11.8, 0.35, 12)
    return prs


def _save(path, write):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with NamedTemporaryFile(dir=path.parent, suffix='.tmp', delete=False) as stream:
        pending = Path(stream.name)
    try:
        write(pending)
        pending.replace(path)
    finally:
        pending.unlink(missing_ok=True)
    return path


def create_chart_style(path):
    return _save(path, _starter().save)


def _prototypes(prs):
    if any('vbaProject' in str(part.partname) or any(rel.is_external for rel in part.rels.values())
           for part in prs.part.package.iter_parts()):
        raise ValueError('외부 링크나 매크로가 없는 편집용 기본 PPTX를 사용하세요.')
    found = {}
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.name not in STYLE_NAMES:
                continue
            if shape.name in found or not shape.has_chart:
                raise ValueError(f'{shape.name}: 같은 이름의 차트가 둘 이상이거나 차트 객체가 아닙니다.')
            chart = shape.chart
            if (chart.chart_type != STYLE_NAMES[shape.name] or len(chart.plots) != 1 or len(chart.series) != 1
                    or any(rel.is_external for rel in chart.part.rels.values()) or chart.part.chart_workbook.xlsx_part is None):
                raise ValueError(f'{shape.name}: 기본 차트 유형과 데이터 계열 1개, 내장 데이터가 필요합니다.')
            if (shape.left < 0 or shape.top < Inches(1.08) or min(shape.width, shape.height) <= 0
                    or shape.left + shape.width > prs.slide_width or shape.top + shape.height > prs.slide_height - Inches(0.48)):
                raise ValueError(f'{shape.name}: 슬라이드의 제목·출처 영역을 제외한 가운데에 차트를 배치하세요.')
            found[shape.name] = shape
    missing = STYLE_NAMES.keys() - found.keys()
    if missing:
        raise ValueError('기본 서식의 차트 이름을 유지하세요: ' + ', '.join(sorted(missing)))
    return found


def register_chart_style(source, target):
    source = Path(source)
    if source.suffix.lower() != '.pptx':
        raise ValueError('기본 차트 서식은 .pptx 파일이어야 합니다.')
    payload = source.read_bytes()
    try:
        _prototypes(Presentation(BytesIO(payload)))
    except (BadZipFile, PackageNotFoundError, KeyError) as error:
        raise ValueError('읽을 수 없는 차트 서식 PPTX입니다.') from error
    return _save(target, lambda path: path.write_bytes(payload))


def _clone_chart(slide, prototype, labels, values, measure):
    source = prototype.chart.part
    package = slide.part.package
    part = ChartPart.load(package.next_partname(ChartPart.partname_template), source.content_type, package, source.blob)
    root = part.chart._chartSpace
    external = root.find(qn('c:externalData'))
    if external is not None:
        root.remove(external)
    # The new chart lives in the same PPTX and shares read-only style/color parts, never the sample workbook.
    for rid, relation in source.rels.items():
        if relation.target_part is not source.chart_workbook.xlsx_part:
            part.rels._rels[rid] = relation
    rid = slide.part.relate_to(part, RT.CHART)
    frame = slide.shapes._add_chart_graphicFrame(rid, prototype.left, prototype.top, prototype.width, prototype.height)
    frame.nvGraphicFramePr.cNvPr.name = prototype.name
    chart = part.chart
    fmt = '#,##0.0"%"' if measure == '%' else '#,##0.0'
    chart.replace_data(_data(labels, values, fmt))
    chart.has_title = False
    series = chart.series[0]
    points = {int(point.find(qn('c:idx')).get('val')): point for point in series._element.findall(qn('c:dPt'))}
    maximum = max(values)
    for point in points.values():
        series._element.remove(point)
    for index, value in enumerate(values):
        style = points.get(int(value == maximum))
        if style is not None:
            point = deepcopy(style)
            point.find(qn('c:idx')).set('val', str(index))
            series._element.insert_element_before(point, 'c:dLbls', 'c:trendline', 'c:errBars', 'c:cat', 'c:val', 'c:extLst')
        else:
            fill = series.points[index].format.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor.from_string('C65B39' if value == maximum else '357B99')
    for label in root.iter(qn('c:dLbl')):
        # A sample label's text is data, but its position and formatting are style.
        text = label.find(qn('c:tx'))
        if text is not None:
            label.remove(text)
    for group in root.iter(qn('c:dLbls')):
        for extension in group.findall(qn('c:extLst')):
            group.remove(extension)  # Linked custom-label ranges would point to the replaced workbook.
    if chart.plots[0].has_data_labels:
        label_format = chart.plots[0].data_labels
        if not label_format.number_format_is_linked:
            label_format.number_format = _number_format(label_format.number_format, measure == '%')
    if chart.chart_type == XL_CHART_TYPE.COLUMN_CLUSTERED:
        axis = chart.value_axis
        if axis.minimum_scale is not None and axis.minimum_scale > min(0, *values):
            axis.minimum_scale = None
        if axis.maximum_scale is not None and axis.maximum_scale < max(0, *values):
            axis.maximum_scale = None
        if not axis.tick_labels.number_format_is_linked:
            axis.tick_labels.number_format = _number_format(axis.tick_labels.number_format, measure == '%')


def _number_format(fmt, percent):
    # Package percentages are already in 0–100 units, so '%' must be a literal.
    return re.sub(r'"[^"]*"|\\.|%', lambda token: ('"%"' if percent else '')
                  if token[0] in ('%', '"%"', r'\%') else token[0], fmt)


def _key(text):
    return re.sub(r'[\s\[\]■▣●]', '', str(text or '')).lower()


def _coord(cell):
    try:
        return coordinate_to_tuple(str(cell or '').replace('$', '').upper())
    except (ValueError, IndexError, KeyError):
        return math.inf, math.inf


def _representative(rows, table):
    all_rows = [r for r in rows if _key(r.get('series_group')) in {'전체', '전체응답자', '합계', '총계', '계', 'total'}]
    if not all_rows:
        return None, '전체 값 확인 필요'
    measures = defaultdict(list)
    for row in all_rows:
        label = _key(row.get('category'))
        if not label or label in {'전체', '합계', '총계', '계', '사례수', '표본수', 'n', 'base'}:
            continue
        measures[str(row.get('measure') or '').strip()].append(row)
    if '%' in measures:
        measure = '%'
    elif len(measures) == 1:
        measure = next(iter(measures))
    else:
        return None, '대표 지표 확인 필요' if measures else '비교 항목 없음'
    points = sorted(measures[measure], key=lambda r: _coord(r.get('source_cell')))
    if len(points) < 2:
        return None, '비교 항목 없음'
    if len({_key(p.get('category')) for p in points}) != len(points):
        return None, '보기 이름 또는 전체 기준 중복 확인 필요'
    values = [p.get('value') for p in points]
    if any(isinstance(v, bool) or not isinstance(v, (int, float)) or not math.isfinite(v) for v in values):
        return None, '누락 또는 숫자가 아닌 보기 값 확인 필요'
    if measure == '%' and not all(0 <= v <= 100 for v in values):
        return None, '비율 범위 확인 필요'
    if table and table.get('matrix'):
        addresses = {_coord(p.get('source_cell')) for p in points}
        row_ids, col_ids = map(set, zip(*addresses))
        for matrix_row in table['matrix']:
            for cell in matrix_row:
                row, col = _coord(cell.get('source_cell'))
                if cell.get('role') == 'value' and (row, col) not in addresses and str(cell.get('display_text') or '').strip() in {'', '-', '–', '—'} and (
                    (len(row_ids) == 1 and row in row_ids and min(col_ids) <= col <= max(col_ids)) or
                    (len(col_ids) == 1 and col in col_ids and min(row_ids) <= row <= max(row_ids))):
                    return None, '집계표의 누락 값 확인 필요'
    labels = [_key(p['category']) for p in points]
    pairs = ({'있음', '없음'}, {'예', '아니오'}, {'예', '아니요'}, {'경험있음', '경험없음'})
    binary = (len(points) == 2 and measure == '%' and abs(sum(values) - 100) <= 0.2000001
              and not any(p.get('multiple_response') for p in points)
              and (set(labels) in pairs or all(p.get('confirmed_binary_share') is True for p in points)))
    return (points, measure, 'RA_CHART_DOUGHNUT' if binary else 'RA_CHART_COLUMN'), ''


def write_chart_review(package, output, style_path=None):
    if style_path is not None and Path(output).resolve() == Path(style_path).resolve():
        raise ValueError('등록한 차트 서식을 결과 파일로 덮어쓸 수 없습니다.')
    prs = Presentation(style_path) if style_path is not None else _starter()
    prototypes = _prototypes(prs)
    sample_slides = list(prs.slides._sldIdLst)
    candidates = defaultdict(list)
    for row in package.get('charts', []):
        candidates[row.get('table_key')].append(row)
    sections = {row['table_key']: row for row in package.get('sections', [])}
    tables = {row['table_key']: row for row in package.get('tables', [])}
    skips = []
    for key in dict.fromkeys((*sections, *tables, *candidates)):
        section, table = sections.get(key, {}), tables.get(key, {})
        title = section.get('title') or table.get('title') or key or '차트'
        if any(_key(t) in PROFILE_TITLES for t in (title, table.get('title'), table.get('table_type'))):
            continue
        entry, reason = _representative(candidates[key], table)
        if not entry:
            skips.append(f'{key} · {title}: {reason}')
            continue
        rows, measure, kind = entry
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        _text(slide, title, 0.65, 0.3, prs.slide_width.inches - 1.3, 0.78, 22)
        _clone_chart(slide, prototypes[kind], [str(p['category']) for p in rows], [p['value'] for p in rows], measure)
        source = '!'.join(x for x in (table.get('source_sheet'), table.get('source_range')) if x)
        footer = ' · '.join(str(s) for s in (key, '전체', f'단위: {measure}' if measure else '', source) if s)
        _text(slide, footer, 0.66, prs.slide_height.inches - 0.48, prs.slide_width.inches - 1.32, 0.33, 11)
    if not skips and len(prs.slides) == len(sample_slides):
        raise ValueError('생성할 수 있는 차트가 없습니다.')
    for index in range(0, len(skips), 12):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        _text(slide, '차트 검토 필요', 0.65, 0.35, prs.slide_width.inches - 1.3, 0.7, 22)
        _text(slide, '\n'.join(skips[index:index + 12]), 0.65, 1.3, prs.slide_width.inches - 1.3, prs.slide_height.inches - 1.7, 15)
    for slide_id in sample_slides:
        prs.part.drop_rel(slide_id.rId)
        prs.slides._sldIdLst.remove(slide_id)
    return _save(output, prs.save)


def run_cli():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('action', choices=('create', 'register', 'review'))
    parser.add_argument('--input', type=Path)
    parser.add_argument('--style', type=Path)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    if args.action != 'create' and args.input is None:
        parser.error('--input is required')
    try:
        if args.action == 'create':
            result = create_chart_style(args.output)
        elif args.action == 'register':
            result = register_chart_style(args.input, args.output)
        else:
            if args.input.resolve() == args.output.resolve():
                raise ValueError('입력 패키지를 결과 파일로 덮어쓸 수 없습니다.')
            result = write_chart_review(json.loads(args.input.read_text(encoding='utf-8-sig')), args.output, args.style)
        print(result)
    except (ValueError, OSError, KeyError, BadZipFile, PackageNotFoundError) as error:
        parser.exit(1, f'{error}\n')


if __name__ == '__main__':
    run_cli()
