from io import BytesIO
import json
from zipfile import ZipFile

import openpyxl
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
import pytest

from report_automation_engine.document_writer import write_document
from report_automation_engine.survey_chart_style import create_chart_style, register_chart_style, write_chart_review


def shapes(path):
    return [shape for slide in Presentation(path).slides for shape in slide.shapes if shape.has_chart]


def point(key, label, value, cell, measure='%', group='전체', **extra):
    return dict(table_key=key, category=label, value=value, source_cell=cell, measure=measure,
                series_group=group, include_chart=False, **extra)


def test_preserves_edited_style_and_replaces_sample_data(tmp_path):
    editable = create_chart_style(tmp_path / 'edit.pptx')
    prs = Presentation(editable)
    column = prs.slides[0].shapes[1].chart
    column.plots[0].gap_width = 215
    column.category_axis.tick_labels.font.name = 'Test Font'
    column.plots[0].data_labels.number_format = '0.00%'
    column.value_axis.maximum_scale = 100
    color = RGBColor(11, 22, 33)
    column.series[0].points[1].format.fill.fore_color.rgb = color
    prs.save(editable)
    snapshot = editable.read_bytes()
    style = register_chart_style(editable, tmp_path / 'profile' / 'style.pptx')
    editable.unlink()
    assert style.read_bytes() == snapshot
    package = {'sections': [{'table_key': key, 'title': key} for key in ('A', 'B', 'C')], 'charts': [
        point('A', '세 번째', 11.7, 'E4'), point('A', '두 번째', 63.3, 'D4'), point('A', '첫 번째\n긴 보기', 25, 'C4'),
        point('A', '하위 배너', 99, 'C5', group='남성'),
        point('B', '있음', 63.3, 'C10'), point('B', '없음', 36.7, 'D10'),
        point('C', '현재', 123.3, 'C20', '백만원'), point('C', '예상', 136.7, 'D20', '백만원'),
    ]}
    source = tmp_path / 'report_package.json'
    source.write_text(json.dumps(package), encoding='utf-8')
    out = write_document(source, tmp_path / 'review.pptx', 'chart_review', chart_style_path=style)
    a, b, c = [shape.chart for shape in shapes(out)]
    assert a.plots[0].gap_width == c.plots[0].gap_width == 215
    assert a.category_axis.tick_labels.font.name == 'Test Font'
    assert [item.label for item in a.plots[0].categories] == ['첫 번째\n긴 보기', '두 번째', '세 번째']
    assert a.series[0].values == (25.0, 63.3, 11.7)
    assert a.series[0].points[1].format.fill.fore_color.rgb == color
    assert a.plots[0].data_labels.number_format == '0.00"%"'
    assert b.chart_type == XL_CHART_TYPE.DOUGHNUT and c.chart_type == XL_CHART_TYPE.COLUMN_CLUSTERED
    assert c.series[0].values == (123.3, 136.7)
    assert a.value_axis.maximum_scale == 100 and c.value_axis.maximum_scale is None
    assert c.plots[0].data_labels.number_format == '0.00'
    assert all(not chart.has_title for chart in (a, b, c))
    with ZipFile(out) as archive:
        books = [name for name in archive.namelist() if name.startswith('ppt/embeddings/')]
        assert len(books) == 3
        for name in books:
            wb = openpyxl.load_workbook(BytesIO(archive.read(name)), read_only=True)
            assert wb.active['B1'].value == '전체'
            wb.close()
    assert style.read_bytes() == snapshot
    with pytest.raises(ValueError, match='덮어쓸'):
        write_chart_review(package, style, style)
    assert style.read_bytes() == snapshot


def test_missing_or_invalid_style_does_not_replace_registered_copy(tmp_path):
    path = create_chart_style(tmp_path / 'valid.pptx')
    target = register_chart_style(path, tmp_path / 'registered.pptx')
    previous = target.read_bytes()
    bad = tmp_path / 'bad.pptx'
    Presentation().save(bad)
    with pytest.raises(ValueError, match='RA_CHART_COLUMN'):
        register_chart_style(bad, target)
    assert target.read_bytes() == previous
    prs = Presentation(path)
    prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].hyperlink.address = 'https://example.invalid'
    prs.save(bad)
    with pytest.raises(ValueError, match='외부 링크'):
        register_chart_style(bad, target)
    assert target.read_bytes() == previous


def test_review_reports_missing_charts_instead_of_guessing(tmp_path):
    package = {'sections': [{'table_key': key, 'title': title} for key, title in (
        ('A', '일반 문항'), ('B', '표본 부족'), ('C', '응답자 구성표'), ('D', '단일 값'), ('E', '누락값'))], 'charts': [
        point('A', '보기1', 50, 'B3'), point('A', '보기2', 50, 'C3'),
        point('B', '보기1', 30, 'C5', group='여성'),
        point('D', '평균', 50, 'B8', '점'),
        point('E', '보기1', 0, 'C10'), point('E', '보기2', None, 'D10'),
    ]}
    out = write_chart_review(package, tmp_path / 'out.pptx')
    assert [shape.chart.chart_type for shape in shapes(out)] == [XL_CHART_TYPE.COLUMN_CLUSTERED]
    text = '\n'.join(shape.text for slide in Presentation(out).slides for shape in slide.shapes if shape.has_text_frame)
    assert all(reason in text for reason in ('전체 값 확인 필요', '비교 항목 없음', '누락'))
    assert '응답자 구성표' not in text
