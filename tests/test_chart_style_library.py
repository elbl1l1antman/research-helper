from contextlib import contextmanager
from io import BytesIO
import json
from pathlib import Path
from types import SimpleNamespace
from zipfile import ZipFile

from lxml import etree
from openpyxl import load_workbook
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_TICK_LABEL_POSITION
from pptx.opc.constants import RELATIONSHIP_TYPE
from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.ns import qn
from pptx.util import Inches
import pytest

from report_automation_engine import chart_style_library as library


def crtx(plot='', extra=''):
    data = BytesIO()
    with ZipFile(data, 'w') as book:
        book.writestr('chart/chart.xml', f'''<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Old project</a:t></a:r></a:p></c:rich></c:tx></c:title>
            <c:plotArea>{plot}</c:plotArea></c:chart>{extra}</c:chartSpace>''')
    return data.getvalue()


def test_crtx_keeps_style_but_not_data_and_supports_only_compatible_defaults():
    column = '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser/><c:gapWidth val="106"/></c:barChart>'
    content, root = library._clean_crtx(crtx(column))
    assert library._template_role(root) == 'RA_CHART_COLUMN'
    assert next(root.iter(qn('c:gapWidth'))).get('val') == '106'
    with ZipFile(BytesIO(content)) as book:
        assert b'Old project' not in book.read('chart/chart.xml')
    with pytest.raises(ValueError, match='원본 차트 데이터'):
        library._clean_crtx(crtx(column, '<c:strCache/>'))
    _, combo = library._clean_crtx(crtx(column + '<c:lineChart><c:ser/></c:lineChart>'))
    with pytest.raises(ValueError, match='단일 계열'):
        library._template_role(combo)


def test_extracts_grouped_chart_without_opening_external_link(tmp_path, monkeypatch):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    group = slide.shapes.add_group_shape()
    data = CategoryChartData()
    data.categories = ('A', 'B')
    data.add_series('전체', (40, 60))
    grouped = group.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED,
                                     Inches(1), Inches(1), Inches(4), Inches(2), data)
    grouped.chart.plots[0].gap_width = 129
    linked = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED,
                                   Inches(1), Inches(4), Inches(4), Inches(2), data)
    linked.chart.part.relate_to('file:///private-source.xlsx', RELATIONSHIP_TYPE.OLE_OBJECT, is_external=True)
    source = tmp_path / 'source.pptx'
    prs.save(source)
    before = source.read_bytes()

    def save(path):
        Path(path).write_bytes(crtx('<c:barChart><c:ser/></c:barChart>'))

    @contextmanager
    def fake_com(path):
        scratch = Presentation(path)
        library.validate_local_deck(scratch)
        assert [s.shape_id for s in library._chart_shapes(scratch.slides[0].shapes)] == [grouped.shape_id]
        chart = SimpleNamespace(Id=grouped.shape_id, HasChart=-1, Chart=SimpleNamespace(SaveChartTemplate=save))
        yield SimpleNamespace(Slides=lambda _: SimpleNamespace(Shapes=[
            SimpleNamespace(HasChart=0, Type=6, GroupItems=[chart])]))

    monkeypatch.setattr(library, '_powerpoint_deck', fake_com)
    output = library.extract_chart_styles(source, tmp_path)
    report = json.loads(output.read_text(encoding='utf-8'))
    assert 'source' not in report
    assert (report['chart_count'], report['exported_count'], report['skipped_count'], report['failed_count']) == (2, 1, 1, 0)
    row = report['charts'][0]
    assert row['layout']['gapWidth'] == ['129']
    assert row['prototype'] == 'RA_CHART_COLUMN'
    assert (output.parent / row['file']).is_file()
    assert (output.parent / 'chart_style_inventory.csv').is_file()
    assert source.read_bytes() == before


def test_default_style_restores_labels_without_changing_visible_axis(tmp_path):
    path = library.create_chart_style(tmp_path / 'base.pptx')
    chart = library._prototypes(Presentation(path))['RA_CHART_COLUMN'].chart
    axis = chart.category_axis
    layout = OxmlElement('c:layout')
    layout.append(OxmlElement('c:manualLayout'))
    plot_area = chart._chartSpace.find('.//' + qn('c:plotArea'))
    current = plot_area.find(qn('c:layout'))
    if current is not None:
        plot_area.remove(current)
    plot_area.insert(0, layout)
    assert not library._ensure_category_labels(chart)
    assert len(layout) == 1
    axis.visible = False
    axis.tick_label_position = XL_TICK_LABEL_POSITION.NONE
    assert library._ensure_category_labels(chart)
    assert axis.visible and axis.tick_label_position == XL_TICK_LABEL_POSITION.LOW
    assert len(layout) == 0
    assert chart.plots[0].gap_width == 125


def test_style_package_removes_metadata_also_from_embedded_workbook():
    book = BytesIO()
    with ZipFile(book, 'w') as archive:
        archive.writestr('docProps/core.xml', '<properties><creator>Private author</creator></properties>')
        archive.writestr('xl/worksheets/sheet1.xml', '<example>60</example>')
    deck = BytesIO()
    with ZipFile(deck, 'w') as archive:
        archive.writestr('docProps/custom.xml', '<properties><p name="Private project">Private value</p></properties>')
        archive.writestr('docProps/thumbnail.jpeg', b'private picture')
        archive.writestr('_rels/.rels', '<Relationships><Relationship Target="docProps/thumbnail.jpeg"/></Relationships>')
        archive.writestr('ppt/theme/theme1.xml', '<theme name="Private theme"><fontScheme name="Private"><font typeface="Example Font"/></fontScheme></theme>')
        archive.writestr('ppt/embeddings/data.xlsx', book.getvalue())
    with ZipFile(BytesIO(library.clean_style_package(deck.getvalue()))) as archive:
        assert 'docProps/thumbnail.jpeg' not in archive.namelist()
        assert len(etree.fromstring(archive.read('_rels/.rels'))) == 0
        assert len(etree.fromstring(archive.read('docProps/custom.xml'))) == 0
        theme = etree.fromstring(archive.read('ppt/theme/theme1.xml'))
        assert theme.get('name') == theme[0].get('name') == 'Chart Theme'
        assert theme[0][0].get('typeface') == 'Example Font'
        with ZipFile(BytesIO(archive.read('ppt/embeddings/data.xlsx'))) as embedded:
            assert len(etree.fromstring(embedded.read('docProps/core.xml'))) == 0
            assert embedded.read('xl/worksheets/sheet1.xml') == b'<example>60</example>'


def test_crtx_replaces_only_root_gallery_preview():
    base = BytesIO(crtx('<c:barChart><c:ser/></c:barChart>'))
    with ZipFile(base, 'a') as archive:
        archive.writestr('_rels/.rels', '<Relationships><Relationship Type="test/image" Target="chart/media/icon.bmp"/></Relationships>')
        archive.writestr('chart/media/icon.bmp', b'old chart image')
        archive.writestr('chart/media/fill.bmp', b'series texture')
    result, _ = library._clean_crtx(base.getvalue())
    with ZipFile(BytesIO(result)) as archive:
        assert archive.read('chart/media/icon.bmp').startswith(b'BM')
        assert archive.read('chart/media/fill.bmp') == b'series texture'


@pytest.mark.parametrize('theme,gap,overlap', [('deep_blue',106,-5), ('sky_blue',100,-10)])
def test_bundled_chart_styles_use_neutral_metadata_and_demo_values(theme, gap, overlap):
    path = Path(__file__).resolve().parents[1] / 'templates' / f'chart_style_{theme}.pptx'
    prototypes = library._prototypes(Presentation(path))
    chart = prototypes['RA_CHART_COLUMN'].chart
    assert (chart.plots[0].gap_width, chart.plots[0].overlap) == (gap, overlap)
    assert chart.series[0].values == (25, 63.3, 11.7)
    for role, shape in prototypes.items():
        with ZipFile(BytesIO(shape.chart.part.chart_workbook.xlsx_part.blob)) as archive:
            assert len(etree.fromstring(archive.read('docProps/core.xml'))) == 0
        workbook = load_workbook(BytesIO(shape.chart.part.chart_workbook.xlsx_part.blob), read_only=True, data_only=True)
        try:
            assert tuple(r[1] for r in list(workbook.active.values)[1:]) == ((25, 63.3, 11.7) if role == 'RA_CHART_COLUMN' else (36.7, 63.3))
        finally:
            workbook.close()
    with ZipFile(path) as archive:
        assert len(etree.fromstring(archive.read('docProps/core.xml'))) == 0
        assert not any(n.startswith(('docProps/thumbnail.', 'ppt/notesSlides/', 'ppt/comments/', 'customXml/')) for n in archive.namelist())
