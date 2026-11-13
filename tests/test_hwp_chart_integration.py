from types import SimpleNamespace

import pytest

from report_automation_engine import hwp_com_writer as writer
from report_automation_engine.report_package import build_preflight, normalize_charts
from report_automation_engine.survey_chart_style import survey_chart_plan


def test_chart_plan_distinguishes_empty_slots_from_exclusions():
    titles = {'A': '서비스 이용', 'B': '응답자 구성표', 'C': '기준 없음', 'D': '한 값', 'E': '긴 보기'}
    rows = [(key, label, value, group) for key, label, value, group in (
        ('A', '있음', 60, '전체'), ('A', '없음', 40, '전체'), ('C', '여성', 20, '연령'),
        ('D', '평균', 5, '전체'), ('E', '긴 보기 ' * 24, 20, '전체'), ('E', '보기 B', 15, '전체'))]
    package = {'sections': [{'table_key': key, 'title': title} for key, title in titles.items()],
               'charts': [dict(table_key=k, category=c, value=v, series_group=g, measure='%', source_cell=f'A{i}')
                          for i, (k, c, v, g) in enumerate(rows, 1)]}
    plan = survey_chart_plan(package)
    assert [(x['key'], x['status']) for x in plan] == [('A', 'ready'), ('C', 'held'), ('D', 'excluded'), ('E', 'review')]
    assert '\n' in plan[-1]['labels'][0]


def test_figure_is_written_after_narrative_and_before_table(monkeypatch):
    events = []
    monkeypatch.setattr(writer, 'insert_text', lambda *_: None)
    monkeypatch.setattr(writer, 'run_action', lambda *_: True)
    monkeypatch.setattr(writer, 'insert_narrative_blocks', lambda *_: events.append('narrative'))
    monkeypatch.setattr(writer, 'current_body_width', lambda *_: 40000)
    monkeypatch.setattr(writer, 'write_table_part', lambda *_: events.append('table'))
    package = {'sections': [{'table_key': 'A', 'title': '문항'}], 'tables': [{'table_key': 'A'}]}
    writer.write_body(None, package, {'sections_written': 0}, write_figure=lambda *args: events.append(('figure', args[1])))
    assert events == ['narrative', ('figure', 'A'), 'table']


@pytest.mark.parametrize('prefix', ['', '3-2', ' 3', '1000'])
def test_invalid_figure_prefix(prefix):
    with pytest.raises(ValueError, match='앞자리'):
        writer.validate_figure_prefix(prefix)


def test_invalid_emf_cannot_replace_output(tmp_path, monkeypatch):
    monkeypatch.setattr(writer.tempfile, 'gettempdir', lambda: str(tmp_path))
    output = tmp_path / 'report.hwpx'
    output.write_bytes(b'previous output')

    def save(path, *_):
        from pathlib import Path

        Path(path).write_bytes(b'invalid EMF')
        return True

    hwp = SimpleNamespace(XHwpDocuments=SimpleNamespace(Active_XHwpDocument=SimpleNamespace(SaveAs=save)), SaveAs=save, Save=save)
    report = {'com': {'steps': []}}

    def invalid(_):
        raise ValueError('not EMF')

    with pytest.raises(writer.HwpWriterError, match='EMF 저장 검증 실패'):
        writer.save_as_hwpx(hwp, output, report, verify=invalid)
    assert output.read_bytes() == b'previous output'
    assert list(tmp_path.iterdir()) == [output]


def test_invalid_chart_value_does_not_block_report():
    qa = []
    charts = normalize_charts([{'table_key': 'T01', 'value': 'NaN', 'include_chart': 'Y'}], qa)
    assert charts[0]['value'] is None
    preflight = build_preflight({'qa': qa, 'charts': charts, 'tables': [], 'sections': []}, [])
    assert preflight['status'] == 'ready_with_warnings'
    assert not preflight['errors']
