from report_automation_engine.report_table_evidence import measure_conflicts_with_title, recover_narrative, sample_count, table_head_and_total
from report_automation_engine.report_table_matrix import make_cell, resolve_table_metadata


def table(title, measure="평균(백만원)", total="■전 체■", value="3,214.0"):
    rows = [
        [("base", "※BASE : 방문 경험자"), ("banner_horizontal", "사례수"), ("banner_horizontal", "기업 매출")],
        [("blank", ""), ("banner_horizontal", "빈도"), ("banner_horizontal", measure)],
        [("banner_vertical", total), ("value", "(1,234)"), ("value", value)],
    ]
    return {"title": title, "matrix": [[make_cell(r, c, role, text, text, "", f"R{r}C{c}", "")
                                       for c, (role, text) in enumerate(row, 1)] for r, row in enumerate(rows, 1)]}


def test_total_and_unit_are_grounded_in_their_own_cells():
    source = table("기업 조사")
    assert resolve_table_metadata(source) == {"base_label": "방문 경험자", "unit": "백만원"}
    assert sample_count(*table_head_and_total(source)) == (1234, "R3C2")
    assert recover_narrative(source, "")[2] == ["R3C3"]
    assert "3,214.0백만원" in recover_narrative(source, "")[0]
    assert recover_narrative(table("기업 조사", total="남성"), "") is None


def test_scale_requires_matching_title_and_header():
    assert "3.92점" in recover_narrative(table("지역서점 (5점평균)", "(5점척도)", value="3.92"), "")[0]
    assert recover_narrative(table("지역서점 (100점평균)", "(5점척도)", value="3.92"), "") is None
    assert recover_narrative(table("지역서점 (5점평균)", "(5점척도)", value="39.2"), "") is None


def test_sample_count_is_not_a_base_population():
    source = table("[ 응 답 자 특 성 표 ]")
    source["matrix"][0][0]["display_text"] = ""
    assert resolve_table_metadata(source)["base_label"] == "확인 필요"
    assert recover_narrative(source, "") == ("전체 사례수는 1,234임.", "profile_count", ["R3C2"])


def test_conflicting_title_does_not_convert_average_to_sum():
    source = table("기업 매출(백만원/합계)")
    assert measure_conflicts_with_title(source)
    text = recover_narrative(source, "")[0]
    assert "3,214.0백만원" in text and "평균" not in text and "합계" not in text
