"""Validate mm settings and apply them to the BODY section through PageSetup."""

import math


PAGE_FIELDS = {
    "paper_width_mm": "PaperWidth", "paper_height_mm": "PaperHeight",
    "top_mm": "TopMargin", "bottom_mm": "BottomMargin",
    "left_mm": "LeftMargin", "right_mm": "RightMargin",
    "gutter_mm": "GutterLen", "header_mm": "HeaderLen", "footer_mm": "FooterLen",
}


def normalize_page_setup(config):
    if not isinstance(config, dict) or set(config) != set(PAGE_FIELDS) | {"orientation"}:
        raise ValueError("편집용지 설정에 필수 항목이 없거나 알 수 없는 항목이 있습니다.")
    result = dict(config)
    for name in PAGE_FIELDS:
        value = result[name]
        if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value) or value < 0:
            raise ValueError(f"편집용지 {name}: 유한한 0 이상의 mm 값이 필요합니다.")
    if result["orientation"] not in ("portrait", "landscape"):
        raise ValueError("편집용지 방향은 portrait 또는 landscape여야 합니다.")
    width, height = result["paper_width_mm"], result["paper_height_mm"]
    if width <= 0 or height <= 0:
        raise ValueError("용지 너비와 높이는 0보다 커야 합니다.")
    if result["orientation"] == "landscape":
        width, height = height, width
    if width <= result["left_mm"] + result["right_mm"] + result["gutter_mm"]:
        raise ValueError("좌우 여백과 제본 여백을 제외한 본문 너비가 없습니다.")
    if height <= result["top_mm"] + result["bottom_mm"]:
        raise ValueError("위아래 여백을 제외한 본문 높이가 없습니다.")
    # HeaderLen/FooterLen are inside the page margins, not additional body deductions.
    if result["header_mm"] > result["top_mm"] or result["footer_mm"] > result["bottom_mm"]:
        raise ValueError("머리말/꼬리말 간격은 위쪽/아래쪽 여백 이하여야 합니다.")
    if any(result[key] > 2147483647 * 25.4 / 7200 for key in PAGE_FIELDS):
        raise ValueError("편집용지 값이 아래한글 정수 단위 범위를 초과합니다.")
    units = {key: round(result[key] * 7200 / 25.4) for key in PAGE_FIELDS}
    w, h = (units["paper_height_mm"], units["paper_width_mm"]) if result["orientation"] == "landscape" else (units["paper_width_mm"], units["paper_height_mm"])
    if w <= units["left_mm"] + units["right_mm"] + units["gutter_mm"] or h <= units["top_mm"] + units["bottom_mm"]:
        raise ValueError("아래한글 단위로 변환한 본문 공간이 없습니다.")
    return result


def apply_page_setup(hwp, config):
    config = normalize_page_setup(config)
    params = hwp.HParameterSet.HSecDef
    if hwp.HAction.GetDefault("PageSetup", params.HSet) is False:
        raise ValueError("현재 편집용지 설정을 읽지 못했습니다.")
    page = params.PageDef
    expected = {field: round(config[name] * 7200 / 25.4) for name, field in PAGE_FIELDS.items()}
    expected.update(Landscape=int(config["orientation"] == "landscape"), GutterType=0)
    for field, value in expected.items():
        setattr(page, field, value)
    # ApplyTo belongs to the action set. 2 preserves other sections (e.g. the cover).
    params.HSet.SetItem("ApplyTo", 2)
    if hwp.HAction.Execute("PageSetup", params.HSet) is False:
        raise ValueError("아래한글이 편집용지 설정을 적용하지 못했습니다.")
    if hwp.HAction.GetDefault("PageSetup", params.HSet) is False:
        raise ValueError("적용된 편집용지 설정을 확인하지 못했습니다.")
    actual = {field: int(getattr(params.PageDef, field)) for field in expected}
    if any(abs(actual[field] - value) > 1 for field, value in expected.items()):
        raise ValueError("요청한 편집용지 설정과 실제 적용값이 다릅니다.")
    width, height = actual["PaperWidth"], actual["PaperHeight"]
    if actual["Landscape"]:
        width, height = height, width
    return {"scope": "body_section", "applied": True, "actual_hwpunit": actual,
            "body_width_hwpunit": width - actual["LeftMargin"] - actual["RightMargin"] - actual["GutterLen"],
            "body_height_hwpunit": height - actual["TopMargin"] - actual["BottomMargin"]}
