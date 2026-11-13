"""Conservative analysis recovery from the exported Excel display contract."""

import re


UNIT = r"백만원|천만원|십억원|억원|만원|천원|원|명|개|건|점|%"
AGGREGATE = re.compile(rf"^(?:[▷▶◀◁\s]*)(평균|합계)(?:[▷▶◀◁\s]*)\(({UNIT})\)$")
HEADER_ROLES = {"header", "banner_horizontal"}
LABEL_ROLES = {"stub", "banner_vertical", "base"}


def table_head_and_total(table):
    headers = []
    for row in table.get("matrix", []):
        anchors = [cell for cell in row if not cell.get("covered_by")]
        if any(cell.get("role") == "value" for cell in anchors):
            labels = [str(cell.get("display_text") or "") for cell in anchors if cell.get("role") in LABEL_ROLES]
            first = next((text for text in labels if text.strip()), "")
            total = re.sub(r"[\s■▣●]", "", first).lower() in {"전체", "전체응답자", "합계", "총계", "total"}
            return headers, anchors if total else []
        headers.append(anchors)
    return headers, []


def aggregate_measure(text):
    match = AGGREGATE.fullmatch(str(text or "").strip())
    return match.groups() if match else None


def sample_count(headers, total):
    if not total:
        return None
    for cell in total:
        if cell.get("role") != "value":
            continue
        text = str(cell.get("display_text") or "").strip()
        col = cell["col"]
        tokens = [re.sub(r"\s+", "", h["display_text"]).lower() for row in headers for h in row
                  if h["col"] == col and h.get("role") in HEADER_ROLES]
        if not any(re.fullmatch(r"(?:사례수|표본수|응답수|n)(?:\(명\))?", t) for t in tokens):
            continue
        if re.fullmatch(r"\([\d,]+\)|[\d,]+", text):
            count = int(text.strip("()").replace(",", ""))
            if count > 0:
                return count, cell.get("source_cell", "")
    return None


def recover_narrative(table, profile):
    headers, total = table_head_and_total(table)
    if not headers or not total:
        return None
    if re.search(r"응답자(?:특성|구성)표?", re.sub(r"\s+", "", table.get("title", ""))):
        evidence = sample_count(headers, total)
        if evidence:
            count, address = evidence
            return finish(f"전체 사례수는 {count:,}", profile), "profile_count", [address]
    points = []
    measures = set()
    scale = re.search(r"(\d+)\s*점\s*평균", table.get("title", ""))
    for cell in total:
        if cell.get("role") != "value":
            continue
        labels = [h["display_text"] for row in headers for h in row if h.get("role") in HEADER_ROLES
                  and h["col"] <= cell["col"] < h["col"] + h.get("colspan", 1) and h.get("display_text")]
        measure = aggregate_measure(labels[-1]) if labels else None
        if not measure and scale and labels and re.fullmatch(rf"\({scale[1]}\s*점\s*(?:척도|평균)\)", labels[-1]):
            measure = (f"{scale[1]}점 척도 평균", "점")
        value = str(cell.get("display_text") or "").strip()
        if not measure or len(labels) < 2 or not re.fullmatch(r"[+-]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?%?", value):
            continue
        if (value.endswith("%") and measure[1] != "%") or (scale and measure[0].endswith("척도 평균") and not 0 <= float(value.replace(",", "")) <= int(scale[1])):
            return None
        measures.add(measure)
        points.append((labels[-2], value, cell.get("source_cell", "")))
    if not points or len(measures) != 1 or len({label for label, _, _ in points}) != len(points):
        return None
    kind, unit = measures.pop()
    # Do not infer a ranking or sum columns that may describe different populations.
    shown = points[:4]
    values = ", ".join(f"‘{label}’ {value.removesuffix('%')}{unit}" for label, value, _ in shown)
    label = "전체 행에 표시된 값은" if measure_conflicts_with_title(table) else f"전체 행의 {kind}은"
    return finish(f"{label} {values}", profile), "aggregate_total", [address for _, _, address in shown]


def measure_conflicts_with_title(table):
    title = str(table.get("title") or "")
    declared = re.search(r"\((?:[^()]*/)?(평균|합계)\)\s*$", title)
    if not declared:
        return False
    headers, _ = table_head_and_total(table)
    actual = {measure[0] for row in headers for cell in row if cell.get("role") in HEADER_ROLES
              and (measure := aggregate_measure(cell.get("display_text")))}
    return len(actual) == 1 and declared[1] not in actual


def finish(text, profile):
    return text + ("입니다." if any(mode in profile for mode in ("서술", "설명")) else "임.")
