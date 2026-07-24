"""Build HWP/HWPX template blueprints from probe reports.

The blueprint is a middle layer between raw HWPX probing and the future writer.
It ranks candidate "question result" blocks so the launcher can ask the user
which template block should be repeated for survey result pages.
"""

from __future__ import annotations

import argparse
import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Tuple


RESULT_CLASSES = {"survey_result_table", "likely_report_table"}
NEGATIVE_TERMS = {
    "목 차",
    "목차",
    "표목차",
    "그림목차",
    "Contents",
    "발 행 일",
    "발 행 처",
    "연구기관",
    "제 출 문",
    "Client Logo",
    "Copyright",
}
PLACEHOLDER_TERMS = {"표제목 입력", "그림제목 입력", "내용입력", "내용을 입력합니다", "입력"}


def build_blueprint(probe_report_path: str | Path, output_path: str | Path | None = None) -> Dict[str, Any]:
    probe_path = Path(probe_report_path).resolve()
    report = load_json(probe_path)
    blueprint: Dict[str, Any] = {
        "schema_version": "1.0",
        "created_at": now(),
        "source_probe_report": str(probe_path),
        "template_type": "survey_report_hwp",
        "status": "ready",
        "templates": [],
        "warnings": [],
        "errors": [],
    }

    for index, template in enumerate(report.get("templates", []), start=1):
        blueprint["templates"].append(build_template_blueprint(template, index))

    if not any(t.get("blocks") for t in blueprint["templates"]):
        blueprint["status"] = "needs_review"
        blueprint["warnings"].append("반복 결과 블록 후보를 찾지 못했습니다.")

    if output_path:
        write_json(output_path, blueprint)
    return blueprint


def build_template_blueprint(template: Dict[str, Any], template_index: int) -> Dict[str, Any]:
    blocks = []
    for fallback_index, table in enumerate(template.get("tables", []), start=1):
        candidate = make_block_candidate(template, table, fallback_index)
        if candidate["score"] >= 35:
            blocks.append(candidate)

    blocks.sort(key=lambda block: block["score"], reverse=True)
    return {
        "template_index": template_index,
        "source_path": template.get("source_path", ""),
        "analyzed_path": template.get("analyzed_path", ""),
        "conversion": template.get("conversion", ""),
        "table_count": template.get("table_count", 0),
        "recommended_block_id": blocks[0]["block_id"] if blocks else "",
        "blocks": blocks[:12],
        "warnings": template.get("warnings", []),
        "errors": template.get("errors", []),
    }


def make_block_candidate(template: Dict[str, Any], table: Dict[str, Any], fallback_index: int) -> Dict[str, Any]:
    table_index = int(table.get("table_index") or fallback_index)
    before_text = [str(value) for value in table.get("before_text", []) if str(value).strip()]
    sample_rows = table.get("sample_rows", [])
    flat_before = " ".join(before_text)
    flat_rows = flatten_rows(sample_rows)
    flat_all = f"{flat_before} {flat_rows}"

    table_kind = infer_table_kind(sample_rows, flat_all)
    sequence = infer_sequence(flat_before, flat_rows)
    score, reasons = score_candidate(table, table_kind, sequence, flat_all)
    title = infer_title(before_text, sample_rows)

    return {
        "block_id": f"T{template.get('source_type', '').replace('.', '').upper() or 'HWP'}_{table_index:03d}",
        "type": "question_result",
        "score": score,
        "score_reasons": reasons,
        "title_candidate": title,
        "table_index": table_index,
        "section": table.get("section", ""),
        "table_kind": table_kind,
        "sequence": sequence,
        "style_source": {
            "table_index": table_index,
            "section": table.get("section", ""),
            "row_count": table.get("row_count", 0),
            "col_count": table.get("col_count", 0),
            "classification": table.get("classification", ""),
        },
        "detected_parts": {
            "has_narrative": has_narrative(flat_before),
            "has_figure_caption": has_figure_caption(flat_before),
            "has_table_caption": has_table_caption(flat_before),
            "has_base_or_unit": has_base_or_unit(flat_before, flat_rows),
            "looks_like_placeholder": looks_like_placeholder(flat_all),
            "looks_like_layout": looks_like_layout(flat_all),
        },
        "preview": {
            "before_text": before_text[-3:],
            "sample_rows": sample_rows[:5],
        },
        "render_contract": {
            "requires": ["section_title", "narrative", "result_table"],
            "optional": ["figure_caption", "chart_placeholder", "table_caption", "base_note", "source_note"],
            "recommended_package_version": "report_package_v2",
        },
    }


def score_candidate(table: Dict[str, Any], table_kind: str, sequence: List[str], flat_all: str) -> Tuple[int, List[str]]:
    score = 0
    reasons: List[str] = []
    classification = str(table.get("classification", ""))
    row_count = int(table.get("row_count") or 0)
    col_count = int(table.get("col_count") or 0)

    if classification in RESULT_CLASSES:
        score += 30
        reasons.append("result_table_class")
    if table_kind in {"modern_n_percent", "legacy_frequency_percent"}:
        score += 25
        reasons.append(table_kind)
    if row_count >= 8 and col_count >= 4:
        score += 15
        reasons.append("large_enough_result_table")
    if has_base_or_unit("", flat_all):
        score += 10
        reasons.append("base_or_unit_detected")
    if has_table_caption(flat_all):
        score += 8
        reasons.append("table_caption_detected")
    if has_figure_caption(flat_all):
        score += 5
        reasons.append("figure_caption_detected")
    if has_narrative(flat_all):
        score += 7
        reasons.append("narrative_detected")

    if looks_like_layout(flat_all):
        score -= 35
        reasons.append("layout_penalty")
    if looks_like_placeholder(flat_all):
        score -= 25
        reasons.append("placeholder_penalty")
    if row_count <= 3:
        score -= 20
        reasons.append("too_small_for_result_table")

    return max(score, 0), reasons


def infer_table_kind(sample_rows: List[List[str]], flat_all: str) -> str:
    normalized = flat_all.replace(" ", "").lower()
    if "사례수" in flat_all and "n" in normalized and "%" in flat_all:
        return "modern_n_percent"
    if "base|" in normalized or "base:" in normalized or "base｜" in normalized:
        return "modern_base_block"
    if "빈도" in flat_all and "%" in flat_all:
        return "legacy_frequency_percent"
    if looks_like_placeholder(flat_all):
        return "placeholder_style_source"
    if sample_rows and len(sample_rows[0]) <= 2 and any("목차" in cell for row in sample_rows for cell in row):
        return "layout_or_toc"
    return "unknown"


def infer_sequence(flat_before: str, flat_rows: str) -> List[str]:
    sequence = ["section_title", "narrative"]
    if has_figure_caption(flat_before):
        sequence.append("figure_caption")
    if has_base_or_unit(flat_before, flat_rows):
        sequence.append("base_note")
    sequence.append("chart_placeholder")
    if has_table_caption(flat_before):
        sequence.append("table_caption")
    sequence.append("result_table")
    sequence.append("source_note")
    return sequence


def infer_title(before_text: List[str], sample_rows: List[List[str]]) -> str:
    caption_title = infer_caption_title(before_text)
    if caption_title:
        return caption_title
    for text in reversed(before_text):
        cleaned = text.strip()
        if not cleaned or len(cleaned) > 120:
            continue
        if any(token in cleaned for token in ("표", "그림", "Base", "단위")):
            continue
        return cleaned
    for row in sample_rows[:2]:
        for cell in row:
            cleaned = str(cell).strip()
            if cleaned and not any(token in cleaned for token in ("Base", "단위", "구분", "사례수", "빈도")):
                return cleaned[:120]
    return ""


def infer_caption_title(before_text: List[str]) -> str:
    patterns = [
        r"\[표[^\]]*\]\s*(.+)",
        r"【표[^】]*】\s*(.+)",
        r"<표[^>]*>\s*(.+)",
        r"표-\s*\]\s*(.+)",
    ]
    stop_tokens = ["(단위", "[단위", "Base", "구분", "빈도", "사례수", " n ", " % "]
    for text in reversed(before_text):
        for pattern in patterns:
            match = re.search(pattern, text)
            if not match:
                continue
            title = match.group(1).strip()
            for token in stop_tokens:
                if token in title:
                    title = title.split(token, 1)[0].strip()
            title = re.sub(r"\s+", " ", title)
            if 2 <= len(title) <= 120:
                return title
    return ""


def flatten_rows(rows: List[List[str]]) -> str:
    return " ".join(" ".join(str(cell) for cell in row) for row in rows)


def has_narrative(text: str) -> bool:
    return any(token in text for token in ("조사한 결과", "나타남", "나타났", "응답", "비율", "가장 높"))


def has_figure_caption(text: str) -> bool:
    return any(token in text for token in ("그림", "[그림", "【그림", "<그림"))


def has_table_caption(text: str) -> bool:
    return any(token in text for token in ("[표", "【표", "<표", "표-"))


def has_base_or_unit(*texts: str) -> bool:
    text = " ".join(texts)
    return any(token in text for token in ("Base", "BASE", "BASE :", "Base |", "[BASE", "단위", "(단위"))


def looks_like_placeholder(text: str) -> bool:
    return any(token in text for token in PLACEHOLDER_TERMS)


def looks_like_layout(text: str) -> bool:
    return any(token in text for token in NEGATIVE_TERMS)


def load_json(path: str | Path) -> Dict[str, Any]:
    return json.loads(Path(path).read_text(encoding="utf-8"))


def write_json(path: str | Path, value: Dict[str, Any]) -> None:
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    Path(path).write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")


def now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def run_cli(argv: List[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Build template blueprint candidates from hwp_template_probe output.")
    parser.add_argument("--probe", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args(argv)

    blueprint = build_blueprint(args.probe, args.output)
    print(
        json.dumps(
            {
                "status": blueprint.get("status"),
                "output": str(Path(args.output).resolve()),
                "template_count": len(blueprint.get("templates", [])),
                "block_count": sum(len(t.get("blocks", [])) for t in blueprint.get("templates", [])),
            },
            ensure_ascii=True,
            indent=2,
        )
    )
    return 0 if blueprint.get("status") in {"ready", "needs_review"} else 2


if __name__ == "__main__":
    raise SystemExit(run_cli())
