"""보고서 자동화 문장 스타일 설정 로더.

이 파일은 보고서 문장 생성 규칙을 코드에서 분리하기 위한 작은 유틸리티다.
엑셀/HWPX 파서가 계속 바뀌더라도, 문장 패턴이나 제외할 항목명은 JSON만 수정해서
실험할 수 있게 하는 것이 목적이다.

현재 프로젝트에서의 위치:
- `report_automation_engine/config/default_style_schema.json`을 기본 설정으로 둔다.
- 외부 EXE 런처에서는 나중에 이 JSON 경로를 옵션으로 넘겨 사용자별 문체를 바꿀 수 있다.
- VBA add-in 자체와는 직접 결합하지 않는다. VBA는 산출 시트 생성, Python은 본문 초안
  생성이라는 경계를 유지해야 기존 매크로와 충돌이 적다.
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List


REQUIRED_TOP_LEVEL_KEYS = {
    "schema_version",
    "report_style",
    "title_rules",
    "ranking_rules",
    "table_rules",
    "characteristic_rules",
}


@dataclass
class StyleConfig:
    """문장 스타일과 표 해석 규칙 묶음.

    세부 필드를 dict로 유지하는 이유는 스타일 JSON을 자주 실험할 수 있게 하기 위해서다.
    대신 `load_style_config()`에서 필수 키와 리스트 타입을 먼저 검증해, 잘못된 설정이
    본문 생성 도중 늦게 터지지 않도록 한다.
    """

    schema_version: str
    report_style: Dict[str, Any]
    title_rules: Dict[str, Any]
    ranking_rules: Dict[str, Any]
    table_rules: Dict[str, Any]
    characteristic_rules: Dict[str, Any]


def _require_keys(payload: Dict[str, Any], keys: set[str], scope: str) -> None:
    missing = sorted(key for key in keys if key not in payload)
    if missing:
        raise ValueError(f"{scope} 필수 키 누락: {', '.join(missing)}")


def _require_list(payload: Dict[str, Any], key: str, scope: str) -> List[str]:
    value = payload.get(key)
    if not isinstance(value, list):
        raise ValueError(f"{scope}.{key} 는 리스트여야 합니다.")
    return [str(item) for item in value]


def load_style_config(path: str | Path) -> StyleConfig:
    config_path = Path(path)
    payload = json.loads(config_path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise ValueError("스타일 설정 파일은 JSON 객체여야 합니다.")

    _require_keys(payload, REQUIRED_TOP_LEVEL_KEYS, "root")

    report_style = payload["report_style"]
    title_rules = payload["title_rules"]
    ranking_rules = payload["ranking_rules"]
    table_rules = payload["table_rules"]
    characteristic_rules = payload["characteristic_rules"]

    for key, scope in (
        ("report_style", "root"),
        ("title_rules", "root"),
        ("ranking_rules", "root"),
        ("table_rules", "root"),
        ("characteristic_rules", "root"),
    ):
        if not isinstance(payload[key], dict):
            raise ValueError(f"{scope}.{key} 는 객체여야 합니다.")

    _require_keys(
        report_style,
        {
            "main_bullet",
            "sub_bullet",
            "top_pattern",
            "follow_pattern",
            "mean_pattern",
            "score_pattern",
            "characteristic_pattern",
            "district_pattern",
        },
        "report_style",
    )
    _require_keys(title_rules, {"strip_prefixes", "strip_markers", "skip_title_keywords"}, "title_rules")
    _require_keys(
        ranking_rules,
        {"exclude_labels", "max_follow_items", "ignore_zero", "deduplicate_labels"},
        "ranking_rules",
    )
    _require_keys(
        table_rules,
        {"case_column_name", "total_label", "mean_keywords", "summary_keywords", "score_keywords", "unit_keywords"},
        "table_rules",
    )
    _require_keys(characteristic_rules, {"preferred_groups", "max_groups"}, "characteristic_rules")

    _require_list(title_rules, "strip_prefixes", "title_rules")
    _require_list(title_rules, "strip_markers", "title_rules")
    _require_list(title_rules, "skip_title_keywords", "title_rules")
    _require_list(ranking_rules, "exclude_labels", "ranking_rules")
    _require_list(table_rules, "mean_keywords", "table_rules")
    _require_list(table_rules, "summary_keywords", "table_rules")
    _require_list(table_rules, "score_keywords", "table_rules")
    _require_list(table_rules, "unit_keywords", "table_rules")
    _require_list(characteristic_rules, "preferred_groups", "characteristic_rules")

    return StyleConfig(
        schema_version=str(payload["schema_version"]),
        report_style=report_style,
        title_rules=title_rules,
        ranking_rules=ranking_rules,
        table_rules=table_rules,
        characteristic_rules=characteristic_rules,
    )


def render_pattern(pattern: str, **kwargs: Any) -> str:
    return pattern.format(**kwargs)


if __name__ == "__main__":
    here = Path(__file__).resolve().parent
    sample_path = here / "style_schema.json"
    config = load_style_config(sample_path)
    print("loaded:", config.schema_version)
    print(
        render_pattern(
            config.report_style["top_pattern"],
            title="업종(세분류)",
            top_label="석유화학(뷰티/화장품)",
            top_value="48.4",
        )
    )
