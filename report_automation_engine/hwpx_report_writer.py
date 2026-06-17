"""HWPX 집계표 기반 보고서 본문 초안 생성기.

사용자가 제공한 `hwpx_report_writer_revised4` 계열 코드를 프로젝트 안으로 편입한
파일이다. 이름은 writer지만 현재 기능의 핵심은 기존 HWPX 안에 들어 있는 표를
읽고, 표별 전체/지역/특성 요약 문장을 TXT로 만드는 것이다.

프로젝트에서의 권장 사용 방식:
- HWPX 템플릿에 직접 본문을 삽입하는 기능은 별도 단계에서 구현한다.
- 이 모듈은 먼저 "HWPX 표 구조를 어떻게 해석할지" 검증하는 분석용 엔진으로 둔다.
- 실제 HWP/HWPX 파일 생성은 한글 COM API 또는 HWPX XML 생성기를 별도로 붙일 때
  이 모듈의 문장 생성 함수를 재사용한다.
"""

from __future__ import annotations

import argparse
import re
import zipfile
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple
import tkinter as tk
from tkinter import filedialog
import xml.etree.ElementTree as ET


K_TOTAL = "전체"
K_TABLE = "<표"
K_FIGURE = "<그림"
K_UNIT = "단위"
K_CASE = "사례수"
K_DISTRICT = "주요 배너별"
K_CHARACTERISTIC = "사업체 특성별"
K_MEAN = "평균"
K_BASE = "Base :"

OUTPUT_HEADER = "자동 생성 보고서 본문"
OUTPUT_DONE = "완료"
TITLE_UNKNOWN = "제목 미확인"


@dataclass
class Block:
    """HWPX 본문 흐름에서 추출한 블록.

    kind가 `paragraph`이면 text만 사용하고, `table`이면 table에 정규화된 표 데이터가
    들어간다. 문단과 표의 순서를 유지해야 "표 바로 위 제목"을 찾을 수 있다.
    """

    kind: str
    text: str = ""
    table: Optional["TableData"] = None


@dataclass
class TableRow:
    """HWPX 표의 한 행을 보고서 문장 생성용으로 정규화한 구조."""

    label: str
    group: str
    subgroup: str
    case_count: Optional[float]
    values: Dict[str, Optional[float]]


@dataclass
class TableData:
    """HWPX 표 하나를 문장 생성에 필요한 속성으로 묶은 구조."""

    title: str
    normalized_title: str
    subtype: str
    unit_text: str
    columns: List[str]
    response_columns: List[str]
    rows: List[TableRow]
    total_row: Optional[TableRow]
    multi_response: bool
    mean_columns: List[str] = field(default_factory=list)


def extract_namespaces(xml_bytes: bytes) -> Dict[str, str]:
    namespaces: Dict[str, str] = {}
    for _, elem in ET.iterparse(BytesIO(xml_bytes), events=("start-ns",)):
        prefix, uri = elem
        namespaces[prefix or "default"] = uri
    return namespaces


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def clean_text(text: str) -> str:
    text = text.replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def parse_number(raw: str) -> Optional[float]:
    value = clean_text(raw).replace(",", "")
    if not value or value == "-":
        return None
    value = value.replace("%", "")
    match = re.search(r"-?\d+(?:\.\d+)?", value)
    if not match:
        return None
    try:
        return float(match.group())
    except ValueError:
        return None


def format_number(value: Optional[float], digits: int = 1) -> str:
    if value is None:
        return "-"
    if abs(value - round(value)) < 1e-9:
        return f"{int(round(value)):,}"
    return f"{value:,.{digits}f}"


def format_percent(value: Optional[float]) -> str:
    if value is None:
        return "-"
    return f"{format_number(value, 1)}%"


def last_meaningful_char(text: str) -> str:
    cleaned = clean_text(text).rstrip("'\"”)]} ")
    return cleaned[-1] if cleaned else ""


def has_batchim(text: str) -> bool:
    char = last_meaningful_char(text)
    if not char:
        return False

    code = ord(char)
    if 0xAC00 <= code <= 0xD7A3:
        return (code - 0xAC00) % 28 != 0

    if char.isdigit():
        return char in {"0", "1", "3", "6", "7", "8"}

    lowered = char.lower()
    if "a" <= lowered <= "z":
        return lowered in {"b", "c", "k", "l", "m", "n", "p", "t"}

    return False


def particle(word: str, with_batchim: str, without_batchim: str) -> str:
    return with_batchim if has_batchim(word) else without_batchim


def topic(word: str) -> str:
    return particle(word, "은", "는")


def subject(word: str) -> str:
    return particle(word, "이", "가")


def quoted(word: str) -> str:
    return f"'{word}'"


def strip_title_prefix(title: str) -> str:
    title = clean_text(title)
    title = re.sub(r"^<[^>]*>\s*", "", title)
    title = re.sub(r"\[\s*Base\s*:[^\]]+\]", "", title)
    title = re.sub(r"\[\s*[^\]]*" + K_UNIT + r"\s*:[^\]]+\]", "", title)
    for marker in (K_CASE, K_TOTAL, "진흥지구명", "사업체 특성"):
        pos = title.find(marker)
        if pos > 0:
            title = title[:pos]
            break
    return clean_text(title)


def normalize_analysis_title(title: str) -> str:
    title = strip_title_prefix(title)
    for prefix in (K_DISTRICT + " ", K_CHARACTERISTIC + " "):
        if title.startswith(prefix):
            title = title[len(prefix) :]
    title = re.sub(
        r"\s+(?:입주년도|종사기간|조직형태|매출액|"
        r"사업체 규모(?:\(종사자 수\))?|"
        r"사업장 면적|계약기간|소유형태|"
        r"업종|사업활동 범위)\s+"
        r"(?:진흥지구명|사업체 특성)$",
        "",
        title,
    )
    return clean_text(title)


def detect_table_subtype(title: str) -> str:
    title = strip_title_prefix(title)
    if K_DISTRICT in title:
        return "district"
    if K_CHARACTERISTIC in title:
        return "characteristic"
    return "overall"


def collect_paragraph_text(elem: ET.Element) -> str:
    parts: List[str] = []
    for node in elem.iter():
        if local_name(node.tag) == "t" and node.text:
            parts.append(node.text)
    return clean_text(" ".join(parts))


def should_skip_table(table_elem: ET.Element, parent_map: Dict[ET.Element, ET.Element]) -> bool:
    row_count = int(parse_number(table_elem.attrib.get("rowCnt", "")) or 0)
    col_count = int(parse_number(table_elem.attrib.get("colCnt", "")) or 0)
    if row_count < 4 or col_count < 3:
        return True

    parent = parent_map.get(table_elem)
    while parent is not None:
        if local_name(parent.tag) in {"header", "footer", "footNote", "endNote"}:
            return True
        parent = parent_map.get(parent)
    return False


def build_table_grid(table_elem: ET.Element) -> List[List[str]]:
    rows = [child for child in table_elem if local_name(child.tag) == "tr"]
    grid: Dict[Tuple[int, int], str] = {}
    max_col = 0

    for row_idx, tr in enumerate(rows):
        col_idx = 0
        cells = [child for child in tr if local_name(child.tag) == "tc"]
        for tc in cells:
            while (row_idx, col_idx) in grid:
                col_idx += 1
            text = collect_paragraph_text(tc)
            col_span = 1
            row_span = 1
            for child in tc:
                if local_name(child.tag) == "cellSpan":
                    col_span = int(child.attrib.get("colSpan", "1"))
                    row_span = int(child.attrib.get("rowSpan", "1"))
                    break
            for rs in range(row_span):
                for cs in range(col_span):
                    grid[(row_idx + rs, col_idx + cs)] = text
            col_idx += col_span
            max_col = max(max_col, col_idx)

    matrix: List[List[str]] = []
    for row_idx in range(len(rows)):
        row = [clean_text(grid.get((row_idx, col_idx), "")) for col_idx in range(max_col)]
        if any(cell for cell in row):
            matrix.append(row)

    if not matrix:
        return []

    keep_indices = [
        idx
        for idx in range(len(matrix[0]))
        if any(idx < len(row) and clean_text(row[idx]) for row in matrix)
    ]
    return [[row[idx] for idx in keep_indices] for row in matrix]


def skip_candidate_title(text: str) -> bool:
    text = clean_text(text)
    if not text:
        return True
    if text.startswith(K_FIGURE) or text.startswith("[") or text.startswith(K_BASE):
        return True
    if K_CASE in text:
        return True
    if text.endswith("나타남"):
        return True
    if "%" in text and text.count(" ") > 5:
        return True
    return False


def find_title_for_table(blocks: Sequence[Block], block_index: int) -> str:
    for offset in range(1, 10):
        idx = block_index - offset
        if idx < 0:
            break
        block = blocks[idx]
        if block.kind != "paragraph":
            continue
        text = clean_text(block.text)
        if text.startswith(K_TABLE):
            return text
        if skip_candidate_title(text):
            continue
        return text
    return TITLE_UNKNOWN


def split_row_label_parts(row: List[str], case_idx: int) -> Tuple[str, str, str]:
    left = [clean_text(cell) for cell in row[:case_idx] if clean_text(cell)]
    if not left:
        return "", "", ""
    if len(left) == 1:
        return left[0], "", left[0]
    group = left[0]
    subgroup = left[-1]
    if group == subgroup:
        subgroup = ""
    return subgroup or group, group, subgroup


def build_table_data(title: str, matrix: List[List[str]]) -> Optional[TableData]:
    if not matrix:
        return None

    unit_text = matrix[0][0] if matrix[0] and K_UNIT in matrix[0][0] else ""
    header_idx = next(
        (idx for idx, row in enumerate(matrix) if any(K_CASE in cell for cell in row)),
        None,
    )
    if header_idx is None:
        return None

    header_row = matrix[header_idx]
    case_idx = next((idx for idx, cell in enumerate(header_row) if K_CASE in cell), None)
    if case_idx is None or case_idx == len(header_row) - 1:
        return None

    columns = [clean_text(cell) for cell in header_row]
    response_columns = [clean_text(cell) for cell in columns[case_idx + 1 :] if clean_text(cell)]
    mean_columns = [col for col in response_columns if K_MEAN in col]
    value_columns = [col for col in response_columns if col not in mean_columns]

    rows: List[TableRow] = []
    total_row: Optional[TableRow] = None
    for raw_row in matrix[header_idx + 1 :]:
        if len(raw_row) <= case_idx:
            continue
        label, group, subgroup = split_row_label_parts(raw_row, case_idx)
        if not label:
            continue
        values: Dict[str, Optional[float]] = {}
        for idx, col_name in enumerate(response_columns, start=case_idx + 1):
            values[col_name] = parse_number(raw_row[idx]) if idx < len(raw_row) else None
        row = TableRow(
            label=label,
            group=group,
            subgroup=subgroup,
            case_count=parse_number(raw_row[case_idx]),
            values=values,
        )
        rows.append(row)
        if label.replace(" ", "") == K_TOTAL:
            total_row = row

    if not rows:
        return None

    numeric_total = 0.0
    if total_row is not None:
        numeric_total = sum((total_row.values.get(col) or 0.0) for col in value_columns)

    return TableData(
        title=strip_title_prefix(title),
        normalized_title=normalize_analysis_title(title),
        subtype=detect_table_subtype(title),
        unit_text=unit_text,
        columns=columns,
        response_columns=response_columns,
        rows=rows,
        total_row=total_row,
        multi_response=numeric_total > 105.0 if value_columns else False,
        mean_columns=mean_columns,
    )


def parse_hwpx_blocks(hwpx_path: Path) -> List[Block]:
    blocks: List[Block] = []
    with zipfile.ZipFile(hwpx_path, "r") as archive:
        section_names = sorted(
            name
            for name in archive.namelist()
            if name.startswith("Contents/section") and name.endswith(".xml")
        )
        for section_name in section_names:
            xml_bytes = archive.read(section_name)
            root = ET.fromstring(xml_bytes)
            _ = extract_namespaces(xml_bytes)
            parent_map = {child: parent for parent in root.iter() for child in parent}

            raw_blocks: List[Block] = []
            table_nodes: List[ET.Element] = []
            for elem in root.iter():
                name = local_name(elem.tag)
                if name == "p":
                    parent = parent_map.get(elem)
                    skip = False
                    while parent is not None:
                        if local_name(parent.tag) in {"tbl", "header", "footer", "footNote", "endNote"}:
                            skip = True
                            break
                        parent = parent_map.get(parent)
                    if skip:
                        continue
                    text = collect_paragraph_text(elem)
                    if text:
                        raw_blocks.append(Block(kind="paragraph", text=text))
                elif name == "tbl":
                    if should_skip_table(elem, parent_map):
                        continue
                    raw_blocks.append(Block(kind="table"))
                    table_nodes.append(elem)

            table_positions = [idx for idx, block in enumerate(raw_blocks) if block.kind == "table"]
            for idx, table_elem in zip(table_positions, table_nodes):
                title = find_title_for_table(raw_blocks, idx)
                table = build_table_data(title, build_table_grid(table_elem))
                if table is not None:
                    raw_blocks[idx].table = table

            for block in raw_blocks:
                if block.kind == "table" and block.table is None:
                    continue
                blocks.append(block)
    return blocks


def choose_value_columns(table: TableData) -> List[str]:
    return [col for col in table.response_columns if col not in table.mean_columns]


def rank_columns(row: TableRow, columns: Sequence[str]) -> List[Tuple[str, float]]:
    ranked: List[Tuple[str, float]] = []
    for col in columns:
        value = row.values.get(col)
        if value is not None:
            ranked.append((col, value))
    ranked.sort(key=lambda item: (-item[1], item[0]))
    return ranked


# -------------------------------------------------------------------------
# NEW: 관심 응답(Focus Column) 추출 로직
# 인지도 조사에서 '모른다'가 1위더라도, 심층 분석에서는 '알고 있다', '예' 등을 기준으로 잡도록 함
# -------------------------------------------------------------------------
def get_focus_column(table: TableData, value_columns: List[str]) -> str:
    # 긍정적 혹은 우리가 중점적으로 분석하고 싶은 응답 키워드 (우선순위)
    priority_keywords = ["알고 있다", "알고있다", "경험 있다", "있다", "예", "필요하다", "만족", "도움"]
    
    for keyword in priority_keywords:
        for col in value_columns:
            if keyword in col:
                return col
                
    # 우선순위 키워드에 해당하는 항목이 없으면 기존대로 합계(전체)에서 가장 높은 비율을 차지한 항목 선택
    if table.total_row is not None:
        ranked = rank_columns(table.total_row, value_columns)
        if ranked:
            return ranked[0][0]
            
    return value_columns[0] if value_columns else ""


def overall_summary(table: TableData) -> str:
    # 전체 요약은 그대로 통계적으로 가장 높은 비율(ex. '모른다' 79.8%)을 먼저 보여줍니다.
    if table.total_row is None:
        return f"{table.normalized_title} 문항의 전체 응답을 요약하지 못했습니다."

    total = table.total_row
    value_columns = choose_value_columns(table)
    ranked = rank_columns(total, value_columns)
    if not ranked:
        return f"{table.normalized_title} 문항의 수치형 응답을 찾지 못했습니다."

    top1 = ranked[0]
    others = ranked[1:4]
    if table.multi_response:
        sentence = (
            f"{table.normalized_title}{topic(table.normalized_title)} 복수응답 기준으로 "
            f"{quoted(top1[0])}{subject(top1[0])} {format_percent(top1[1])}로 가장 높게 나타남"
        )
    else:
        sentence = (
            f"{table.normalized_title}{topic(table.normalized_title)} {quoted(top1[0])}{subject(top1[0])} "
            f"{format_percent(top1[1])}로 가장 높게 나타남"
        )

    if others:
        joined = ", ".join(f"'{name}'({format_percent(value)})" for name, value in others)
        sentence += f"\n다음으로 {joined} 순으로 나타남"

    if table.mean_columns:
        mean_parts = [
            f"{mean_col}{topic(mean_col)} {format_number(total.values.get(mean_col), 1)}"
            for mean_col in table.mean_columns
        ]
        sentence += "\n" + ", ".join(mean_parts) + "으로 집계됨"

    return sentence


def district_summary(table: TableData) -> str:
    value_columns = choose_value_columns(table)
    if not value_columns:
        return ""

    # 분석 기준이 될 관심 응답(예: "알고 있다") 컬럼 획득
    focus_col = get_focus_column(table, value_columns)
    if not focus_col:
        return ""

    district_rows = [row for row in table.rows if row.label.replace(" ", "") != K_TOTAL]
    
    # 해당 관심 응답(focus_col)의 비율이 높은 순서대로 주요 배너 항목을 정렬
    ranked_districts = []
    for row in district_rows:
        val = row.values.get(focus_col)
        if val is not None and val > 0:
            ranked_districts.append((row.label, val))
            
    # 비율 내림차순 정렬
    ranked_districts.sort(key=lambda x: x[1], reverse=True)
    
    # 상위 3개 배너 항목 추출
    top_districts = ranked_districts[:3]
    if not top_districts:
        return ""
        
    parts = [f"'{name}'({format_percent(val)})" for name, val in top_districts]
    joined = ", ".join(parts)
    
    return f"주요 배너별로 보면, {quoted(focus_col)} 응답은 {joined} 등의 순으로 나타남"


def select_characteristic_groups(table: TableData) -> List[str]:
    # 이전 수정사항: 사업장 소유형태 제외 및 배너변수 최대 7개로 확대
    preferred = [
        "입주",
        "입주년도",
        "업종 운영기간",
        "조직형태",
        "사업체 규모 (종사자 수)",
        "사업체 규모",
        "종사자 수",
        "2025년 매출액",
        "2018년 매출액",
        "사업장 면적",
    ]

    groups: List[str] = []
    seen = set()
    for wanted in preferred:
        for row in table.rows:
            # 소유형태 배제
            if "소유형태" in row.group.replace(" ", ""):
                continue
                
            if row.group == wanted and wanted not in seen:
                seen.add(wanted)
                groups.append(wanted)
    if groups:
        return groups[:7]

    for row in table.rows:
        if row.group and row.group not in seen:
            if "소유형태" not in row.group.replace(" ", ""):
                seen.add(row.group)
                groups.append(row.group)
    return groups[:7]


def characteristic_summary(table: TableData) -> str:
    value_columns = choose_value_columns(table)
    if not value_columns:
        return ""

    # 분석 기준이 될 관심 응답(예: "알고 있다") 컬럼 획득
    focus_col = get_focus_column(table, value_columns)
    if not focus_col:
        return ""

    descriptions = []
    for group in select_characteristic_groups(table):
        candidates = [
            row
            for row in table.rows
            if row.group == group and row.label.replace(" ", "") != K_TOTAL
        ]
        if not candidates:
            continue
        
        # 각 배너 변수(그룹) 내에서 관심 응답(focus_col) 비율이 가장 높은 하위 속성 찾기
        best = max(candidates, key=lambda row: row.values.get(focus_col) or -1.0)
        val = best.values.get(focus_col)
        
        if val is not None and val > 0:
            subgroup = best.subgroup or best.label
            descriptions.append(f"{group} '{subgroup}'({format_percent(val)})")

    if not descriptions:
        return ""

    joined = ", ".join(descriptions)
    return (
        f"사업체 특성별로 보면, {quoted(focus_col)} 응답은 "
        f"{joined}에서 높게 나타남"
    )


def render_question_section(title: str, tables: Sequence[TableData]) -> str:
    overall_table = next((table for table in tables if table.subtype == "overall"), None)
    district_table = next((table for table in tables if table.subtype == "district"), None)
    characteristic_table = next((table for table in tables if table.subtype == "characteristic"), None)

    primary = overall_table or district_table or characteristic_table
    if primary is None:
        return ""

    lines = [title, "", overall_summary(primary)]
    if district_table is not None:
        text = district_summary(district_table)
        if text:
            lines.extend(["", text])
    if characteristic_table is not None:
        text = characteristic_summary(characteristic_table)
        if text:
            lines.extend(["", text])
    return "\n".join(lines).strip()


def group_tables(blocks: Sequence[Block]) -> List[Tuple[str, List[TableData]]]:
    groups: List[Tuple[str, List[TableData]]] = []
    current_title = ""
    current_tables: List[TableData] = []

    for block in blocks:
        if block.kind != "table" or block.table is None:
            continue
        table = block.table
        if current_title and table.normalized_title != current_title:
            groups.append((current_title, current_tables))
            current_tables = []
        current_title = table.normalized_title
        current_tables.append(table)

    if current_tables:
        groups.append((current_title, current_tables))
    return groups


def render_report(hwpx_path: Path) -> str:
    blocks = parse_hwpx_blocks(hwpx_path)
    sections = []
    for title, tables in group_tables(blocks):
        section = render_question_section(title, tables)
        if section:
            sections.append(section)
    return f"{hwpx_path.name} {OUTPUT_HEADER}\n\n" + "\n\n".join(sections)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="HWPX 집계표를 읽어 보고서 본문 초안을 생성합니다."
    )
    parser.add_argument("hwpx", nargs="?", type=Path, help="입력 HWPX 파일 경로")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="출력 TXT 파일 경로. 생략 시 _자동작성.txt로 저장합니다.",
    )
    args = parser.parse_args()

    hwpx_path = args.hwpx
    if hwpx_path is None:
        root = tk.Tk()
        root.withdraw()
        selected = filedialog.askopenfilename(
            title="HWPX 파일 선택",
            filetypes=[("HWPX files", "*.hwpx"), ("All files", "*.*")],
        )
        root.destroy()
        if not selected:
            print("선택된 파일이 없습니다.")
            return
        hwpx_path = Path(selected)

    if not hwpx_path.exists():
        raise FileNotFoundError(f"HWPX 파일을 찾을 수 없습니다: {hwpx_path}")

    report_text = render_report(hwpx_path)
    output_path = args.output or hwpx_path.with_name(f"{hwpx_path.stem}_자동작성.txt")
    output_path.write_text(report_text, encoding="utf-8")
    print(f"{OUTPUT_DONE}: {output_path}")


if __name__ == "__main__":
    main()
