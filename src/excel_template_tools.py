import copy
import json
import shutil
from io import BytesIO
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Any

import pandas as pd
from openpyxl import load_workbook
from pydantic import BaseModel, ConfigDict


DEFAULT_MAPPING = {
    "sheet_name": "Sheet1",
    "report_date_cell": "C2",
    "data_table": {
        "header_row": 4,
        "data_start_row": 5,
        "data_end_row": 8,
        "columns": {
            "media": "A",
            "impression": "B",
            "click": "C",
            "cost": "D",
            "conversion": "E",
        },
    },
    "formula_columns": ["F", "G"],
    "total_row": 9,
    "note_cell": "A12",
}


class MappingColumns(BaseModel):
    model_config = ConfigDict(extra="forbid")

    media: str
    impression: str
    click: str
    cost: str
    conversion: str


class MappingTable(BaseModel):
    model_config = ConfigDict(extra="forbid")

    header_row: int
    data_start_row: int
    data_end_row: int
    columns: MappingColumns


class TemplateMapping(BaseModel):
    model_config = ConfigDict(extra="forbid")

    sheet_name: str
    report_date_cell: str
    data_table: MappingTable
    formula_columns: list[str]
    total_row: int
    note_cell: str


def extract_template_structure(workbook_bytes: bytes, max_rows: int = 30, max_cols: int = 15) -> dict[str, Any]:
    """Extract visible template structure for AI-assisted mapping."""
    wb = load_workbook(BytesIO(workbook_bytes), data_only=False)
    result: dict[str, Any] = {"sheets": []}

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        sheet_info: dict[str, Any] = {
            "name": sheet_name,
            "dimensions": ws.dimensions,
            "merged_ranges": [str(r) for r in ws.merged_cells.ranges],
            "cells": [],
        }

        for row in ws.iter_rows(min_row=1, max_row=max_rows, min_col=1, max_col=max_cols):
            for cell in row:
                if cell.value is None:
                    continue

                cell_info: dict[str, Any] = {
                    "addr": cell.coordinate,
                    "row": cell.row,
                    "col": cell.column,
                    "value": str(cell.value),
                }
                if cell.font and cell.font.bold:
                    cell_info["bold"] = True
                if cell.fill and cell.fill.start_color and cell.fill.start_color.rgb:
                    rgb = cell.fill.start_color.rgb
                    if isinstance(rgb, str) and rgb not in ("00000000", "FFFFFFFF"):
                        cell_info["fill"] = rgb
                if cell.number_format and cell.number_format != "General":
                    cell_info["fmt"] = cell.number_format
                sheet_info["cells"].append(cell_info)

        result["sheets"].append(sheet_info)

    return result


def build_ai_prompt(structure: dict[str, Any]) -> str:
    """Build a prompt that asks an LLM to infer replaceable data locations."""
    sheet_blocks = []

    for index, sheet in enumerate(structure["sheets"], start=1):
        grid_lines = []
        for cell in sheet["cells"]:
            flags = []
            if cell.get("bold"):
                flags.append("bold")
            if cell.get("fill"):
                flags.append(f"fill={cell['fill'][-6:]}")
            if cell.get("fmt"):
                flags.append(f"fmt={cell['fmt']}")
            flag_str = f" [{', '.join(flags)}]" if flags else ""
            grid_lines.append(f"  {cell['addr']}: \"{cell['value']}\"{flag_str}")

        merged = "\n".join(f"  - {m}" for m in sheet["merged_ranges"]) or "  (없음)"
        cells = "\n".join(grid_lines) or "  (값이 있는 셀 없음)"
        sheet_blocks.append(
            f"""[시트 {index}]
시트명: {sheet['name']}
데이터 영역: {sheet['dimensions']}

병합된 셀:
{merged}

셀 값 및 서식:
{cells}"""
        )

    sheet_names = ", ".join(sheet["name"] for sheet in structure["sheets"])

    return f"""다음은 광고 보고서 엑셀 템플릿의 전체 워크북 구조입니다.

[워크북 전체 시트 수] {len(structure["sheets"])}
[전체 시트] {sheet_names}

{chr(10).join(sheet_blocks)}

위 템플릿에서 매일 교체되어야 할 동적 데이터의 셀 위치를 추론하여
아래 JSON 스키마로 답해주세요. 설명 없이 JSON만 출력하세요.

{{
  "sheet_name": "데이터를 교체할 시트명",
  "report_date_cell": "보고일자가 들어갈 셀 주소",
  "data_table": {{
    "header_row": 헤더 행 번호,
    "data_start_row": 데이터 시작 행,
    "data_end_row": 데이터 끝 행,
    "columns": {{
      "media": "매체명 컬럼 letter",
      "impression": "노출수 컬럼 letter",
      "click": "클릭수 컬럼 letter",
      "cost": "비용 컬럼 letter",
      "conversion": "전환수 컬럼 letter"
    }}
  }},
  "formula_columns": ["수식이 들어있어 건드리면 안 되는 컬럼 letter들"],
  "total_row": 합계 행 번호,
  "note_cell": "비고/메모 입력 셀 주소"
}}
"""


def generate_mapping_with_gemini(
    structure: dict[str, Any],
    api_key: str,
    model: str = "gemini-2.5-flash",
) -> dict[str, Any]:
    """Ask the Gemini API to infer a template mapping from extracted workbook structure."""
    try:
        from google import genai
    except ImportError as exc:
        raise ImportError(
            "Gemini SDK를 불러오지 못했습니다. Streamlit Cloud가 requirements.txt의 "
            "google-genai 패키지를 다시 설치하도록 앱을 재부팅하거나 redeploy하세요."
        ) from exc

    client = genai.Client(api_key=api_key)
    prompt = build_ai_prompt(structure)
    system_instruction = (
        "You infer Excel template mappings. Return only fields that identify where daily data "
        "should be written. Preserve formulas and total rows by excluding them from editable "
        "data columns. Choose the most likely sheet when a workbook has multiple sheets."
    )

    response = client.models.generate_content(
        model=model,
        contents=prompt,
        config={
            "system_instruction": system_instruction,
            "response_mime_type": "application/json",
            "response_schema": TemplateMapping,
        },
    )

    parsed = response.parsed
    if isinstance(parsed, TemplateMapping):
        return parsed.model_dump()
    return TemplateMapping.model_validate_json(response.text).model_dump()


def read_data_table(uploaded_file: Any) -> pd.DataFrame:
    """Read daily replacement rows from CSV or Excel."""
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    return pd.read_excel(uploaded_file)


def normalize_daily_data(df: pd.DataFrame, report_date: str, note: str) -> dict[str, Any]:
    rows = df.where(pd.notnull(df), None).to_dict(orient="records")
    return {"report_date": report_date, "rows": rows, "note": note}


def snapshot_design(workbook_bytes: bytes, max_rows: int = 40, max_cols: int = 15) -> dict[str, Any]:
    wb = load_workbook(BytesIO(workbook_bytes), data_only=False)
    snap: dict[str, Any] = {
        "sheet_names": wb.sheetnames,
        "sheets": {},
    }

    for ws in wb.worksheets:
        sheet_snap: dict[str, Any] = {
            "merged_ranges": sorted(str(r) for r in ws.merged_cells.ranges),
            "column_widths": {k: v.width for k, v in ws.column_dimensions.items() if v.width},
            "row_heights": {k: v.height for k, v in ws.row_dimensions.items() if v.height},
            "chart_count": len(ws._charts),
            "cells": {},
        }
        for row in ws.iter_rows(min_row=1, max_row=max_rows, min_col=1, max_col=max_cols):
            for cell in row:
                info: dict[str, Any] = {}
                if cell.font:
                    info["font_bold"] = cell.font.bold
                    info["font_color"] = (
                        str(cell.font.color.rgb) if cell.font.color and cell.font.color.type == "rgb" else None
                    )
                    info["font_name"] = cell.font.name
                    info["font_size"] = cell.font.size
                if cell.fill and cell.fill.start_color:
                    info["fill"] = str(cell.fill.start_color.rgb)
                if cell.border:
                    info["has_border"] = bool(cell.border.left and cell.border.left.style)
                if cell.number_format:
                    info["fmt"] = cell.number_format
                if cell.alignment:
                    info["align"] = (cell.alignment.horizontal, cell.alignment.vertical)
                sheet_snap["cells"][cell.coordinate] = info
        snap["sheets"][ws.title] = sheet_snap
    return snap


def compare_snapshots(before: dict[str, Any], after: dict[str, Any]) -> dict[str, Any]:
    structural_diffs = []
    if before["sheet_names"] != after["sheet_names"]:
        structural_diffs.append("시트 목록 또는 순서가 변경되었습니다.")

    style_changed = []
    for sheet_name, before_sheet in before["sheets"].items():
        after_sheet = after["sheets"].get(sheet_name)
        if not after_sheet:
            structural_diffs.append(f"{sheet_name} 시트가 사라졌습니다.")
            continue

        if before_sheet["merged_ranges"] != after_sheet["merged_ranges"]:
            structural_diffs.append(f"{sheet_name} 시트의 병합 영역이 변경되었습니다.")
        if before_sheet["column_widths"] != after_sheet["column_widths"]:
            structural_diffs.append(f"{sheet_name} 시트의 컬럼 너비가 변경되었습니다.")
        if before_sheet["row_heights"] != after_sheet["row_heights"]:
            structural_diffs.append(f"{sheet_name} 시트의 행 높이가 변경되었습니다.")
        if before_sheet["chart_count"] != after_sheet["chart_count"]:
            structural_diffs.append(f"{sheet_name} 시트의 차트 개수가 변경되었습니다.")

        for addr, before_info in before_sheet["cells"].items():
            after_info = after_sheet["cells"].get(addr, {})
            for key in ["font_bold", "font_color", "fill", "fmt", "align", "has_border"]:
                if before_info.get(key) != after_info.get(key):
                    style_changed.append(
                        {
                            "sheet": sheet_name,
                            "cell": addr,
                            "property": key,
                            "before": before_info.get(key),
                            "after": after_info.get(key),
                        }
                    )

    return {
        "structural_diffs": structural_diffs,
        "style_diffs_count": len(style_changed),
        "style_diffs_sample": style_changed[:10],
    }


def inject_data(template_bytes: bytes, mapping: dict[str, Any], data: dict[str, Any]) -> bytes:
    """Copy the template and replace values only, preserving styles, formulas, and charts."""
    with NamedTemporaryFile(suffix=".xlsx", delete=False) as template_file:
        template_path = Path(template_file.name)
        template_file.write(template_bytes)

    output_path = template_path.with_name(f"{template_path.stem}_output.xlsx")
    try:
        shutil.copy(template_path, output_path)
        wb = load_workbook(output_path, data_only=False)
        sheet_name = mapping.get("sheet_name")
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active

        if mapping.get("report_date_cell"):
            ws[mapping["report_date_cell"]] = data.get("report_date")

        table = mapping["data_table"]
        cols = table["columns"]
        start_row = int(table["data_start_row"])
        end_row = int(table["data_end_row"])

        for index, row_data in enumerate(data["rows"]):
            excel_row = start_row + index
            if excel_row > end_row:
                break
            for field, col_letter in cols.items():
                if field in row_data:
                    ws[f"{col_letter}{excel_row}"] = row_data[field]

        if mapping.get("note_cell"):
            ws[mapping["note_cell"]] = data.get("note")

        wb.save(output_path)
        return output_path.read_bytes()
    finally:
        template_path.unlink(missing_ok=True)
        output_path.unlink(missing_ok=True)


def validate_mapping(mapping_text: str) -> dict[str, Any]:
    mapping = json.loads(mapping_text)
    required = ["report_date_cell", "data_table", "note_cell"]
    missing = [key for key in required if key not in mapping]
    if missing:
        raise ValueError(f"매핑 JSON에 필수 키가 없습니다: {', '.join(missing)}")
    if "columns" not in mapping["data_table"]:
        raise ValueError("data_table.columns가 필요합니다.")
    return mapping


def mapping_as_text(mapping: dict[str, Any] | None = None) -> str:
    return json.dumps(copy.deepcopy(mapping or DEFAULT_MAPPING), ensure_ascii=False, indent=2)
