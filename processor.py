from __future__ import annotations

import math
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import excel2img


SEVERITY_COLORS: Dict[str, Optional[str]] = {
    "very low": "00B050",
    "low": "00B050",
    "medium": "FFFF00",
    "high": "F4B183",
    "very high": "FF0000",
    "critical": "FF0000",
    "unknown": None,
    "unknow": None,
}


HEADER_FILL = PatternFill(fill_type="solid", fgColor="D9E2F3")
HEADER_FONT = Font(bold=True, size=13)

THIN_BORDER = Border(
    left=Side(style="thin", color="000000"),
    right=Side(style="thin", color="000000"),
    top=Side(style="thin", color="000000"),
    bottom=Side(style="thin", color="000000"),
)



# Utility functions


def normalize_text(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def find_column_index(ws, column_name: str) -> Optional[int]:
    for col in range(1, ws.max_column + 1):
        if normalize_text(ws.cell(1, col).value).lower() == column_name.lower():
            return col
    return None


def delete_column_if_needed(ws, column_name: str, should_delete: bool) -> None:
    idx = find_column_index(ws, column_name)
    if idx and should_delete:
        ws.delete_cols(idx, 1)


def all_values_are_na(values: List[str]) -> bool:
    cleaned = [v for v in values if v != ""]
    if not cleaned:
        return False
    return all(v.upper() == "N/A" for v in cleaned)


def all_values_empty(values: List[str]) -> bool:
    return all(v == "" for v in values)



# Formatting


def apply_basic_formatting(ws):

    for col in range(1, ws.max_column + 1):
        header = ws.cell(1, col)
        header.fill = HEADER_FILL
        header.font = HEADER_FONT
        header.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        header.border = THIN_BORDER

    for row in range(2, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row, col)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.border = THIN_BORDER


def color_severity_column(ws):

    idx = find_column_index(ws, "Severity")
    if not idx:
        return

    for row in range(2, ws.max_row + 1):

        cell = ws.cell(row, idx)
        text = normalize_text(cell.value).lower()

        fill = None

        if "very low" in text:
            fill = SEVERITY_COLORS["very low"]

        elif "very high" in text:
            fill = SEVERITY_COLORS["very high"]

        elif "critical" in text:
            fill = SEVERITY_COLORS["critical"]

        elif text in ["unknown", "unknow"]:
            fill = None

        elif "medium" in text:
            fill = SEVERITY_COLORS["medium"]

        elif "high" in text:
            fill = SEVERITY_COLORS["high"]

        elif "low" in text:
            fill = SEVERITY_COLORS["low"]

        if fill:
            cell.fill = PatternFill(fill_type="solid", fgColor=fill)



# Layout optimisation


def optimize_layout(ws):

    preferred_widths = {
        "Created": 18,
        "Position in current video": 16,
        "Position from first video": 16,
        "Position from the start (m)": 12,
        "Distance from the previous manhole (m)": 18,
        "Code": 8,
        "Characteristic 1": 12,
        "Characteristic 2": 12,
        "Observation type": 34,
        "Clockface references": 12,
        "Continuing defect": 12,
        "End of": 12,
        "Observation step": 12,
        "Note": 24,
        "Severity": 12,
        "Longitude": 12,
        "Latitude": 12,
    }

    min_width = 8
    max_width_default = 18

    for col in range(1, ws.max_column + 1):

        header = normalize_text(ws.cell(1, col).value)

        if header in preferred_widths:
            ws.column_dimensions[get_column_letter(col)].width = preferred_widths[header]
            continue

        max_len = len(header)

        for row in range(2, ws.max_row + 1):

            value = normalize_text(ws.cell(row, col).value)

            line_max = max((len(part) for part in value.splitlines()), default=0)

            max_len = max(max_len, line_max)

        estimated = min(max(max_len + 2, min_width), max_width_default)

        ws.column_dimensions[get_column_letter(col)].width = estimated



# Main processing function


def process_excel(
    input_xlsx: str,
    output_xlsx: str,
    output_png: str,
    sheet_name: Optional[str] = None
):

    wb = load_workbook(input_xlsx)

    ws = wb[sheet_name] if sheet_name else wb.active


    # delete Video column
    delete_column_if_needed(ws, "Video", True)


    # Observation step
    idx = find_column_index(ws, "Observation step")

    if idx:
        values = [normalize_text(ws.cell(r, idx).value) for r in range(2, ws.max_row + 1)]
        delete_column_if_needed(ws, "Observation step", all_values_are_na(values))


    # Note column
    idx = find_column_index(ws, "Note")

    if idx:
        values = [normalize_text(ws.cell(r, idx).value) for r in range(2, ws.max_row + 1)]
        delete_column_if_needed(ws, "Note", all_values_empty(values))


    apply_basic_formatting(ws)

    color_severity_column(ws)

    optimize_layout(ws)


    wb.save(output_xlsx)


    # export PNG using Excel formatting
    excel2img.export_img(
        output_xlsx,
        output_png,
        sheet_name=sheet_name or ws.title
    )


    return output_xlsx, output_png
