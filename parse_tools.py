# === parse_tools.py (formerly bounding_box.py) ===
# parse_range and presence-matrix logic commented out; not used in current app.
# Presence matrix for this app is built in main.py: get_presence_matrix_for_sheet().

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries, get_column_letter
import re

# from tools.parse_tools.cell_matrix_builder import build_cell_matrix
# from tools.parse_tools.formatting_extractor import extract_formatting

# Regular expression for finding cell references in formulas (used if parse_range is restored)
CELL_REF_PATTERN = re.compile(r"\b([A-Z]{1,3}[0-9]{1,7})\b")

def get_presence_matrix_for_sheet(file_path: str, sheet_name: str, cell_range: str):
    """
    Build 0/1 presence matrix for a sheet range (1 = cell has value or formula, 0 = empty).
    Returns list of lists [row][col] 0-indexed, matching block_generation expectations.
    """
    min_col, min_row, max_col, max_row = range_boundaries(cell_range)
    wb = load_workbook(file_path, data_only=False)
    try:
        ws = wb[sheet_name]
        P = []
        for r in range(min_row, max_row + 1):
            row_vals = []
            for c in range(min_col, max_col + 1):
                cell = ws.cell(row=r, column=c)
                has_content = (
                    cell.value is not None
                    and str(cell.value).strip() != ""
                ) or cell.data_type == "f"
                row_vals.append(1 if has_content else 0)
            P.append(row_vals)
        return P
    finally:
        wb.close()


def bbox_to_excel_range(bbox: "BBox", base_row: int = 1, base_col: int = 1) -> str:
    """Convert 0-indexed BBox to Excel range string (e.g. A1:D10)."""
    r1 = bbox.r1 + base_row
    c1 = bbox.c1 + base_col
    r2 = bbox.r2 + base_row
    c2 = bbox.c2 + base_col
    return f"{get_column_letter(c1)}{r1}:{get_column_letter(c2)}{r2}"



# def parse_range(file_path: str, sheet_name: str, cell_range: str, parse_tools: list) -> dict:
#     """
#     Parses a specific range in an Excel sheet using the specified parse tools.
#
#     Args:
#         file_path: Path to the Excel file
#         sheet_name: Name of the sheet to parse
#         cell_range: Excel range to parse (e.g., "A1:C10")
#         parse_tools: List of parse tools to apply. Options:
#             - "cell_matrix": Build 0/1 matrix for the range
#             - "formatting": Extract cell formatting
#             - "formula": Extract formulas and formula links
#             - "cell_data": Extract cell values, formulas, and formatting
#
#     Returns:
#         dict: Parsed data with structure:
#         {
#             "cell_matrix": [[0, 1, 0, ...], ...],  # if "cell_matrix" in parse_tools
#             "formatting": {...},  # if "formatting" in parse_tools
#             "formulas": [...],  # if "formula" in parse_tools
#             "cell_data": [...],  # if "cell_data" in parse_tools
#         }
#     """
#     # Load workbook
#     wb = load_workbook(file_path, data_only=False)
#     wb_values = load_workbook(file_path, data_only=True)
#
#     try:
#         ws = wb[sheet_name]
#         ws_values = wb_values[sheet_name]
#
#         # Parse the range
#         min_col, min_row, max_col, max_row = range_boundaries(cell_range)
#
#         result = {}
#
#         # Apply cell_matrix tool
#         if "cell_matrix" in parse_tools:
#             # Build DataFrame for the range
#             df = pd.DataFrame([[cell.value for cell in row] for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col)])
#             matrix = build_cell_matrix(df)
#             result["cell_matrix"] = matrix
#
#         # Apply formatting tool
#         if "formatting" in parse_tools:
#             sheet_formatting = extract_formatting(file_path)
#             range_formatting = {}
#             for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
#                 for cell in row:
#                     coord = cell.coordinate
#                     if coord in sheet_formatting.get(sheet_name, {}):
#                         range_formatting[coord] = sheet_formatting[sheet_name][coord]
#             result["formatting"] = range_formatting
#
#         # Apply formula tool
#         if "formula" in parse_tools:
#             sheet_formulas = []
#             for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
#                 for cell in row:
#                     if cell.data_type == "f":  # is formula
#                         target = cell.coordinate
#                         formula_str = str(cell.value)
#                         sources = set(CELL_REF_PATTERN.findall(formula_str))
#                         for source in sources:
#                             sheet_formulas.append({
#                                 "from": source,
#                                 "to": target,
#                                 "formula": formula_str,
#                                 "sheet": sheet_name
#                             })
#             result["formulas"] = sheet_formulas
#
#         # Apply cell_data tool (extracts values, formulas, and formatting together)
#         if "cell_data" in parse_tools:
#             # Get formatting if not already extracted
#             if "formatting" not in result:
#                 sheet_formatting = extract_formatting(file_path)
#             else:
#                 sheet_formatting = {sheet_name: result["formatting"]}
#
#             cell_data = []
#             for row, row_values in zip(
#                 ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col),
#                 ws_values.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col)
#             ):
#                 for cell, cell_value in zip(row, row_values):
#                     coord = cell.coordinate
#                     if cell.value is None or str(cell.value).strip() == "":
#                         continue  # Skip empty cells
#
#                     if cell.data_type == "f":
#                         value = cell_value.value
#                         formula = cell.value
#                     else:
#                         value = cell.value
#                         formula = None
#
#                     cell_data.append({
#                         "cell": coord,
#                         "value": value,
#                         "formula": formula,
#                         "formatting": sheet_formatting.get(sheet_name, {}).get(coord, {})
#                     })
#             result["cell_data"] = cell_data
#
#         return result
#
#     finally:
#         wb.close()
#         wb_values.close()
