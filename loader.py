# === loader.py (full loader commented out; app uses simple loader in main.py only, no VBA) ===

# from pathlib import Path
# from openpyxl import load_workbook
# from openpyxl.utils import get_column_letter
# import pandas as pd
#
# try:
#     from tools.parse_tools.cell_matrix_builder import build_cell_matrix
# except ImportError:
#     build_cell_matrix = None
# try:
#     from tools.parse_tools.cell_bounding_box import get_bounding_cell_range
# except ImportError:
#     get_bounding_cell_range = None
# try:
#     from tools.parse_tools.image_cleaner import run_macro_cleaning
# except ImportError:
#     run_macro_cleaning = None
# from block_generation import generate_block_candidates
#
#
# def _bbox_to_range(bbox, base_row=1, base_col=1):
#     r1 = bbox.r1 + base_row
#     c1 = bbox.c1 + base_col
#     r2 = bbox.r2 + base_row
#     c2 = bbox.c2 + base_col
#     return f"{get_column_letter(c1)}{r1}:{get_column_letter(c2)}{r2}"
#
#
# def parse_excel_file(file_path: str) -> dict:
#     cleaned_path = run_macro_cleaning(file_path) if run_macro_cleaning else file_path
#     wb = load_workbook(cleaned_path, data_only=False)
#     result_by_sheet = {}
#     for sheet_name in wb.sheetnames:
#         if sheet_name == "Shape_Info":
#             continue
#         ws = wb[sheet_name]
#         if build_cell_matrix is not None:
#             df = pd.DataFrame([[cell.value for cell in row] for row in ws.iter_rows()])
#             full_matrix = build_cell_matrix(df)
#         else:
#             full_matrix = _matrix_from_worksheet(ws)
#         if get_bounding_cell_range is not None:
#             box_range = get_bounding_cell_range(full_matrix)
#         else:
#             box_range = _bounding_range_from_matrix(full_matrix)
#         candidates = generate_block_candidates(sheet_id=sheet_name, presence_matrix=full_matrix)
#         blocks = [{ "range": _bbox_to_range(bc.bbox), "cells": bc.n_cells_on, "density": round(bc.density, 3), "provenance": bc.provenance } for bc in candidates]
#         result_by_sheet[sheet_name] = { "box": box_range, "blocks": blocks }
#     wb.close()
#     return result_by_sheet

# def _matrix_from_worksheet(ws): ...
# def _bounding_range_from_matrix(matrix): ...

# --- generate_sheet_summaries (already commented below) ---

# def generate_sheet_summaries(file_path: str, sheet_names: list = None) -> dict:
#     """ ... """
#     ...
