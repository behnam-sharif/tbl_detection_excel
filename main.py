# === app/main.py ===
import os
import sys
import tempfile
import importlib.util
from pathlib import Path

# Ensure project root (folder containing main.py) is on path for imports
_root = Path(__file__).resolve().parent
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries, get_column_letter

# Optional: full loader (requires tools.* and excel_parse.parser)
try:
    from loader import parse_excel_file
except ImportError:
    parse_excel_file = None

# Block candidates: load from same directory as this file (unique name to avoid cache conflicts)
generate_block_candidates = None
BlockCandidate = None
BBox = None
_block_file = _root / "block_generation.py"
if _block_file.exists():
    try:
        spec = importlib.util.spec_from_file_location("block_generation_impl", _block_file)
        _mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(_mod)
        generate_block_candidates = getattr(_mod, "generate_block_candidates", None)
        BlockCandidate = getattr(_mod, "BlockCandidate", None)
        BBox = getattr(_mod, "BBox", None)
    except Exception:
        pass
if generate_block_candidates is None and str(_root) in sys.path:
    try:
        import block_generation as _bg
        generate_block_candidates = getattr(_bg, "generate_block_candidates", None)
        BlockCandidate = getattr(_bg, "BlockCandidate", None)
        BBox = getattr(_bg, "BBox", None)
    except Exception:
        pass


def _fallback_parse_excel_file(file_path: str) -> dict:
    """
    Minimal parse when loader is not available: load with openpyxl,
    return sheet names and used range as bounding box per sheet.
    """
    wb = load_workbook(file_path, data_only=True)
    result_by_sheet = {}
    for sheet_name in wb.sheetnames:
        if sheet_name == "Shape_Info":
            continue
        ws = wb[sheet_name]
        if ws.max_row and ws.max_column:
            from openpyxl.utils import get_column_letter
            box = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
        else:
            box = "A1:A1"
        result_by_sheet[sheet_name] = {"box": box}
    wb.close()
    return result_by_sheet


def get_parser():
    if parse_excel_file is not None:
        return parse_excel_file
    return _fallback_parse_excel_file


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


# Initialize session state
if "excel_parsed" not in st.session_state:
    st.session_state.excel_parsed = False
if "current_file" not in st.session_state:
    st.session_state.current_file = None
if "parsed_sheets_info" not in st.session_state:
    st.session_state.parsed_sheets_info = None
if "cleaned_file_path" not in st.session_state:
    st.session_state.cleaned_file_path = None
if "block_candidates" not in st.session_state:
    st.session_state.block_candidates = {}  # sheet_name -> list of BlockCandidate

st.title("XGEN: Excel BIM Intelligence")

# Sidebar navigation
st.sidebar.title("Navigation")
page = st.sidebar.radio(
    "Select Section",
    ["Upload Excel", "View Sheets", "Block Candidates"],
)

# Upload Excel Section
if page == "Upload Excel":
    st.header("Upload Excel File")

    # Show existing uploaded file info if available
    if st.session_state.excel_parsed and st.session_state.parsed_sheets_info:
        current = st.session_state.current_file
        name = Path(current).name if current else "N/A"
        st.info(f"Current file: `{name}`")
        st.markdown("### Parsed Sheets")
        for sheet_name, sheet_info in st.session_state.parsed_sheets_info.items():
            st.markdown(f"**{sheet_name}**: Bounding box: `{sheet_info['box']}`")
        st.markdown("---")

    uploaded_file = st.file_uploader(
        "Upload your BIM Excel file",
        type=["xlsx", "xlsm", "xlsb"],
        key="excel_uploader",
    )

    if uploaded_file:
        ext = Path(uploaded_file.name).suffix.lower()
        os.makedirs("data/uploads", exist_ok=True)

        with tempfile.NamedTemporaryFile(
            delete=False, dir="data/uploads", suffix=ext
        ) as tmp:
            tmp.write(uploaded_file.getbuffer())
            file_path = tmp.name
            st.session_state.current_file = file_path

        st.success(
            f"Excel file uploaded as `{Path(file_path).name}` and saved to temp folder."
        )

        try:
            parse_fn = get_parser()
            parsed = parse_fn(file_path)
            st.session_state.excel_parsed = True
            st.session_state.parsed_sheets_info = parsed
            # Store cleaned file path if full loader created data/cleaned/cleaned.xlsx
            cleaned_path = Path("data/cleaned/cleaned.xlsx")
            if cleaned_path.exists():
                st.session_state.cleaned_file_path = str(cleaned_path)
            else:
                st.session_state.cleaned_file_path = file_path

            st.markdown("### Parsed Sheets")
            for sheet_name, sheet_info in parsed.items():
                st.markdown(f"**{sheet_name}**: Bounding box: `{sheet_info['box']}`")
        except Exception as e:
            st.warning(f"Could not parse Excel file: {e}")
            st.session_state.excel_parsed = False
            st.session_state.parsed_sheets_info = None

# View Sheets Section
elif page == "View Sheets":
    st.header("View Sheets")

    if not st.session_state.excel_parsed:
        st.warning("Please upload and parse an Excel file first.")
    else:
        try:
            file_path = (
                st.session_state.cleaned_file_path
                or st.session_state.current_file
            )
            wb = load_workbook(file_path, data_only=True)
            sheet_names = [
                n for n in wb.sheetnames if n != "Shape_Info"
            ]
            wb.close()
        except Exception as e:
            st.error(f"Error loading Excel file: {e}")
            sheet_names = []

        if not sheet_names:
            st.warning("No sheets found in the Excel file.")
        else:
            selected_sheet = st.selectbox(
                "Select a sheet:",
                sheet_names,
                key="view_sheet_select",
            )
            if selected_sheet:
                wb = load_workbook(
                    st.session_state.cleaned_file_path
                    or st.session_state.current_file,
                    data_only=True,
                )
                ws = wb[selected_sheet]
                rows = list(ws.iter_rows(values_only=True))
                wb.close()
                if rows:
                    st.dataframe(rows[:100], use_container_width=True)
                    if len(rows) > 100:
                        st.caption(f"Showing first 100 of {len(rows)} rows.")
                else:
                    st.info("Sheet is empty.")

# Block Candidates Section
elif page == "Block Candidates":
    st.header("Block Candidates")

    if not st.session_state.excel_parsed:
        st.warning("Please upload and parse an Excel file first.")
    elif generate_block_candidates is None or BBox is None:
        st.warning(
            "Block generation module not available. Ensure `block_generation.py` "
            "exists in the same folder as `main.py`."
        )
    else:
        try:
            file_path = (
                st.session_state.cleaned_file_path
                or st.session_state.current_file
            )
            wb = load_workbook(file_path, data_only=True)
            sheet_names = [
                n for n in wb.sheetnames if n != "Shape_Info"
            ]
            wb.close()
        except Exception as e:
            st.error(f"Error loading Excel file: {e}")
            sheet_names = []

        if not sheet_names:
            st.warning("No sheets found in the Excel file.")
        else:
            col1, col2 = st.columns(2)
            with col1:
                do_bridging = st.checkbox("Bridge small gaps", value=True, key="block_do_bridging")
                min_on_cells = st.number_input("Min cells per block", min_value=1, value=4, key="block_min_cells")
            with col2:
                connectivity = st.radio("Connectivity", [4, 8], horizontal=True, key="block_connectivity")
                bridging_rows = st.number_input("Bridge gap (rows)", min_value=0, value=1, key="block_bridge_r")
                bridging_cols = st.number_input("Bridge gap (cols)", min_value=0, value=1, key="block_bridge_c")

            if st.button("Generate block candidates for all sheets", key="block_generate_btn"):
                st.session_state.block_candidates = {}
                progress = st.progress(0.0, text="Processing sheets...")
                n_sheets = len(sheet_names)
                for idx, sheet_name in enumerate(sheet_names):
                    progress.progress((idx + 0.5) / n_sheets, text=f"Processing '{sheet_name}'...")
                    try:
                        box_str = st.session_state.parsed_sheets_info.get(sheet_name, {}).get("box", "A1:ZZ1000")
                        min_col, min_row, max_col, max_row = range_boundaries(box_str)
                        P = get_presence_matrix_for_sheet(file_path, sheet_name, box_str)
                        if not P:
                            st.session_state.block_candidates[sheet_name] = {
                                "candidates": [],
                                "base_row": min_row,
                                "base_col": min_col,
                            }
                        else:
                            candidates = generate_block_candidates(
                                sheet_id=sheet_name,
                                presence_matrix=P,
                                do_bridging=do_bridging,
                                bridging_rows=bridging_rows,
                                bridging_cols=bridging_cols,
                                connectivity=connectivity,
                                min_on_cells=min_on_cells,
                            )
                            st.session_state.block_candidates[sheet_name] = {
                                "candidates": candidates,
                                "base_row": min_row,
                                "base_col": min_col,
                            }
                    except Exception as e:
                        st.session_state.block_candidates[sheet_name] = {
                            "candidates": [],
                            "base_row": 1,
                            "base_col": 1,
                            "error": str(e),
                        }
                progress.progress(1.0, text="Done.")
                st.success(f"Block candidates generated for {n_sheets} sheet(s).")
                st.rerun()

            st.markdown("---")
            st.markdown("### Block candidates by sheet")

            if not st.session_state.block_candidates:
                st.info("Click **Generate block candidates for all sheets** to detect blocks on every sheet.")
            else:
                for sheet_name in sheet_names:
                    if sheet_name not in st.session_state.block_candidates:
                        continue
                    info = st.session_state.block_candidates[sheet_name]
                    candidates = info.get("candidates", [])
                    err = info.get("error")
                    base_row = info.get("base_row", 1)
                    base_col = info.get("base_col", 1)
                    n_blocks = len(candidates)
                    with st.expander(f"**{sheet_name}** — {n_blocks} block(s)", expanded=(n_blocks > 0)):
                        if err:
                            st.error(err)
                        if not candidates:
                            st.caption("No blocks detected.")
                        else:
                            for i, bc in enumerate(candidates, 1):
                                excel_range = bbox_to_excel_range(bc.bbox, base_row=base_row, base_col=base_col)
                                st.markdown(f"**Block {i}:** `{excel_range}` — {bc.n_cells_on} cells, density {bc.density:.2f} ({bc.provenance})")
