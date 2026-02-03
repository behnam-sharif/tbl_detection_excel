"""
Parse Excel files in the current folder.
Uses openpyxl (same approach as xlgen): load workbooks, iterate sheets and rows.
"""
from pathlib import Path

try:
    from openpyxl import load_workbook
    from openpyxl.worksheet.worksheet import Worksheet
except ImportError:
    load_workbook = None
    Worksheet = None


# Supported extensions (openpyxl handles .xlsx and .xlsm)
EXCEL_EXTENSIONS = (".xlsx", ".xlsm")


def excel_files_in_folder(folder: str | Path | None = None):
    """Yield (path, filename) for each Excel file in folder. Default: current directory."""
    folder = Path(folder or Path.cwd())
    if not folder.is_dir():
        return
    for path in sorted(folder.iterdir()):
        if path.is_file() and path.suffix.lower() in EXCEL_EXTENSIONS:
            yield path, path.name


def open_workbook(path: str | Path, read_only: bool = False, data_only: bool = True):
    """Open an Excel workbook. data_only=True reads computed values; read_only saves memory for large files."""
    if load_workbook is None:
        raise RuntimeError("openpyxl is required. Install with: pip install openpyxl")
    return load_workbook(path, read_only=read_only, data_only=data_only)


def iter_sheets(workbook):
    """Yield (sheet_name, sheet) for each sheet in the workbook."""
    for name in workbook.sheetnames:
        yield name, workbook[name]


def iter_rows(sheet: "Worksheet", min_row: int = 1, max_row: int | None = None):
    """Yield each row as a tuple of cell values (same row-by-row approach as xlgen)."""
    if sheet is None:
        return
    for row in sheet.iter_rows(min_row=min_row, max_row=max_row, values_only=True):
        yield row


def parse_folder(folder: str | Path | None = None, data_only: bool = True):
    """
    Read all Excel files in folder. Yields (filepath, workbook) so caller can
    iterate sheets and rows. Workbooks are opened with data_only=True by default.
    """
    for path, _ in excel_files_in_folder(folder):
        wb = open_workbook(path, read_only=False, data_only=data_only)
        yield path, wb
        wb.close()


def main():
    """Example: list sheets and first few rows of each Excel file in the folder."""
    if load_workbook is None:
        print("Install openpyxl: pip install openpyxl")
        return

    folder = Path(__file__).resolve().parent
    for path, wb in parse_folder(folder):
        print(f"\n--- {path.name} ---")
        for sheet_name, sheet in iter_sheets(wb):
            print(f"  Sheet: {sheet_name}")
            for i, row in enumerate(iter_rows(sheet)):
                if i >= 5:  # first 5 rows per sheet
                    print("    ...")
                    break
                print("   ", row)


if __name__ == "__main__":
    main()
