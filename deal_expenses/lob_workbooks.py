"""Create refreshed LOB-specific copies of a combined deal expense workbook."""

from __future__ import annotations

from collections.abc import Callable
from io import BytesIO
from pathlib import Path
import shutil
import sys
from zipfile import ZIP_STORED, ZipFile


LOB_NAMES = ("ECM", "DCM", "M&A", "SPG", "Lev Finance", "Struc Finance")
CONCUR_REPORT_SHEET = "Concur Report"
SAP_REPORT_SHEETS = ("SAP Report", "SAP Invoices Report")
XL_UP = -4162
XL_TO_LEFT = -4159


def build_output_archive(output_paths: list[Path], folder_name: str) -> bytes:
    """Return the seven generated workbooks in one folder within a ZIP archive."""
    archive = BytesIO()
    with ZipFile(archive, "w", compression=ZIP_STORED) as zip_file:
        for output_path in output_paths:
            zip_file.write(output_path, arcname=f"{folder_name}/{output_path.name}")
    return archive.getvalue()


def refresh_workbook_pivots(workbook) -> None:
    """Refresh PivotTables without reloading workbook queries or connections."""
    pivot_tables = []
    refreshed_cache_indexes: set[int] = set()
    for sheet_index in range(1, workbook.Worksheets.Count + 1):
        worksheet_pivots = workbook.Worksheets(sheet_index).PivotTables()
        for pivot_index in range(1, worksheet_pivots.Count + 1):
            pivot_table = worksheet_pivots.Item(pivot_index)
            pivot_tables.append(pivot_table)
            cache_index = int(pivot_table.CacheIndex)
            if cache_index not in refreshed_cache_indexes:
                pivot_table.PivotCache().Refresh()
                refreshed_cache_indexes.add(cache_index)
    for pivot_table in pivot_tables:
        pivot_table.RefreshTable()


def _normalized(value: object) -> str:
    return " ".join(str(value or "").strip().split()).casefold()


def _rows(values: object) -> list[tuple[object, ...]]:
    if values in (None, ""):
        return []
    if not isinstance(values, tuple):
        return [(values,)]
    if values and not isinstance(values[0], tuple):
        return [tuple(values)]
    return [tuple(row) for row in values]


def _header_column(worksheet, header: str) -> int:
    last_column = worksheet.Cells(1, worksheet.Columns.Count).End(XL_TO_LEFT).Column
    headers = _rows(worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(1, last_column)).Value2)[0]
    matches = [column for column, value in enumerate(headers, start=1) if _normalized(value) == _normalized(header)]
    if len(matches) != 1:
        description = "missing" if not matches else "duplicated"
        raise ValueError(f"The LOB column is {description} in worksheet {worksheet.Name!r}.")
    return matches[0]


def _contiguous_ranges(rows: list[int]) -> list[tuple[int, int]]:
    if not rows:
        return []
    ranges: list[tuple[int, int]] = []
    start = previous = rows[0]
    for row in rows[1:]:
        if row == previous + 1:
            previous = row
            continue
        ranges.append((start, previous))
        start = previous = row
    ranges.append((start, previous))
    return ranges


def _delete_non_lob_rows(worksheet, lob_name: str) -> int:
    lob_column = _header_column(worksheet, "LOB")
    last_row = worksheet.Cells(worksheet.Rows.Count, lob_column).End(XL_UP).Row
    if last_row < 2:
        return 0
    values = worksheet.Range(worksheet.Cells(2, lob_column), worksheet.Cells(last_row, lob_column)).Value2
    if last_row == 2:
        lob_values = [values]
    elif isinstance(values, tuple):
        lob_values = [value[0] if isinstance(value, tuple) else value for value in values]
    else:
        lob_values = [values]
    rows_to_delete = [row for row, value in enumerate(lob_values, start=2) if _normalized(value) != _normalized(lob_name)]
    for first_row, last_row in reversed(_contiguous_ranges(rows_to_delete)):
        worksheet.Rows(f"{first_row}:{last_row}").Delete()
    return len(rows_to_delete)


def _sap_report_worksheet(workbook):
    for sheet_name in SAP_REPORT_SHEETS:
        try:
            return workbook.Worksheets(sheet_name)
        except Exception:
            continue
    raise ValueError(f"Workbook is missing a SAP report worksheet ({', '.join(SAP_REPORT_SHEETS)}).")


def create_lob_workbooks(
    master_path: Path,
    output_paths: dict[str, Path],
    on_progress: Callable[[str, int, int], None] | None = None,
) -> dict[str, dict[str, int]]:
    """Create one filtered and refreshed workbook per configured LOB."""
    if sys.platform != "win32":
        raise RuntimeError("Excel workbook processing requires Windows and Microsoft Excel.")
    if set(output_paths) != set(LOB_NAMES):
        raise ValueError("LOB output paths must be supplied for every configured LOB.")
    try:
        import pythoncom
        import win32com.client as win32
    except ImportError as error:
        raise RuntimeError("pywin32 is required. Install it with 'pip install pywin32'.") from error

    excel = workbook = None
    metrics: dict[str, dict[str, int]] = {}
    pythoncom.CoInitialize()
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        for index, lob_name in enumerate(LOB_NAMES, start=1):
            if on_progress is not None:
                on_progress(lob_name, index, len(LOB_NAMES))
            output_path = output_paths[lob_name]
            output_path.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(master_path, output_path)
            workbook = excel.Workbooks.Open(str(output_path.resolve()), ReadOnly=False)
            if workbook.ReadOnly:
                raise RuntimeError(f"Excel opened LOB output workbook as read-only: {output_path.name}")
            concur_deleted = _delete_non_lob_rows(workbook.Worksheets(CONCUR_REPORT_SHEET), lob_name)
            sap_deleted = _delete_non_lob_rows(_sap_report_worksheet(workbook), lob_name)
            refresh_workbook_pivots(workbook)
            workbook.Save()
            workbook.Close(SaveChanges=False)
            workbook = None
            metrics[lob_name] = {"concur_deleted": concur_deleted, "sap_deleted": sap_deleted}
        return metrics
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()