"""Read rendered Excel pivot tables for display in the Streamlit application."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import sys


@dataclass(frozen=True)
class PivotPreview:
    """One rendered Excel pivot table, including its displayed column headings."""

    name: str
    headers: list[str]
    rows: list[list[object]]
    emphasis_rows: list[bool]


def _range_values(range_) -> list[list[object]]:
    """Return every displayed cell in a COM range as a two-dimensional list."""
    row_count = range_.Rows.Count
    column_count = range_.Columns.Count
    raw_values = range_.Value2
    if row_count == 1 and column_count == 1:
        return [[raw_values]]
    if row_count == 1:
        return [list(raw_values)]
    if column_count == 1:
        return [[value] for value in raw_values]
    return [list(row) for row in raw_values]


def _headers(values: list[object]) -> list[str]:
    """Create non-empty, unique labels suitable for a dataframe."""
    used: dict[str, int] = {}
    headers = []
    for index, value in enumerate(values, start=1):
        base = str(value).strip() if value not in (None, "") else f"Column {index}"
        count = used.get(base, 0) + 1
        used[base] = count
        headers.append(base if count == 1 else f"{base} ({count})")
    return headers


def _is_blank(row: list[object]) -> bool:
    return all(value in (None, "") for value in row)


def _table_rows(values: list[list[object]]) -> list[list[object]]:
    """Exclude report filters and retain the header row plus displayed pivot data."""
    for index, row in enumerate(values):
        if _is_blank(row):
            return values[index + 1 :]
    return values


def _is_emphasis_row(row: list[object]) -> bool:
    label = next((str(value).strip() for value in row if value not in (None, "")), "")
    return label.casefold() in {"sancap", "bsny", "grand total"}


def _preview_pivot(pivot_table) -> PivotPreview | None:
    values = _table_rows(_range_values(pivot_table.TableRange2))
    if not values:
        return None
    rows = values[1:]
    return PivotPreview(
        name=str(pivot_table.Name),
        headers=_headers(values[0]),
        rows=rows,
        emphasis_rows=[_is_emphasis_row(row) for row in rows],
    )


def extract_pivot_previews(workbook_path: Path, sheet_names: list[str], limit: int = 1) -> dict[str, list[PivotPreview]]:
    """Read up to ``limit`` rendered pivot tables from each requested worksheet.

    The workbook is opened read-only because the writers have already refreshed and
    saved the pivot caches before this display-only step runs.
    """
    if sys.platform != "win32":
        raise RuntimeError("Pivot previews require Windows and Microsoft Excel.")
    try:
        import pythoncom
        import win32com.client as win32
    except ImportError as error:
        raise RuntimeError("Pivot previews require pywin32.") from error

    excel = workbook = None
    pythoncom.CoInitialize()
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Open(str(workbook_path.resolve()), ReadOnly=True, UpdateLinks=0)
        previews: dict[str, list[PivotPreview]] = {}
        for sheet_name in sheet_names:
            try:
                worksheet = workbook.Worksheets(sheet_name)
            except Exception:
                previews[sheet_name] = []
                continue
            pivot_tables = [worksheet.PivotTables(index) for index in range(1, worksheet.PivotTables().Count + 1)]
            pivot_tables.sort(key=lambda pivot_table: (pivot_table.TableRange2.Row, pivot_table.TableRange2.Column))
            previews[sheet_name] = [
                preview
                for pivot_table in pivot_tables[:limit]
                if (preview := _preview_pivot(pivot_table)) is not None
            ]
        return previews
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()