"""Excel COM writer for replacing Uber Report data in an existing master workbook."""

from __future__ import annotations

from collections import defaultdict
from pathlib import Path
import sys
from typing import Any

from deal_expenses.hr_refresh import HrRefreshPlan, build_hr_refresh_plan, normalize_network_id
from deal_expenses.models import UberRunRequest
from deal_expenses.sources.uber import REQUIRED_HEADERS, normalize_uber_header
from deal_expenses.uber_validation import UBER_HEADER_ROW, UBER_REPORT_SHEET, cost_center_columns, in_scope_column, master_headers, required_column


XL_UP = -4162
XL_TO_LEFT = -4159
XL_PASTE_FORMATS = -4122
XL_CALCULATION_AUTOMATIC = -4105


def _rows(values: Any) -> list[tuple[Any, ...]]:
    if values in (None, ""):
        return []
    if not isinstance(values, tuple):
        return [(values,)]
    if values and not isinstance(values[0], tuple):
        return [tuple(values)]
    return [tuple(row) for row in values]


def _last_used_row(worksheet, column: int, header_row: int) -> int:
    return max(header_row, worksheet.Cells(worksheet.Rows.Count, column).End(XL_UP).Row)


def _com_headers(worksheet, header_row: int) -> dict[str, list[int]]:
    last_column = worksheet.Cells(header_row, worksheet.Columns.Count).End(XL_TO_LEFT).Column
    values = _rows(worksheet.Range(worksheet.Cells(header_row, 1), worksheet.Cells(header_row, last_column)).Value2)[0]
    headers: dict[str, list[int]] = defaultdict(list)
    for column, value in enumerate(values, start=1):
        header = normalize_uber_header(value)
        if header:
            headers[header].append(column)
    return headers


def _first_formula(worksheet, column: int, first_row: int, last_row: int) -> str | None:
    for row in range(first_row, last_row + 1):
        formula = worksheet.Cells(row, column).FormulaR1C1
        if isinstance(formula, str) and formula.startswith("="):
            return formula
    return None


def _normalized_value(value: Any) -> str:
    return normalize_network_id(value)


class UberExcelComWorkbookWriter:
    """Replace Uber raw data and regenerate workbook helper columns."""

    def __init__(self) -> None:
        if sys.platform != "win32":
            raise RuntimeError("Excel workbook processing requires Windows and Microsoft Excel.")

    @staticmethod
    def _source_rows(worksheet, headers: dict[str, list[int]]) -> list[tuple[Any, ...]]:
        last_row = max(_last_used_row(worksheet, columns[0], 1) for columns in headers.values())
        return _rows(worksheet.Range(worksheet.Cells(2, 1), worksheet.Cells(last_row, max(max(columns) for columns in headers.values()))).Value2)

    @staticmethod
    def _select_source_worksheet(workbook):
        required_headers = {normalize_uber_header(header) for header in REQUIRED_HEADERS}
        matches = []
        for index in range(1, workbook.Worksheets.Count + 1):
            worksheet = workbook.Worksheets(index)
            headers = _com_headers(worksheet, 1)
            if any(len(columns) != 1 for columns in headers.values()):
                continue
            if required_headers.issubset(headers):
                matches.append(worksheet)
        if len(matches) != 1:
            names = ", ".join(worksheet.Name for worksheet in matches) or "none"
            raise ValueError(f"Expected exactly one Uber source worksheet with required headers; found {names}.")
        return matches[0]

    @staticmethod
    def _hr_lookups(workbook) -> tuple[dict[str, str], dict[str, str]]:
        matches = [workbook.Worksheets(index) for index in range(1, workbook.Worksheets.Count + 1) if workbook.Worksheets(index).Name.startswith("HR HC Combined")]
        if len(matches) != 1:
            raise ValueError("Expected exactly one HR HC Combined worksheet.")
        worksheet = matches[0]
        headers = _com_headers(worksheet, 1)
        employee_id_column = required_column(headers, "Network ID", worksheet.Name)
        name_column = required_column(headers, "Name", worksheet.Name)
        cost_center_column = required_column(headers, "Cost Center Name", worksheet.Name)
        last_row = _last_used_row(worksheet, employee_id_column, 1)
        values = _rows(worksheet.Range(worksheet.Cells(2, 1), worksheet.Cells(last_row, max(max(columns) for columns in headers.values()))).Value2)
        by_id: dict[str, str] = {}
        name_cost_centers: dict[str, set[str]] = defaultdict(set)
        for row in values:
            employee_id = _normalized_value(row[employee_id_column - 1])
            name = normalize_uber_header(row[name_column - 1])
            cost_center = str(row[cost_center_column - 1] or "").strip()
            if employee_id and cost_center and employee_id not in by_id:
                by_id[employee_id] = cost_center
            if name and cost_center:
                name_cost_centers[name].add(cost_center)
        by_name = {name: next(iter(cost_centers)) for name, cost_centers in name_cost_centers.items() if len(cost_centers) == 1}
        return by_id, by_name

    @staticmethod
    def _refresh_hr_worksheet(workbook, plan: HrRefreshPlan) -> None:
        worksheet = workbook.Worksheets(plan.master_sheet_name)
        header_row = plan.master_header_row
        key_column = plan.master_headers["network id"]
        final_column = max(plan.master_headers.values())
        old_last_row = _last_used_row(worksheet, key_column, header_row)
        final_row = header_row + len(plan.retained_rows)
        duplicate_template = _first_formula(worksheet, plan.duplicate_column, header_row + 1, old_last_row) if plan.duplicate_column else None
        if old_last_row > header_row:
            worksheet.Range(worksheet.Cells(old_last_row, 1), worksheet.Cells(old_last_row, final_column)).Copy()
            worksheet.Range(worksheet.Cells(header_row + 1, 1), worksheet.Cells(final_row, final_column)).PasteSpecial(Paste=XL_PASTE_FORMATS)
            worksheet.Application.CutCopyMode = False
            worksheet.Range(worksheet.Cells(header_row + 1, 1), worksheet.Cells(max(old_last_row, final_row), final_column)).ClearContents()
        for header, column in plan.master_headers.items():
            if column == plan.duplicate_column:
                continue
            values = tuple((row[column - 1],) for row in plan.retained_rows)
            worksheet.Range(worksheet.Cells(header_row + 1, column), worksheet.Cells(final_row, column)).Value2 = values
        if plan.duplicate_column and duplicate_template:
            worksheet.Range(worksheet.Cells(header_row + 1, plan.duplicate_column), worksheet.Cells(final_row, plan.duplicate_column)).FormulaR1C1 = duplicate_template
        for index in range(1, worksheet.ListObjects.Count + 1):
            table = worksheet.ListObjects(index)
            if table.Range.Row == header_row and table.Range.Column <= key_column < table.Range.Column + table.Range.Columns.Count:
                table.Resize(worksheet.Range(worksheet.Cells(header_row, table.Range.Column), worksheet.Cells(final_row, table.Range.Column + table.Range.Columns.Count - 1)))
                break

    def write(self, request: UberRunRequest) -> dict[str, int | float]:
        try:
            import pythoncom
            import win32com.client as win32
        except ImportError as error:
            raise RuntimeError("pywin32 is required. Install it with 'pip install pywin32'.") from error

        plan = build_hr_refresh_plan(request.master_path, request.hr_source_path)
        excel = master_workbook = source_workbook = None
        succeeded = False
        try:
            pythoncom.CoInitialize()
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            excel.EnableEvents = False
            master_workbook = excel.Workbooks.Open(str(request.master_path.resolve()))
            source_workbook = excel.Workbooks.Open(str(request.source_path.resolve()), ReadOnly=True)
            self._refresh_hr_worksheet(master_workbook, plan)
            target = master_workbook.Worksheets(UBER_REPORT_SHEET)
            source = self._select_source_worksheet(source_workbook)
            target_headers = _com_headers(target, UBER_HEADER_ROW)
            source_headers = _com_headers(source, 1)
            source_rows = self._source_rows(source, source_headers)
            if not source_rows:
                raise ValueError("Uber source has no data rows.")
            employee_id_target = required_column(target_headers, "Employee ID", target.Name)
            transaction_target = required_column(target_headers, "Transaction Amount USD", target.Name)
            full_name_column = required_column(target_headers, "Full Name", target.Name)
            final_cost_center_column = required_column(target_headers, "Final Cost Center", target.Name)
            formula_columns = [
                full_name_column,
                final_cost_center_column,
                required_column(target_headers, "LOB", target.Name),
                in_scope_column(target_headers, target.Name),
                required_column(target_headers, "Company Name", target.Name),
            ]
            cost_center_id_column, cost_center_name_column = cost_center_columns(target, target_headers)
            first_data_row = UBER_HEADER_ROW + 1
            old_last_row = _last_used_row(target, transaction_target, UBER_HEADER_ROW)
            final_row = UBER_HEADER_ROW + len(source_rows)
            final_column = max(max(columns) for columns in target_headers.values())
            formulas = {column: _first_formula(target, column, first_data_row, old_last_row) for column in formula_columns}
            missing = [str(target.Cells(UBER_HEADER_ROW, column).Value2) for column, formula in formulas.items() if formula is None]
            if missing:
                raise ValueError("Missing helper formula template for: " + ", ".join(missing))
            if final_row > old_last_row:
                target.Range(target.Cells(first_data_row, 1), target.Cells(first_data_row, final_column)).Copy()
                target.Range(target.Cells(old_last_row + 1, 1), target.Cells(final_row, final_column)).PasteSpecial(Paste=XL_PASTE_FORMATS)
                target.Application.CutCopyMode = False
            extent = max(old_last_row, final_row)
            for header, columns in source_headers.items():
                target_column = required_column(target_headers, header, target.Name)
                target.Range(target.Cells(first_data_row, target_column), target.Cells(extent, target_column)).ClearContents()
                target.Range(target.Cells(first_data_row, target_column), target.Cells(final_row, target_column)).Value2 = tuple((row[columns[0] - 1],) for row in source_rows)
            by_id, by_name = self._hr_lookups(master_workbook)
            source_id_column = required_column(source_headers, "Employee ID", source.Name)
            source_first_name_column = required_column(source_headers, "First Name", source.Name)
            source_last_name_column = required_column(source_headers, "Last Name", source.Name)
            id_cost_centers = []
            name_cost_centers = []
            for row in source_rows:
                id_cost_center = by_id.get(_normalized_value(row[source_id_column - 1]), "")
                full_name = normalize_uber_header(f"{row[source_first_name_column - 1] or ''} {row[source_last_name_column - 1] or ''}")
                id_cost_centers.append(id_cost_center)
                name_cost_centers.append("" if id_cost_center else by_name.get(full_name, ""))
            for column, values in ((cost_center_id_column, id_cost_centers), (cost_center_name_column, name_cost_centers)):
                target.Range(target.Cells(first_data_row, column), target.Cells(extent, column)).ClearContents()
                target.Range(target.Cells(first_data_row, column), target.Cells(final_row, column)).Value2 = tuple((value,) for value in values)
            for column, formula in formulas.items():
                target.Range(target.Cells(first_data_row, column), target.Cells(extent, column)).ClearContents()
                target.Range(target.Cells(first_data_row, column), target.Cells(final_row, column)).FormulaR1C1 = formula
            excel.Calculation = XL_CALCULATION_AUTOMATIC
            excel.CalculateFullRebuild()
            final_cost_centers = _rows(target.Range(target.Cells(first_data_row, final_cost_center_column), target.Cells(final_row, final_cost_center_column)).Value2)
            unassigned = sum(value[0] in (None, "") for value in final_cost_centers)
            master_workbook.Save()
            succeeded = True
            return {"imported_rows": len(source_rows), "unassigned_final_cost_center_rows": unassigned, "unassigned_final_cost_center_percent": unassigned * 100 / len(source_rows)}
        finally:
            if source_workbook is not None:
                source_workbook.Close(SaveChanges=False)
            if master_workbook is not None:
                master_workbook.Close(SaveChanges=succeeded)
            if excel is not None:
                excel.Quit()
            if 'pythoncom' in locals():
                pythoncom.CoUninitialize()
