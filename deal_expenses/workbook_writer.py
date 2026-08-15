"""Excel COM implementation for refreshing the master workbook."""

from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path
import shutil
import sys

from deal_expenses.models import RunRequest
from deal_expenses.sources.base import EXPENSE_HEADER, SOURCE_TO_MASTER, YEAR_HEADER, is_reporting_year, normalize_header
from deal_expenses.validation import MasterWorkbookValidator


XL_UP = -4162
XL_TO_LEFT = -4159
XL_CALCULATION_MANUAL = -4135
XL_CALCULATION_AUTOMATIC = -4105
XL_PASTE_VALUES = -4163
XL_SHEET_VERY_HIDDEN = 2


class WorkbookWriter(ABC):
    """Writes one validated run into an output workbook."""

    @abstractmethod
    def write(self, request: RunRequest) -> dict[str, int | float]:
        """Create the output workbook and return final metrics."""


class ExcelComWorkbookWriter(WorkbookWriter):
    """Windows Microsoft Excel writer preserving formulas, tables, and formatting."""

    master_sheet_name = "Concur Report"
    bsny_sheet_name = "Concur"
    helper_headers = MasterWorkbookValidator.helper_headers

    def __init__(self) -> None:
        if sys.platform != "win32":
            raise RuntimeError("Excel workbook processing requires Windows and Microsoft Excel.")

    @staticmethod
    def _read_headers(worksheet) -> dict[str, list[int]]:
        last_column = worksheet.Cells(1, worksheet.Columns.Count).End(XL_TO_LEFT).Column
        raw_headers = worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(1, last_column)).Value2
        header_values = list(raw_headers[0]) if isinstance(raw_headers, tuple) else [raw_headers]
        headers: dict[str, list[int]] = {}
        for column_number, raw_header in enumerate(header_values, start=1):
            header = normalize_header(raw_header)
            if header:
                headers.setdefault(header, []).append(column_number)
        return headers

    @staticmethod
    def _unique_header_column(headers: dict[str, list[int]], header: str, worksheet) -> int:
        columns = headers.get(header, [])
        if len(columns) != 1:
            description = "missing" if not columns else f"repeated in columns {columns}"
            raise ValueError(f"Required header '{header}' is {description} in '{worksheet.Name}'.")
        return columns[0]

    @staticmethod
    def _last_used_row(worksheet) -> int:
        return worksheet.Cells(worksheet.Rows.Count, 1).End(XL_UP).Row

    @staticmethod
    def _column_values(worksheet, column_number: int, last_row: int) -> list[object]:
        if last_row < 2:
            return []
        raw_values = worksheet.Range(worksheet.Cells(2, column_number), worksheet.Cells(last_row, column_number)).Value2
        if last_row == 2:
            return [raw_values]
        return [row[0] for row in raw_values]

    def _source_rows_for_year(self, worksheet, headers: dict[str, list[int]], year: int) -> list[int]:
        last_row = self._last_used_row(worksheet)
        year_column = self._unique_header_column(headers, YEAR_HEADER, worksheet)
        years = self._column_values(worksheet, year_column, last_row)
        return [row_number for row_number, value in enumerate(years, start=2) if is_reporting_year(value, year)]

    def _values_for_rows(self, worksheet, column_number: int, rows: list[int]) -> list[object]:
        if not rows:
            return []
        values = self._column_values(worksheet, column_number, max(rows))
        return [values[row_number - 2] for row_number in rows]

    @staticmethod
    def _numeric_total(values: list[object]) -> float:
        return sum(float(value) for value in values if value not in (None, ""))

    @staticmethod
    def _paste_column_from_staging(staging_sheet, worksheet, start_row: int, column_number: int, values: list[object]) -> None:
        if not values:
            return
        staging_range = staging_sheet.Range(staging_sheet.Cells(1, 1), staging_sheet.Cells(len(values), 1))
        staging_range.Value2 = tuple((value,) for value in values)
        staging_range.Copy()
        worksheet.Range(
            worksheet.Cells(start_row, column_number),
            worksheet.Cells(start_row + len(values) - 1, column_number),
        ).PasteSpecial(Paste=XL_PASTE_VALUES)

    def _build_column_mappings(self, master_headers, bsny_headers, sancap_headers) -> list[tuple[int, int, int]]:
        master_to_source = {target: source for source, target in SOURCE_TO_MASTER.items()}
        source_occurrences: dict[str, int] = {}
        mappings = []
        for master_header, master_columns in master_headers.items():
            if master_header in self.helper_headers:
                continue
            source_header = master_to_source.get(master_header, master_header)
            for master_column in master_columns:
                occurrence = source_occurrences.get(source_header, 0)
                source_occurrences[source_header] = occurrence + 1
                for source_name, source_headers in (("BSNY", bsny_headers), ("SanCap", sancap_headers)):
                    available_columns = source_headers.get(source_header, [])
                    if len(available_columns) <= occurrence:
                        raise ValueError(
                            f"{source_name} is missing occurrence {occurrence + 1} of source header "
                            f"'{source_header}' required for master column {master_column}."
                        )
                mappings.append((master_column, bsny_headers[source_header][occurrence], sancap_headers[source_header][occurrence]))
        return mappings

    @staticmethod
    def _resize_containing_table(worksheet, column_number: int, final_row: int) -> None:
        for table_index in range(1, worksheet.ListObjects.Count + 1):
            table = worksheet.ListObjects(table_index)
            first_column = table.Range.Column
            last_column = first_column + table.Range.Columns.Count - 1
            if first_column <= column_number <= last_column:
                table.Resize(worksheet.Range(worksheet.Cells(table.Range.Row, first_column), worksheet.Cells(final_row, last_column)))
                return

    def write(self, request: RunRequest) -> dict[str, int | float]:
        """Refresh the master using BSNY rows first and SanCap rows second."""
        try:
            import pythoncom
            import win32com.client as win32
        except ImportError as error:
            raise RuntimeError("pywin32 is required. Install it with 'pip install pywin32'.") from error

        bsny_path = request.source_paths["bsny"]
        sancap_path = request.source_paths["sancap"]
        request.output_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(request.master_path, request.output_path)

        excel = master_workbook = bsny_workbook = sancap_workbook = staging_sheet = None
        refresh_succeeded = False
        original_autofill = None
        com_initialized = False
        try:
            pythoncom.CoInitialize()
            com_initialized = True
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            excel.EnableEvents = False
            master_workbook = excel.Workbooks.Open(str(request.output_path.resolve()))
            bsny_workbook = excel.Workbooks.Open(str(bsny_path.resolve()), ReadOnly=True)
            sancap_workbook = excel.Workbooks.Open(str(sancap_path.resolve()), ReadOnly=True)
            excel.Calculation = XL_CALCULATION_MANUAL
            original_autofill = excel.AutoCorrect.AutoFillFormulasInLists
            excel.AutoCorrect.AutoFillFormulasInLists = False

            master_sheet = master_workbook.Worksheets(self.master_sheet_name)
            bsny_sheet = bsny_workbook.Worksheets(self.bsny_sheet_name)
            sancap_sheet = sancap_workbook.Worksheets(1)
            staging_sheet = master_workbook.Worksheets.Add(After=master_workbook.Worksheets(master_workbook.Worksheets.Count))
            staging_sheet.Visible = XL_SHEET_VERY_HIDDEN

            master_headers = self._read_headers(master_sheet)
            bsny_headers = self._read_headers(bsny_sheet)
            sancap_headers = self._read_headers(sancap_sheet)
            helper_columns = {header: self._unique_header_column(master_headers, header, master_sheet) for header in self.helper_headers}
            for target_header in SOURCE_TO_MASTER.values():
                self._unique_header_column(master_headers, target_header, master_sheet)
            master_year_column = self._unique_header_column(master_headers, YEAR_HEADER, master_sheet)
            master_expense_column = self._unique_header_column(master_headers, EXPENSE_HEADER, master_sheet)
            column_mappings = self._build_column_mappings(master_headers, bsny_headers, sancap_headers)

            bsny_rows = self._source_rows_for_year(bsny_sheet, bsny_headers, request.reporting_year)
            sancap_rows = self._source_rows_for_year(sancap_sheet, sancap_headers, request.reporting_year)
            if not bsny_rows or not sancap_rows:
                raise ValueError(f"No {request.reporting_year} data found in one or both source workbooks. No data was cleared.")

            first_data_row = 2
            final_data_row = len(bsny_rows) + len(sancap_rows) + 1
            current_last_row = self._last_used_row(master_sheet)
            helper_formulas = {header: master_sheet.Cells(first_data_row, column).FormulaR1C1 for header, column in helper_columns.items()}
            missing_formulas = [header for header, formula in helper_formulas.items() if not str(formula).startswith("=")]
            if missing_formulas:
                raise ValueError("The formula template is missing in row 2 for: " + ", ".join(missing_formulas))

            self._resize_containing_table(master_sheet, master_expense_column, final_data_row)
            for master_column, _, _ in column_mappings:
                master_sheet.Range(master_sheet.Cells(first_data_row, master_column), master_sheet.Cells(current_last_row, master_column)).ClearContents()

            expected_expense_values = None
            for master_column, bsny_column, sancap_column in column_mappings:
                combined_values = self._values_for_rows(bsny_sheet, bsny_column, bsny_rows) + self._values_for_rows(sancap_sheet, sancap_column, sancap_rows)
                self._paste_column_from_staging(staging_sheet, master_sheet, first_data_row, master_column, combined_values)
                if master_column == master_expense_column:
                    expected_expense_values = combined_values
                if expected_expense_values is not None:
                    actual_values = self._column_values(master_sheet, master_expense_column, final_data_row)
                    if round(self._numeric_total(actual_values), 2) != round(self._numeric_total(expected_expense_values), 2):
                        raise AssertionError(f"Writing master column {master_column} changed expense values.")

            if expected_expense_values is None:
                raise AssertionError("Expense column was not included in the master/source mapping.")
            for helper_header, formula in helper_formulas.items():
                helper_column = helper_columns[helper_header]
                formula_range = master_sheet.Range(master_sheet.Cells(first_data_row, helper_column), master_sheet.Cells(final_data_row, helper_column))
                formula_range.Clear()
                formula_range.FormulaR1C1 = formula

            excel.Calculation = XL_CALCULATION_AUTOMATIC
            excel.CalculateFullRebuild()
            written_expense_values = self._column_values(master_sheet, master_expense_column, final_data_row)
            expected_bsny_total = self._numeric_total(self._values_for_rows(bsny_sheet, self._unique_header_column(bsny_headers, EXPENSE_HEADER, bsny_sheet), bsny_rows))
            expected_sancap_total = self._numeric_total(self._values_for_rows(sancap_sheet, self._unique_header_column(sancap_headers, EXPENSE_HEADER, sancap_sheet), sancap_rows))
            actual_bsny_total = self._numeric_total(written_expense_values[:len(bsny_rows)])
            actual_sancap_total = self._numeric_total(written_expense_values[len(bsny_rows):])
            if round(actual_bsny_total + actual_sancap_total, 2) != round(expected_bsny_total + expected_sancap_total, 2):
                raise AssertionError("Expense total mismatch after workbook refresh.")
            if not is_reporting_year(master_sheet.Cells(first_data_row, master_year_column).Value2, request.reporting_year):
                raise AssertionError("The first output row is not a BSNY reporting-year record.")
            if not is_reporting_year(master_sheet.Cells(first_data_row + len(bsny_rows), master_year_column).Value2, request.reporting_year):
                raise AssertionError("The first SanCap output row is not a reporting-year record.")
            for helper_header, helper_column in helper_columns.items():
                if not master_sheet.Cells(final_data_row, helper_column).HasFormula:
                    raise AssertionError(f"Helper formula was not filled to the last row: {helper_header}")

            staging_sheet.Visible = True
            staging_sheet.Delete()
            staging_sheet = None
            master_workbook.Save()
            refresh_succeeded = True
            return {
                "bsny_rows": len(bsny_rows),
                "sancap_rows": len(sancap_rows),
                "total_expense": actual_bsny_total + actual_sancap_total,
            }
        finally:
            if bsny_workbook is not None:
                bsny_workbook.Close(SaveChanges=False)
            if sancap_workbook is not None:
                sancap_workbook.Close(SaveChanges=False)
            if staging_sheet is not None:
                try:
                    staging_sheet.Visible = True
                    staging_sheet.Delete()
                except Exception:
                    pass
            if master_workbook is not None:
                master_workbook.Close(SaveChanges=refresh_succeeded)
            if excel is not None:
                if original_autofill is not None:
                    excel.AutoCorrect.AutoFillFormulasInLists = original_autofill
                excel.Quit()
            if com_initialized:
                pythoncom.CoUninitialize()
