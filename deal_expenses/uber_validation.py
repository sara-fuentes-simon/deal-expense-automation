"""Structural validation for the Uber section of the master workbook."""

from __future__ import annotations

from collections import defaultdict
from pathlib import Path

from openpyxl import load_workbook

from deal_expenses.hr_refresh import build_hr_refresh_plan
from deal_expenses.models import ValidationResult
from deal_expenses.sources.uber import normalize_uber_header, worksheet_headers


UBER_REPORT_SHEET = "Uber Report"
UBER_HEADER_ROW = 6
HELPER_HEADERS = ("Full Name", "Final Cost Center", "LOB", "Company Name")
IN_SCOPE_HEADERS = ("In Scope", "In Scope?")
HR_REQUIRED_HEADERS = ("Employee ID", "Name", "Cost Center Name")


def master_headers(worksheet) -> dict[str, list[int]]:
    headers: dict[str, list[int]] = defaultdict(list)
    for column, cell in enumerate(worksheet[UBER_HEADER_ROW], start=1):
        header = normalize_uber_header(cell.value)
        if header:
            headers[header].append(column)
    return headers


def required_column(headers: dict[str, list[int]], header: str, worksheet_name: str) -> int:
    columns = headers.get(normalize_uber_header(header), [])
    if len(columns) != 1:
        description = "missing" if not columns else f"repeated in columns {columns}"
        raise ValueError(f"Required header '{header}' is {description} in '{worksheet_name}'.")
    return columns[0]


def in_scope_column(headers: dict[str, list[int]], worksheet_name: str) -> int:
    matches = [column for header in IN_SCOPE_HEADERS for column in headers.get(normalize_uber_header(header), [])]
    if len(matches) != 1:
        raise ValueError(f"Required header 'In Scope' or 'In Scope?' is missing or repeated in '{worksheet_name}'.")
    return matches[0]


def cost_center_columns(worksheet, headers: dict[str, list[int]]) -> tuple[int, int]:
    columns = headers.get(normalize_uber_header("Cost Center"), [])
    if len(columns) != 2:
        raise ValueError(f"Expected exactly two 'Cost Center' columns in '{worksheet.title}'.")
    if hasattr(worksheet, "Cells"):
        labels = {normalize_uber_header(worksheet.Cells(UBER_HEADER_ROW - 1, column).Value2): column for column in columns}
    else:
        labels = {normalize_uber_header(worksheet.cell(UBER_HEADER_ROW - 1, column).value): column for column in columns}
    id_column = labels.get("using employee id")
    name_column = labels.get("using employee name")
    if id_column is None or name_column is None:
        raise ValueError("The two Cost Center columns must be labeled 'Using Employee ID' and 'Using Employee Name' above the header row.")
    return id_column, name_column


class UberMasterWorkbookValidator:
    """Verify the master workbook can safely receive an Uber refresh."""

    def validate(self, master_path: Path) -> list[ValidationResult]:
        try:
            workbook = load_workbook(master_path, read_only=True, data_only=False)
            try:
                if UBER_REPORT_SHEET not in workbook.sheetnames:
                    raise ValueError(f"Worksheet '{UBER_REPORT_SHEET}' was not found.")
                hr_sheets = [worksheet for worksheet in workbook.worksheets if worksheet.title.startswith("HR HC Combined")]
                if len(hr_sheets) != 1:
                    raise ValueError(f"Expected exactly one worksheet beginning 'HR HC Combined'; found {[sheet.title for sheet in hr_sheets]}.")
                uber_sheet = workbook[UBER_REPORT_SHEET]
                headers = master_headers(uber_sheet)
                for header in (*HELPER_HEADERS, "Employee ID", "Transaction Amount USD"):
                    required_column(headers, header, uber_sheet.title)
                in_scope = in_scope_column(headers, uber_sheet.title)
                id_cost_center, name_cost_center = cost_center_columns(uber_sheet, headers)
                formula_columns = [required_column(headers, header, uber_sheet.title) for header in HELPER_HEADERS] + [in_scope]
                missing_templates = [str(uber_sheet.cell(UBER_HEADER_ROW, column).value) for column in formula_columns if not isinstance(uber_sheet.cell(UBER_HEADER_ROW + 1, column).value, str) or not uber_sheet.cell(UBER_HEADER_ROW + 1, column).value.startswith("=")]
                if missing_templates:
                    raise ValueError("Missing row 7 formula template for: " + ", ".join(missing_templates))
                if id_cost_center == name_cost_center:
                    raise ValueError("Cost Center helper columns must be distinct.")
                hr_headers = {normalize_uber_header(cell.value) for cell in hr_sheets[0][1]}
                missing_hr = [header for header in (*HR_REQUIRED_HEADERS, "Network ID", "Effective Date") if normalize_uber_header(header) not in hr_headers]
                if missing_hr:
                    raise ValueError("HR HC Combined is missing required column(s): " + ", ".join(missing_hr))
            finally:
                workbook.close()
        except Exception as error:
            return [ValidationResult("Uber master workbook", False, str(error))]
        return [ValidationResult("Uber master workbook", True, "Workbook structure and helper formulas are valid.")]


def validate_headcount_source(master_path: Path, source_path: Path) -> ValidationResult:
    """Check the Headcount source can be merged into the master HR worksheet."""
    try:
        plan = build_hr_refresh_plan(master_path, source_path)
    except Exception as error:
        return ValidationResult("Headcount HR source", False, str(error))
    return ValidationResult(
        "Headcount HR source",
        True,
        f"{plan.source_rows:,} source row(s) will refresh '{plan.master_sheet_name}'; {plan.superseded_rows:,} superseded row(s) will be removed.",
    )


def validate_source_compatibility(master_path: Path, source_path: Path, adapter) -> ValidationResult:
    """Require every source field to have one non-helper target column."""
    try:
        master_workbook = load_workbook(master_path, read_only=True, data_only=False)
        source_workbook = load_workbook(source_path, read_only=True, data_only=False)
        try:
            target_headers = master_headers(master_workbook[UBER_REPORT_SHEET])
            source_headers = worksheet_headers(adapter.select_worksheet(source_workbook))
            source_only = sorted(header for header in source_headers if header not in target_headers)
            repeated_targets = sorted(header for header in source_headers if len(target_headers.get(header, [])) != 1)
            if source_only or repeated_targets:
                messages = []
                if source_only:
                    messages.append("Source-only columns: " + ", ".join(source_only))
                if repeated_targets:
                    messages.append("Target columns missing or repeated: " + ", ".join(repeated_targets))
                raise ValueError("; ".join(messages))
        finally:
            source_workbook.close()
            master_workbook.close()
    except Exception as error:
        return ValidationResult("Uber source mapping", False, str(error))
    return ValidationResult("Uber source mapping", True, "All Uber source columns map to Uber Report by header name.")
