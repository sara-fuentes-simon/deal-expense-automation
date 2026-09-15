"""Application service that coordinates one Uber workbook refresh."""

from __future__ import annotations

from collections.abc import Iterator

from deal_expenses.models import ProgressEvent, RunResult, UberRunRequest, ValidationResult
from deal_expenses.uber_validation import UberMasterWorkbookValidator, validate_headcount_source, validate_source_compatibility
from deal_expenses.uber_workbook_writer import UberExcelComWorkbookWriter


class UberExpensePipeline:
    """Run Uber preflight checks before refreshing the Uber Report worksheet."""

    def __init__(self, adapter, master_validator: UberMasterWorkbookValidator, workbook_writer: UberExcelComWorkbookWriter) -> None:
        self._adapter = adapter
        self._master_validator = master_validator
        self._workbook_writer = workbook_writer
        self.result = RunResult(success=False)
        self.metrics: dict[str, int | float] = {}

    def run(self, request: UberRunRequest) -> Iterator[ProgressEvent]:
        validations: list[ValidationResult] = []
        summaries = []
        self.metrics = {}
        try:
            yield ProgressEvent("Uber preflight", "Checking Uber workbook structure.", 85)
            request.validate_paths_exist()
            for check in self._master_validator.validate(request.master_path):
                validations.append(check)
            if not all(check.passed for check in validations):
                self.result = RunResult(False, validations=validations, error_message=validations[-1].message)
                yield ProgressEvent("Stopped", self.result.error_message, 100)
                return

            yield ProgressEvent("Headcount validation", "Checking the Headcount HR source worksheet.", 87)
            headcount_check = validate_headcount_source(request.master_path, request.hr_source_path)
            validations.append(headcount_check)
            if not headcount_check.passed:
                self.result = RunResult(False, validations=validations, error_message=headcount_check.message)
                yield ProgressEvent("Stopped", self.result.error_message, 100)
                return

            yield ProgressEvent("Uber source validation", "Checking the Uber source worksheet.", 88)
            source_checks = self._adapter.validate(request.source_path)
            validations.extend(source_checks)
            if not all(check.passed for check in source_checks):
                self.result = RunResult(False, validations=validations, error_message=source_checks[0].message)
                yield ProgressEvent("Stopped", self.result.error_message, 100)
                return
            mapping_check = validate_source_compatibility(request.master_path, request.source_path, self._adapter)
            validations.append(mapping_check)
            if not mapping_check.passed:
                self.result = RunResult(False, validations=validations, error_message=mapping_check.message)
                yield ProgressEvent("Stopped", self.result.error_message, 100)
                return
            summaries.append(self._adapter.summarize(request.source_path, request.reporting_year))

            yield ProgressEvent("Uber Excel refresh", "Refreshing HR Headcount, then Uber Report, in Microsoft Excel.", 93)
            self.metrics = self._workbook_writer.write(request)
            validations.append(
                ValidationResult(
                    "Uber Final Cost Center coverage",
                    True,
                    f"{self.metrics['unassigned_final_cost_center_rows']:,} of {self.metrics['imported_rows']:,} rows ({self.metrics['unassigned_final_cost_center_percent']:.2f}%) have no Final Cost Center.",
                )
            )
            self.result = RunResult(True, request.master_path, validations, summaries, float(summaries[0].expense_total))
            yield ProgressEvent("Complete", "Uber data refreshed and validated.", 100)
        except Exception as error:
            self.result = RunResult(False, validations=validations, source_summaries=summaries, error_message=str(error))
            yield ProgressEvent("Stopped", str(error), 100)
