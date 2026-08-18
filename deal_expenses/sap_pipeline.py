"""Application service that coordinates one SAP workbook refresh."""

from __future__ import annotations

from collections.abc import Iterator

from deal_expenses.models import ProgressEvent, RunResult, SapRunRequest, ValidationResult
from deal_expenses.sap_validation import SapMasterWorkbookValidator
from deal_expenses.sap_workbook_writer import SapExcelComWorkbookWriter
from deal_expenses.sources import SapSourceAdapter


class SapExpensePipeline:
    """Run SAP preflight checks before delegating workbook mutation to a writer."""

    def __init__(
        self,
        adapters: list[SapSourceAdapter],
        master_validator: SapMasterWorkbookValidator,
        workbook_writer: SapExcelComWorkbookWriter,
    ) -> None:
        self._adapters = {adapter.key: adapter for adapter in adapters}
        self._master_validator = master_validator
        self._workbook_writer = workbook_writer
        self.result = RunResult(success=False)
        self.metrics: dict[str, int] = {}

    def run(self, request: SapRunRequest) -> Iterator[ProgressEvent]:
        validations: list[ValidationResult] = []
        summaries = []
        self.metrics = {}
        try:
            yield ProgressEvent("SAP preflight", f"Checking SAP workbooks for {request.reporting_year} data.", 55)
            request.validate_paths_exist()
            master_checks = self._master_validator.validate(request.master_path)
            validations.extend(master_checks)
            if not all(check.passed for check in master_checks):
                self.result = RunResult(False, validations=validations, error_message=master_checks[0].message)
                yield ProgressEvent("Stopped", self.result.error_message, 100)
                return

            for index, (key, adapter) in enumerate(self._adapters.items(), start=1):
                if key not in request.source_paths:
                    raise ValueError(f"No upload was supplied for {adapter.display_name}.")
                yield ProgressEvent(
                    "SAP source validation",
                    f"Checking {adapter.display_name} for {request.reporting_year} data.",
                    55 + index * 10,
                )
                source_checks = adapter.validate(request.source_paths[key])
                validations.extend(source_checks)
                if not all(check.passed for check in source_checks):
                    self.result = RunResult(False, validations=validations, error_message=source_checks[0].message)
                    yield ProgressEvent("Stopped", self.result.error_message, 100)
                    return
                summaries.append(adapter.summarize(request.source_paths[key], request.reporting_year))

            yield ProgressEvent("SAP Excel refresh", f"Refreshing SAP data for {request.reporting_year} in Microsoft Excel.", 80)
            self.metrics = self._workbook_writer.write(request)
            self.result = RunResult(
                success=True,
                output_path=request.master_path,
                validations=validations,
                source_summaries=summaries,
                total_expense=sum(summary.expense_total for summary in summaries),
            )
            yield ProgressEvent("Complete", f"SAP {request.reporting_year} data refreshed and validated.", 100)
        except Exception as error:
            self.result = RunResult(
                success=False,
                validations=validations,
                source_summaries=summaries,
                error_message=str(error),
            )
            yield ProgressEvent("Stopped", str(error), 100)
