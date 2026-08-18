"""Application service that coordinates one Concur workbook refresh."""

from __future__ import annotations

from collections.abc import Iterator

from deal_expenses.concur_validation import ConcurMasterWorkbookValidator
from deal_expenses.models import ProgressEvent, RunRequest, RunResult, ValidationResult
from deal_expenses.sources.base_concur import SourceAdapter
from deal_expenses.concur_workbook_writer import ConcurWorkbookWriter


class ConcurExpensePipeline:
    """Run Concur preflight checks before delegating workbook mutation to a writer."""

    def __init__(
        self,
        adapters: list[SourceAdapter],
        master_validator: ConcurMasterWorkbookValidator,
        workbook_writer: ConcurWorkbookWriter,
    ) -> None:
        self._adapters = {adapter.key: adapter for adapter in adapters}
        self._master_validator = master_validator
        self._workbook_writer = workbook_writer
        self.result = RunResult(success=False)

    def run(self, request: RunRequest) -> Iterator[ProgressEvent]:
        validations: list[ValidationResult] = []
        summaries = []
        try:
            yield ProgressEvent("Concur preflight", f"Checking Concur workbooks for {request.reporting_year} data.", 10)
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
                    "Concur source validation",
                    f"Checking {adapter.display_name} for {request.reporting_year} data.",
                    10 + index * 15,
                )
                source_checks = adapter.validate(request.source_paths[key])
                validations.extend(source_checks)
                if not all(check.passed for check in source_checks):
                    self.result = RunResult(False, validations=validations, error_message=source_checks[0].message)
                    yield ProgressEvent("Stopped", self.result.error_message, 100)
                    return
                summaries.append(adapter.summarize(request.source_paths[key], request.reporting_year))

            yield ProgressEvent(
                "Concur Excel refresh",
                f"Refreshing Concur data for {request.reporting_year} in Microsoft Excel.",
                50,
            )
            metrics = self._workbook_writer.write(request)
            total_expense = float(metrics["total_expense"])
            self.result = RunResult(
                success=True,
                output_path=request.output_path,
                validations=validations,
                source_summaries=summaries,
                total_expense=total_expense,
            )
            yield ProgressEvent("Complete", f"Concur {request.reporting_year} data refreshed and validated.", 100)
        except Exception as error:
            self.result = RunResult(
                success=False,
                validations=validations,
                source_summaries=summaries,
                error_message=str(error),
            )
            yield ProgressEvent("Stopped", str(error), 100)
