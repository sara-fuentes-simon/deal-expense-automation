from pathlib import Path

import pytest

from deal_expenses.models import RunRequest, SapRunRequest, SourceSummary, ValidationResult
from deal_expenses.concur_pipeline import ConcurExpensePipeline
from deal_expenses.sap_pipeline import SapExpensePipeline
from deal_expenses.sources import BsnyConcurAdapter, SanCapConcurAdapter
from deal_expenses.concur_validation import ConcurMasterWorkbookValidator
from deal_expenses.concur_workbook_writer import ConcurWorkbookWriter


class StubWriter(ConcurWorkbookWriter):
    def write(self, request: RunRequest) -> dict[str, int | float]:
        return {"bsny_rows": 1, "sancap_rows": 1, "total_expense": 30.0}


class StubSapAdapter:
    def __init__(self, key: str) -> None:
        self.key = key
        self.display_name = f"{key.upper()} SAP report"

    def validate(self, source_path: Path) -> list[ValidationResult]:
        return [ValidationResult(self.display_name, True, "Workbook structure is valid.")]

    def summarize(self, source_path: Path, year: int) -> SourceSummary:
        return SourceSummary(self.key, self.display_name, 1, 10.0)


class StubSapMasterValidator:
    def validate(self, master_path: Path) -> list[ValidationResult]:
        return [ValidationResult("SAP master workbook", True, "Workbook structure is valid.")]


class StubSapWriter:
    def write(self, request: SapRunRequest) -> dict[str, int]:
        return {"bsny_rows": 1, "sancap_rows": 1, "appended_rows": 2}


@pytest.mark.parametrize("reporting_year", [2025, 2026])
def test_pipeline_completes_preflight_with_sample_workbooks(reporting_year: int):
    root = Path("samples")
    request = RunRequest(
        master_path=root / "Expense Report Relating to deals 01-01-26 to 6-30-26_V8 - COMBINED.xlsx",
        source_paths={
            "bsny": root / "BSNY - SAP & Concur Repoort May 2025 - June 2026.xlsx",
            "sancap": root / "SanCap - Expense Report USA May 2025 - JUNE 2026.xlsx",
        },
        output_path=Path("outputs/test.xlsx"),
        reporting_year=reporting_year,
    )
    pipeline = ConcurExpensePipeline(
        [BsnyConcurAdapter(), SanCapConcurAdapter()],
        ConcurMasterWorkbookValidator(),
        StubWriter(),
    )

    events = list(pipeline.run(request))

    assert pipeline.result.success
    assert pipeline.result.total_expense == 30.0
    assert events[-1].step == "Complete"
    assert all(str(reporting_year) in event.message for event in events)


def test_sap_pipeline_completes_preflight_with_stubbed_workbooks(tmp_path: Path):
    master_path = tmp_path / "master.xlsx"
    bsny_path = tmp_path / "bsny_sap.xlsx"
    sancap_path = tmp_path / "sancap_sap.xlsx"
    for path in (master_path, bsny_path, sancap_path):
        path.touch()
    request = SapRunRequest(
        master_path=master_path,
        source_paths={"bsny": bsny_path, "sancap": sancap_path},
        reporting_year=2026,
    )
    pipeline = SapExpensePipeline(
        [StubSapAdapter("bsny"), StubSapAdapter("sancap")],
        StubSapMasterValidator(),
        StubSapWriter(),
    )

    events = list(pipeline.run(request))

    assert pipeline.result.success
    assert pipeline.metrics == {"bsny_rows": 1, "sancap_rows": 1, "appended_rows": 2}
    assert events[-1].step == "Complete"
    assert all("2026" in event.message for event in events)
