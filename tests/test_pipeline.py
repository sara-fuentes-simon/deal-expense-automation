from pathlib import Path

from deal_expenses.models import RunRequest
from deal_expenses.pipeline import DealExpensePipeline
from deal_expenses.sources import BsnyConcurAdapter, SanCapAdapter
from deal_expenses.validation import MasterWorkbookValidator
from deal_expenses.workbook_writer import WorkbookWriter


class StubWriter(WorkbookWriter):
    def write(self, request: RunRequest) -> dict[str, int | float]:
        return {"bsny_rows": 1, "sancap_rows": 1, "total_expense": 30.0}


def test_pipeline_completes_preflight_with_sample_workbooks():
    root = Path("samples")
    request = RunRequest(
        master_path=root / "Expense Report Relating to deals 01-01-26 to 6-30-26_V8 - COMBINED.xlsx",
        source_paths={
            "bsny": root / "BSNY - SAP & Concur Repoort May 2025 - June 2026.xlsx",
            "sancap": root / "SanCap - Expense Report USA May 2025 - JUNE 2026.xlsx",
        },
        output_path=Path("outputs/test.xlsx"),
    )
    pipeline = DealExpensePipeline(
        [BsnyConcurAdapter(), SanCapAdapter()],
        MasterWorkbookValidator(),
        StubWriter(),
    )

    events = list(pipeline.run(request))

    assert pipeline.result.success
    assert pipeline.result.total_expense == 30.0
    assert events[-1].step == "Complete"
