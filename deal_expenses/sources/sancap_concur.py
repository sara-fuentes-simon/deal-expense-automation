"""Adapter for the SanCap Concur source workbook."""

from __future__ import annotations

from deal_expenses.sources.base_concur import SourceAdapter


class SanCapConcurAdapter(SourceAdapter):
    """Read Concur source rows from the first SanCap worksheet."""

    key = "sancap"
    display_name = "SanCap Concur report"

    def select_worksheet(self, workbook):
        return workbook.worksheets[0]
