"""Adapter for the SanCap expense report source workbook."""

from __future__ import annotations

from deal_expenses.sources.base import SourceAdapter


class SanCapAdapter(SourceAdapter):
    key = "sancap"
    display_name = "SanCap expense report"

    def select_worksheet(self, workbook):
        return workbook.worksheets[0]
