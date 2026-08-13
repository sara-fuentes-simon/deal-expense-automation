"""Adapter for the BSNY Concur source workbook."""

from __future__ import annotations

from deal_expenses.sources.base import SourceAdapter


class BsnyConcurAdapter(SourceAdapter):
    key = "bsny"
    display_name = "BSNY Concur report"
    sheet_name = "Concur"

    def select_worksheet(self, workbook):
        if self.sheet_name not in workbook.sheetnames:
            raise ValueError(f"Worksheet '{self.sheet_name}' was not found.")
        return workbook[self.sheet_name]
