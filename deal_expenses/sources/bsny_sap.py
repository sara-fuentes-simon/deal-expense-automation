"""Adapter for the BSNY SAP source workbook."""

from __future__ import annotations

from deal_expenses.sources.base_sap import SapSourceAdapter


class BsnySapAdapter(SapSourceAdapter):
    """Read SAP source rows from the BSNY SAP worksheet."""

    key = "bsny"
    display_name = "BSNY SAP report"
    sheet_name = "SAP"

    def select_worksheet(self, workbook):
        if self.sheet_name not in workbook.sheetnames:
            raise ValueError(f"Worksheet '{self.sheet_name}' was not found.")
        return workbook[self.sheet_name]
