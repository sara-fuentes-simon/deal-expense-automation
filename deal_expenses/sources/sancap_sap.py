"""Adapter for the SanCap SAP source workbook."""

from __future__ import annotations

from deal_expenses.sources.base_sap import SapSourceAdapter


class SanCapSapAdapter(SapSourceAdapter):
    """Read SAP source rows from the sole SanCap SAP worksheet."""

    key = "sancap"
    display_name = "SanCap SAP report"

    def select_worksheet(self, workbook):
        if len(workbook.worksheets) != 1:
            raise ValueError("The SANCAP SAP report must contain exactly one worksheet.")
        return workbook.worksheets[0]
