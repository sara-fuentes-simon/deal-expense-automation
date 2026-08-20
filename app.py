"""Streamlit interface for the deal expense workbook refresh."""

from __future__ import annotations

from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd
import streamlit as st

from deal_expenses.models import RunRequest, SapRunRequest, sanitize_output_filename
from deal_expenses.pivot_preview import PivotPreview, extract_pivot_previews
from deal_expenses.concur_pipeline import ConcurExpensePipeline
from deal_expenses.sap_pipeline import SapExpensePipeline
from deal_expenses.sap_validation import SapMasterWorkbookValidator
from deal_expenses.sap_workbook_writer import SapExcelComWorkbookWriter
from deal_expenses.sources import BsnyConcurAdapter, BsnySapAdapter, SanCapConcurAdapter, SanCapSapAdapter
from deal_expenses.concur_validation import ConcurMasterWorkbookValidator
from deal_expenses.concur_workbook_writer import ConcurExcelComWorkbookWriter


st.set_page_config(page_title="Deal Expenses Automation", page_icon=":material/receipt_long:", layout="wide")


@st.cache_resource
def build_concur_pipeline() -> ConcurExpensePipeline:
    return ConcurExpensePipeline(
        adapters=[BsnyConcurAdapter(), SanCapConcurAdapter()],
        master_validator=ConcurMasterWorkbookValidator(),
        workbook_writer=ConcurExcelComWorkbookWriter(),
    )


@st.cache_resource
def build_sap_pipeline() -> SapExpensePipeline:
    return SapExpensePipeline(
        adapters=[BsnySapAdapter(), SanCapSapAdapter()],
        master_validator=SapMasterWorkbookValidator(),
        workbook_writer=SapExcelComWorkbookWriter(),
    )


def save_upload(upload, directory: Path, filename: str) -> Path:
    path = directory / filename
    path.write_bytes(upload.getbuffer())
    return path


def render_pivot_previews() -> None:
    previews: dict[str, list[PivotPreview]] = st.session_state.get("pivot_previews", {})
    preview_error: str | None = st.session_state.get("pivot_preview_error")
    if not previews and preview_error is None:
        return

    st.subheader("Analysis pivots")
    if preview_error is not None:
        st.warning(f"Pivot previews could not be loaded: {preview_error}")
        return

    for column, sheet_name, title in zip(
        st.columns(2),
        ("Concur Analysis - PIVOTS", "SAP Invoices Analysis - PIVOTS"),
        ("Concur expenses by LOB", "SAP invoices by LOB"),
        strict=True,
    ):
        with column:
            st.markdown(f"**{title}**")
            sheet_previews = previews.get(sheet_name, [])
            if not sheet_previews:
                st.info("No pivot table was found on this worksheet.")
                continue
            preview = sheet_previews[0]
            dataframe = pd.DataFrame(preview.rows, columns=preview.headers)
            st.dataframe(
                dataframe.style.apply(
                    lambda row: ["font-weight: bold" if preview.emphasis_rows[row.name] else ""] * len(row),
                    axis=1,
                ).format(precision=2, thousands=","),
                hide_index=True,
                height=min(420, 72 + len(dataframe) * 35),
            )


def render_results() -> None:
    result = st.session_state.get("run_result")
    if result is None:
        return

    if not result.success:
        st.error(result.error_message or "The workbook refresh did not complete.")
    else:
        st.success("The master workbook was refreshed and validated.")
        columns = st.columns(4)
        columns[0].metric("Reporting year", st.session_state["run_reporting_year"])
        columns[1].metric("BSNY rows", f"{result.source_summaries[0].row_count:,}")
        columns[2].metric("SanCap rows", f"{result.source_summaries[1].row_count:,}")
        columns[3].metric("Expense total", f"${result.total_expense:,.2f}")
        sap_result: dict[str, int] | None = st.session_state.get("sap_result")
        if sap_result is not None:
            sap_columns = st.columns(3)
            sap_columns[0].metric("SAP BSNY rows", f"{sap_result['bsny_rows']:,}")
            sap_columns[1].metric("SAP SanCap rows", f"{sap_result['sancap_rows']:,}")
            sap_columns[2].metric("New SAP rows", f"{sap_result['appended_rows']:,}")
        render_pivot_previews()
        st.download_button(
            "Download combined refreshed workbook",
            data=st.session_state["output_bytes"],
            file_name=st.session_state["output_filename"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

    with st.expander("Validation details"):
        for validation in result.validations:
            if validation.passed:
                st.success(f"{validation.name}: {validation.message}")
            else:
                st.error(f"{validation.name}: {validation.message}")
        for validation in st.session_state.get("sap_validations", []):
            if validation.passed:
                st.success(f"{validation.name}: {validation.message}")
            else:
                st.error(f"{validation.name}: {validation.message}")


def main() -> None:
    st.title("Deal Expenses Automation")
    st.caption("Refresh the Concur and SAP sections of one deal expense master workbook.")

    master_upload = st.file_uploader("Master workbook", type=["xlsx"], key="master")

    concur_column, sap_column = st.columns(2)
    with concur_column:
        st.subheader("Concur Sources")
        bsny_upload = st.file_uploader("BSNY Concur report", type=["xlsx"], key="bsny")
        sancap_upload = st.file_uploader("SanCap Concur report", type=["xlsx"], key="sancap")
    with sap_column:
        st.subheader("SAP Sources")
        bsny_sap_upload = st.file_uploader("BSNY SAP report", type=["xlsx"], key="bsny_sap")
        sancap_sap_upload = st.file_uploader("SanCap SAP report", type=["xlsx"], key="sancap_sap")

    settings_column, _ = st.columns(2)
    with settings_column:
        output_filename = st.text_input("Output master filename", value="refreshed_deal_expenses.xlsx")
        reporting_year = st.number_input("Reporting year", value=date.today().year, step=1, format="%d")

    run_requested = st.button("Refresh combined master workbook", type="primary", width="stretch")
    if run_requested:
        if not all([master_upload, bsny_upload, sancap_upload, bsny_sap_upload, sancap_sap_upload]):
            st.error("Upload the master workbook and all Concur and SAP source reports before running.")
            return
        try:
            clean_output_filename = sanitize_output_filename(output_filename)
        except ValueError as error:
            st.error(str(error))
            return

        st.session_state.pop("run_result", None)
        st.session_state.pop("output_bytes", None)
        st.session_state.pop("run_reporting_year", None)
        st.session_state.pop("sap_result", None)
        st.session_state.pop("sap_validations", None)
        st.session_state.pop("pivot_previews", None)
        st.session_state.pop("pivot_preview_error", None)
        with TemporaryDirectory(prefix="deal_expenses_") as temporary_directory:
            directory = Path(temporary_directory)
            request = RunRequest(
                master_path=save_upload(master_upload, directory, "master.xlsx"),
                source_paths={
                    "bsny": save_upload(bsny_upload, directory, "bsny.xlsx"),
                    "sancap": save_upload(sancap_upload, directory, "sancap.xlsx"),
                },
                output_path=directory / clean_output_filename,
                reporting_year=int(reporting_year),
            )
            pipeline = build_concur_pipeline()
            progress = st.progress(0, text=f"Preparing {request.reporting_year} refresh.")
            with st.status(f"Refreshing {request.reporting_year} workbook", expanded=True) as status:
                for event in pipeline.run(request):
                    progress.progress(event.percent, text=event.message)
                    status.write(event.message)
                if pipeline.result.success:
                    sap_request = SapRunRequest(
                        master_path=request.output_path,
                        source_paths={
                            "bsny": save_upload(bsny_sap_upload, directory, "bsny_sap.xlsx"),
                            "sancap": save_upload(sancap_sap_upload, directory, "sancap_sap.xlsx"),
                        },
                        reporting_year=request.reporting_year,
                    )
                    sap_pipeline = build_sap_pipeline()
                    for event in sap_pipeline.run(sap_request):
                        progress.progress(event.percent, text=event.message)
                        status.write(event.message)
                    if not sap_pipeline.result.success:
                        pipeline.result.success = False
                        pipeline.result.error_message = f"SAP refresh failed: {sap_pipeline.result.error_message}"
                        st.session_state["sap_validations"] = sap_pipeline.result.validations
                        status.update(label="Refresh stopped", state="error", expanded=True)
                    else:
                        sap_result = sap_pipeline.metrics
                        status.write(f"Appended {sap_result['appended_rows']:,} SAP row(s).")
                        status.update(label=f"{request.reporting_year} combined refresh complete", state="complete", expanded=False)
                        try:
                            st.session_state["pivot_previews"] = extract_pivot_previews(
                                request.output_path,
                                ["Concur Analysis - PIVOTS", "SAP Invoices Analysis - PIVOTS"],
                            )
                        except RuntimeError as error:
                            st.session_state["pivot_preview_error"] = str(error)
                        st.session_state["output_bytes"] = request.output_path.read_bytes()
                        st.session_state["output_filename"] = clean_output_filename
                        st.session_state["run_reporting_year"] = request.reporting_year
                        st.session_state["sap_result"] = sap_result
                        st.session_state["sap_validations"] = sap_pipeline.result.validations
                else:
                    status.update(label="Refresh stopped", state="error", expanded=True)
            st.session_state["run_result"] = pipeline.result

    render_results()


if __name__ == "__main__":
    main()
