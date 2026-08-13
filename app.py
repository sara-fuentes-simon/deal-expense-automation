"""Streamlit interface for the deal expense workbook refresh."""

from __future__ import annotations

from pathlib import Path
from tempfile import TemporaryDirectory

import streamlit as st

from deal_expenses.models import RunRequest, sanitize_output_filename
from deal_expenses.pipeline import DealExpensePipeline
from deal_expenses.sources import BsnyConcurAdapter, SanCapAdapter
from deal_expenses.validation import MasterWorkbookValidator
from deal_expenses.workbook_writer import ExcelComWorkbookWriter


st.set_page_config(page_title="Deal Expenses Automation", page_icon=":material/receipt_long:", layout="wide")


@st.cache_resource
def build_pipeline() -> DealExpensePipeline:
    return DealExpensePipeline(
        adapters=[BsnyConcurAdapter(), SanCapAdapter()],
        master_validator=MasterWorkbookValidator(),
        workbook_writer=ExcelComWorkbookWriter(),
    )


def save_upload(upload, directory: Path, filename: str) -> Path:
    path = directory / filename
    path.write_bytes(upload.getbuffer())
    return path


def render_results() -> None:
    result = st.session_state.get("run_result")
    if result is None:
        return

    if not result.success:
        st.error(result.error_message or "The workbook refresh did not complete.")
    else:
        st.success("The master workbook was refreshed and validated.")
        columns = st.columns(3)
        columns[0].metric("BSNY rows", f"{result.source_summaries[0].row_count:,}")
        columns[1].metric("SanCap rows", f"{result.source_summaries[1].row_count:,}")
        columns[2].metric("Expense total", f"${result.total_expense:,.2f}")
        st.download_button(
            "Download refreshed workbook",
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


def main() -> None:
    st.title("Deal Expenses Automation")
    st.caption("Refresh the deal expense master workbook from BSNY and SanCap source reports.")

    left_column, right_column = st.columns(2)
    with left_column:
        master_upload = st.file_uploader("Master workbook", type=["xlsx"], key="master")
        bsny_upload = st.file_uploader("BSNY Concur report", type=["xlsx"], key="bsny")
    with right_column:
        sancap_upload = st.file_uploader("SanCap expense report", type=["xlsx"], key="sancap")
        output_filename = st.text_input("Output master filename", value="refreshed_deal_expenses.xlsx")

    run_requested = st.button("Refresh master workbook", type="primary", use_container_width=True)
    if run_requested:
        if not all([master_upload, bsny_upload, sancap_upload]):
            st.error("Upload the master workbook, BSNY report, and SanCap report before running.")
            return
        try:
            clean_output_filename = sanitize_output_filename(output_filename)
        except ValueError as error:
            st.error(str(error))
            return

        st.session_state.pop("run_result", None)
        st.session_state.pop("output_bytes", None)
        with TemporaryDirectory(prefix="deal_expenses_") as temporary_directory:
            directory = Path(temporary_directory)
            request = RunRequest(
                master_path=save_upload(master_upload, directory, "master.xlsx"),
                source_paths={
                    "bsny": save_upload(bsny_upload, directory, "bsny.xlsx"),
                    "sancap": save_upload(sancap_upload, directory, "sancap.xlsx"),
                },
                output_path=directory / clean_output_filename,
            )
            pipeline = build_pipeline()
            progress = st.progress(0, text="Preparing refresh.")
            with st.status("Refreshing workbook", expanded=True) as status:
                for event in pipeline.run(request):
                    progress.progress(event.percent, text=event.message)
                    status.write(event.message)
                if pipeline.result.success:
                    status.update(label="Refresh complete", state="complete", expanded=False)
                    st.session_state["output_bytes"] = request.output_path.read_bytes()
                    st.session_state["output_filename"] = clean_output_filename
                else:
                    status.update(label="Refresh stopped", state="error", expanded=True)
            st.session_state["run_result"] = pipeline.result

    render_results()


if __name__ == "__main__":
    main()
