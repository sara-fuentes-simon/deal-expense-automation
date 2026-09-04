# Deal Expenses Automation

Local Streamlit application for refreshing the Concur and SAP sections of one deal expense master workbook.

## Requirements

- Windows
- Microsoft Excel installed
- Python 3.11 or newer
- The uploaded workbooks must not be open in Excel while the refresh runs

## Install

```powershell
python -m venv venv
.\venv\Scripts\python.exe -m pip install --upgrade pip
.\venv\Scripts\python.exe -m pip install -r requirements.txt
```

If `win32com` is unavailable after installation, run:

```powershell
.\venv\Scripts\python.exe -m pip install pywin32
```

## Run

```powershell
.\venv\Scripts\streamlit.exe run app.py
```

The browser opens the local frontend, normally at `http://localhost:8501`.

1. Upload the master workbook.
2. Upload BSNY and SanCap reports for both Concur and SAP.
3. Enter the reporting year and desired `.xlsx` output filename.
4. Select **Refresh combined master workbook**.
5. Review progress, validation details, and the Concur/SAP pivot previews, then download the combined workbook.

The application refreshes Concur first, then appends validated SAP rows to that same temporary workbook. It refreshes the workbook's pivot caches before saving, so the downloaded file opens with current pivot results. The SAP refresh requires the master workbook to contain `SAP Report` or `SAP Invoices Report`, an `In Scope CCs` sheet, and the expected SAP master headers.

Uploaded files are stored only in a temporary directory while the refresh runs. The completed workbook is retained in the browser session until the page is refreshed or a new refresh starts.

## Test

```powershell
.\venv\Scripts\python.exe -m pytest tests -q
```
