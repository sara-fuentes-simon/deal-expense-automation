# Deal Expenses Automation

Local Streamlit application for refreshing the Concur and SAP sections of one deal expense master workbook.

## Requirements

- Windows
- Microsoft Excel installed
- Python 3.11 or newer
- The uploaded workbooks must not be open in Excel while the refresh runs

## Install

```powershell
.\venv\Scripts\python.exe -m pip install -r requirements.txt
```

If `win32com` is unavailable after installation:

```powershell
.\venv\Scripts\python.exe -m pip install pywin32
```

## Run

```powershell
.\venv\Scripts\streamlit.exe run app.py
```

Upload the master workbook, BSNY and SanCap Concur reports, and BSNY and SanCap SAP reports. Enter the desired `.xlsx` output name and download one combined refreshed workbook.

The application refreshes Concur first, then appends validated SAP rows to that same temporary workbook. The SAP refresh requires the master to contain `SAP Report` or `SAP Invoices Report`, an `In Scope CCs` sheet, and the expected SAP master headers. Uploaded files are kept only in a temporary directory during the refresh; the completed workbook is retained in the browser session for download.

## Test

```powershell
.\venv\Scripts\python.exe -m pytest tests -q
```
