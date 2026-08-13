# Deal Expenses Automation

Local Streamlit application for refreshing a deal expense master workbook from BSNY and SanCap reports.

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

Upload the master workbook, BSNY Concur report, and SanCap expense report. Enter the desired `.xlsx` output name and download the validated result.

The application uses Microsoft Excel automation to retain workbook tables, formatting, and helper formulas. Uploaded files are kept only in a temporary directory during the refresh; the completed workbook is retained in the browser session for download.

## Test

```powershell
.\venv\Scripts\python.exe -m pytest tests -q
```
