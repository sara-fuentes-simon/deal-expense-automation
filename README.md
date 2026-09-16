# Deal Expenses Automation

Local Streamlit application for refreshing the Concur, SAP, and Uber sections of one deal expense master workbook.

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
2. Upload BSNY and SanCap reports for both Concur and SAP, plus the Uber report.
3. Enter the reporting year and the desired master and LOB output filenames.
4. Select **Refresh combined master workbook**.
5. Review progress, validation details, Uber Final Cost Center coverage, and the Concur/SAP pivot previews, then download the ZIP archive containing the master workbook and one workbook for each LOB.

The application refreshes Concur first, appends validated SAP rows, then replaces Uber Report raw data in that same temporary workbook. Uber data maps by header, uses HR `Employee ID` before a Full Name fallback for Cost Center, and reports rather than blocks rows without a Final Cost Center. It refreshes the master workbook's pivot caches after Uber processing, then creates refreshed copies for ECM, DCM, M&A, SPG, Lev Finance, and Struc Finance. Each LOB copy keeps only matching LOB rows in `Concur Report` and `SAP Report` (or `SAP Invoices Report`) before its pivot refresh. The SAP refresh requires the master workbook to contain `SAP Report` or `SAP Invoices Report`, an `In Scope CCs` sheet, and the expected SAP master headers.

Uploaded files are stored only in a temporary directory while the refresh runs. The completed workbook is retained in the browser session until the page is refreshed or a new refresh starts.

## Test

```powershell
.\venv\Scripts\python.exe -m pytest tests -q
```
