import win32com.client
import os

filepath = os.path.abspath(r"samples\Expense Report Relating to deals 01-01-26 to 6-30-26_V8 - COMBINED.xlsx")
excel = win32com.client.Dispatch("Excel.Application")
try:
    print(f"Opening workbook: {filepath}")
    wb = excel.Workbooks.Open(filepath, ReadOnly=True)
    sheets = ["Concur Analysis - PIVOTS", "SAP Invoices Analysis - PIVOTS"]
    for s_name in sheets:
        try:
            ws = wb.Sheets(s_name)
            print(f"\nSheet: {s_name}")
            pivots = ws.PivotTables()
            if pivots.Count == 0:
                print("  No Pivot Tables found.")
            for i in range(1, pivots.Count + 1):
                pt = pivots.Item(i)
                addr = pt.TableRange2.Address
                row_start = pt.TableRange2.Row
                col_start = pt.TableRange2.Column
                row_count = pt.TableRange2.Rows.Count
                col_count = pt.TableRange2.Columns.Count
                print(f"  Pivot Name: {pt.Name}")
                print(f"    TableRange2.Address: {addr}")
                print(f"    Start Row: {row_start}, Column: {col_start}")
                print(f"    Row Count: {row_count}, Column Count: {col_count}")
        except Exception as sheet_err:
            print(f"\nError reading sheet {s_name}: {sheet_err}")
finally:
    try:
        wb.Close(SaveChanges=False)
    except Exception:
        pass
    excel.Quit()
