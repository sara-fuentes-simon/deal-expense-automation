import os
import win32com.client

abs_path = os.path.abspath(r"samples/Expense Report Relating to deals 01-01-26 to 6-30-26_V8 - COMBINED.xlsx")
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    print(f"Opening workbook in read-only mode: {abs_path}")
    wb = excel.Workbooks.Open(abs_path, ReadOnly=True)
    sheets_to_check = ["Concur Analysis - PIVOTS", "SAP Invoices Analysis - PIVOTS"]
    
    for sheet_name in sheets_to_check:
        try:
            ws = wb.Sheets(sheet_name)
        except Exception as e:
            print(f"Sheet {sheet_name} not found: {e}")
            continue
            
        pivots = ws.PivotTables()
        if pivots.Count == 0:
            print(f"No PivotTables found on sheet: {sheet_name}")
            continue
            
        # Get first PivotTable
        pt = pivots.Item(1)
        pt_name = pt.Name
        tr2 = pt.TableRange2
        addr = tr2.Address
        
        print("\n" + "="*80)
        print(f"Sheet: {sheet_name} | PivotTable: {pt_name} | Address: {addr}")
        print("="*80)
        
        rows_count = tr2.Rows.Count
        cols_count = tr2.Columns.Count
        
        vals = tr2.Value
        
        for r_idx in range(1, rows_count + 1):
            row_range = tr2.Rows(r_idx)
            first_cell = row_range.Cells(1, 1)
            is_bold = first_cell.Font.Bold
            
            if isinstance(vals, tuple):
                if len(vals) > 0 and isinstance(vals[0], tuple):
                    row_vals = vals[r_idx - 1]
                else:
                    row_vals = vals if r_idx == 1 else []
            else:
                row_vals = [vals] if r_idx == 1 else []
                
            row_vals_clean = ["" if x is None else x for x in row_vals]
            print(f"Row {r_idx:02d} | Bold={is_bold} | {row_vals_clean}")
            
except Exception as e:
    print(f"An error occurred: {e}")
finally:
    try:
        wb.Close(SaveChanges=False)
    except Exception as e:
        pass
    try:
        excel.Quit()
    except Exception as e:
        pass
