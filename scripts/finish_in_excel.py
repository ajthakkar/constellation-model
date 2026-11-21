"""Open the workbook in Excel, add the sensitivity Data Tables, calculate, and save.

openpyxl can write formulas but can't create Excel's What-If Data Tables or calculate
anything, so this step drives the installed copy of Excel (Windows only).
Run after build_model.py: python scripts/finish_in_excel.py
"""
import json
from pathlib import Path

import pythoncom
import win32com.client

ROOT = Path(__file__).resolve().parent.parent
BOOK = ROOT / "model" / "CEG-valuation.xlsx"


def main():
    tables = json.loads((ROOT / "model" / "tables.json").read_text())
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(str(BOOK))
        ws = wb.Worksheets("Valuation")
        for name, (rng, row_input, col_input) in tables.items():
            target = ws.Range(rng)
            if row_input and col_input:
                target.Table(ws.Range(row_input), ws.Range(col_input))
            else:
                target.Table(pythoncom.Empty, ws.Range(col_input))
            print(f"data table {name}: {rng}")
        xl.CalculateFull()
        wb.Worksheets("Summary").Activate()
        wb.Worksheets("Summary").Range("A1").Select()
        wb.Save()
        wb.Close(SaveChanges=False)
    finally:
        xl.Quit()
    print(f"calculated and saved {BOOK}")


if __name__ == "__main__":
    main()
