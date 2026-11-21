"""Read the calculated workbook back: report any Excel errors and print the key outputs."""
import sys
sys.stdout.reconfigure(encoding="utf-8")
from pathlib import Path

from openpyxl import load_workbook

BOOK = Path(__file__).resolve().parent.parent / "model" / "CEG-valuation.xlsx"
wb = load_workbook(BOOK, data_only=True)
formulas = load_workbook(BOOK)

errors, empty = [], []
for ws in wb.worksheets:
    for row in formulas[ws.title].iter_rows():
        for cell in row:
            v = ws[cell.coordinate].value
            if isinstance(v, str) and v.startswith("#"):
                errors.append(f"{ws.title}!{cell.coordinate} {v}")
            if isinstance(cell.value, str) and cell.value.startswith("=") and v is None:
                empty.append(f"{ws.title}!{cell.coordinate}")
print(f"errors: {len(errors)}", errors[:20])
print(f"formulas with no value: {len(empty)}", empty[:10])

v = wb["Valuation"]
i = wb["Inputs"]
c = wb["Comps"]
s = wb["Pitch scorecard"]
print(f"beta raw {i['C39'].value:.2f} adj {i['C40'].value:.2f}  rf {i['C38'].value:.2%}  ke {i['C42'].value:.2%}")
print("EPS path:", [round(v.cell(16, col).value, 2) for col in range(3, 15)])
print(f"DCF {v['C29'].value:.2f}  TV share {v['C30'].value:.0%}  DCF P/E {v['C31'].value:.1f}x  "
      f"12m target {v['C32'].value:.2f}  price {v['C33'].value}  P/E now {v['C36'].value:.1f}x")
print("scenarios (DCF, target, EPS27, P/E, upside):")
for r in (73, 74, 75):
    print("  ", v[f"B{r}"].value, [round(v.cell(r, col).value, 3) for col in range(4, 9)])
print("break-even growth:", v["C95"].value)
print("ke x g table:")
for r in range(41, 47):
    print("  ", [v.cell(r, col).value if isinstance(v.cell(r, col).value, str) else round(v.cell(r, col).value or 0, 3) for col in range(3, 9)])
print("comps:")
for r in range(5, 12):
    print("  ", c[f"B{r}"].value, [c.cell(r, col).value if isinstance(c.cell(r, col).value, str) or c.cell(r, col).value is None else round(c.cell(r, col).value, 2) for col in range(3, 16)])
for r in range(1, 60):
    if c[f"B{r}"].value and "peer multiples" in str(c[f"B{r}"].value).lower():
        print("  ", c[f"B{r+1}"].value, round(c[f"C{r+1}"].value, 2), "|", c[f"B{r+2}"].value, round(c[f"C{r+2}"].value, 2))
print("scorecard:", [(s[f"B{r}"].value, round(s[f"E{r}"].value, 4)) for r in range(5, 11) if s[f"E{r}"].value is not None])
print("attribution:", [(s[f"B{r}"].value, s[f"C{r}"].value, s[f"D{r}"].value, s[f"E{r}"].value) for r in range(14, 18)])
sm = wb["Summary"]
for r in range(16, 23):
    print("FF", sm[f"B{r}"].value, sm[f"C{r}"].value, sm[f"D{r}"].value, sm[f"F{r}"].value)
for r in range(44, 49):
    print(sm[f"B{r}"].value)
