"""Build the Constellation Energy valuation workbook (model/CEG-valuation.xlsx).

Every output is an Excel formula; the only typed-in numbers are inputs, each with its source.
Excel itself then adds the sensitivity Data Tables and calculates the file (scripts/finish_in_excel.py).

Run: python scripts/build_model.py && python scripts/finish_in_excel.py
"""
import json
from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
OUT = ROOT / "model" / "CEG-valuation.xlsx"

VAL_DATE = date(2026, 11, 20)
RELEASE_Q3 = "Constellation Q3 2026 earnings release (8-K, Nov 9, 2026), https://www.sec.gov/Archives/edgar/data/1868275/000186827526000116/ceg-20261109991.htm"
RELEASE_Q2 = "Constellation Q2 2026 earnings release (8-K, Aug 6, 2026), https://www.sec.gov/Archives/edgar/data/1868275/000186827526000097/ceg-20260806991.htm"
RELEASE_Q4_25 = "Constellation Q4 2025 earnings release (8-K, Feb 24, 2026), https://www.sec.gov/Archives/edgar/data/1868275/000186827526000029/ceg-20260224991.htm"
RELEASE_Q4_24 = "Constellation Q4 2024 earnings release (8-K, Feb 18, 2025), https://www.sec.gov/Archives/edgar/data/1868275/000186827525000020/ceg-20250218991.htm"
YAHOO = "Yahoo Finance daily close, downloaded by scripts/fetch_data.py"

# ── Styles ──────────────────────────────────────────────────────────────────────
FONT = "Arial"
BLUE = Font(name=FONT, size=10, color="0000FF")
BLACK = Font(name=FONT, size=10, color="000000")
GREEN = Font(name=FONT, size=10, color="008000")
BOLD = Font(name=FONT, size=10, bold=True)
TITLE = Font(name=FONT, size=14, bold=True)
SUB = Font(name=FONT, size=10, italic=True, color="595959")
HEAD = Font(name=FONT, size=10, bold=True, color="FFFFFF")
HEAD_FILL = PatternFill("solid", fgColor="1F3864")
SECTION_FILL = PatternFill("solid", fgColor="D9E1F2")
KEY_FILL = PatternFill("solid", fgColor="FFFF00")
OVERRIDE_FILL = PatternFill("solid", fgColor="F2F2F2")
OUT_FILL = PatternFill("solid", fgColor="E2EFDA")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
TOP = Border(top=Side(style="thin", color="000000"))

USD = '$#,##0.00;($#,##0.00);"-"'
USD0 = '$#,##0;($#,##0);"-"'
MM = '#,##0;(#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'
MULT = '0.0"x";(0.0"x");"-"'
NUM2 = '0.00'
DATE_FMT = "mmm d, yyyy"


def style(cell, font=BLACK, fmt=None, fill=None, bold=False, align=None):
    cell.font = Font(name=font.name, size=font.size, color=font.color, bold=bold or font.bold, italic=font.italic)
    if fmt:
        cell.number_format = fmt
    if fill:
        cell.fill = fill
    if align:
        cell.alignment = Alignment(horizontal=align)
    return cell


def note(cell, text):
    cell.comment = Comment(text, "Model")
    cell.comment.width = 360
    cell.comment.height = 110


def header_row(ws, row, labels, start_col=2):
    for i, label in enumerate(labels):
        c = ws.cell(row=row, column=start_col + i, value=label)
        style(c, HEAD, fill=HEAD_FILL, align="center" if i else None)


def section(ws, row, text, width=8, start_col=2):
    for col in range(start_col, start_col + width):
        ws.cell(row=row, column=col).fill = SECTION_FILL
    style(ws.cell(row=row, column=start_col, value=text), BOLD, fill=SECTION_FILL)


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


# ── Data from SEC XBRL (peers) ──────────────────────────────────────────────────
def months(start, end):
    s, e = start.split("-"), end.split("-")
    return (int(e[0]) - int(s[0])) * 12 + int(e[1]) - int(s[1]) + 1


def xbrl(ticker, concept, kind, end=None):
    """Latest-filed value of a concept. kind: 'FY' (12 months ending end), '9M', 'Q', or 'bal' (instant)."""
    facts = json.loads((DATA / "sec" / f"{ticker}.json").read_text())["facts"]["us-gaap"]
    rows = [r for r in list(facts[concept]["units"].values())[0] if r.get("form") in ("10-K", "10-Q")]
    want = {"FY": 12, "9M": 9, "Q": 3}
    best = {}
    for r in rows:
        if kind == "bal":
            if "start" in r:
                continue
        elif "start" not in r or months(r["start"], r["end"]) != want[kind]:
            continue
        best[r["end"]] = r  # later filings overwrite earlier ones
    if end:
        r = best[end]
    else:
        r = best[max(best)]
    return r["val"], f'SEC XBRL {ticker} {concept}, period ending {r["end"]}, from {r["form"]} filed {r["filed"]}'


def price(ticker, day):
    d = pd.read_csv(DATA / "prices" / f"{ticker}.csv", index_col="date")
    return float(d.loc[day, "close"]), float(d.loc[day, "adjclose"])


def main():
    wb = Workbook()
    summary = wb.active
    summary.title = "Summary"
    inputs = wb.create_sheet("Inputs")
    val = wb.create_sheet("Valuation")
    comps = wb.create_sheet("Comps")
    card = wb.create_sheet("Pitch scorecard")
    market = wb.create_sheet("Market data")
    sources = wb.create_sheet("Sources")
    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False
    source_rows = []

    def inp(ws, ref, value, fmt, src, key=False):
        c = ws[ref]
        c.value = value
        style(c, BLUE, fmt, KEY_FILL if key else None)
        note(c, f"Source: {src}")
        source_rows.append((ws.title, ref, src))
        return c

    # ── Inputs ──────────────────────────────────────────────────────────────────
    ws = inputs
    widths(ws, {"A": 2, "B": 46, "C": 14, "D": 62})
    style(ws["B1"], TITLE).value = "Inputs"
    style(ws["B2"], SUB).value = "Blue = typed-in input with its source (hover a cell for the full citation). Black = formula. Green = link to another sheet."

    section(ws, 4, "Dates and prices", 3)
    ws["B5"] = "Valuation date"
    inp(ws, "C5", VAL_DATE, DATE_FMT, "Last close in the downloaded price data")
    ws["B6"] = "CEG share price at valuation date ($)"
    inp(ws, "C6", price("CEG", "2026-11-20")[0], USD, f"{YAHOO}, Nov 20, 2026")
    ws["B7"] = "Pitch date"
    inp(ws, "C7", date(2026, 9, 10), DATE_FMT, "My Constellation Energy stock pitch (Sep 2026)")
    ws["B8"] = "Price in the pitch ($)"
    inp(ws, "C8", 285.93, USD, "My pitch: 'The 9/10/26 closing price of $285.93'. Yahoo's close that day is $285.97.")
    ws["B9"] = "Pitch 12-month price target ($)"
    inp(ws, "C9", 355, USD, "My pitch: '12 month target of $355'")
    ws["B10"] = "52-week low ($)"
    lo52 = float(pd.read_csv(DATA / "prices" / "CEG.csv", index_col="date").loc["2025-11-21":, "close"].min())
    hi52 = float(pd.read_csv(DATA / "prices" / "CEG.csv", index_col="date").loc["2025-11-21":, "close"].max())
    inp(ws, "C10", round(lo52, 2), USD, f"{YAHOO}, lowest close Nov 21, 2025 to Nov 20, 2026")
    ws["B11"] = "52-week high ($)"
    inp(ws, "C11", round(hi52, 2), USD, f"{YAHOO}, highest close Nov 21, 2025 to Nov 20, 2026")

    section(ws, 13, "Balance sheet at Sep 30, 2026 ($ millions)", 3)
    rows = [
        ("Diluted shares (millions)", 354, MM, f"{RELEASE_Q3}: diluted shares, three months ended Sep 30, 2026"),
        ("Cash and cash equivalents", 603, MM, f"{RELEASE_Q3}: balance sheet"),
        ("Short-term borrowings", 3989, MM, f"{RELEASE_Q3}: balance sheet"),
        ("Long-term debt due within one year", 359, MM, f"{RELEASE_Q3}: balance sheet"),
        ("Long-term debt", 19039, MM, f"{RELEASE_Q3}: balance sheet"),
        ("Noncontrolling interests", 327, MM, f"{RELEASE_Q3}: balance sheet"),
    ]
    for i, (label, v, fmt, src) in enumerate(rows):
        ws[f"B{14 + i}"] = label
        inp(ws, f"C{14 + i}", v, fmt, src)
    ws["B20"] = "Total debt"
    style(ws["C20"], fmt=MM).value = "=C16+C17+C18"
    ws["B21"] = "Net debt"
    style(ws["C21"], fmt=MM).value = "=C20-C15"
    ws["B22"] = "Market capitalization"
    style(ws["C22"], fmt=MM).value = "=C6*C14"
    ws["B23"] = "Enterprise value"
    style(ws["C23"], fmt=MM).value = "=C22+C21+C19"

    section(ws, 25, "Adjusted (non-GAAP) operating EPS ($ per share)", 3)
    eps_rows = [
        ("2024 actual", 8.67, RELEASE_Q4_24),
        ("2025 actual", 9.39, RELEASE_Q4_25),
        ("2026 guidance at pitch, low", 11.50, f"{RELEASE_Q2}: 'Raising full-year ... guidance range to $11.50 - $12.50'"),
        ("2026 guidance at pitch, high", 12.50, RELEASE_Q2),
        ("2026 guidance now, low", 12.50, f"{RELEASE_Q3}: 'narrowing full-year ... guidance range to $12.50 - $12.80'"),
        ("2026 guidance now, high", 12.80, RELEASE_Q3),
    ]
    for i, (label, v, src) in enumerate(eps_rows):
        ws[f"B{26 + i}"] = label
        inp(ws, f"C{26 + i}", v, USD, src)
    ws["B32"] = "2026 guidance midpoint at pitch"
    style(ws["C32"], fmt=USD).value = "=AVERAGE(C28:C29)"
    ws["B33"] = "2026 guidance midpoint now"
    style(ws["C33"], fmt=USD).value = "=AVERAGE(C30:C31)"
    ws["B34"] = "Management's target for adjusted EPS growth through 2029"
    inp(ws, "C34", 0.20, PCT, "Q2 2026 investor materials, as cited in my pitch ('targets 20%+ annual growth rate through 2029'). Not re-checked against the presentation.")
    ws["B35"] = "2026 dividend per share ($)"
    style(ws["C35"], fmt=USD).value = "=0.4265*4"
    note(ws["C35"], f"Quarterly dividend of $0.4265 declared Feb 2026, times 4. Source: {RELEASE_Q4_25}")

    section(ws, 37, "Cost of equity (CAPM)", 3)
    ws["B38"] = "Risk-free rate: 10-year Treasury yield"
    rf = pd.read_csv(DATA / "dgs10.csv").dropna()
    rf = rf[rf.iloc[:, 1] != "."]
    rf_last = rf.iloc[-1]
    inp(ws, "C38", float(rf_last.iloc[1]) / 100, PCT, f"FRED series DGS10, {rf_last.iloc[0]}, https://fred.stlouisfed.org/series/DGS10")
    ws["B39"] = "Beta, raw (2 years of weekly returns vs. S&P 500)"
    style(ws["C39"], GREEN, NUM2).value = "='Market data'!H5"
    ws["B40"] = "Beta, adjusted toward 1 (Blume: 0.67 × raw + 0.33)"
    style(ws["C40"], fmt=NUM2).value = "=0.67*C39+0.33"
    ws["B41"] = "Equity risk premium"
    inp(ws, "C41", 0.045, PCT, "Assumption. Commonly used values run from about 4% to 6%; 4.5% is near the middle.", key=True)
    ws["B42"] = "Cost of equity"
    style(ws["C42"], fmt=PCT, bold=True).value = "=C38+C40*C41"
    for r in range(5, 43):
        if ws[f"B{r}"].value and not ws[f"B{r}"].font.bold:
            style(ws[f"B{r}"])

    # ── Market data: weekly returns for beta ───────────────────────────────────
    ws = market
    widths(ws, {"A": 2, "B": 13, "C": 13, "D": 13, "E": 12, "F": 12, "G": 3, "H": 30})
    style(ws["B1"], TITLE).value = "Market data"
    style(ws["B2"], SUB).value = "Weekly closes adjusted for dividends and splits (Friday), from Yahoo Finance. Used to estimate beta."
    header_row(ws, 4, ["Week ending", "CEG", "S&P 500 (SPY)", "CEG return", "SPY return"])
    ceg = pd.read_csv(DATA / "prices" / "CEG.csv", parse_dates=["date"]).set_index("date").adjclose
    spy = pd.read_csv(DATA / "prices" / "SPY.csv", parse_dates=["date"]).set_index("date").adjclose
    weekly = pd.concat([ceg, spy], axis=1, keys=["ceg", "spy"], sort=True).dropna().resample("W-FRI").last()
    weekly = weekly.loc["2024-11-15":"2026-11-20"]
    first = 5
    for i, (d, row) in enumerate(weekly.iterrows()):
        r = first + i
        style(ws.cell(r, 2, d.date()), BLUE, DATE_FMT)
        style(ws.cell(r, 3, round(float(row.ceg), 4)), BLUE, USD)
        style(ws.cell(r, 4, round(float(row.spy), 4)), BLUE, USD)
        if i:
            style(ws.cell(r, 5, f"=C{r}/C{r - 1}-1"), fmt=PCT)
            style(ws.cell(r, 6, f"=D{r}/D{r - 1}-1"), fmt=PCT)
    last = first + len(weekly) - 1
    source_rows.append(("Market data", f"B{first}:D{last}", YAHOO))
    style(ws["H4"], HEAD, fill=HEAD_FILL).value = "Beta"
    style(ws["H5"], fmt=NUM2, bold=True, fill=OUT_FILL).value = f"=SLOPE(E{first + 1}:E{last},F{first + 1}:F{last})"
    style(ws["H6"], SUB).value = f"=COUNT(E{first + 1}:E{last})&\" weekly returns\""
    style(ws["H7"], SUB).value = "SLOPE = covariance with the market ÷ market variance"

    # ── Valuation ──────────────────────────────────────────────────────────────
    ws = val
    widths(ws, {"A": 2, "B": 52, **{get_column_letter(c): 13 for c in range(3, 15)}, "O": 3, "P": 40})
    style(ws["B1"], TITLE).value = "Valuation"
    style(ws["B2"], SUB).value = "Dollars per share unless noted. Pick a scenario in C5. Gray cells in column H override one assumption; leave them blank (the sensitivity tables use them)."

    section(ws, 4, "Scenarios", 7)
    header_row(ws, 4, ["Scenarios", "Bear", "Base", "Bull", "Live", "", "Override"])
    ws["B5"] = "Scenario in use (1 = Bear, 2 = Base, 3 = Bull)"
    style(ws["C5"], BLUE, "0", KEY_FILL).value = 2
    scen = [
        (6, "Adjusted EPS growth, 2027–2029", (0.08, 0.15, 0.20), PCT,
         "Bull = management's 20% target. Base trims it to 15%. Bear assumes AI power demand disappoints."),
        (7, "Cash conversion (equity cash flow ÷ adjusted EPS)", (0.75, 0.85, 0.95), PCT,
         "Share of adjusted earnings left for shareholders after the capital spending needed to deliver the growth. Nuclear fuel and maintenance roughly offset depreciation, so conversion is high; growth projects (uprates, the Crane restart) take a share."),
        (8, "Long-run growth after 2036", (0.02, 0.03, 0.035), PCT,
         "Growth fades in a straight line from the 2029 rate to this rate by 2036, then continues forever."),
        (9, "Target P/E on 2027 EPS (12-month target)", (15.0, 19.0, 24.0), MULT,
         "Bear: Vistra/NRG territory. Base: a modest premium for contracted nuclear output. Bull: about where CEG traded at the time of the pitch."),
    ]
    for r, label, (b, m, u), fmt, why in scen:
        ws[f"B{r}"] = label
        for col, v in zip("CDE", (b, m, u)):
            style(ws[f"{col}{r}"], BLUE, fmt, KEY_FILL).value = v
        style(ws[f"F{r}"], fmt=fmt, bold=True).value = f"=IF(ISBLANK($H{r}),INDEX(C{r}:E{r},$C$5),$H{r})"
        style(ws[f"H{r}"], BLUE, fmt, OVERRIDE_FILL)
        style(ws[f"P{r}"], SUB).value = why
    ws["B10"] = "Cost of equity"
    for col in "CDE":
        style(ws[f"{col}10"], GREEN, PCT).value = "=Inputs!$C$42"
    style(ws["F10"], fmt=PCT, bold=True).value = "=IF(ISBLANK($H10),Inputs!$C$42,$H10)"
    style(ws["H10"], BLUE, PCT, OVERRIDE_FILL)
    style(ws["P10"], SUB).value = "From CAPM on the Inputs sheet; the same in every scenario."

    # Projection: columns C (2025A) .. N (2036E)
    section(ws, 12, "Equity cash flow projection", 13)
    years = list(range(2025, 2037))
    cols = [get_column_letter(3 + i) for i in range(len(years))]
    style(ws["B13"], BOLD).value = "Fiscal year"
    for col, y in zip(cols, years):
        style(ws[f"{col}13"], BOLD, align="right").value = f"{y}{'A' if y == 2025 else 'E'}"
        style(ws[f"{col}14"], SUB, "0").value = y
    style(ws["B14"], SUB).value = "Year (number, used by formulas)"
    ws["B15"] = "Adjusted EPS growth"
    ws["B16"] = "Adjusted EPS ($)"
    ws["B17"] = "Equity cash flow per share ($)"
    ws["B18"] = "Year end"
    ws["B19"] = "Years from valuation date"
    ws["B20"] = "Discount factor"
    ws["B21"] = "Present value of cash flow ($)"
    style(ws["C16"], GREEN, USD).value = "=Inputs!C27"
    style(ws["D16"], GREEN, USD).value = "=Inputs!C33"
    style(ws["D15"], fmt=PCT).value = "=D16/C16-1"
    for col in cols[2:]:  # 2027E onward
        prev = get_column_letter(ws[f"{col}13"].column - 1)
        style(ws[f"{col}15"], fmt=PCT).value = f"=IF({col}14<=2029,$F$6,$F$6+($F$8-$F$6)*({col}14-2029)/7)"
        style(ws[f"{col}16"], fmt=USD).value = f"={prev}16*(1+{col}15)"
        style(ws[f"{col}17"], fmt=USD).value = f"={col}16*$F$7"
        style(ws[f"{col}18"], fmt="mm/dd/yy").value = f"=DATE({col}14,12,31)"
        style(ws[f"{col}19"], fmt=NUM2).value = f"=({col}18-Inputs!$C$5)/365"
        style(ws[f"{col}20"], fmt="0.000").value = f"=1/(1+$F$10)^{col}19"
        style(ws[f"{col}21"], fmt=USD).value = f"={col}17*{col}20"
    note(ws["D16"], "2026 = midpoint of management's current guidance. Cash flows start in 2027; the few weeks left in 2026 are left out, which slightly understates value.")
    ws["B23"] = "Terminal value at end of 2036 ($)"
    style(ws["N23"], fmt=USD).value = "=IF($F$10>$F$8,N17*(1+$F$8)/($F$10-$F$8),NA())"
    ws["B24"] = "Present value of terminal value ($)"
    style(ws["N24"], fmt=USD).value = "=N23*N20"
    note(ws["N23"], "Gordon growth: next year's cash flow ÷ (cost of equity − long-run growth).")

    section(ws, 26, "Value per share", 13)
    out = [
        (27, "DCF: present value of 2027–2036 cash flows", "=SUM(E21:N21)", USD),
        (28, "DCF: present value of terminal value", "=N24", USD),
        (29, "DCF value per share", "=C27+C28", USD),
        (30, "Terminal value share of DCF value", "=C28/C29", PCT),
        (31, "DCF value as a P/E on 2027 EPS", "=C29/E16", MULT),
        (32, "12-month target: 2027 EPS × target P/E", "=E16*$F$9", USD),
        (33, "Current share price", "=Inputs!C6", USD),
        (34, "Upside to DCF value", "=C29/C33-1", PCT),
        (35, "Upside to 12-month target", "=C32/C33-1", PCT),
        (36, "Current price as a P/E on 2027 EPS", "=C33/E16", MULT),
    ]
    for r, label, f, fmt in out:
        ws[f"B{r}"] = label
        c = style(ws[f"C{r}"], GREEN if "Inputs!" in f else BLACK, fmt, OUT_FILL if r in (29, 32) else None, bold=r in (29, 32))
        c.value = f

    # Sensitivity tables (cells filled by Excel Data Tables in finish_in_excel.py)
    def table_frame(top, left_label, top_label, rows, cols, row_fmt, col_fmt, corner, title):
        section(ws, top - 2, title, 9)
        style(ws.cell(top - 1, 4, top_label), SUB)
        style(ws.cell(top + 1, 2, left_label), SUB)
        c = ws.cell(top, 3, corner)
        style(c, fmt=USD, bold=True)
        for j, v in enumerate(cols):
            style(ws.cell(top, 4 + j, v), BLUE, col_fmt, align="center")
        for i, v in enumerate(rows):
            style(ws.cell(top + 1 + i, 3, v), BLUE, row_fmt, align="center")
            for j in range(len(cols)):
                ws.cell(top + 1 + i, 4 + j).number_format = USD0
                ws.cell(top + 1 + i, 4 + j).border = BOX
        rng = f"C{top}:{get_column_letter(3 + len(cols))}{top + len(rows)}"
        body = f"D{top + 1}:{get_column_letter(3 + len(cols))}{top + len(rows)}"
        ws.conditional_formatting.add(body, CellIsRule(operator="greaterThanOrEqual", formula=["$C$33"], fill=OUT_FILL))
        return rng

    tables = {}
    tables["ke_g"] = (table_frame(41, "Cost of equity ↓", "Long-run growth after 2036 →",
                                  [0.09, 0.10, 0.11, 0.12, 0.13], [0.02, 0.025, 0.03, 0.035, 0.04], PCT, PCT, "=C29",
                                  "DCF value per share: cost of equity vs. long-run growth (green = at or above today's price)"), "H8", "H10")
    tables["g_conv"] = (table_frame(51, "EPS growth 2027–29 ↓", "Cash conversion →",
                                    [0.05, 0.10, 0.15, 0.20, 0.25], [0.65, 0.75, 0.85, 0.95, 1.0], PCT, PCT, "=C29",
                                    "DCF value per share: near-term growth vs. cash conversion"), "H7", "H6")
    tables["g_pe"] = (table_frame(61, "EPS growth 2027–29 ↓", "Target P/E on 2027 EPS →",
                                  [0.05, 0.10, 0.15, 0.20, 0.25], [14.0, 17.0, 20.0, 23.0, 26.0], PCT, MULT, "=C32",
                                  "12-month target: near-term growth vs. P/E"), "H9", "H6")

    # Scenario summary (one-variable data table on the scenario switch)
    section(ws, 70, "All three scenarios side by side", 9)
    header_row(ws, 71, ["Scenario", "", "DCF value", "12-month target", "2027 EPS", "DCF as P/E", "Upside (DCF)"])
    for j, f in enumerate(["=C29", "=C32", "=E16", "=C31", "=C34"]):
        style(ws.cell(72, 4 + j, f), SUB, [USD, USD, USD, MULT, PCT][j])
    style(ws["B72"], SUB).value = "(formula row read by the table)"
    for i, (n, name) in enumerate([(1, "Bear"), (2, "Base"), (3, "Bull")]):
        r = 73 + i
        ws[f"B{r}"] = name
        style(ws[f"C{r}"], BLUE, "0", align="center").value = n
        for j, fmt in enumerate([USD, USD, USD, MULT, PCT]):
            ws.cell(r, 4 + j).number_format = fmt
            ws.cell(r, 4 + j).border = BOX
    tables["scen"] = ("C72:H75", None, "C5")

    # Break-even growth (one-variable data table on growth)
    section(ws, 78, "What growth does today's price assume? DCF value by 2027–29 EPS growth", 9)
    header_row(ws, 79, ["EPS growth 2027–29", "", "DCF value", "vs. price"])
    style(ws["D80"], SUB, USD).value = "=C29"
    style(ws["E80"], SUB, PCT).value = "=C29/C33-1"
    style(ws["B80"], SUB).value = "(formula row read by the table)"
    grid = [round(0.0 + 0.025 * i, 3) for i in range(13)]
    for i, g in enumerate(grid):
        r = 81 + i
        style(ws[f"C{r}"], BLUE, PCT, align="center").value = g
        ws[f"D{r}"].number_format = USD
        ws[f"E{r}"].number_format = PCT
        ws[f"D{r}"].border = BOX
        ws[f"E{r}"].border = BOX
    ws.conditional_formatting.add(f"D81:D{80 + len(grid)}", CellIsRule(operator="greaterThanOrEqual", formula=["$C$33"], fill=OUT_FILL))
    tables["breakeven"] = (f"C80:E{80 + len(grid)}", None, "H6")
    style(ws["B95"]).value = "Break-even growth (where the DCF value first reaches today's price)"
    style(ws["C95"], fmt=PCT, bold=True, fill=OUT_FILL).value = (
        f'=IFERROR(INDEX(C81:C{80 + len(grid)},MATCH(TRUE,INDEX(D81:D{80 + len(grid)}>=C33,0),0)),"above 30%")'
    )
    note(ws["C95"], "The lowest growth rate in the table above whose DCF value is at least today's price, holding the other assumptions at the live scenario.")
    for r in range(5, 96):
        b = ws[f"B{r}"]
        if b.value and b.font.color is None:
            style(b)

    # ── Comps ────────────────────────────────────────────────────────────────
    ws = comps
    widths(ws, {"A": 2, "B": 30, **{get_column_letter(c): 12 for c in range(3, 18)}})
    style(ws["B1"], TITLE).value = "Comparable companies"
    style(ws["B2"], SUB).value = ("Trailing 12 months to Sep 30, 2026, on reported (GAAP) numbers. GAAP results for power producers swing with "
                                  "hedge gains and losses, so treat these as a rough cross-check, not a valuation.")
    header_row(ws, 4, ["Company", "Price ($)", "Diluted shares (mm)", "Market cap ($mm)", "Total debt ($mm)", "Cash ($mm)",
                       "Preferred + NCI ($mm)", "Enterprise value ($mm)", "LTM revenue ($mm)", "LTM EBITDA ($mm)",
                       "LTM EPS ($)", "EV / EBITDA", "P / E", "EV / revenue"])
    companies = ["CEG", "VST", "NRG", "TLN"]
    names = {"CEG": "Constellation (CEG)", "VST": "Vistra (VST)", "NRG": "NRG Energy (NRG)", "TLN": "Talen Energy (TLN)"}

    # Raw inputs block: FY2025, 9M 2026, 9M 2025 for each metric
    raw_top = 14
    section(ws, raw_top - 1, "Raw inputs ($mm except per share): LTM = FY2025 + 9 months 2026 − 9 months 2025", 15)
    header_row(ws, raw_top, ["Company / metric", "FY2025", "9M 2026", "9M 2025", "LTM"])
    metrics = [("Revenue", "Revenues"), ("Operating income", "OperatingIncomeLoss"),
               ("Depreciation & amortization", None), ("Diluted EPS ($)", "EarningsPerShareDiluted")]
    ceg_9m = {  # Q3 2026 release (9 months ended Sep 30), not yet in the SEC's XBRL feed
        "Revenue": (28176, 19459), "Operating income": (5404, 2488), "Depreciation & amortization": (1340, 743),
        "Diluted EPS ($)": (11.06, 6.02),
    }
    ceg_fy = {"Revenue": ("Revenues", None), "Operating income": ("OperatingIncomeLoss", None),
              "Depreciation & amortization": ("CostOfGoodsAndServicesSoldDepreciationAndAmortization", None),
              "Diluted EPS ($)": ("EarningsPerShareDiluted", None)}
    ltm_ref = {}
    r = raw_top + 1
    for t in companies:
        style(ws[f"B{r}"], BOLD).value = names[t]
        r += 1
        for label, concept in metrics:
            ws[f"B{r}"] = f"   {label}"
            style(ws[f"B{r}"])
            fmt = USD if "EPS" in label else MM
            scale = 1 if "EPS" in label else 1e6
            if t == "CEG":
                fy_v, fy_src = xbrl(t, ceg_fy[label][0], "FY", "2025-12-31")
                inp(ws, f"C{r}", round(fy_v / scale, 2), fmt, fy_src)
                inp(ws, f"D{r}", ceg_9m[label][0], fmt, f"{RELEASE_Q3}, nine months ended Sep 30, 2026")
                inp(ws, f"E{r}", ceg_9m[label][1], fmt, f"{RELEASE_Q3}, nine months ended Sep 30, 2025")
            else:
                c = concept or "DepreciationDepletionAndAmortization"
                for col, kind, end in (("C", "FY", "2025-12-31"), ("D", "9M", "2026-09-30"), ("E", "9M", "2025-09-30")):
                    v, src = xbrl(t, c, kind, end)
                    inp(ws, f"{col}{r}", round(v / scale, 2), fmt, src)
            style(ws[f"F{r}"], fmt=fmt).value = f"=C{r}+D{r}-E{r}"
            ltm_ref[(t, label)] = f"F{r}"
            r += 1

    balance = {
        "CEG": dict(shares=("=Inputs!C14", None), debt=("=Inputs!C20", None), cash=("=Inputs!C15", None), pref=("=Inputs!C19", None)),
    }
    for i, t in enumerate(companies):
        row = 5 + i
        ws[f"B{row}"] = names[t]
        style(ws[f"B{row}"], BOLD if t == "CEG" else BLACK)
        p, _ = price(t, "2026-11-20")
        if t == "CEG":
            style(ws[f"C{row}"], GREEN, USD).value = "=Inputs!C6"
            style(ws[f"D{row}"], GREEN, MM).value = "=Inputs!C14"
            style(ws[f"F{row}"], GREEN, MM).value = "=Inputs!C20"
            style(ws[f"G{row}"], GREEN, MM).value = "=Inputs!C15"
            style(ws[f"H{row}"], GREEN, MM).value = "=Inputs!C19"
        else:
            inp(ws, f"C{row}", p, USD, f"{YAHOO}, Nov 20, 2026")
            sh, s_src = xbrl(t, "WeightedAverageNumberOfDilutedSharesOutstanding", "Q", "2026-09-30")
            inp(ws, f"D{row}", round(sh / 1e6, 1), MM, s_src)
            if t == "VST":
                d1, src1 = xbrl(t, "LongTermDebt", "bal", "2026-09-30")
                debt, dsrc = d1, src1 + " (total, including current portion)"
            elif t == "NRG":
                d1, src1 = xbrl(t, "LongTermDebtAndCapitalLeaseObligations", "bal", "2026-09-30")
                d2, _ = xbrl(t, "LongTermDebtAndCapitalLeaseObligationsCurrent", "bal", "2026-09-30")
                debt, dsrc = d1 + d2, src1 + " plus the current portion"
            else:
                d1, src1 = xbrl(t, "LongTermDebtNoncurrent", "bal", "2026-09-30")
                d2, _ = xbrl(t, "LongTermDebtCurrent", "bal", "2026-09-30")
                debt, dsrc = d1 + d2, src1 + " plus the current portion"
            inp(ws, f"F{row}", round(debt / 1e6), MM, dsrc)
            cash, csrc = xbrl(t, "CashAndCashEquivalentsAtCarryingValue", "bal", "2026-09-30")
            inp(ws, f"G{row}", round(cash / 1e6), MM, csrc)
            pref = 0.0
            parts = []
            for concept in ("PreferredStockValue", "MinorityInterest"):
                try:
                    v, s = xbrl(t, concept, "bal", "2026-09-30")
                    pref += v
                    parts.append(s)
                except KeyError:
                    pass
            inp(ws, f"H{row}", round(pref / 1e6), MM, "; ".join(parts) or "None reported")
        style(ws[f"E{row}"], fmt=MM).value = f"=C{row}*D{row}"
        style(ws[f"I{row}"], fmt=MM).value = f"=E{row}+F{row}-G{row}+H{row}"
        style(ws[f"J{row}"], fmt=MM).value = f"={ltm_ref[(t, 'Revenue')]}"
        style(ws[f"K{row}"], fmt=MM).value = f"={ltm_ref[(t, 'Operating income')]}+{ltm_ref[(t, 'Depreciation & amortization')]}"
        style(ws[f"L{row}"], fmt=USD).value = f"={ltm_ref[(t, 'Diluted EPS ($)')]}"
        style(ws[f"M{row}"], fmt=MULT).value = f'=IF(K{row}>0,I{row}/K{row},"NM")'
        style(ws[f"N{row}"], fmt=MULT).value = f'=IF(L{row}>0,C{row}/L{row},"NM")'
        style(ws[f"O{row}"], fmt=MULT).value = f"=I{row}/J{row}"
    ws["B10"] = "Peer median (VST, NRG, TLN)"
    style(ws["B10"], BOLD)
    for col in "MNO":
        style(ws[f"{col}10"], fmt=MULT, bold=True).value = f"=MEDIAN({col}6:{col}8)"
    note(ws["M8"], "Talen's trailing GAAP EBITDA is tiny because of hedge losses and acquisition effects, so its multiple is an outlier. The median limits its pull.")
    ws["B11"] = "CEG premium to peer median"
    for col in "MNO":
        style(ws[f"{col}11"], fmt=PCT).value = f'=IFERROR({col}5/{col}10-1,"NM")'

    imp_top = raw_top + 1 + len(companies) * (len(metrics) + 1) + 2
    section(ws, imp_top, "What CEG would be worth at peer multiples", 6)
    ws[f"B{imp_top + 1}"] = "At peer median EV / EBITDA ($ per share)"
    style(ws[f"C{imp_top + 1}"], fmt=USD, fill=OUT_FILL).value = "=(M10*K5-F5+G5-H5)/D5"
    ws[f"B{imp_top + 2}"] = "At peer median P / E ($ per share)"
    style(ws[f"C{imp_top + 2}"], fmt=USD, fill=OUT_FILL).value = "=N10*L5"
    ws[f"B{imp_top + 3}"] = "Note"
    style(ws[f"C{imp_top + 3}"], SUB).value = ("CEG's trailing numbers include only nine months of Calpine (acquired January 2026), "
                                               "and its premium reflects long-term contracted nuclear output. Peers here are gas-heavy.")
    comps_imp = (f"Comps!C{imp_top + 1}", f"Comps!C{imp_top + 2}")

    # ── Pitch scorecard ─────────────────────────────────────────────────────
    ws = card
    widths(ws, {"A": 2, "B": 46, "C": 14, "D": 14, "E": 14, "F": 50})
    style(ws["B1"], TITLE).value = "Pitch scorecard"
    style(ws["B2"], SUB).value = "How the September 2026 call has done so far, and why. Total returns include dividends."
    header_row(ws, 4, ["Total return since the pitch", "Sep 10, 2026", "Nov 20, 2026", "Return"])
    for i, t in enumerate(["CEG", "SPY", "VST", "NRG", "TLN"]):
        r = 5 + i
        ws[f"B{r}"] = {"SPY": "S&P 500 (SPY)", **names}.get(t, t)
        style(ws[f"B{r}"], BOLD if t == "CEG" else BLACK)
        _, a0 = price(t, "2026-09-10")
        _, a1 = price(t, "2026-11-20")
        inp(ws, f"C{r}", round(a0, 4), USD, f"{YAHOO}: close adjusted for dividends, Sep 10, 2026")
        inp(ws, f"D{r}", round(a1, 4), USD, f"{YAHOO}: close adjusted for dividends, Nov 20, 2026")
        style(ws[f"E{r}"], fmt=PCT, bold=t == "CEG").value = f"=D{r}/C{r}-1"
    ws["B10"] = "CEG return relative to the S&P 500"
    style(ws["E10"], fmt=PCT, bold=True).value = "=(1+E5)/(1+E6)-1"

    section(ws, 12, "Why the price fell: earnings up, multiple down", 5)
    header_row(ws, 13, ["", "At pitch", "Now", "Change"])
    ws["B14"] = "Share price ($)"
    style(ws["C14"], GREEN, USD).value = "=Inputs!C8"
    style(ws["D14"], GREEN, USD).value = "=Inputs!C6"
    ws["B15"] = "2026 adjusted EPS guidance midpoint ($)"
    style(ws["C15"], GREEN, USD).value = "=Inputs!C32"
    style(ws["D15"], GREEN, USD).value = "=Inputs!C33"
    ws["B16"] = "P/E on 2026 guidance midpoint"
    style(ws["C16"], fmt=MULT).value = "=C14/C15"
    style(ws["D16"], fmt=MULT).value = "=D14/D15"
    for r in (14, 15, 16):
        style(ws[f"E{r}"], fmt=PCT).value = f"=D{r}/C{r}-1"
    ws["B17"] = "Check: (1 + EPS change) × (1 + P/E change) − 1"
    style(ws["E17"], fmt=PCT).value = "=(1+E15)*(1+E16)-1"
    style(ws["F15"], SUB).value = "Guidance was raised twice; Q3 adjusted EPS was $5.06 vs. $3.04 a year earlier."
    style(ws["F16"], SUB).value = "The market now pays less for each dollar of earnings."
    style(ws["F17"], SUB).value = "Equals the price change: the whole drop is the multiple."
    note(ws["F15"], f"Source: {RELEASE_Q3}")

    section(ws, 19, "The pitch's exit criteria, checked against filings", 5)
    header_row(ws, 20, ["Exit if…", "Status", "", "", "Evidence"])
    checks = [
        ("A major power purchase agreement is canceled", "Not triggered",
         "No cancellation reported. Q3 release: new 20-year, 270 MW agreement and a long-term collaboration with Google."),
        ("The NRC delays the Crane restart past 2028", "Not triggered",
         "Q2 release: NRC approved Crane's fuel license and FERC granted the interconnection waiver. No delay reported since."),
        ("Hyperscaler capital spending stalls noticeably", "Warning signs",
         "Not yet visible in the hyperscalers' filings, but the market is pricing it in: see the AI and chip stocks below."),
    ]
    for i, (a, b, c) in enumerate(checks):
        r = 21 + i
        ws[f"B{r}"] = a
        style(ws[f"C{r}"], BOLD).value = b
        style(ws[f"F{r}"], SUB).value = c
        ws[f"F{r}"].alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 30
    section(ws, 25, "Market signals on AI spending, same period", 5)
    header_row(ws, 26, ["Total return since the pitch", "Sep 10, 2026", "Nov 20, 2026", "Return"])
    for i, (t, label) in enumerate([("NVDA", "Nvidia (NVDA)"), ("SMH", "Semiconductor ETF (SMH)"), ("QQQ", "Nasdaq-100 (QQQ)")]):
        r = 27 + i
        ws[f"B{r}"] = label
        _, a0 = price(t, "2026-09-10")
        _, a1 = price(t, "2026-11-20")
        inp(ws, f"C{r}", round(a0, 4), USD, f"{YAHOO}: close adjusted for dividends, Sep 10, 2026")
        inp(ws, f"D{r}", round(a1, 4), USD, f"{YAHOO}: close adjusted for dividends, Nov 20, 2026")
        style(ws[f"E{r}"], fmt=PCT).value = f"=D{r}/C{r}-1"

    section(ws, 31, "Decision as of Nov 20, 2026", 5)
    style(ws["B32"], BOLD).value = "Recommendation"
    style(ws["C32"], BOLD, fill=KEY_FILL).value = "Exit"
    reasons = [
        "The pitch's exit rules were about the drivers of the thesis: new power deals, power prices, and policy.",
        "Nvidia and chip stocks have fallen sharply since the pitch (above), which says the market doubts the AI spending behind the power-demand story. The November elections add policy uncertainty for nuclear.",
        "The company is still executing (guidance raised twice), but the reason to own it was AI-driven demand, and that is the part in question. Today's price already assumes growth below management's target, so the upside case needs the AI story to recover.",
    ]
    for i, text in enumerate(reasons):
        r = 33 + i
        c = ws[f"B{r}"]
        c.value = text
        style(c)
        ws.merge_cells(f"B{r}:F{r}")
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 30

    for r in range(4, 33):
        b = ws[f"B{r}"]
        if b.value and b.font.color is None:
            style(b)

    # ── Summary ──────────────────────────────────────────────────────────────
    ws = summary
    widths(ws, {"A": 2, "B": 40, "C": 13, "D": 13, "E": 13, "F": 13, "G": 3, "H": 11, "I": 11})
    style(ws["B1"], TITLE).value = "Constellation Energy (NASDAQ: CEG): valuation model"
    style(ws["B2"], SUB).value = ("Built to test my September 2026 long pitch with real numbers. Student work, not investment advice. "
                                  "Every number traces to an input on the Inputs or Comps sheet, each with its source.")
    section(ws, 4, "Snapshot", 5)
    snap = [
        ("Valuation date", "=Inputs!C5", DATE_FMT),
        ("Share price ($)", "=Inputs!C6", USD),
        ("Pitch price, Sep 10, 2026 ($)", "=Inputs!C8", USD),
        ("Total return since the pitch", "='Pitch scorecard'!E5", PCT),
        ("S&P 500 total return, same period", "='Pitch scorecard'!E6", PCT),
        ("2026 adjusted EPS guidance midpoint ($)", "=Inputs!C33", USD),
        ("Cost of equity", "=Inputs!C42", PCT),
        ("Scenario in use", '=INDEX({"Bear","Base","Bull"},Valuation!C5)', None),
    ]
    for i, (label, f, fmt) in enumerate(snap):
        r = 5 + i
        ws[f"B{r}"] = label
        style(ws[f"B{r}"])
        style(ws[f"C{r}"], GREEN, fmt, align="right").value = f

    section(ws, 14, "Valuation range ($ per share)", 5)
    header_row(ws, 15, ["Method", "Low", "High", "Range", "Base / point"])
    ff = [
        ("52-week trading range", "=Inputs!C10", "=Inputs!C11", "=Inputs!C6"),
        ("DCF, bear to bull scenario", "=Valuation!D73", "=Valuation!D75", "=Valuation!D74"),
        ("DCF, cost of equity 10–12%", "=MIN(Valuation!D43:H45)", "=MAX(Valuation!D43:H45)", "=Valuation!F44"),
        ("12-month P/E target, bear to bull", "=Valuation!E73", "=Valuation!E75", "=Valuation!E74"),
        ("Peer multiples (GAAP, trailing)", f"=MIN({comps_imp[0]},{comps_imp[1]})", f"=MAX({comps_imp[0]},{comps_imp[1]})",
         f"=AVERAGE({comps_imp[0]},{comps_imp[1]})"),
        ("Pitch target", "=Inputs!C9", "=Inputs!C9", "=Inputs!C9"),
        ("Today's price", "=Inputs!C6", "=Inputs!C6", "=Inputs!C6"),
    ]
    for i, (label, lo, hi, mid) in enumerate(ff):
        r = 16 + i
        ws[f"B{r}"] = label
        style(ws[f"B{r}"])
        style(ws[f"C{r}"], GREEN, USD0).value = lo
        style(ws[f"D{r}"], GREEN, USD0).value = hi
        style(ws[f"E{r}"], fmt=USD0).value = f"=D{r}-C{r}"
        style(ws[f"F{r}"], GREEN, USD0).value = mid
        # Chart helpers: a single value is drawn as a $4-wide tick so it stays visible.
        style(ws[f"H{r}"], SUB, USD0).value = f"=IF(E{r}<4,C{r}-2,C{r})"
        style(ws[f"I{r}"], SUB, USD0).value = f"=MAX(E{r},4)"
    style(ws["H15"], SUB).value = "Chart start"
    style(ws["I15"], SUB).value = "Chart width"
    last_ff = 15 + len(ff)
    chart = BarChart()
    chart.type = "bar"
    chart.grouping = "stacked"
    chart.overlap = 100
    chart.title = "Valuation range vs. today's price"
    chart.y_axis.title = "$ per share"
    cats = Reference(ws, min_col=2, min_row=16, max_row=last_ff)
    base = Reference(ws, min_col=8, min_row=15, max_row=last_ff)
    span = Reference(ws, min_col=9, min_row=15, max_row=last_ff)
    chart.add_data(base, titles_from_data=True)
    chart.add_data(span, titles_from_data=True)
    chart.set_categories(cats)
    chart.series[0].graphicalProperties.noFill = True
    chart.series[0].graphicalProperties.line.noFill = True
    chart.series[1].graphicalProperties.solidFill = "1F3864"
    chart.legend = None
    chart.roundedCorners = False
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.title.overlay = False
    chart.y_axis.title.overlay = False
    chart.x_axis.scaling.orientation = "maxMin"
    chart.height = 7.5
    chart.width = 17
    ws.add_chart(chart, "B25")

    section(ws, 43, "What the model says", 5)
    lines = [
        '="1. The business beat the pitch: 2026 guidance midpoint rose "&TEXT(\'Pitch scorecard\'!E15,"0.0%")&", yet the stock fell "&TEXT(-\'Pitch scorecard\'!E14,"0.0%")&". The P/E on 2026 earnings fell from "&TEXT(\'Pitch scorecard\'!C16,"0.0")&"x to "&TEXT(\'Pitch scorecard\'!D16,"0.0")&"x."',
        '="2. Vistra and NRG fell about as much (VST "&TEXT(\'Pitch scorecard\'!E7,"0%")&", NRG "&TEXT(\'Pitch scorecard\'!E8,"0%")&"), so this was a sector-wide repricing, not a CEG-specific miss."',
        '="3. At an "&TEXT(Inputs!C42,"0.0%")&" cost of equity, today\'s price is justified on cash flows only if adjusted EPS grows about "&TEXT(Valuation!C95-0.025,"0%")&"–"&TEXT(Valuation!C95,"0.0%")&" a year through 2029, below management\'s 20% target (Valuation sheet, break-even table)."',
        '="4. Base case: 12-month target $"&TEXT(Valuation!E74,"0")&" ("&TEXT(Valuation!E74/Inputs!C6-1,"+0%;-0%")&"); DCF value $"&TEXT(Valuation!D74,"0")&". The gap is the question: is a "&TEXT(Valuation!F9,"0")&"x P/E justified by growth after 2029?"',
        '="5. Decision (Nov 20, 2026): exit. Nvidia is down "&TEXT(-\'Pitch scorecard\'!E27,"0%")&" and chip stocks "&TEXT(-\'Pitch scorecard\'!E28,"0%")&" since the pitch, so the AI spending behind the thesis is in doubt, and policy after the November elections is less certain (Pitch scorecard sheet)."',
    ]
    for i, f in enumerate(lines):
        c = ws[f"B{44 + i}"]
        c.value = f
        style(c)
        ws.merge_cells(f"B{44 + i}:F{44 + i}")
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[44 + i].height = 30

    # ── Sources ──────────────────────────────────────────────────────────────
    ws = sources
    widths(ws, {"A": 2, "B": 18, "C": 14, "D": 120})
    style(ws["B1"], TITLE).value = "Sources"
    style(ws["B2"], SUB).value = "Every typed-in number in the workbook. Hover a blue cell to see the same citation."
    header_row(ws, 4, ["Sheet", "Cell", "Source"])
    for i, (sheet, ref, src) in enumerate(source_rows):
        style(ws.cell(5 + i, 2, sheet))
        style(ws.cell(5 + i, 3, ref))
        style(ws.cell(5 + i, 4, src))

    from openpyxl.worksheet.datavalidation import DataValidation

    dv = DataValidation(type="list", formula1='"1,2,3"', allow_blank=False)
    val.add_data_validation(dv)
    dv.add("C5")

    for ws in wb.worksheets:
        ws.sheet_properties.tabColor = {"Summary": "1F3864", "Inputs": "0000FF"}.get(ws.title, "A6A6A6")
        if ws.title == "Comps":
            ws.freeze_panes = "C5"
        # One font everywhere: any cell still in the default font gets Arial 10, keeping bold/italic/color.
        for row in ws.iter_rows():
            for cell in row:
                f = cell.font
                if f.name != FONT:
                    cell.font = Font(name=FONT, size=10, bold=f.bold, italic=f.italic, color=f.color)

    OUT.parent.mkdir(exist_ok=True)
    wb.save(OUT)
    (ROOT / "model" / "tables.json").write_text(json.dumps(tables, indent=1))
    print(f"wrote {OUT} with {len(source_rows)} sourced inputs")


if __name__ == "__main__":
    main()
