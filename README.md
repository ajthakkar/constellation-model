# Constellation Energy (CEG): pitch and valuation model

In September 2026 I pitched Constellation Energy as a long at $285.93 with a 12-month target of $355 ([the pitch](pitch/Constellation-Energy-pitch.pdf)). This repo turns that pitch into an Excel model built from SEC filings, and checks how the call has done.

**[Download the model (CEG-valuation.xlsx)](model/CEG-valuation.xlsx)**. Every output is a live formula, and every typed-in number has its source in a cell comment and on the Sources sheet.

## What the model finds (as of Nov 20, 2026)

| | |
|---|---|
| Total return since the pitch | **−16.9%** (S&P 500: −8.7%; Vistra −21.5%; NRG −21.0%; Talen +6.9%) |
| 2026 adjusted EPS guidance | **$12.00 → $12.65** midpoint (raised twice; Q3 adjusted EPS $5.06 vs. $3.04 a year earlier) |
| P/E on 2026 guidance | **23.8x → 18.8x** |
| Base-case value | DCF **$234**, 12-month target **$276** (19x 2027 EPS), vs. a $237.41 price |
| Bull case (management's 20% growth) | DCF **$345**, 12-month target **$364** |
| Bear case (8% growth) | DCF **$138**, 12-month target **$205** |
| Growth today's price needs | about **15–17.5%** a year through 2029, below management's 20% target |

**The business beat the pitch; the stock didn't.** The price change splits exactly into earnings and multiple: (1 + 5.4%) × (1 − 21.2%) − 1 = −17.0%. Every bit of the drop came from investors paying less per dollar of earnings, and the gas-heavy peers fell just as hard, so this looks like a sector-wide repricing rather than anything specific to Constellation.

## The workbook

| Sheet | What's on it |
|---|---|
| Summary | Snapshot, valuation ranges, a football-field chart, and the four takeaways (all formulas) |
| Inputs | Prices, the Sep 30, 2026 balance sheet, adjusted EPS history and guidance, and the cost of equity (CAPM) |
| Valuation | Bear / base / bull scenarios, a 10-year equity cash flow DCF, the 12-month P/E target, and five Excel Data Tables: cost of equity × long-run growth, growth × cash conversion, growth × P/E, all three scenarios, and the break-even growth rate |
| Comps | Vistra, NRG, and Talen on trailing GAAP numbers: EV/EBITDA, P/E, EV/revenue, and what CEG would be worth at peer multiples |
| Pitch scorecard | Returns since the pitch, the earnings-vs-multiple split, and the pitch's three exit criteria checked against filings |
| Market data | Two years of weekly prices; beta is a `SLOPE` formula over the returns |
| Sources | All 96 typed-in numbers with their sources |

Formatting follows banking convention: blue = input, black = formula, green = link to another sheet, yellow = key assumption.

## Method, and why

- **Adjusted EPS, not GAAP.** Constellation's GAAP results swing with mark-to-market hedge gains and losses (Q2 2026: GAAP EPS $1.42 vs. adjusted $2.55), and its reported operating cash flow has been negative in some years because of how receivables are financed. So the model starts from management's adjusted EPS, the number the market prices, and converts it to cash for shareholders with an explicit, sensitized assumption (85% in the base case).
- **A 10-year equity DCF with a fading growth rate.** Growth runs at the scenario rate through 2029, fades in a straight line to the long-run rate by 2036, then grows forever (Gordon growth). Nuclear plants are long-lived and contracted, so a 5-year model would cut off most of the value.
- **Cost of equity from CAPM:** 10-year Treasury 5.04% (FRED) + Blume-adjusted beta 1.36 (raw 1.54 from two years of weekly returns vs. the S&P 500) × a 4.5% equity risk premium = 11.2%. The high risk-free rate is a big part of why the DCF is demanding.
- **The P/E target is the pitch's own method made explicit:** 2027 EPS × a target multiple. The pitch's $355 implied about 24.7x 2027 EPS on management's growth path.
- **Comps are a cross-check, not a valuation.** On trailing GAAP numbers CEG trades at 14.1x EV/EBITDA vs. a 9.5x peer median, a 49% premium. Peers are gas-heavy, CEG's trailing year has only nine months of Calpine, and the premium is the bet on long-term contracted nuclear output.

## Limits

- Adjusted EPS guidance and the 20% growth target are management's numbers; the 20% target comes from the Q2 2026 materials as cited in my pitch and isn't re-checked here.
- Cash conversion is a judgment call. The growth × conversion table shows how much it matters.
- The DCF ignores the few weeks left in 2026, which slightly understates value.
- Comps use GAAP results, which are noisy for power producers; peer forward estimates aren't freely available.

## Rebuild it

Windows with Excel installed (Excel adds the Data Tables and calculates the file):

```bash
python -m venv .venv && .venv/Scripts/activate
pip install openpyxl pandas requests pywin32
python scripts/fetch_data.py        # SEC filings, prices, Treasury yields
python scripts/build_model.py       # writes model/CEG-valuation.xlsx with formulas
python scripts/finish_in_excel.py   # adds Data Tables, calculates, saves
python scripts/check_model.py       # scans every cell for errors and prints the outputs
```

Sources: SEC EDGAR company facts (XBRL) and Constellation's 8-K earnings releases (Feb 2025, Feb 2026, Aug 2026, Nov 2026); Yahoo Finance prices; FRED series DGS10.

*A student stock pitch and model, not investment advice.*
