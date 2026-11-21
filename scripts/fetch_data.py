"""Download everything the model uses, from free public sources.

- SEC EDGAR "company facts": every number each company has reported in its 10-K and 10-Q
  filings, in machine-readable form (XBRL). https://www.sec.gov/edgar/sec-api-documentation
- Yahoo Finance chart data: daily closing prices for the stocks and the S&P 500 ETF (SPY).
- FRED (St. Louis Fed): the 10-year US Treasury yield, for the risk-free rate.

Run: python scripts/fetch_data.py   (re-downloads everything; takes ~30 seconds)
"""
import json
import time
from pathlib import Path

import requests

DATA = Path(__file__).resolve().parent.parent / "data"
# The SEC asks automated clients to identify themselves in the User-Agent.
SEC_HEADERS = {"User-Agent": "constellation-model student research admin@example.com"}

COMPANIES = {  # ticker: SEC CIK number
    "CEG": 1868275,  # Constellation Energy
    "VST": 1692819,  # Vistra
    "NRG": 1013871,  # NRG Energy
    "TLN": 1622536,  # Talen Energy
}
PRICES = ["CEG", "VST", "NRG", "TLN", "SPY", "NVDA", "SMH", "QQQ"]


def main():
    (DATA / "sec").mkdir(parents=True, exist_ok=True)
    (DATA / "prices").mkdir(parents=True, exist_ok=True)

    for ticker, cik in COMPANIES.items():
        url = f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik:010d}.json"
        r = requests.get(url, headers=SEC_HEADERS, timeout=60)
        r.raise_for_status()
        (DATA / "sec" / f"{ticker}.json").write_text(r.text, encoding="utf-8")
        print(f"SEC facts  {ticker}: {len(r.content) // 1024} KB")
        time.sleep(0.5)  # the SEC allows 10 requests a second; stay well under

    for ticker in PRICES:
        # Yahoo Finance's public chart endpoint: daily close, plus close adjusted for dividends and splits.
        url = f"https://query1.finance.yahoo.com/v8/finance/chart/{ticker}?range=5y&interval=1d"
        r = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=60)
        r.raise_for_status()
        res = r.json()["chart"]["result"][0]
        times = res["timestamp"]
        close = res["indicators"]["quote"][0]["close"]
        adj = res["indicators"]["adjclose"][0]["adjclose"]
        rows = ["date,close,adjclose"]
        for t, c, a in zip(times, close, adj):
            if c is not None and a is not None:
                rows.append(f"{time.strftime('%Y-%m-%d', time.gmtime(t - 5 * 3600))},{c:.4f},{a:.4f}")
        (DATA / "prices" / f"{ticker}.csv").write_text("\n".join(rows) + "\n", encoding="utf-8")
        print(f"Prices     {ticker}: {len(rows) - 1} days, last {rows[-1]}")

    r = requests.get("https://fred.stlouisfed.org/graph/fredgraph.csv?id=DGS10", timeout=60)
    r.raise_for_status()
    (DATA / "dgs10.csv").write_text(r.text, encoding="utf-8")
    print(f"10-year Treasury: {r.text.count(chr(10))} days")

    (DATA / "SOURCES.json").write_text(
        json.dumps({"downloaded": time.strftime("%Y-%m-%d"), "sec_ciks": COMPANIES, "prices": PRICES}, indent=1)
    )


if __name__ == "__main__":
    main()
