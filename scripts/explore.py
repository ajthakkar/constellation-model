"""Print the reported values of key line items, to see what's available. Usage: explore.py TICKER [concept ...]"""
import json
import sys
from pathlib import Path

DATA = Path(__file__).resolve().parent.parent / "data"
ticker = sys.argv[1]
facts = json.loads((DATA / "sec" / f"{ticker}.json").read_text())["facts"]["us-gaap"]
wanted = sys.argv[2:] or [
    "Revenues", "RevenueFromContractWithCustomerIncludingAssessedTax", "OperatingIncomeLoss",
    "DepreciationDepletionAndAmortization", "DepreciationAndAmortization", "NetIncomeLoss",
    "EarningsPerShareDiluted", "NetCashProvidedByUsedInOperatingActivities",
    "PaymentsToAcquirePropertyPlantAndEquipment", "PaymentsToAcquireNuclearFuel",
    "LongTermDebt", "LongTermDebtNoncurrent", "LongTermDebtCurrent", "CashAndCashEquivalentsAtCarryingValue",
    "WeightedAverageNumberOfDilutedSharesOutstanding", "InterestExpense", "InterestExpenseNonoperating",
    "IncomeTaxExpenseBenefit", "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
]
for c in wanted:
    if c not in facts:
        print(f"{c}: (not reported)")
        continue
    units = facts[c]["units"]
    for unit, rows in units.items():
        # Keep the latest-filed value for each fiscal period: full years (FY) and quarters.
        seen = {}
        for r in rows:
            fp, fy, form = r.get("fp"), r.get("fy"), r.get("form")
            if form not in ("10-K", "10-Q"):
                continue
            start, end = r.get("start", ""), r["end"]
            key = (start, end)
            seen[key] = r["val"]
        annual = {e[:4]: v for (s, e), v in seen.items() if s and int(e[:4]) - int(s[:4]) >= 0 and (int(e[5:7]) - int(s[5:7]) + 12 * (int(e[:4]) - int(s[:4]))) == 11}
        quarters = {e: v for (s, e), v in seen.items() if s and (int(e[5:7]) - int(s[5:7]) + 12 * (int(e[:4]) - int(s[:4]))) == 2}
        instants = {e: v for (s, e), v in seen.items() if not s}
        def fmt(d, n):
            return "  ".join(f"{k}:{v/1e6:,.0f}" if abs(v) > 1e4 else f"{k}:{v}" for k, v in sorted(d.items())[-n:])
        print(f"{c} [{unit}]")
        if annual: print("   FY  ", fmt(annual, 5))
        if quarters: print("   Q   ", fmt(quarters, 6))
        if instants: print("   bal ", fmt(instants, 5))
