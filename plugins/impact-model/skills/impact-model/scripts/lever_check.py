#!/usr/bin/env python3
"""Challenge the lever: which P&L line has the most room to improve against peers?

    python scripts/lever_check.py peers.csv --target "Client Co" --emphasis revenue
    python scripts/lever_check.py peers.csv --target "Client Co" --emphasis revenue --json lever.json

peers.csv (one row per company and year, values in one currency, e.g. from an
Orbis/SABI peer export):
    company,year,revenue,cogs,personnel,other_opex,da,employees
Cost columns are positive numbers. Any extra numeric column is treated as a
cost line. The latest year per company is used; revenue growth uses the first
and last year available.

For each cost line: share of revenue for the target vs the peer median and best
quartile, and the money gap at the target's revenue. For revenue: target CAGR vs
peer median CAGR, and revenue per employee. The line with the largest money gap
to the peer median is "the most room"; if it differs from --emphasis the script
says so.
"""
from __future__ import annotations

import argparse
import csv
import json
import statistics
import sys

COST_HINT = {"cogs": "cost of sales", "personnel": "personnel costs", "other_opex": "other operating costs",
             "da": "depreciation & amortisation"}


def q(xs, p):
    xs = sorted(xs)
    if not xs:
        return None
    k = (len(xs) - 1) * p
    f = int(k)
    c = min(f + 1, len(xs) - 1)
    return xs[f] + (xs[c] - xs[f]) * (k - f)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("csv")
    ap.add_argument("--target", required=True)
    ap.add_argument("--emphasis", default="revenue", help="line the strategy emphasises: revenue or a cost column")
    ap.add_argument("--json")
    a = ap.parse_args()
    rows = list(csv.DictReader(open(a.csv)))
    if not rows:
        sys.exit("empty CSV")
    num = lambda x: float(x) if x not in (None, "") else None
    by = {}
    for r in rows:
        by.setdefault(r["company"], []).append({k: (v if k == "company" else num(v)) for k, v in r.items()})
    for c in by:
        by[c].sort(key=lambda r: r["year"])
    if a.target not in by:
        sys.exit(f"target '{a.target}' not in CSV companies: {sorted(by)}")
    peers = [c for c in by if c != a.target]
    if len(peers) < 3:
        print(f"WARNING: only {len(peers)} peers; use 3-5 for a defensible benchmark")
    cols = [k for k in rows[0] if k not in ("company", "year", "revenue", "employees")]
    t = by[a.target][-1]
    out = []
    print(f"Lever check: {a.target} ({int(t['year'])}) vs {len(peers)} peers: {', '.join(peers)}\n")
    print(f"{'line':28} {'target':>8} {'peer med':>9} {'best q':>8} {'gap to median':>15}")
    for col in cols:
        if t.get(col) is None or not t["revenue"]:
            continue
        ts = t[col] / t["revenue"]
        ps = [by[c][-1][col] / by[c][-1]["revenue"] for c in peers if by[c][-1].get(col) is not None]
        if not ps:
            continue
        med, best = statistics.median(ps), q(ps, 0.25)
        gap = (ts - med) * t["revenue"]
        out.append({"line": col, "target_pct": ts, "peer_median_pct": med, "best_quartile_pct": best, "gap": gap,
                    "gap_best": (ts - best) * t["revenue"]})
        print(f"{COST_HINT.get(col, col):28} {ts:8.1%} {med:9.1%} {best:8.1%} {gap:15,.1f}")

    def cagr(series):
        s = [r for r in series if r.get("revenue")]
        if len(s) < 2:
            return None
        n = s[-1]["year"] - s[0]["year"]
        return (s[-1]["revenue"] / s[0]["revenue"]) ** (1 / n) - 1 if n > 0 else None
    tg = cagr(by[a.target])
    pg = [g for g in (cagr(by[c]) for c in peers) if g is not None]
    if tg is not None and pg:
        med = statistics.median(pg)
        gap = (med - tg) * t["revenue"]
        out.append({"line": "revenue", "target_pct": tg, "peer_median_pct": med, "best_quartile_pct": q(pg, 0.75),
                    "gap": gap, "gap_best": (q(pg, 0.75) - tg) * t["revenue"]})
        print(f"{'revenue growth (CAGR)':28} {tg:8.1%} {med:9.1%} {q(pg, 0.75):8.1%} {gap:15,.1f}  (per year)")
    if t.get("employees"):
        rpe = t["revenue"] / t["employees"]
        pr = [by[c][-1]["revenue"] / by[c][-1]["employees"] for c in peers if by[c][-1].get("employees")]
        if pr:
            print(f"{'revenue per employee':28} {rpe:8,.2f} {statistics.median(pr):9,.2f}")
    out.sort(key=lambda x: -x["gap"])
    top = out[0] if out else None
    if top:
        name = COST_HINT.get(top["line"], top["line"])
        print(f"\nMost room vs peer median: {name} ({top['gap']:,.1f} at {a.target}'s revenue)")
        agree = top["line"] == a.emphasis
        msg = (f"The strategy emphasises {a.emphasis}; " +
               ("that is also where the peer gap is largest - the lever is confirmed."
                if agree else f"but the largest gap is {name}. Say so in the deck, or add an initiative on it."))
        print(msg)
        if a.json:
            json.dump({"lines": out, "most_room": top["line"], "emphasis": a.emphasis, "confirmed": agree,
                       "answer": f"Against {len(peers)} peers ({', '.join(peers)}), the largest gap to the peer "
                                 f"median is {name} ({top['gap']:,.1f}). {msg}",
                       "cell": f"lever_check on {a.csv}"}, open(a.json, "w"), indent=2)


if __name__ == "__main__":
    main()
