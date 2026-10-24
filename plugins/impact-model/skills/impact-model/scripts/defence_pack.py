#!/usr/bin/env python3
"""Draft the defence pack (markdown) from a built model and its spec.

    python scripts/defence_pack.py Model.xlsx spec.json Defence_Pack.md [--lever lever.json]

Produces: one-slide impact bridge (table + bridge.png if matplotlib is present),
ten likely judge questions with 30-second answers and the cell behind each,
and the three weakest assumptions (sensitivity x weak grade) with the NPV if
each is wrong and the break-even value. Numbers come from the model itself
(recalculated with the built-in engine), never retyped.
Requires: openpyxl. Optional: matplotlib (bridge chart).
"""
from __future__ import annotations

import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xleval import Model, is_num, run_big_stack  # noqa: E402

GRADE_W = {c: i for i, c in enumerate("ABCDEF")}


def find_label(m, sheet, text, col=2):
    for (sh, r, c), v in m.cells.items():
        if sh == sheet and c == col and isinstance(v, str) and v.startswith(text):
            return r
    return None


# register header (row 5 of Inputs) -> field name; matched on the start of the header text
REG_FIELDS = {"input": "name", "unit": "unit", "type": "type", "low": "low", "base": "base", "high": "high",
              "source": "source", "date": "date", "link": "link", "grade": "grade",
              "benchmark low": "bl", "benchmark median": "bmed", "benchmark high": "bh",
              "benchmark source": "bsrc", "justification": "just", "key": "key", "revision": "rev"}


def register_columns(ws):
    col = {}
    for c in range(1, ws.max_column + 1):
        h = str(ws.cell(5, c).value or "").lower()
        for pre, f in REG_FIELDS.items():
            if h.startswith(pre) and f not in col:
                col[f] = c
    return col


class Pack:
    def __init__(self, xlsx, spec):
        self.m = Model(xlsx)
        self.spec = spec
        self.cur = spec["meta"]["currency"]
        self.ctx = self.m.context()
        self.inputs = {}
        ws = self.m.wb["Inputs"]
        self.col = col = register_columns(ws)
        for r in range(6, ws.max_row + 1):
            key = ws.cell(r, col["key"]).value
            if key:
                self.inputs[key] = {k: ws.cell(r, c).value for k, c in col.items()}
                self.inputs[key]["row"] = r
        self.calc_row = {}
        for label in ("NPV of strategy", "IRR of strategy", "Horizon-only NPV", "PV of terminal value", "Payback (years", "WACC", "Cost of equity",
                      "Cumulative incremental revenue", "Model revenue in base year", "Reported revenue in base year"):
            self.calc_row[label] = find_label(self.m, "Calculations", label)

    def v(self, sheet, r, c, ctx=None):
        return (ctx or self.ctx).get((sheet, r, c))

    def calc(self, label, ctx=None):
        if label not in self.calc_row:
            self.calc_row[label] = find_label(self.m, "Calculations", label)
        r = self.calc_row[label]
        return self.v("Calculations", r, 5, ctx) if r else None

    def npv_with(self, key, value):
        r = self.inputs[key]["row"]
        ctx = self.m.context({("Inputs", r, 6): value})
        return self.calc("NPV of strategy", ctx)

    # -- formatting
    def money(self, x):
        return f"{x:,.1f} {self.cur}" if is_num(x) else str(x)

    def val(self, key, x=None):
        d = self.derived(key) if key not in self.inputs else None
        if d:   # a derived input lives on Calculations section 1b (id in column A, value in E)
            ws = self.m.wb["Calculations"]
            r = next((r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == key), None)
            v = self.v("Calculations", r, 5) if r else None
            u = d.get("unit", "")
            return (f"{v:.1%}" if u.startswith("%") else f"{v:,.4g} {u}".strip()) if is_num(v) else key
        i = self.inputs[key]
        x = i["base"] if x is None else x
        if (i["unit"] or "").startswith("%"):
            return f"{x:.1%}"
        unit = "" if (i["unit"] or "") in ("x", "abs") else (i["unit"] or "")
        return f"{x:,.4g} {unit}".strip()

    def table(self, sheet):
        """Evaluated rows of a sheet: [(row, [values in columns B..L])], blanks dropped at the end of a row."""
        if sheet not in self.m.wb.sheetnames:
            return []
        ws = self.m.wb[sheet]
        out = []
        for r in range(1, ws.max_row + 1):
            vals = []
            for c in range(2, 13):
                raw = ws.cell(r, c).value
                vals.append(self.v(sheet, r, c) if isinstance(raw, str) and raw.startswith("=") else raw)
            while vals and vals[-1] in (None, ""):
                vals.pop()
            out.append((r, vals))
        return out

    def source(self, key):
        """Source of an input, or of a derived id (built from its arguments' sources)."""
        if key in self.inputs:
            return self.inputs[key]["source"]
        d = next((x for x in self.spec.get("derived", []) if x["id"] == key), None)
        if not d:
            return key
        return f"{d['name']} = " + " and ".join(f"{self.inputs.get(x, {}).get('name', x)} ({self.source(x)})"
                                                 for x in d["args"])

    def derived(self, key):
        return next((x for x in self.spec.get("derived", []) if x["id"] == key), None)

    def cite(self, key):
        d = self.derived(key) if key not in self.inputs else None
        if d:
            return f"derived ({d['name']}) from " + "; ".join(self.cite(x) for x in d["args"])
        i = self.inputs[key]
        return f"`Inputs!F{i['row']}` - {i['source']} ({i['date']}; grade {i['grade']})"

    # -- analyses
    def tornado(self):
        m = self.m
        ws = m.wb["Sensitivity"]
        r0 = find_label(m, "Sensitivity", "Live NPV")
        rows = []
        r = r0 + 1
        base = self.calc("NPV of strategy")
        while isinstance(ws.cell(r, 2).value, str) and ws.cell(r, 2).value.startswith("=Inputs"):
            ir = int(ws.cell(r, 2).value.split("$B$")[1])
            key = m.wb["Inputs"].cell(ir, self.col["key"]).value
            lo, hi = self.v("Sensitivity", r, 4), self.v("Sensitivity", r, 6)
            rows.append({"key": key, "lo": lo, "hi": hi, "swing": abs(hi - lo), "row": r})
            r += 1
        rows.sort(key=lambda x: -x["swing"])
        return base, rows

    def breakeven(self, key):
        i = self.inputs[key]
        lo, hi = i["low"], i["high"]
        span = (hi - lo) or abs(i["base"]) or 1.0
        a, b = lo - 3 * span, hi + 3 * span
        if lo >= 0:
            a = max(a, 0.0)  # inputs that are never negative (costs, uplifts, p, q) stay on their side of zero
        if (i["unit"] or "").startswith("%") or i["type"] in ("uplift",):
            a = max(a, -0.99)
        # scan for a sign change (values that break the model, e.g. p <= 0, are skipped), then bisect
        n = 24
        pts = [a + (b - a) * k / n for k in range(n + 1)]
        vals = [(x, self.npv_with(key, x)) for x in pts]
        vals = [(x, f) for x, f in vals if is_num(f)]
        br = next(((x0, f0, x1, f1) for (x0, f0), (x1, f1) in zip(vals, vals[1:]) if f0 * f1 <= 0), None)
        if not br:
            return None
        a, fa, b, fb = br
        base = i["base"]
        for _ in range(50):
            mid = (a + b) / 2
            fm = self.npv_with(key, mid)
            if not is_num(fm):
                break
            if fa * fm <= 0:
                b, fb = mid, fm
            else:
                a, fa = mid, fm
        be = (a + b) / 2
        f_be = self.npv_with(key, be)
        scale = max(abs(fa), abs(fb), 1e-9)
        if not is_num(f_be) or abs(f_be) > 0.01 * scale:
            return None   # a pole (e.g. terminal growth reaching the WACC), not a break-even
        if is_num(base) and base > 0 and be < 0:
            return "unreachable"  # would need a negative cost / rate: no plausible value turns NPV positive
        return be

    def scenarios(self):
        r = find_label(self.m, "Sensitivity", "Live (current switches)")
        return {k: self.v("Sensitivity", r + j, 4) for j, k in ((1, "low"), (2, "base"), (3, "high"))}, r

    def bridge(self, start=5):
        ws = self.m.wb["Bridge"]
        out = []
        r = start
        while ws.cell(r, 2).value:
            out.append((ws.cell(r, 2).value, self.v("Bridge", r, 3), r))
            r += 1
        return out

    def ebit_bridge(self):
        r = find_label(self.m, "Bridge", "Do nothing (no incremental EBIT)")
        return self.bridge(r) if r else []


def weakness(grade):
    g = (grade or "F6").upper().replace(" ", "")
    try:
        return (GRADE_W[g[0]] + int(g[1]) - 1) / 10.0 + 0.1
    except (KeyError, ValueError, IndexError):
        return 1.1


def draw_bridge(items, cur, path, title="Revenue bridge: baseline to target"):
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
    except ImportError:
        return False
    labels = [x[0] for x in items]
    vals = [x[1] for x in items]
    fig, ax = plt.subplots(figsize=(10, 4.8))
    run = 0.0
    for k, (lab, v) in enumerate(zip(labels, vals)):
        if k == 0 or k == len(vals) - 1:
            ax.bar(k, v, color="#7F7F7F")
            run = v
            ax.text(k, v, f"{v:,.0f}", ha="center", va="bottom", fontsize=9)
        else:
            bottom = run if v >= 0 else run + v
            ax.bar(k, abs(v), bottom=bottom, color="#70AD47" if v >= 0 else "#C00000")
            ax.text(k, run + max(v, 0), f"{v:+,.0f}", ha="center", va="bottom", fontsize=9)
            run += v
    if vals[0] > 0 and min(vals) >= 0:
        ax.set_ylim(min(vals[0], vals[-1]) * 0.97, max(vals[0], vals[-1], run) * 1.02)
    ax.axhline(0, color="#404040", lw=0.6)
    ax.set_xticks(range(len(labels)))
    ax.set_xticklabels([l.replace(": ", ":\n") for l in labels], fontsize=8)
    ax.set_ylabel(cur)
    ax.spines[["top", "right"]].set_visible(False)
    ax.set_title(title, fontsize=12)
    fig.tight_layout()
    fig.savefig(path, dpi=160)
    plt.close(fig)
    return True


def fin_margin_name(spec):
    return spec.get("finance", {}).get("margin")


def num(x, unit=""):
    if not is_num(x):
        return str(x if x is not None else "-")
    if unit in ("%", "pp"):
        return f"{x * 100:+.1f} pp" if unit == "pp" else f"{x:.1%}"
    return f"{x:,.1f}" if abs(x) < 10000 else f"{x:,.0f}"


def objectives_section(P, spec, L, fy):
    """Section 1: the case objectives first (base case vs with strategies), then the framework views."""
    rows = P.table("Objectives")
    if not rows:
        return
    L.append(f"## 1. Impact on the case objectives ({fy}, base case = do nothing)\n")
    head = next((i for i, (_, v) in enumerate(rows) if v and v[0] == "Objective"), None)
    kpis = []
    if head is not None:
        L.append("| Objective | KPI | Base case | With strategies | Change | Change % | Better? |")
        L.append("|---|---|---:|---:|---:|---:|---|")
        for r, v in rows[head + 1:]:
            if not v:
                break
            v = v + [None] * (9 - len(v))
            unit = v[2]
            pct = v[6] if is_num(v[6]) else "n/a"
            L.append(f"| {v[0]} | {v[1]} ({unit}) | {num(v[3], '%' if unit == 'pp' else '')} | "
                     f"{num(v[4], '%' if unit == 'pp' else '')} | {num(v[5], unit if unit == 'pp' else '')} | "
                     f"{pct if not is_num(pct) else f'{pct:+.1%}'} | {v[8]} |")
            kpis.append((v[1], unit, v[5], v[8]))
        L.append(f"\n_Source: `Objectives!B{rows[head][0]}` (links to Calculations section 10)._\n")
    worse = [k for k in kpis if k[3] == "worse"]
    if worse:
        L.append("> **Trade-off to explain:** " + "; ".join(f"{k[0]} moves the wrong way ({num(k[2], k[1])})"
                                                         for k in worse) + ". Say why it is worth it, or which "
                 "strategy causes it (table below).\n")
    ch = next((i for i, (_, v) in enumerate(rows) if v and str(v[0]).startswith("Contribution")), None)
    if ch is not None:
        hdr = rows[ch + 1][1]
        L.append(f"**How each strategy moves the objectives** (change in {fy} vs base case)\n")
        L.append("| " + " | ".join(str(x) for x in hdr) + " |")
        L.append("|" + "---|" * 2 + "---:|" * (len(hdr) - 2))
        for r, v in rows[ch + 2:]:
            if not v:
                break
            L.append("| " + " | ".join(num(x) if is_num(x) else str(x or "") for x in v) + " |")
        L.append("")
    fw = P.table("Frameworks")
    if not fw:
        return
    L.append("**The same impact through the consulting frameworks** (`Frameworks` sheet; every view adds up to "
             "the same totals)\n")
    i = 0
    while i < len(fw):
        r, v = fw[i]
        title = str(v[0]) if v else ""
        if title[:2].rstrip(".").isdigit() and i + 2 < len(fw) and fw[i + 2][1] and \
                fw[i + 2][1][0] in ("Category", "Lever", "Perspective"):
            hdr = fw[i + 2][1]
            body = []
            j = i + 3
            while j < len(fw) and fw[j][1] and fw[j][1][0] != "Total":
                body.append(fw[j][1])
                j += 1
            if hdr[0] == "Lever":
                parts = [f"{b[0].split(' (')[0].split(', after')[0]} {num(b[1])}" for b in body
                         if is_num(b[1]) and abs(b[1]) > 1e-9]
                L.append(f"- **{title.split(':')[0][3:]}:** " + "; ".join(parts) + f" ({P.cur}).")
            elif hdr[0] == "Category":
                col = 2 if len(hdr) > 2 else 1   # the profit column (column 1 is cumulative revenue)
                live = [b for b in body if any(is_num(x) and abs(x) > 1e-9 for x in b[1:len(hdr) - 1])]
                parts = []
                for b in live:
                    cells = [f"{hdr[k].split(' (')[0].split(': ', 1)[-1].replace(f' {fy}', '')} {num(b[k])}"
                             for k in range(1, len(hdr) - 1) if is_num(b[k]) and abs(b[k]) > 1e-9]
                    parts.append(f"{b[0].split(' (')[0]}: " + ", ".join(cells))
                L.append(f"- **{title.split(':')[0][3:]}:** " + "; ".join(parts) + ".")
            i = j
        i += 1
    L.append("")


def build_pack(xlsx, spec, out, lever=None):
    P = Pack(xlsx, spec)
    meta = spec["meta"]
    fy = meta["base_year"] + meta["horizon_years"]
    by = {i["id"]: i for i in spec["inputs"]}
    inits = [dict(it, pillar=p["name"]) for p in spec["pillars"] for it in p.get("initiatives", [])]
    npv = P.calc("NPV of strategy")
    irr = P.calc("IRR of strategy")
    pb = P.calc("Payback (years")
    wacc = P.calc("WACC")
    base, torn = P.tornado()
    scen, srow = P.scenarios()
    br = P.bridge()
    lastc = 7 + (meta["base_year"] - meta.get("first_hist_year", meta["base_year"])) + meta["horizon_years"]
    ramps, slow = [], []
    for it in inits:
        rp = (it.get("ramp") or {}).get("type", "step")
        adr = find_label(P.m, "Calculations", f"Adoption ({rp} curve): {it['name']}")
        a_last = P.v("Calculations", adr, lastc) if adr else None
        if rp == "bass":
            how = f"Bass S-curve (p = {P.val(it['ramp']['p'])}, q = {P.val(it['ramp']['q'])})"
        elif rp == "logistic":
            how = "logistic S-curve"
        elif rp == "linear":
            how = f"linear rollout over {P.inputs[it['ramp']['years_to_peak']]['base']:g} years" \
                if it['ramp']['years_to_peak'] in P.inputs else "linear rollout"
        else:
            how = "full effect from launch (step)"
        ramps.append(f"{it['name']}: {how}" + (f", {a_last:.0%} of peak by {fy}" if is_num(a_last) else ""))
        if is_num(a_last) and a_last < 0.5 and (it.get("uplift") or it.get("model") == "build"):
            slow.append(f"{it['name']} ({a_last:.0%})")
    L = []
    L.append(f"# Defence pack - {meta.get('title', meta['company'])}")
    L.append(f"_Generated from `{os.path.basename(xlsx)}` (all numbers recalculated from the model). "
             f"Currency: {P.cur}. Base year {meta['base_year']}, horizon to {fy}._\n")
    objectives_section(P, spec, L, fy)
    L.append("## 1b. Revenue and EBIT bridge\n")
    png = os.path.join(os.path.dirname(os.path.abspath(out)), "bridge.png")
    if draw_bridge(br, P.cur, png):
        L.append(f"![Revenue bridge](bridge.png)\n")
    L.append(f"| Step | {fy} revenue ({P.cur}) | Cell |\n|---|---:|---|")
    for lab, v, r in br:
        L.append(f"| {lab} | {v:,.1f} | `Bridge!C{r}` |")
    eb = P.ebit_bridge()
    if eb:
        png2 = os.path.join(os.path.dirname(os.path.abspath(out)), "ebit_bridge.png")
        L.append(f"\n**Value bridge: incremental EBIT {fy} by initiative** (cost savings and running costs show "
                 f"here, not in revenue)\n")
        if draw_bridge(eb, P.cur, png2, f"Incremental EBIT {fy} by initiative ({P.cur})"):
            L.append("![EBIT bridge](ebit_bridge.png)\n")
        L.append(f"| Step | {fy} EBIT ({P.cur}) | Cell |\n|---|---:|---|")
        for lab, v, r in eb:
            L.append(f"| {lab} | {v:,.1f} | `Bridge!C{r}` |")
    up = br[-1][1] / br[0][1] - 1
    L.append(f"\n**Headline:** the strategy adds {br[-1][1] - br[0][1]:,.0f} {P.cur} of revenue in {fy} "
             f"(+{up:.1%} vs do-nothing). NPV {P.money(npv)} at a {wacc:.1%} WACC; IRR "
             f"{irr if not is_num(irr) else f'{irr:.1%}'} (horizon cash flows); payback "
             f"{pb if not is_num(pb) else f'{pb:.1f} years'}. Of the NPV, {P.money(P.calc('PV of terminal value'))} "
             f"is terminal value; horizon-only NPV is {P.money(P.calc('Horizon-only NPV'))}. "
             f"Scenarios: low {P.money(scen['low'])} / base {P.money(scen['base'])} / high {P.money(scen['high'])} "
             f"(`Sensitivity!D{srow + 1}:D{srow + 3}`).\n")
    if is_num(irr) and is_num(npv) and irr < 0 < npv:
        L.append("> **IRR is negative while NPV is positive:** IRR uses the horizon cash flows only, and the NPV "
                 "includes the terminal value. Quote the horizon-only NPV next to it.\n")
    tvv = P.calc("PV of terminal value")
    if is_num(tvv) and is_num(npv) and npv > 0 and tvv > 0.75 * npv:
        L.append(f"> **Terminal value carries {tvv / npv:.0%} of the NPV.** Expect the question 'what happens after "
                 f"year {meta['horizon_years']}?': show the horizon-only NPV and the fade assumption.\n")
    if slow:
        L.append(f"> **Still ramping at the horizon:** {', '.join(slow)} below 50% of peak effect in {fy}. The "
                 f"terminal value capitalises a level that is still growing: lengthen the horizon or say so.\n")
    if is_num(npv) and npv < 0:
        L.append("> **Read this before presenting:** the base case destroys value. Say so, show what has to be "
                 "true to turn it positive (break-even values below), or rework the initiatives. Do not hide it.\n")
    # ---------------------------------------------------------------- questions
    Q = []
    rev = next((a for a in spec.get("actuals", []) if a.get("role") == "revenue"), None)
    mb, rb = P.calc("Model revenue in base year"), P.calc("Reported revenue in base year")
    if rev:
        Q.append(("Where does your baseline come from?",
                  f"From reported accounts, not press coverage: {rev['name']} {meta['base_year']} = "
                  f"{rb:,.1f} {P.cur} ({rev['source']}, grade {rev['grade']}). Our driver tree rebuilds it to "
                  f"{mb:,.1f} ({abs(mb / rb - 1):.2%} gap), so every uplift starts from the real P&L.",
                  f"`Actuals` row for {rev['id']}; calibration check on `Checks`"))
    bg = P.calc("Baseline revenue CAGR")
    if rev and is_num(bg):
        hc, hlo, hhi = P.calc(f"{rev['name']}: CAGR"), P.calc("History band, low"), P.calc("History band, high")
        q1, q3 = P.calc(f"{rev['name']}: YoY growth, lower percentile"), P.calc(f"{rev['name']}: YoY growth, upper")
        ys = sorted(int(y) for y in rev.get("values", {}))
        shock = meta.get("shock_years") or []
        inside = is_num(hlo) and is_num(hhi) and hlo <= bg <= hhi
        txt = (f"Our do-nothing revenue grows {bg:.1%} a year to {fy}. The company's own record {ys[0]}-{ys[-1]}: "
               f"CAGR {hc:.1%}, year-on-year growth {q1:.1%} to {q3:.1%} (interquartile)" if is_num(hc) and
               is_num(q1) else f"Our do-nothing revenue grows {bg:.1%} a year to {fy}")
        if shock and is_num(hc):
            txt += (f", leaving out the shock years {', '.join(map(str, shock))} and the rebound after them"
                    + (f" ({meta['shock_note']})" if meta.get("shock_note") else ""))
        br = meta.get("baseline_bridge")
        if inside:
            txt += ". It sits inside that band, so the strategy's uplift is not hiding in an optimistic baseline."
        elif br:
            steps = []
            for k, item in enumerate(br.get("items", [])):
                v = P.calc(f"Bridge to history: {item['name']}")
                steps.append(f"{item['name']} {v:+.1%}" if is_num(v) else item["name"])
            gap = P.calc("Bridge gap")
            st0 = P.calc("Bridge to history: start")
            txt += (f". It sits outside that band, and we bridge it step by step: "
                    + (f"history {st0:.1%}, " if is_num(st0) else "") + ", ".join(steps)
                    + (f" = {bg:.1%} (gap {gap:.1%})" if is_num(gap) else "") + "."
                    + (f" Why: {meta['baseline_justification']}" if meta.get("baseline_justification") else ""))
        else:
            txt += ". It sits outside that band and is not yet bridged to it: [team: add the bridge before presenting]."
        linked = [i for i in spec["inputs"] if i.get("history")]
        if linked:
            txt += " Base-year levels and growth linked to the actuals: " + ", ".join(i["name"] for i in linked) + "."
        notes = [i for i in spec["inputs"] if i.get("history_note")]
        for i in notes[:2]:
            txt += f" {i['name']}: {i['history_note'].rstrip('.')}."
        bt = P.calc("Back-test: largest yearly gap")
        if is_num(bt):
            ws = P.m.wb["Calculations"]
            tested = [str(ws.cell(r, 2).value)[len("Back-test: "):-len(" from driver history")]
                      for r in range(1, ws.max_row + 1) if str(ws.cell(r, 2).value or "").startswith("Back-test: ")
                      and str(ws.cell(r, 2).value).endswith(" from driver history")]
            untested = [x["name"] for x in spec["streams"] if x["name"] not in tested]
            txt += (f" Run on the reported driver history, our driver tree reproduces past "
                    f"{' and '.join(tested) or 'revenue'} within {bt:.2%} in every year"
                    + (f" ({', '.join(untested)} not back-tested: no independent driver history)." if untested
                       else "."))
        mg = [a for a in spec.get("actuals", []) if a.get("role") in ("ebitda", "ebit", "gross_profit")]
        if mg and fin_margin_name(spec):
            m = P.calc(f"{mg[0]['name']}: median % of revenue")
            if is_num(m):
                txt += (f" The company's {mg[0]['name']} margin has a median of {m:.1%}; the contribution margin on "
                        f"incremental revenue ({P.val(spec['finance']['margin'])}) is higher because fixed costs do "
                        f"not grow with the extra sales.")
        Q.append((f"Why does your baseline grow {bg:.1%} a year when the company grew "
                  f"{hc:.1%}?" if is_num(hc) else "Is your baseline consistent with the company's history?", txt,
                  "`Calculations` sections 0 (history) and 9b-9c; `Checks` rows 'Base case vs history' and 'Back-test'"))
    names = {it["id"]: it["name"] for it in inits}
    for o in spec.get("objectives", []):
        for k in (o.get("kpis") or [])[:1]:   # the first KPI of each non-financial objective
            eff = "; ".join(f"{names.get(e['initiative'], e['initiative'])}: its modelled {e['quantity']}"
                            + "".join(f" x {P.inputs.get(x, {}).get('name', x)} ({P.val(x)})" for x in e.get("factor") or [])
                            + "".join(f" / {P.inputs.get(x, {}).get('name', x)}" for x in e.get("per") or [])
                            for e in k["effects"])
            base = (" x ".join(f"{P.inputs.get(x, {}).get('name', x)} ({P.val(x)})" for x in k.get("baseline") or [])
                    or "zero (only the change is counted)")
            cites = [P.cite(x) for e in k["effects"] for x in (e.get("factor") or [])][:2] + \
                [P.cite(x) for x in (k.get("baseline") or [])][:1]
            nm = k["name"][0].lower() + k["name"][1:] if k["name"][1:2].islower() else k["name"]
            Q.append((f"How did you measure {nm}?",
                      f"From the model's own quantities, not a separate estimate. Base case: {base}. Change: {eff}. "
                      f"Every factor is a sourced, graded input, so it moves with the scenarios.",
                      "; ".join(cites) or "`Calculations` section 10"))
    ups = [t for t in torn if by.get(t["key"], {}).get("type") == "uplift"]
    for t in ups[:1 if is_num(bg) and rev else 2]:
        i = P.inputs[t["key"]]
        bench = f"{i['bl']:.1%}-{i['bh']:.1%}" if is_num(i["bl"]) and is_num(i["bh"]) else "n/a"
        Q.append((f"Why is {i['name'].lower()} of {P.val(t['key'])} credible?",
                  f"It sits inside the reference class of {bench} ({i['bsrc']}). Source: {i['source']}. "
                  f"If it came in at {P.val(t['key'], i['low'])} instead, NPV moves to {P.money(t['lo'])}.",
                  P.cite(t["key"])))
    if ramps:
        Q.append(("How fast do the effects ramp up - aren't you assuming day-one impact?",
                  "No. Each initiative ramps on its own curve: " + "; ".join(ramps) + ".",
                  "adoption rows on `Calculations` (section 3 / 4c); ramp inputs on `Inputs`"))
    Q.append(("Aren't you double counting between pillars?",
              "No. Every driver-tree leaf has exactly one owning initiative or an explicit split rule, the "
              "incremental revenue of each stream is split across leaves with log shares that add to 100%, and "
              "two live checks prove initiatives and pillars add back to the stream total every year.",
              "`Driver tree` sheet; `Checks` rows 'Initiatives add up...' and 'Pillars add up...'"))
    Q.append(("What if it works less well than you hope?",
              f"Our low case moves every ranged input to its pessimistic end at once: NPV {P.money(scen['low'])} "
              f"(base {P.money(scen['base'])}, high {P.money(scen['high'])}). The tornado shows which single input "
              f"matters most.", f"`Sensitivity!D{srow + 1}`"))
    fin = spec["finance"]
    parts = [f"rf {P.val(fin['rf'])}", f"beta {P.val(fin['beta'])}", f"ERP {P.val(fin['erp'])}"]
    if fin.get("crp"):
        parts.append(f"country risk {P.val(fin['crp'])}")
    Q.append(("How did you get your discount rate?",
              f"WACC {wacc:.1%} built bottom-up: " + ", ".join(parts) + f", cost of debt {P.val(fin['kd'])}, "
              f"debt weight {P.val(fin['debt_weight'])}, tax {P.val(fin['tax'])}. Each input is sourced: "
              f"{P.inputs[fin['erp']]['source']}.", f"`Calculations` WACC block; {P.cite(fin['beta'])}"))
    costs = [k for k in P.inputs if by.get(k, {}).get("type") == "cost"]
    if costs:
        desc = "; ".join(f"{P.inputs[k]['name']}: {P.val(k)}" for k in costs)
        Q.append(("What does it cost and when does it pay back?",
                  f"Costs are in the cash flow, not left out: {desc}. Payback: "
                  f"{pb if not is_num(pb) else f'{pb:.1f} years after the base year'}.",
                  "`Calculations` section 7 (costs) and payback row"))
    ceil = [lf for lf in spec["leaves"] if lf.get("ceiling")]
    if ceil:
        lf = ceil[0]
        pr_ = find_label(P.m, "Calculations", f"Penetration of addressable market (% of ceiling): {lf['name']}")
        lastc = 7 + (meta["base_year"] - meta.get("first_hist_year", meta["base_year"])) + meta["horizon_years"]
        pen = P.v("Calculations", pr_, lastc) if pr_ else None
        builds = [it for it in inits if it.get("model") == "build"]
        shared = [it["name"] for it in builds if it["volume"].get("capacity_of") == lf["id"]]
        unshared = [it["name"] for it in builds if it["kind"] == "revenue" and
                    it["volume"].get("capacity_of") != lf["id"]]
        txt = (f"{lf['name']} is capped by its ceiling ({P.source(lf['ceiling'])}); by {fy} it reaches "
               f"{pen:.1%} of it." if is_num(pen) else f"{lf['name']} is capped by its ceiling.")
        if shared:
            txt += f" {', '.join(shared)} draw on the same ceiling, and a combined penetration check covers them."
        if unshared:
            txt += (f" {', '.join(unshared)} are capped by their own ceilings only: check that the total still "
                    f"fits capacity before presenting.")
        txt += " Every capped volume has a check that blocks anything above 100% of its ceiling."
        Q.append(("Is the market (or capacity) big enough for this?", txt, P.cite(lf["ceiling"])))
    if torn:
        t = torn[0]
        be = P.breakeven(t["key"])
        Q.append(("Which single number matters most?",
                  f"{P.inputs[t['key']]['name']}: moving it across its range swings NPV by {P.money(t['swing'])} "
                  f"({P.money(t['lo'])} to {P.money(t['hi'])})." +
                  (f" NPV turns zero at {P.val(t['key'], be)}." if is_num(be) else
                   " No plausible value of it alone turns NPV positive." if be == "unreachable" else ""),
                  f"{P.cite(t['key'])}; tornado on `Sensitivity`"))
    if lever:
        Q.append(("Why this lever and not another?",
                  lever.get("answer", ""), lever.get("cell", "lever_check output")))
    elif len(Q) < 10:
        Q.append(("Why did you pick this lever over others?",
                  "Benchmarked against 3-5 peers on Orbis/SABI: see the lever check. (Run scripts/lever_check.py "
                  "once the peer P&L export is in; until then this answer is open.)", "Data Request Sheet"))
    L.append("## 2. Ten questions a judge is most likely to ask\n")
    for k, (q, a, c) in enumerate(Q[:10], 1):
        L.append(f"**Q{k}. {q}**  \n{a}  \n_Source cell: {c}_\n")
    # ---------------------------------------------------------------- weakest
    L.append("## 3. Three weakest assumptions (sensitivity x weak grade)\n")
    scored = []
    for t in torn:
        i = P.inputs[t["key"]]
        scored.append((t["swing"] * weakness(i["grade"]), t, i))
    scored.sort(key=lambda x: -x[0])
    L.append(f"| # | Assumption | Base | Grade | NPV if at low / high | Break-even for NPV = 0 | Cell |")
    L.append("|---|---|---|---|---|---|---|")
    for k, (sc, t, i) in enumerate(scored[:3], 1):
        be = P.breakeven(t["key"])
        bes = P.val(t["key"], be) if is_num(be) else ("not reachable alone" if be == "unreachable" else "none in range")
        L.append(f"| {k} | {i['name']} | {P.val(t['key'])} | {i['grade']} | {P.money(t['lo'])} / "
                 f"{P.money(t['hi'])} | {bes} | "
                 f"`Inputs!F{i['row']}` |")
    L.append("")
    for k, (sc, t, i) in enumerate(scored[:3], 1):
        L.append(f"- **{i['name']}** - source: {i['source']} ({i['grade']}). If wrong at the pessimistic end NPV "
                 f"becomes {P.money(min(t['lo'], t['hi']))}. Strengthen it with: "
                 f"{by.get(t['key'], {}).get('request', {}).get('database', 'a better-graded source from references/source-map.md')}.")
    with open(out, "w") as f:
        f.write("\n".join(L) + "\n")
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("xlsx")
    ap.add_argument("spec")
    ap.add_argument("out")
    ap.add_argument("--lever", help="JSON written by lever_check.py --json")
    a = ap.parse_args()
    with open(a.spec) as f:
        spec = json.load(f)
    m_ = spec["meta"]   # same default as build_model.normalise: the timeline starts at the earliest reported year
    m_.setdefault("first_hist_year", min([int(y) for x in spec.get("actuals", []) for y in x.get("values", {})
                                          if int(y) <= int(m_["base_year"])] + [int(m_["base_year"])]))
    lever = None
    if a.lever:
        with open(a.lever) as f:
            lever = json.load(f)
    run_big_stack(build_pack, a.xlsx, spec, a.out, lever)
    print(f"defence pack written to {a.out}")


if __name__ == "__main__":
    main()
