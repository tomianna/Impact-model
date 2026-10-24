#!/usr/bin/env python3
"""Build a FAST-standard impact model (.xlsx) from a JSON spec.

    python scripts/build_model.py spec.json out/Model.xlsx
    python scripts/build_model.py spec.json out/Model.xlsx --requests out/Data_Requests.xlsx
    python scripts/build_model.py --templates templates/        # regenerate blank templates

Sheets: Cover | Inputs (assumptions register) | Actuals | Calculations | Outputs |
Bridge | Sensitivity | Checks | Driver tree | Data requests.
Every number lives once on Inputs/Actuals with source, date, link and A-F/1-6
grade; every other cell is a live formula. Scenarios, tornado and two-way tables
are Excel Data Tables, recalculated by Excel on open.
The spec format is documented in references/spec-format.md.
Requires: openpyxl.
"""
from __future__ import annotations

import argparse
import copy
import json
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from openpyxl import Workbook  # noqa: E402
from openpyxl.chart import BarChart, LineChart, Reference  # noqa: E402
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side  # noqa: E402
from openpyxl.utils import get_column_letter as L  # noqa: E402
from openpyxl.workbook.defined_name import DefinedName  # noqa: E402
from openpyxl.worksheet.formula import DataTableFormula  # noqa: E402

EXEMPT = {"actual", "constant", "plan", "conversion", "control"}   # no range required
TYPES = EXEMPT | {"assumption", "uplift", "cost"}
FIRST_TL_COL = 7  # timeline starts in column G on every time-based sheet
# Assumptions Register columns (the hand-off format every agent reads; see references/agents.md)
REG_HEAD = ["ID", "Input", "Unit", "Type", "Low", "Base", "High", "Polarity", "Source", "Date",
            "Link / database path", "Grade", "Benchmark low (p25)", "Benchmark median", "Benchmark high (p75)",
            "Benchmark source", "Justification (if above p75)", "Used by", "Key", "Revision"]

BLUE = Font(color="1F4E9E")
BOLD = Font(bold=True)
TITLE = Font(bold=True, size=14)
HEAD_FILL = PatternFill("solid", fgColor="1F3864")
HEAD_FONT = Font(bold=True, color="FFFFFF")
SEC_FILL = PatternFill("solid", fgColor="D9E1F2")
IN_FILL = PatternFill("solid", fgColor="FFF2CC")
CTL_FILL = PatternFill("solid", fgColor="FCE4D6")
OUT_FILL = PatternFill("solid", fgColor="E2EFDA")
THIN = Side(style="thin", color="BFBFBF")


from objectives import ObjectivesMixin, SpecError, normalise_objectives  # noqa: E402


def ids_of(x):
    """An input reference may be one id or a list of ids multiplied together (quantity x unit cost)."""
    if not x:
        return []
    return [x] if isinstance(x, str) else list(x)


BUILD_KINDS = ("revenue", "saving")
DERIVED_OPS = ("prod", "sum", "sub", "div")   # sub: a - b - c ...; div: a / b / c ...


# Reported history (spec "actuals"): roles say what a series is, so the model can derive growth, margins and a back-test
HIST_ROLES = ("revenue", "gross_profit", "ebitda", "ebit", "ebt", "net_profit", "cost", "driver", "other")
MARGIN_ROLES = ("gross_profit", "ebitda", "ebit", "ebt", "net_profit", "cost")
HIST_STATS = ("last", "cagr", "median_growth", "median_margin")
HIST_LABEL = {"last": "last reported value", "cagr": "CAGR", "median_growth": "median YoY growth",
              "median_margin": "median % of revenue"}


def pct_inc(xs, k):
    """Excel PERCENTILE.INC."""
    s = sorted(xs)
    h = (len(s) - 1) * k
    lo = int(h)
    return s[lo] + (h - lo) * (s[min(lo + 1, len(s) - 1)] - s[lo])


def shock_hit(y, shock):
    """A YoY change is distorted when its year or the year before is a shock year (the trough and the rebound)."""
    return y in shock or y - 1 in shock


def cagr_ends(ys, shock):
    """CAGR endpoints: the first and last reported years that are not shock years."""
    clean = [y for y in ys if y not in shock]
    return (clean[0], clean[-1]) if len(clean) >= 2 else (None, None)


def series_stats(a, rev=None, shock=()):
    """Growth, CAGR, quartiles and (for P&L lines) margins of one actuals series, as the workbook computes them.
    Shock years (meta.shock_years) are left out of the medians, quartiles and the CAGR endpoints."""
    v = {int(y): float(x) for y, x in a.get("values", {}).items() if isinstance(x, (int, float))}
    ys = sorted(v)
    st = {"years": ys, "values": v, "last": v[ys[-1]] if ys else None, "last_year": ys[-1] if ys else None}
    g_all = {y: v[y] / v[y - 1] - 1 for y in ys if y - 1 in v and v[y - 1] != 0}
    g = {y: x for y, x in g_all.items() if not shock_hit(y, shock)}
    st["growth_all"], st["growth"] = g_all, g
    y0, y1 = cagr_ends(ys, shock)
    st["cagr_years"] = (y0, y1)
    if y0 is not None and v[y0] > 0 and v[y1] > 0:
        st["cagr"] = (v[y1] / v[y0]) ** (1 / (y1 - y0)) - 1
    if g:
        gs = list(g.values())
        st["median_growth"], st["g_p25"], st["g_p75"] = pct_inc(gs, 0.5), pct_inc(gs, 0.25), pct_inc(gs, 0.75)
    if rev is not None and a.get("role") in MARGIN_ROLES:
        rv = {int(y): float(x) for y, x in rev.get("values", {}).items() if isinstance(x, (int, float))}
        m = {y: v[y] / rv[y] for y in ys if rv.get(y) and y not in shock}
        st["margin"] = m
        if m:
            ms = list(m.values())
            st["median_margin"], st["m_p25"], st["m_p75"] = pct_inc(ms, 0.5), pct_inc(ms, 0.25), pct_inc(ms, 0.75)
    return st


def init_input_ids(it):
    """Every input id an initiative reads."""
    out = [it.get(k) for k in ("launch_year", "decay", "margin", "opex_growth", "capex_years", "asset_life")]
    out += ids_of(it.get("uplift")) + ids_of(it.get("opex")) + ids_of(it.get("capex")) + ids_of(it.get("scale"))
    out += [v for k, v in (it.get("ramp") or {}).items() if k != "type"]
    if it.get("model") == "build":
        out += ids_of(it["volume"]["drivers"]) + [it["volume"].get("ceiling"), it["volume"].get("ceiling_growth"),
                                                   it["volume"].get("growth"), it["price"].get("base"),
                                                   it["price"].get("growth")]
    return [x for x in out if x]


# ----------------------------------------------------------------------------- spec handling
def load_spec(path):
    with open(path) as f:
        return json.load(f)


def normalise(spec):
    """Fill defaults, add builder-owned inputs, validate. Returns (spec, warnings)."""
    s = copy.deepcopy(spec)
    warn = []
    meta = s.setdefault("meta", {})
    for k in ("company", "currency", "base_year", "horizon_years"):
        if k not in meta:
            raise SpecError(f"meta.{k} is required")
    meta.setdefault("title", f"{meta['company']} - strategy impact model")
    by_ = int(meta["base_year"])
    yrs = [int(y) for a in spec.get("actuals", []) for y in a.get("values", {}) if int(y) <= by_]
    meta.setdefault("first_hist_year", min(yrs + [by_]))   # show all reported history
    meta.setdefault("as_of", "")
    asof = meta["as_of"]
    inputs = s.setdefault("inputs", [])
    ids = [i["id"] for i in inputs]
    dup = {x for x in ids if ids.count(x) > 1}
    if dup:
        raise SpecError(f"duplicate input ids: {sorted(dup)}")
    by = {i["id"]: i for i in inputs}

    def add(i):
        if i["id"] not in by:
            inputs.append(i)
            by[i["id"]] = i

    conv = dict(type="constant", source="Modelling convention (this model)", date=asof,
                link="Cover sheet", grade="A1", section="Model settings")
    add(dict(id="first_hist_year", name="First historical year shown", unit="year",
             base=meta["first_hist_year"], **conv))
    add(dict(id="base_year", name="Base year (last reported actuals)", unit="year", base=meta["base_year"], **conv))
    add(dict(id="growth_cap", name="Growth cap: max YoY growth of target revenue", unit="%",
             base=meta.get("growth_cap", 0.25), **conv))
    add(dict(id="calib_tol", name="Tolerance: baseline vs reported revenue", unit="%",
             base=meta.get("calibration_tolerance", 0.02), **conv))
    add(dict(id="sum_tol", name="Tolerance: totals vs sum of parts", unit="abs",
             base=meta.get("sum_tolerance", 1e-6), **conv))
    add(dict(id="tie_step", name="Tornado tie-break step (keeps ranks unique when two swings are equal)",
             unit="abs", base=1e-9, **conv))
    pillars = s.setdefault("pillars", [])
    pids = {p["id"] for p in pillars}
    inits = [dict(i, pillar=p["id"]) for p in pillars for i in p.get("initiatives", [])]
    s["_initiatives"] = inits
    for it in inits:
        ly = it.get("launch_year")
        if isinstance(ly, (int, float)):
            lid = f"{it['id']}_launch"
            add(dict(id=lid, name=f"Launch year: {it['name']}", unit="year", base=int(ly), type="plan",
                     source="Team implementation plan", date=asof, link="Strategy deck / roadmap",
                     grade="A1", section=f"Initiatives: {it['pillar']}"))
            it["launch_year"] = lid
        if it.get("pillar") not in pids:
            raise SpecError(f"initiative {it['id']}: unknown pillar")
    leaves = {lf["id"]: lf for lf in s.get("leaves", [])}
    streams = s.get("streams", [])
    if not streams or not leaves:
        raise SpecError("spec needs at least one stream and its leaves")
    for st in streams:
        for lf in st["leaves"]:
            if lf not in leaves:
                raise SpecError(f"stream {st['id']}: unknown leaf {lf}")
    # one owner per leaf, or an explicit split rule (no euro counted twice)
    owners = {}
    for it in inits:
        if it.get("model") == "build":
            if it.get("kind") not in BUILD_KINDS:
                raise SpecError(f"initiative {it['id']}: a build initiative needs kind 'revenue' (new stream) or "
                                f"'saving' (cost avoided)")
            vol, pr = it.get("volume") or {}, it.get("price") or {}
            if not ids_of(vol.get("drivers")):
                raise SpecError(f"initiative {it['id']}: build needs volume.drivers (input ids multiplied together, "
                                f"e.g. [\"pv_mwp\", \"pv_yield\"])")
            if not pr.get("base"):
                raise SpecError(f"initiative {it['id']}: build needs price.base (value per unit of volume)")
            if it.get("leaf") or it.get("uplift"):
                raise SpecError(f"initiative {it['id']}: a build initiative has volume x price, not leaf/uplift")
            it["leaf"] = None
            continue
        if not it.get("uplift"):
            if not (it.get("opex") or it.get("capex")):
                raise SpecError(f"initiative {it['id']}: needs an uplift on a leaf, a build (volume x price), or "
                                f"(cost/saving-only) opex/capex")
            it["leaf"] = None
            continue
        if it.get("leaf") not in leaves:
            raise SpecError(f"initiative {it['id']}: leaf '{it.get('leaf')}' is not a driver-tree leaf")
        owners.setdefault(it["leaf"], []).append(it["id"])
    rules = s.setdefault("split_rules", {})
    for lf, own in owners.items():
        if len(own) > 1 and rules.get(lf) != "additive":
            raise SpecError(f"leaf '{lf}' is moved by {own}: add split_rules: {{\"{lf}\": \"additive\"}} "
                            f"(uplifts add; incremental revenue split pro rata to each uplift) or give it one owner")
    s["_owners"] = owners

    derived = {}
    for d in s.setdefault("derived", []):
        if d.get("op") not in DERIVED_OPS:
            raise SpecError(f"derived {d.get('id')}: op must be one of {', '.join(DERIVED_OPS)}")
        if d["id"] in by or d["id"] in derived:
            raise SpecError(f"derived {d['id']}: id already used")
        for x in d.get("args", []):
            if x not in by and x not in derived:
                raise SpecError(f"derived {d['id']}: argument '{x}' is not an input or an earlier derived id")
        if len(d.get("args", [])) < (1 if d["op"] in ("prod", "sum") else 2):
            raise SpecError(f"derived {d['id']}: needs more arguments for '{d['op']}'")
        derived[d["id"]] = d
    s["_derived"] = derived

    def need(iid, ctx):
        if iid is None:
            return
        if iid not in by and iid not in derived:
            raise SpecError(f"{ctx}: input '{iid}' is not defined in inputs[] or derived[]")

    for lf in leaves.values():
        need(lf.get("base"), f"leaf {lf['id']}.base")
        if not lf.get("base"):
            raise SpecError(f"leaf {lf['id']} needs a base input")
        for k in ("growth", "inflation", "ceiling", "ceiling_growth"):
            need(lf.get(k), f"leaf {lf['id']}.{k}")
    for st in streams:
        need(st.get("scale"), f"stream {st['id']}.scale")
    for it in inits:
        for k in ("launch_year", "decay", "margin", "opex_growth", "capex_years", "asset_life"):
            need(it.get(k), f"initiative {it['id']}.{k}")
        for k in ("opex", "capex", "uplift"):
            for x in ids_of(it.get(k)):
                need(x, f"initiative {it['id']}.{k}")
        if it.get("model") == "build":
            for x in ids_of(it["volume"]["drivers"]) + ids_of(it.get("scale")):
                need(x, f"initiative {it['id']}.volume.drivers/scale")
            for k in ("ceiling", "ceiling_growth", "growth"):
                need(it["volume"].get(k), f"initiative {it['id']}.volume.{k}")
            cap = it["volume"].get("capacity_of")
            if cap and not (cap in leaves and leaves[cap].get("ceiling")):
                raise SpecError(f"initiative {it['id']}: volume.capacity_of '{cap}' must be a leaf with a ceiling")
            for k in ("base", "growth"):
                need(it["price"].get(k), f"initiative {it['id']}.price.{k}")
        r = it.setdefault("ramp", {"type": "step"})
        for k, v in r.items():
            if k != "type":
                need(v, f"initiative {it['id']}.ramp.{k}")
        req = {"bass": ("p", "q"), "logistic": ("midpoint", "steepness"), "linear": ("years_to_peak",), "step": ()}
        if r["type"] not in req:
            raise SpecError(f"initiative {it['id']}: ramp type must be bass, logistic, linear or step")
        for k in req[r["type"]]:
            if k not in r:
                raise SpecError(f"initiative {it['id']}: {r['type']} ramp needs '{k}'")
        if not it.get("launch_year"):
            raise SpecError(f"initiative {it['id']} needs launch_year")
    fin = s.setdefault("finance", {})
    for k in ("margin", "tax", "rf", "beta", "erp", "kd", "debt_weight"):
        if k not in fin:
            raise SpecError(f"finance.{k} is required (WACC from sourced inputs)")
        need(fin[k], f"finance.{k}")
    for k in ("crp", "terminal_growth", "terminal_switch"):
        need(fin.get(k), f"finance.{k}")
    if not fin.get("terminal_growth"):
        add(dict(id="tv_g", name="Terminal growth of incremental FCF after the horizon", unit="%", type="assumption",
                 low=-0.05, base=-0.02, high=0.01, date=asof, grade="C3", section="Finance & WACC",
                 source="Modelling choice: effects fade 2%/yr after the horizon (low: 5% fade as competitors copy; "
                        "high: +1%); Gordon growth model", link="Damodaran, Investment Valuation (3rd ed.), ch. 12"))
        fin["terminal_growth"] = "tv_g"
    if not fin.get("terminal_switch"):
        add(dict(id="tv_on", name="Include terminal value (1 = yes, 0 = horizon only)", unit="switch",
                 type="control", base=meta.get("terminal_value", 1), date=asof, grade="A1",
                 section="Finance & WACC", source="Modelling choice (horizon-only NPV is also reported)",
                 link="Outputs: headline results"))
        fin["terminal_switch"] = "tv_on"
    # reported history: validate, compute stats, link inputs that take their value from history
    acts = s.setdefault("actuals", [])
    aby = {}
    fy = int(meta["first_hist_year"])
    for a in acts:
        if a["id"] in aby:
            raise SpecError(f"duplicate actuals id: {a['id']}")
        if a.get("op"):   # derived series (e.g. members per club = members / clubs): a formula row on Actuals
            if a["op"] not in DERIVED_OPS:
                raise SpecError(f"actuals {a['id']}: op must be one of {', '.join(DERIVED_OPS)}")
            args = a.get("args") or []
            if len(args) < 2 or any(x not in aby for x in args):
                raise SpecError(f"actuals {a['id']}: a derived series needs 2+ args that are earlier actuals ids")
            if a.get("values"):
                raise SpecError(f"actuals {a['id']}: give op/args or values, not both")
            common = set.intersection(*(set(aby[x]["values"]) for x in args))
            vals = {}
            for y in sorted(common, key=int):
                xs = [float(aby[x]["values"][y]) for x in args]
                if a["op"] == "div" and any(x == 0 for x in xs[1:]):
                    continue
                r = xs[0]
                for x in xs[1:]:
                    r = {"prod": r * x, "sum": r + x, "sub": r - x, "div": r / x}[a["op"]]
                vals[y] = r
            a["values"] = vals
            sym = {"prod": " x ", "sum": " + ", "sub": " - ", "div": " / "}[a["op"]]
            a.setdefault("source", "Derived on the Actuals sheet: " + sym.join(aby[x]["name"] for x in args))
            grades = [str(aby[x].get("grade", "")) for x in args if aby[x].get("grade")]
            if grades:
                a.setdefault("grade", max(g[0] for g in grades) + max(g[1:] for g in grades))
            a.setdefault("date", max(str(aby[x].get("date", "")) for x in args))
            a.setdefault("unit", "")
        aby[a["id"]] = a
        a.setdefault("role", "other")
        if a["role"] not in HIST_ROLES:
            raise SpecError(f"actuals {a['id']}: role must be one of {', '.join(HIST_ROLES)}")
        if a["role"] == "driver" and a.get("leaf") not in leaves:
            raise SpecError(f"actuals {a['id']}: a driver series needs leaf: <driver-tree leaf id> (history of that leaf)")
        if a.get("stream") and a["stream"] not in {x["id"] for x in streams}:
            raise SpecError(f"actuals {a['id']}: stream '{a['stream']}' is not a revenue stream")
        late = [y for y in a.get("values", {}) if int(y) > int(meta["base_year"])]
        if late:
            warn.append(f"actuals {a['id']}: values after the base year ({', '.join(late)}) are ignored - move the "
                        f"base year forward if they are reported")
        a["values"] = {y: x for y, x in a.get("values", {}).items() if fy <= int(y) <= int(meta["base_year"])}
    shock = sorted({int(y) for y in meta.get("shock_years", [])})
    hist_years = {int(y) for a in acts for y in a["values"]}
    bad = [y for y in shock if y not in hist_years]
    if bad:
        raise SpecError(f"meta.shock_years {bad}: not a reported year in actuals")
    if shock and not meta.get("shock_note"):
        warn.append("meta.shock_years without meta.shock_note - say what the shock was (it is the source on Actuals)")
    s["_shock"] = shock
    revs = [a for a in acts if a["role"] == "revenue"]
    if len(revs) > 1:
        raise SpecError("only one actuals series can have role 'revenue' (total reported revenue)")
    rev = revs[0] if revs else None
    stats = {a["id"]: series_stats(a, rev, shock) for a in acts}
    s["_hist"] = stats
    if not rev:
        warn.append("no reported revenue history (actuals with role 'revenue') - checker check 13 will fail")
    elif len(stats[rev["id"]]["years"]) < 3:
        warn.append(f"only {len(stats[rev['id']]['years'])} years of reported revenue - pull 5 (checker needs 3+)")
    elif stats[rev["id"]]["last_year"] != int(meta["base_year"]):
        warn.append(f"reported revenue ends in {stats[rev['id']]['last_year']}, not the base year")
    if rev and shock and len(stats[rev["id"]]["growth"]) < 2:
        raise SpecError(f"shock years {shock} leave fewer than 2 clean YoY changes of revenue - pull more history "
                        f"or drop a shock year")
    if rev and len(stats[rev["id"]]["growth"]) >= 2:
        hconv = dict(conv, section="Model settings")
        add(dict(id="hist_p_low", name="History band: lower percentile of YoY growth", unit="%", base=0.25, **hconv))
        add(dict(id="hist_p_high", name="History band: upper percentile of YoY growth", unit="%", base=0.75, **hconv))
        add(dict(id="hist_tol", name="History band: tolerance around the band (baseline growth vs history)",
                 unit="%", base=meta.get("history_tolerance", 0.01), **hconv))
    for i in inputs:
        h = i.get("history")
        if not h:
            continue
        for k in ("source", "date", "grade", "link"):
            if i.get(k) == "":
                del i[k]   # an empty field means "use the series' own"
        a = aby.get(h.get("series"))
        if not a:
            raise SpecError(f"input {i['id']}: history.series '{h.get('series')}' is not an actuals id")
        stat = h.get("stat")
        if stat not in HIST_STATS:
            raise SpecError(f"input {i['id']}: history.stat must be one of {', '.join(HIST_STATS)}")
        st = stats[a["id"]]
        if stat != "last" and "hist_p_low" not in by:
            raise SpecError(f"input {i['id']}: history stats need 3+ years of reported revenue (role 'revenue')")
        if stat == "median_margin" and a["role"] not in MARGIN_ROLES:
            raise SpecError(f"input {i['id']}: median_margin needs a P&L series ({', '.join(MARGIN_ROLES)})")
        val = st.get(stat)
        if val is None:
            raise SpecError(f"input {i['id']}: cannot compute {stat} of {a['id']} (needs 2+ consecutive positive years)")
        lo, hi = {"last": (val, val), "median_margin": (st.get("m_p25"), st.get("m_p75"))}.get(
            stat, (st.get("g_p25"), st.get("g_p75")))
        i["base"] = val
        i.setdefault("type", "actual" if stat == "last" else "assumption")
        i["_hist"] = {"series": a["id"], "stat": stat, "low": "low" not in i, "high": "high" not in i}
        i.setdefault("low", min(val, lo))
        i.setdefault("high", max(val, hi))
        if stat != "last" and i["high"] - i["low"] < 0.005:
            warn.append(f"{i['id']}: history gives a narrow range ({i['low']:.2%} to {i['high']:.2%}) - type low/high "
                        f"from the outlook (checker needs a real range)")
        ys = list(st["cagr_years"]) if stat == "cagr" else st["years"]
        i.setdefault("source", f"History: {HIST_LABEL[stat]} of {a['name']} {ys[0]}-{ys[-1]}"
                               + (f", shock years {', '.join(map(str, shock))} excluded" if shock and stat != "last" else "")
                               + " "
                               f"(live link to the Actuals sheet). Data: {a.get('source', '')}")
        i.setdefault("date", a.get("date", asof))
        i.setdefault("grade", a.get("grade", ""))
        i.setdefault("link", a.get("link") or "Actuals sheet")
    # bridge from history to the baseline: replaces a free-text justification with numbers that must add up
    br = meta.get("baseline_bridge")
    bridge_refs = []
    if br:
        if br.get("from", "cagr") not in ("cagr", "median_growth"):
            raise SpecError("meta.baseline_bridge.from must be 'cagr' or 'median_growth'")
        if not rev or "hist_p_low" not in by:
            raise SpecError("meta.baseline_bridge needs 3+ years of reported revenue (role 'revenue')")
        if not br.get("items"):
            raise SpecError("meta.baseline_bridge needs items: [{name, plus: [refs], minus: [refs]}]")
        for it_ in br["items"]:
            if not it_.get("name") or not (it_.get("plus") or it_.get("minus")):
                raise SpecError("each baseline_bridge item needs a name and plus and/or minus refs")
            for ref in (it_.get("plus") or []) + (it_.get("minus") or []) + \
                    ([it_["weight"]] if it_.get("weight") else []):
                if str(ref).startswith("hist:"):
                    _, sid, stat = (str(ref).split(":") + ["", ""])[:3]
                    if sid not in aby or stat not in ("cagr", "median_growth", "median_margin", "last") or \
                            stats[sid].get(stat) is None:
                        raise SpecError(f"baseline_bridge ref '{ref}': use hist:<actuals id>:<cagr|median_growth|"
                                        f"median_margin|last> with a computable statistic")
                elif ref in by or ref in derived:
                    bridge_refs.append(ref)
                else:
                    raise SpecError(f"baseline_bridge ref '{ref}' is not an input, a derived id or hist:<id>:<stat>")
    elif meta.get("baseline_justification"):
        warn.append("meta.baseline_justification alone does not pass the history check any more - add "
                    "meta.baseline_bridge (history growth + steps = baseline growth); see references/methods.md")
    s["_bridge_refs"] = bridge_refs
    # back-test: streams whose every leaf has a driver history
    drv = {}
    for a in acts:
        if a["role"] == "driver":
            drv.setdefault(a["leaf"], a["id"])
    s["_drivers"] = drv
    def roots(aid):
        a = aby[aid]
        return set().union(*(roots(x) for x in a["args"])) if a.get("op") else {aid}
    bt = []
    for st in streams:
        if not all(x in drv for x in st["leaves"]):
            continue
        tested = {a["id"] for a in acts if a.get("stream") == st["id"]} | ({rev["id"]} if rev else set())
        circ = sorted({drv[x] for x in st["leaves"] if roots(drv[x]) & tested})
        if circ:
            warn.append(f"stream {st['id']}: driver history {', '.join(circ)} is derived from the revenue it would be "
                        f"tested against - left out of the back-test (it would match by construction)")
        else:
            bt.append(st["id"])
    s["_backtest"] = bt
    if bt and (rev or any(a.get("stream") for a in acts)):
        add(dict(id="backtest_tol", name="Tolerance: back-test of the driver tree on reported history", unit="%",
                 base=meta.get("backtest_tolerance", 0.05), **dict(conv, section="Model settings")))
    for lf in leaves.values():
        if lf["id"] in drv:
            b = by.get(lf["base"])
            if b is not None and not (b.get("history") or {}).get("series") == drv[lf["id"]]:
                warn.append(f"leaf {lf['id']}: has driver history ({drv[lf['id']]}) but its base input {lf['base']} is "
                            f"typed - link it: \"history\": {{\"series\": \"{drv[lf['id']]}\", \"stat\": \"last\"}}")
            gs = [by[x] for x in (lf.get("growth"), lf.get("inflation")) if x in by]
            if not any(g.get("history") or g.get("history_note") for g in gs):
                warn.append(f"leaf {lf['id']}: do-nothing growth is not tied to its history ({drv[lf['id']]}) - use "
                            f"history stat cagr/median_growth, or write history_note on the growth input saying why not")

    # input defaults & warnings
    for i in inputs:
        i.setdefault("type", "assumption")
        if i["type"] not in TYPES:
            raise SpecError(f"input {i['id']}: type must be one of {sorted(TYPES)}")
        i.setdefault("unit", "")
        if "base" not in i:
            raise SpecError(f"input {i['id']} has no base value")
        if i["type"] in EXEMPT:
            i.setdefault("low", i["base"])
            i.setdefault("high", i["base"])
        else:
            if "low" not in i or "high" not in i:
                warn.append(f"{i['id']}: no low/high range (actuary rule) - checker will fail")
                i.setdefault("low", i["base"])
                i.setdefault("high", i["base"])
            if i["low"] > i["high"]:
                i["low"], i["high"] = i["high"], i["low"]
        i.setdefault("polarity", -1 if i["type"] == "cost" else 1)
        for k in ("source", "date", "grade"):
            if not i.get(k):
                warn.append(f"{i['id']}: missing {k} - checker will fail")
        g = str(i.get("grade", "")).upper().replace(" ", "")
        if g and (len(g) != 2 or g[0] not in "ABCDEF" or g[1] not in "123456"):
            warn.append(f"{i['id']}: grade '{g}' is not A-F + 1-6")
        if i["type"] == "uplift":
            if i.get("bench_high") is None:
                warn.append(f"{i['id']}: uplift without reference-class benchmark - checker will fail")
            elif i["base"] > i["bench_high"] and not i.get("justification"):
                warn.append(f"{i['id']}: base {i['base']} above benchmark high {i['bench_high']} with no "
                            f"justification - justify or pull back (outside view)")
    last_y = int(meta["base_year"]) + int(meta["horizon_years"])
    asof_y = int(str(asof)[:4]) if str(asof)[:4].isdigit() else None
    for it in inits:
        ly = by.get(it["launch_year"], {}).get("base")
        if isinstance(ly, (int, float)):
            if asof_y and ly <= asof_y:
                warn.append(f"initiative {it['id']}: launches in {int(ly)}, not after the as-of year {asof_y} - "
                            f"nothing can start in the past; move it to {asof_y + 1} or later")
            n = by.get(it.get("capex_years"), {}).get("base", 1) if it.get("capex") else 0
            if it.get("capex") and ly + n - 1 > last_y:
                warn.append(f"initiative {it['id']}: capex runs to {int(ly + n - 1)}, past the horizon ({last_y}) - "
                            f"that capex is not counted; extend horizon_years or shorten capex_years")
        if it.get("capex") and not it.get("asset_life"):
            warn.append(f"initiative {it['id']}: capex without asset_life - the terminal value then assumes no "
                        f"replacement capex; give asset_life")
        p = by.get(it["ramp"].get("p"))
        if p and p["low"] <= 0:
            raise SpecError(f"{p['id']}: Bass p must be > 0 across its range")
    for lf in leaves.values():
        if lf["base"] in by and by[lf["base"]]["base"] <= 0:
            raise SpecError(f"leaf {lf['id']}: base value must be > 0")
        if lf.get("ceiling") in by and lf["base"] in by and by[lf["ceiling"]]["base"] < by[lf["base"]]["base"]:
            warn.append(f"leaf {lf['id']}: base-year value already above its ceiling - check market size")
    used = set()
    for lf in leaves.values():
        used |= {lf.get(k) for k in ("base", "growth", "inflation", "ceiling", "ceiling_growth")}
    for st in streams:
        used.add(st.get("scale"))
    for it in inits:
        used |= set(init_input_ids(it))
    used |= set(fin.values())
    used |= set(s.get("_bridge_refs", []))
    used |= normalise_objectives(s, by, derived, inits, leaves, warn)
    used |= {"first_hist_year", "base_year", "growth_cap", "calib_tol", "sum_tol", "tie_step",
             "hist_p_low", "hist_p_high", "hist_tol", "backtest_tol"}
    for d in reversed(list(derived.values())):
        if d["id"] in used:
            used |= set(d["args"])
        else:
            warn.append(f"derived {d['id']}: defined but not used by the model")
    for i in inputs:
        if i["id"] not in used:
            warn.append(f"{i['id']}: defined but not used by the model - remove it or wire it in")
    s["_leaves"] = leaves
    s["_by"] = by
    return s, warn


# ----------------------------------------------------------------------------- builder
class Builder(ObjectivesMixin):
    HEAD_FONT, HEAD_FILL, BOLD, OUT_FILL = HEAD_FONT, HEAD_FILL, BOLD, OUT_FILL

    def __init__(self, spec, tornado_ids=None, two_way=None):
        self.s = spec
        self.by = spec["_by"]
        self.meta = spec["meta"]
        self.cur = self.meta["currency"]
        self.years = list(range(int(self.meta["first_hist_year"]),
                                int(self.meta["base_year"]) + int(self.meta["horizon_years"]) + 1))
        self.base_year = int(self.meta["base_year"])
        self.tornado_ids = tornado_ids
        self.two_way = two_way
        self.wb = Workbook()
        self.wb.remove(self.wb.active)
        self.sh = {}
        for name in ("Cover", "Objectives", "Frameworks", "Inputs", "Actuals", "Calculations", "Outputs", "Bridge",
                     "Sensitivity", "Checks",
                     "Driver tree", "Data requests"):
            self.sh[name] = self.wb.create_sheet(name)
        self.inrow = {}      # input id -> Inputs row
        self.live = {}       # input id -> 'Calculations!$E$r'
        self.rows = {}       # named Calculations rows
        self.crow = 8
        self.ref = {}        # named scalar cells for outputs
        self.record = {}     # positions for the second pass
        self.adrow = {}      # initiative id -> adoption row

    # -- column helpers
    def col(self, year):
        return FIRST_TL_COL + self.years.index(year)

    @property
    def base_col(self):
        return self.col(self.base_year)

    @property
    def last_col(self):
        return self.col(self.years[-1])

    @property
    def proj_years(self):
        return [y for y in self.years if y > self.base_year]

    def fmt_for(self, unit):
        u = (unit or "").lower()
        if u in ("%", "pct", "percent") or u.startswith("%"):
            return "0.0%"
        if u == "year":
            return "0"
        if u == "abs":
            return "0.0E+00"
        if self.cur.lower() in u or u in ("currency",):
            return "#,##0.0;[Red]-#,##0.0"
        return "#,##0.00;[Red]-#,##0.00"

    # -- sheet scaffolding
    def title(self, ws, text, sub=None):
        ws["B1"] = text
        ws["B1"].font = TITLE
        if sub:
            ws["B2"] = sub
            ws["B2"].font = Font(italic=True, color="595959")
        ws.column_dimensions["A"].width = 6
        ws.column_dimensions["B"].width = 52
        ws.column_dimensions["C"].width = 12
        ws.column_dimensions["D"].width = 14
        ws.column_dimensions["E"].width = 14
        ws.column_dimensions["F"].width = 4

    def timeline_header(self, ws, own=False):
        """Rows 3 (labels 2024A/2025E) and 4 (years) in the same columns on every sheet."""
        ws["B3"], ws["B4"] = "Period", "Year"
        for y in self.years:
            c = L(self.col(y))
            if own:
                if y == self.years[0]:
                    ws[f"{c}4"] = f"=Inputs!$F${self.inrow['first_hist_year']}"
                else:
                    ws[f"{c}4"] = f"={L(self.col(y) - 1)}4+1"
                ws[f"{c}3"] = f'={c}4&IF({c}4<=Inputs!$F${self.inrow["base_year"]},"A","E")'
            else:
                ws[f"{c}3"] = f"=Calculations!{c}3"
                ws[f"{c}4"] = f"=Calculations!{c}4"
            for r in (3, 4):
                ws[f"{c}{r}"].font = HEAD_FONT
                ws[f"{c}{r}"].fill = HEAD_FILL
                ws[f"{c}{r}"].alignment = Alignment(horizontal="center")
            ws[f"{c}4"].number_format = "0"
            ws.column_dimensions[c].width = 11
        for r in (3, 4):
            ws[f"B{r}"].font = HEAD_FONT
            ws[f"B{r}"].fill = HEAD_FILL

    def section(self, ws, r, text):
        ws[f"B{r}"] = text
        ws[f"B{r}"].font = BOLD
        for c in range(2, max(FIRST_TL_COL + len(self.years), 6)):
            ws.cell(r, c).fill = SEC_FILL

    # ----------------------------------------------------------------------- Inputs
    def build_inputs(self):
        ws = self.sh["Inputs"]
        self.title(ws, "Inputs - assumptions register",
                   "Every number used by the model, once. Blue = typed input. Grade: source reliability A-F, "
                   "information credibility 1-6 (A1 best).")
        heads = REG_HEAD
        widths = [6, 46, 12, 11, 11, 11, 11, 9, 46, 12, 40, 7, 11, 11, 11, 36, 40, 26, 14, 10]
        for j, (h, w) in enumerate(zip(heads, widths), 1):
            c = ws.cell(5, j, h)
            c.font, c.fill = HEAD_FONT, HEAD_FILL
            c.alignment = Alignment(wrap_text=True, vertical="center")
            ws.column_dimensions[L(j)].width = w
        ws.freeze_panes = "C6"
        used_by = self.used_by()
        order = ["Baseline drivers", "Market ceilings", "Revenue scaling"]
        order += [f"Initiatives: {p['id']}" for p in self.s["pillars"]]
        order += ["Costs", "Finance & WACC", "Model settings"]
        groups = {}
        for i in self.s["inputs"]:
            sec = i.get("section") or self.default_section(i)
            groups.setdefault(sec, []).append(i)
        r = 6
        n = 0
        pname = {f"Initiatives: {p['id']}": f"Initiatives - {p['name']}" for p in self.s["pillars"]}
        for sec in order + sorted(set(groups) - set(order)):
            if sec not in groups:
                continue
            ws.cell(r, 2, pname.get(sec, sec)).font = BOLD
            for c in range(1, len(REG_HEAD) + 1):
                ws.cell(r, c).fill = SEC_FILL
            r += 1
            for i in groups[sec]:
                n += 1
                i["_n"] = n
                self.inrow[i["id"]] = r
                vals = [n, i["name"], i["unit"], i["type"], i["low"], i["base"], i["high"], i["polarity"],
                        i.get("source", ""), i.get("date", ""), i.get("link", ""), i.get("grade", ""),
                        i.get("bench_low"), i.get("bench_median"), i.get("bench_high"), i.get("bench_source", ""),
                        i.get("justification", ""), used_by.get(i["id"], ""), i["id"], i.get("revision", "")]
                for j, v in enumerate(vals, 1):
                    c = ws.cell(r, j, v)
                    if j in (5, 6, 7, 13, 14, 15):
                        c.number_format = self.fmt_for(i["unit"])
                    if j in (5, 6, 7):
                        c.font, c.fill = BLUE, IN_FILL
                    if j in (9, 11, 16, 17):
                        c.alignment = Alignment(wrap_text=False)
                r += 1
        self.inputs_last = r - 1

    def default_section(self, i):
        sid = i["id"]
        fin = self.s["finance"]
        if sid in fin.values():
            return "Finance & WACC"
        for lf in self.s["_leaves"].values():
            if sid in (lf.get("ceiling"), lf.get("ceiling_growth")):
                return "Market ceilings"
            if sid in (lf.get("base"), lf.get("growth"), lf.get("inflation")):
                return "Baseline drivers"
        for st in self.s["streams"]:
            if sid == st.get("scale"):
                return "Revenue scaling"
        for it in self.s["_initiatives"]:
            if sid in ids_of(it.get("opex")) + ids_of(it.get("capex")) and \
                    not (it.get("model") == "build" and sid in ids_of(it["volume"]["drivers"])):
                return "Costs"
            if sid in init_input_ids(it):
                return f"Initiatives: {it['pillar']}"
        return "Other assumptions"

    def used_by(self):
        u = {}

        def tag(iid, who):
            if iid:
                u.setdefault(iid, [])
                if who not in u[iid]:
                    u[iid].append(who)
        for lf in self.s["_leaves"].values():
            for k in ("base", "growth", "inflation", "ceiling", "ceiling_growth"):
                tag(lf.get(k), f"leaf {lf['id']}")
        for it in self.s["_initiatives"]:
            for x in init_input_ids(it):
                tag(x, it["id"])
        for k, v in self.s["finance"].items():
            tag(v, "finance")
        return {k: ", ".join(v) for k, v in u.items()}

    # ----------------------------------------------------------------------- Actuals
    def build_actuals(self):
        ws = self.sh["Actuals"]
        self.title(ws, "Actuals - reported financials", "Base year comes from reported accounts, never a press summary.")
        heads = ["ID", "Item", "Unit", "Source", "Grade", "Date"]
        for j, h in enumerate(heads, 1):
            c = ws.cell(5, j, h)
            c.font, c.fill = HEAD_FONT, HEAD_FILL
        ws.column_dimensions["D"].width = 50
        ws.column_dimensions["E"].width = 7
        ws.column_dimensions["F"].width = 12
        for y in self.years:
            c = L(self.col(y))
            ws[f"{c}5"] = f"=Calculations!{c}3"
            ws[f"{c}5"].font, ws[f"{c}5"].fill = HEAD_FONT, HEAD_FILL
            ws.column_dimensions[c].width = 11
        self.actrow = {}
        r = 6
        sym = {"prod": "*", "sum": "+", "sub": "-", "div": "/"}
        for a in self.s.get("actuals", []):
            self.actrow[a["id"]] = r
            for j, v in enumerate([a["id"], a["name"], a.get("unit", self.cur), a.get("source", ""),
                                   a.get("grade", ""), a.get("date", "")], 1):
                ws.cell(r, j, v)
            for y, v in a.get("values", {}).items():
                y = int(y)
                if y in self.years and y <= self.base_year:
                    if a.get("op"):   # derived series: a live formula on the reported rows
                        cl = L(self.col(y))
                        c = ws.cell(r, self.col(y), "=" + sym[a["op"]].join(f"{cl}{self.actrow[x]}" for x in a["args"]))
                    else:
                        c = ws.cell(r, self.col(y), v)
                        c.font, c.fill = BLUE, IN_FILL
                    c.number_format = self.fmt_for(a.get("unit", self.cur))
            r += 1
        self.shockrow = None
        if self.s.get("_shock"):
            r += 1
            self.shockrow = r
            for j, v in enumerate(["shock_year", "Shock year (1 = left out of history statistics and the history band)",
                                   "flag", "Modelling choice: " + self.meta.get("shock_note", "shock years"), "A1",
                                   self.meta.get("as_of", "")], 1):
                ws.cell(r, j, v)
            for y in self.years:
                if y <= self.base_year:
                    c = ws.cell(r, self.col(y), 1 if y in self.s["_shock"] else 0)
                    c.font, c.fill = BLUE, IN_FILL

    # ----------------------------------------------------------------------- Calculations helpers
    def L_(self, key):
        return self.live[key]

    def nm(self, iid):
        x = self.by.get(iid) or self.s["_derived"].get(iid) or {}
        return x.get("name", iid)

    def put_row(self, label, unit, fn, cols="proj", name=None, ident=None, fmt=None, sheet="Calculations"):
        """Write one time-series row. fn(col_letter, prev_letter) -> formula string or None."""
        ws = self.sh[sheet]
        r = self.crow if sheet == "Calculations" else self.orow
        ws.cell(r, 2, label)
        ws.cell(r, 3, unit)
        if ident:
            ws.cell(r, 1, ident)
        if cols == "proj":
            years = self.proj_years
        elif cols == "base+":
            years = [y for y in self.years if y >= self.base_year]
        elif cols == "hist":
            years = [y for y in self.years if y <= self.base_year]
        else:
            years = self.years
        for y in years:
            c = self.col(y)
            f = fn(L(c), L(c - 1), y)
            if f is None:
                continue
            cell = ws.cell(r, c, f)
            cell.number_format = fmt or self.fmt_for(unit)
        if name:
            self.rows[name] = r if sheet == "Calculations" else (sheet, r)
        if sheet == "Calculations":
            self.crow += 1
        else:
            self.orow += 1
        return r

    def put_scalar(self, label, unit, formula, name, sheet="Calculations"):
        ws = self.sh[sheet]
        r = self.crow if sheet == "Calculations" else self.orow
        ws.cell(r, 2, label)
        ws.cell(r, 3, unit)
        c = ws.cell(r, 5, formula)
        c.number_format = self.fmt_for(unit)
        c.font = BOLD
        self.ref[name] = f"{sheet}!$E${r}"
        if sheet == "Calculations":
            self.crow += 1
        else:
            self.orow += 1
        return r

    def blank(self, n=1, sheet="Calculations"):
        if sheet == "Calculations":
            self.crow += n
        else:
            self.orow += n

    def csec(self, text, sheet="Calculations"):
        r = self.crow if sheet == "Calculations" else self.orow
        self.section(self.sh[sheet], r, text)
        self.blank(1, sheet)

    def R(self, name, col):
        return f"{col}${self.rows[name]}" if False else f"{col}{self.rows[name]}"

    # ----------------------------------------------------------------------- Calculations
    def build_calcs(self):
        ws = self.sh["Calculations"]
        self.title(ws, "Calculations", "No typed numbers on this sheet: every cell is a formula fed by Inputs, "
                                       "Actuals or the Sensitivity controls.")
        self.timeline_header(ws, own=True)
        ws["B5"] = "Years from base year (t)"
        for y in self.years:
            c = L(self.col(y))
            ws[f"{c}5"] = f"={c}4-Inputs!$F${self.inrow['base_year']}"
            ws[f"{c}5"].number_format = "0"
        ws.freeze_panes = ws.cell(6, FIRST_TL_COL)
        self.crow = 7
        self.build_history()
        # live inputs
        self.csec("1. Live inputs (Base case unless a scenario, tornado or sensitivity switch is set)")
        sens = "Sensitivity"
        scen, torn = f"{sens}!$E$5", f"{sens}!$E$6"
        self.ovr = {}
        ovr_ids = []
        for a, b in (self.two_way or []):
            for x in (a, b):
                if x not in ovr_ids:
                    ovr_ids.append(x)
        for k, iid in enumerate(ovr_ids):
            self.ovr[iid] = f"{sens}!$E${7 + k}"
        for i in self.s["inputs"]:
            r = self.inrow[i["id"]]
            lo, ba, hi, po, idc = (f"Inputs!${c}${r}" for c in ("E", "F", "G", "H", "A"))
            if i["type"] in EXEMPT:
                f = f"={ba}"
            else:
                core = f"CHOOSE({scen},IF({po}>0,{lo},{hi}),{ba},IF({po}>0,{hi},{lo}))"
                f = f"IF({torn}={idc},{hi},IF({torn}=-{idc},{lo},{core}))"
                if i["id"] in self.ovr:
                    o = self.ovr[i["id"]]
                    f = f"IF(ISNUMBER({o}),{o},{f})"
                f = "=" + f
            row = self.crow
            ws.cell(row, 1, i["id"])
            ws.cell(row, 2, i["name"])
            ws.cell(row, 3, i["unit"])
            ws.cell(row, 4, f"Inputs row {r}")
            c = ws.cell(row, 5, f)
            c.number_format = self.fmt_for(i["unit"])
            self.live[i["id"]] = f"$E${row}"
            self.crow += 1
        self.blank()
        if self.s["_derived"]:
            self.csec("1b. Derived inputs (formulas on inputs, so they move with their parents in every scenario)")
            sym = {"prod": "*", "sum": "+", "sub": "-", "div": "/"}
            for d in self.s["_derived"].values():
                row = self.crow
                ws.cell(row, 1, d["id"])
                ws.cell(row, 2, d["name"])
                ws.cell(row, 3, d.get("unit", ""))
                ws.cell(row, 4, f"{d['op']}: " + ", ".join(d["args"]))
                c = ws.cell(row, 5, "=" + sym[d["op"]].join(self.live[x] for x in d["args"]))
                c.number_format = self.fmt_for(d.get("unit", ""))
                self.live[d["id"]] = f"$E${row}"
                self.crow += 1
            self.blank()
        # WACC
        fin = self.s["finance"]
        lv = lambda k: self.live[fin[k]]
        self.csec("2. Discount rate (WACC from sourced inputs)")
        ke = f"={lv('rf')}+{lv('beta')}*{lv('erp')}" + (f"+{lv('crp')}" if fin.get("crp") else "")
        self.put_scalar("Cost of equity (CAPM" + (" + country risk premium)" if fin.get("crp") else ")"), "%", ke, "ke")
        self.put_scalar("After-tax cost of debt", "%", f"={lv('kd')}*(1-{lv('tax')})", "kd_at")
        self.put_scalar("WACC", "%", f"=(1-{lv('debt_weight')})*{self.local('ke')}+{lv('debt_weight')}*"
                                     f"{self.local('kd_at')}", "wacc")
        self.blank()
        # leaves
        self.csec("3. Driver tree leaves: baseline, ceiling, initiative uplift (S-curve x decay), with strategy")
        self.leafrows = {}
        for lf in self.s["_leaves"].values():
            self.build_leaf(lf)
        # streams
        self.csec("4. Revenue streams = product of leaves; incremental revenue split by leaf (log shares)")
        self.streamrows = {}
        for st in self.s["streams"]:
            self.build_stream(st)
        # by initiative
        inits = self.s["_initiatives"]
        self.csec("5. Incremental revenue by initiative (each leaf's share split pro rata to its uplifts)")
        self.initrev = {}
        for it in inits:
            if it.get("model") == "build":
                if it["kind"] == "revenue":
                    v = self.buildrows[it["id"]]["value"]
                    self.initrev[it["id"]] = self.put_row(f"Incremental revenue: {it['name']}", self.cur,
                                                          lambda c, p, y, v=v: f"={c}{v}", ident=it["id"])
                continue
            lf = it["leaf"]
            if lf is None:
                continue
            parts = []
            for st in self.s["streams"]:
                if lf in st["leaves"]:
                    via = self.streamrows[st["id"]]["via"][lf]
                    tot = self.leafrows[lf]["total_uplift"]
                    up = self.leafrows[lf]["uplift"][it["id"]]
                    parts.append((via, tot, up))

            def fn(c, p, y, parts=parts):
                terms = [f"IF({c}{t}=0,0,{c}{v}*{c}{u}/{c}{t})" for v, t, u in parts]
                return "=" + ("+".join(terms) if terms else "0")
            self.initrev[it["id"]] = self.put_row(f"Incremental revenue: {it['name']}", self.cur, fn, ident=it["id"])
        rr = list(self.initrev.values())
        self.rows["inc_total"] = self.put_row("Total incremental revenue", self.cur,
                                              lambda c, p, y: f"=SUM({c}{rr[0]}:{c}{rr[-1]})" if rr else "=0")
        self.blank()
        self.csec("6. Incremental revenue by pillar")
        self.pillrev = {}
        for pl in self.s["pillars"]:
            rows = [self.initrev[it["id"]] for it in inits if it["pillar"] == pl["id"] and it["id"] in self.initrev]
            self.pillrev[pl["id"]] = self.put_row(
                f"Incremental revenue: {pl['name']}", self.cur,
                lambda c, p, y, rows=rows: "=" + ("+".join(f"{c}{r}" for r in rows) if rows else "0"),
                ident=pl["id"])
        fr, lr = self.pillrev[self.s["pillars"][0]["id"]], self.pillrev[self.s["pillars"][-1]["id"]]
        self.rows["pill_total"] = self.put_row("Total incremental revenue (pillar view)", self.cur,
                                               lambda c, p, y: f"=SUM({c}{fr}:{c}{lr})")
        self.blank()
        savers = [it for it in inits if it.get("model") == "build" and it["kind"] == "saving"]
        self.saverow = {}
        if savers:
            self.csec("6b. Cost savings by initiative (avoided cost: flows straight to EBIT)")
            for it in savers:
                v = self.buildrows[it["id"]]["value"]
                self.saverow[it["id"]] = self.put_row(f"Cost saving: {it['name']}", self.cur,
                                                      lambda c, p, y, v=v: f"={c}{v}", ident=it["id"])
            sr = list(self.saverow.values())
            self.rows["savings"] = self.put_row("Total cost savings", self.cur,
                                                lambda c, p, y: f"=SUM({c}{sr[0]}:{c}{sr[-1]})")
            self.blank()
        # costs
        self.csec("7. Initiative costs (opex from launch; capex one-off in the launch year; quantity x unit cost)")
        opex_rows, capex_rows = [], []
        self.opexrow, self.capexrow, self.deprow = {}, {}, {}
        for it in inits:
            if it.get("opex"):
                o, ly = self.prod(it["opex"]), self.live[it["launch_year"]]
                og = self.live.get(it.get("opex_growth"))
                gm = f"*(1+{og})^({{c}}$4-{ly})" if og else ""
                if it.get("opex_ramp") and it["id"] in self.adrow:
                    gm += f"*{{c}}{self.adrow[it['id']]}"
                opex_rows.append(self.put_row(f"Opex: {it['name']}", self.cur,
                                              lambda c, p, y, o=o, ly=ly, gm=gm: f"=IF({c}$4>={ly},{o}{gm.format(c=c)},0)",
                                              ident=it["id"]))
                self.opexrow[it["id"]] = opex_rows[-1]
        if opex_rows:
            self.rows["opex"] = self.put_row("Total opex", self.cur,
                                             lambda c, p, y: f"=SUM({c}{opex_rows[0]}:{c}{opex_rows[-1]})")
        self.blank()
        for it in inits:
            if it.get("capex"):
                k, ly = self.prod(it["capex"]), self.live[it["launch_year"]]
                n = self.live.get(it.get("capex_years"))
                if n:
                    fn = lambda c, p, y, k=k, ly=ly, n=n: f"=IF(AND({c}$4>={ly},{c}$4<{ly}+{n}),{k}/{n},0)"
                else:
                    fn = lambda c, p, y, k=k, ly=ly: f"=IF({c}$4={ly},{k},0)"
                capex_rows.append(self.put_row(f"Capex: {it['name']}", self.cur, fn, ident=it["id"]))
                self.capexrow[it["id"]] = capex_rows[-1]
        if capex_rows:
            self.rows["capex"] = self.put_row("Total capex", self.cur,
                                              lambda c, p, y: f"=SUM({c}{capex_rows[0]}:{c}{capex_rows[-1]})")
        self.blank()
        deps = [it for it in inits if it.get("asset_life") and it["id"] in self.capexrow]
        if deps:
            self.csec("7b. Depreciation of initiative capex (straight line; tax shield only, not cash)")
            cums = {}
            for it in deps:
                cx = self.capexrow[it["id"]]
                cums[it["id"]] = self.put_row(f"Cumulative capex: {it['name']}", self.cur,
                                              lambda c, p, y, cx=cx: "=0" if y == self.base_year
                                              else f"={p}{self.crow}+{c}{cx}", cols="base+", ident=it["id"])
            self.blank()
            bcl = L(self.base_col)
            for it in deps:
                life, cum = self.live[it["asset_life"]], cums[it["id"]]
                self.deprow[it["id"]] = self.put_row(
                    f"Depreciation (straight line, never above capex spent): {it['name']}", self.cur,
                    lambda c, p, y, cum=cum, life=life: "=0" if y == self.base_year else
                    f"=MIN({c}{cum}/{life},{c}{cum}-SUM(${bcl}{self.crow}:{p}{self.crow}))", cols="base+",
                    ident=it["id"])
            dr = list(self.deprow.values())
            self.rows["dep"] = self.put_row("Total depreciation of initiative capex", self.cur,
                                            lambda c, p, y: f"=SUM({c}{dr[0]}:{c}{dr[-1]})")
            self.blank()
        # cash flow & valuation
        m, tax = self.live[fin["margin"]], self.live[fin["tax"]]
        it_ = self.rows["inc_total"]
        self.csec("8a. Incremental EBIT by initiative (revenue x its margin + cost saving - opex)")
        self.initcon, self.initebit = {}, {}
        for it in inits:
            mi = self.live[it["margin"]] if it.get("margin") else m
            parts = []
            if it["id"] in self.initrev:
                parts.append(f"{{c}}{self.initrev[it['id']]}*{mi}")
            if it["id"] in self.saverow:
                parts.append(f"{{c}}{self.saverow[it['id']]}")
            expr = "+".join(parts) or "0"
            self.initcon[it["id"]] = self.put_row(f"Contribution: {it['name']}", self.cur,
                                                  lambda c, p, y, expr=expr: "=" + expr.format(c=c), ident=it["id"])
        self.blank()
        for it in inits:
            o = self.opexrow.get(it["id"])
            cr = self.initcon[it["id"]]
            self.initebit[it["id"]] = self.put_row(f"Incremental EBIT: {it['name']}", self.cur,
                                                   lambda c, p, y, cr=cr, o=o: f"={c}{cr}" + (f"-{c}{o}" if o else ""),
                                                   ident=it["id"])
        self.blank()
        self.pillebit = {}
        for pl in self.s["pillars"]:
            rows = [self.initebit[it["id"]] for it in inits if it["pillar"] == pl["id"]]
            self.pillebit[pl["id"]] = self.put_row(
                f"Incremental EBIT: {pl['name']}", self.cur,
                lambda c, p, y, rows=rows: "=" + ("+".join(f"{c}{r}" for r in rows) if rows else "0"), ident=pl["id"])
        pe = list(self.pillebit.values())
        self.rows["ebit_pillars"] = self.put_row("Total incremental EBIT (pillar view)", self.cur,
                                                 lambda c, p, y: f"=SUM({c}{pe[0]}:{c}{pe[-1]})")
        self.blank()
        self.csec("8. Incremental profit, cash flow and valuation")
        ic = list(self.initcon.values())
        self.rows["contrib"] = self.put_row("Incremental contribution (revenue x margin + cost savings)", self.cur,
                                            lambda c, p, y: f"=SUM({c}{ic[0]}:{c}{ic[-1]})")
        ox = (lambda c: f"-{c}{self.rows['opex']}") if "opex" in self.rows else (lambda c: "")
        self.rows["ebit"] = self.put_row("Incremental EBIT", self.cur,
                                         lambda c, p, y: f"={c}{self.rows['contrib']}{ox(c)}")
        dp = (lambda c: f"-{c}{self.rows['dep']}") if "dep" in self.rows else (lambda c: "")
        self.rows["taxr"] = self.put_row("Tax on incremental EBIT" + (" less depreciation" if "dep" in self.rows else ""),
                                         self.cur, lambda c, p, y: f"=({c}{self.rows['ebit']}{dp(c)})*{tax}")
        self.rows["nopat"] = self.put_row("Incremental NOPAT", self.cur,
                                          lambda c, p, y: f"={c}{self.rows['ebit']}-{c}{self.rows['taxr']}")
        cx = (lambda c: f"-{c}{self.rows['capex']}") if "capex" in self.rows else (lambda c: "")
        self.rows["fcf"] = self.put_row("Incremental free cash flow", self.cur,
                                        lambda c, p, y: f"={c}{self.rows['nopat']}{cx(c)}")
        w = self.local("wacc")
        self.rows["df"] = self.put_row("Discount factor (end of year)", "x", lambda c, p, y: f"=1/(1+{w})^{c}$5",
                                       fmt="0.000")
        self.rows["pv"] = self.put_row("PV of incremental free cash flow", self.cur,
                                       lambda c, p, y: f"={c}{self.rows['fcf']}*{c}{self.rows['df']}")
        self.rows["cum"] = self.put_row("Cumulative incremental free cash flow", self.cur,
                                        lambda c, p, y: "=0" if y == self.base_year else
                                        f"={p}{self.crow}+{c}{self.rows['fcf']}", cols="base+")
        cum, fcf = self.rows["cum"], self.rows["fcf"]
        self.rows["payb"] = self.put_row("Payback crossing (years after base year)", "years",
                                         lambda c, p, y: f"=IF(AND({c}{cum}>=0,{p}{cum}<0),({c}$5-1)+(-{p}{cum})/"
                                                         f"{c}{fcf},0)", fmt="0.0")
        self.blank()
        fc, lc = L(self.col(self.proj_years[0])), L(self.last_col)
        g, sw = self.live[fin["terminal_growth"]], self.live[fin["terminal_switch"]]
        self.tv_g = g
        self.put_scalar("Horizon-only NPV (no terminal value)", self.cur,
                        f"=SUM({fc}{self.rows['pv']}:{lc}{self.rows['pv']})", "npv_h")
        # terminal value on a normalised final year: growth capex is one-off, replacement capex = depreciation
        dep = self.rows.get("dep")
        if "capex" in self.rows:
            self.put_scalar(f"Normalised final-year FCF for the terminal value ({self.years[-1]} FCF + capex - "
                            f"replacement capex = depreciation)", self.cur,
                            f"={lc}{fcf}+{lc}{self.rows['capex']}" + (f"-{lc}{dep}" if dep else ""), "fcf_tv")
            tvb = self.local("fcf_tv")
        else:
            tvb = f"{lc}{fcf}"
        self.put_scalar("PV of terminal value (switch x Gordon growth on normalised final-year FCF)", self.cur,
                        f"={sw}*{tvb}*(1+{g})/({w}-{g})*{lc}{self.rows['df']}", "tv")
        self.put_scalar("NPV of strategy (incremental, at WACC)", self.cur,
                        f"={self.local('npv_h')}+{self.local('tv')}", "npv")
        self.put_scalar("IRR of strategy", "%", f'=IFERROR(IRR({fc}{fcf}:{lc}{fcf}),"n/a")', "irr")
        self.put_scalar("Payback (years after base year)", "years",
                        f'=IF(MIN({fc}{cum}:{lc}{cum})>=0,"no investment phase (cash flow never negative)",'
                        f'IF({lc}{cum}<0,"beyond horizon",'
                        f'SUM({fc}{self.rows["payb"]}:{lc}{self.rows["payb"]})))', "payback")
        self.put_scalar("Cumulative incremental revenue over horizon", self.cur,
                        f"=SUM({fc}{it_}:{lc}{it_})", "cum_rev")
        self.put_scalar(f"Incremental revenue in {self.years[-1]}", self.cur, f"={lc}{it_}", "rev_final")
        self.put_scalar(f"Incremental EBIT in {self.years[-1]}", self.cur, f"={lc}{self.rows['ebit']}", "ebit_final")
        self.blank()
        self.csec("9. Value by initiative (standalone cash flow; initiatives add up to the strategy NPV)")
        self.initnpv = {}
        dfr = self.rows["df"]
        for it in inits:
            e, d, k = self.initebit[it["id"]], self.deprow.get(it["id"]), self.capexrow.get(it["id"])
            fr = self.put_row(f"Free cash flow: {it['name']}", self.cur,
                              lambda c, p, y, e=e, d=d, k=k: f"={c}{e}-({c}{e}" + (f"-{c}{d}" if d else "") +
                              f")*{tax}" + (f"-{c}{k}" if k else ""), ident=it["id"])
            tvf = f"({lc}{fr}" + (f"+{lc}{k}" if k else "") + (f"-{lc}{d}" if d else "") + ")"
            r = self.put_scalar(f"NPV: {it['name']}", self.cur,
                                f"=SUMPRODUCT({fc}{fr}:{lc}{fr},{fc}{dfr}:{lc}{dfr})+{sw}*{tvf}*(1+{g})/({w}-{g})*"
                                f"{lc}{dfr}", f"npv_{it['id']}")
            self.sh["Calculations"].cell(r, 1, it["id"])
            self.initnpv[it["id"]] = r
        nr_ = list(self.initnpv.values())
        self.put_scalar("Sum of initiative NPVs", self.cur, "=" + "+".join(f"E{r}" for r in nr_), "npv_sum")
        self.blank()
        bc = L(self.base_col)
        self.put_scalar(f"Baseline revenue in {self.years[-1]}", self.cur, f"={lc}{self.rows['base_total']}",
                        "base_final")
        self.put_scalar(f"Target revenue in {self.years[-1]}", self.cur, f"={lc}{self.rows['target_total']}",
                        "target_final")
        self.put_scalar(f"Model revenue in base year {self.base_year}", self.cur, f"={bc}{self.rows['base_total']}",
                        "model_base")
        rev = self.revenue_actual()
        if rev:
            self.put_scalar(f"Reported revenue in base year {self.base_year}", self.cur,
                            f"=Actuals!{bc}{self.actrow[rev['id']]}", "reported_base")
        self.blank()
        self.build_history_tests()
        self.build_objectives_calc()
        for k in ("npv", "irr", "payback"):
            self.sh["Calculations"][self.ref[k].split("!")[1].replace("$", "")].fill = OUT_FILL

    def local(self, name):
        return self.ref[name].split("!")[1]

    # ----------------------------------------------------------------------- History
    def hist_cols(self, a):
        return [y for y in self.years if y <= self.base_year and str(y) in {str(k) for k in a.get("values", {})}]

    def build_history(self):
        """Section 0: every reported series -> YoY growth, CAGR, quartiles, margins. History-linked inputs and the
        baseline-vs-history check read these cells, so the base case moves when the actuals are updated."""
        acts = [a for a in self.s.get("actuals", []) if len(self.hist_cols(a)) >= 2]
        if not acts:
            return
        self.csec("0. Reported history: growth, CAGR, quartiles and margins (feeds history-linked inputs)")
        I = "Inputs!"
        plo, phi = (f"{I}$F${self.inrow[k]}" if k in self.inrow else None for k in ("hist_p_low", "hist_p_high"))
        rev = self.revenue_actual()
        st_all = self.s["_hist"]
        h0, h1 = L(self.col(self.years[0])), L(self.base_col)
        for a in acts:
            ar, ys, st = self.actrow[a["id"]], self.hist_cols(a), st_all[a["id"]]
            A = lambda c: f"Actuals!{c}{ar}"
            gr = None
            sk = self.shockrow
            S = lambda c: f"Actuals!{c}${sk}"
            if st["growth_all"]:
                gr = self.put_row(f"{a['name']}: YoY growth", "%",
                                  lambda c, p, y: f"={A(c)}/{A(p)}-1" if y in st["growth_all"] else None,
                                  cols="hist", ident=a["id"])
                if sk:   # the statistics read a copy without the trough and the rebound
                    g0 = gr
                    gr = self.put_row(f"{a['name']}: YoY growth excluding shock years", "%",
                                      lambda c, p, y: f'=IF(OR({S(c)}=1,{S(p)}=1),"",{c}{g0})'
                                      if y in st["growth_all"] else None, cols="hist", ident=a["id"])
            y0, y1 = st["cagr_years"]
            lc = L(self.col(ys[-1]))
            if "cagr" in st:
                f0, f1 = L(self.col(y0)), L(self.col(y1))
                self.put_scalar(f"{a['name']}: CAGR {y0}-{y1}" + (" (first and last non-shock years)" if sk else ""),
                                "%", f"=({A(f1)}/{A(f0)})^(1/({f1}$4-{f0}$4))-1", f"h_{a['id']}_cagr")
            if gr and plo:
                gy = sorted(st["growth_all"])
                rg = f"{L(self.col(gy[0]))}{gr}:{L(self.col(gy[-1]))}{gr}"
                self.put_scalar(f"{a['name']}: median YoY growth", "%", f"=MEDIAN({rg})", f"h_{a['id']}_median_growth")
                self.put_scalar(f"{a['name']}: YoY growth, lower percentile", "%", f"=PERCENTILE({rg},{plo})",
                                f"h_{a['id']}_g_p25")
                self.put_scalar(f"{a['name']}: YoY growth, upper percentile", "%", f"=PERCENTILE({rg},{phi})",
                                f"h_{a['id']}_g_p75")
            if rev and st.get("margin") and plo:
                rr = self.actrow[rev["id"]]
                mr = self.put_row(f"{a['name']}: % of revenue", "%",
                                  lambda c, p, y: (f'=IF({S(c)}=1,"",{A(c)}/Actuals!{c}{rr})' if sk else
                                                   f"={A(c)}/Actuals!{c}{rr}") if y in self.hist_cols(rev)
                                  and y in self.hist_cols(a) else None, cols="hist", ident=a["id"])
                my = [y for y in ys if y in self.hist_cols(rev)]
                rg = f"{L(self.col(my[0]))}{mr}:{L(self.col(my[-1]))}{mr}"
                self.put_scalar(f"{a['name']}: median % of revenue", "%", f"=MEDIAN({rg})", f"h_{a['id']}_median_margin")
                self.put_scalar(f"{a['name']}: % of revenue, lower percentile", "%", f"=PERCENTILE({rg},{plo})",
                                f"h_{a['id']}_m_p25")
                self.put_scalar(f"{a['name']}: % of revenue, upper percentile", "%", f"=PERCENTILE({rg},{phi})",
                                f"h_{a['id']}_m_p75")
            self.ref[f"h_{a['id']}_last"] = f"Actuals!${lc}${ar}"
            self.blank()

    def link_history_inputs(self):
        """Inputs that take their value from history: Base (and Low/High unless typed) become live links."""
        ws = self.sh["Inputs"]
        lohi = {"last": (None, None), "median_margin": ("m_p25", "m_p75")}
        for i in self.s["inputs"]:
            h = i.get("_hist")
            if not h:
                continue
            r, sid, stat = self.inrow[i["id"]], h["series"], h["stat"]
            base = self.ref[f"h_{sid}_{stat}"]
            lo, hi = lohi.get(stat, ("g_p25", "g_p75"))
            cells = {6: f"={base}"}
            if lo and h["low"]:
                cells[5] = f"=MIN({base},{self.ref[f'h_{sid}_{lo}']})"
            if hi and h["high"]:
                cells[7] = f"=MAX({base},{self.ref[f'h_{sid}_{hi}']})"
            if stat == "last":
                cells[5] = cells[7] = f"={base}"
            for c, f in cells.items():
                x = ws.cell(r, c, f)
                x.font, x.fill = Font(color="000000"), PatternFill()

    def build_history_tests(self):
        """Baseline growth vs the historical band, and the back-test of the driver tree on reported history."""
        rev = self.revenue_actual()
        I = "Inputs!"
        bc, lc = L(self.base_col), L(self.last_col)
        if rev and f"h_{rev['id']}_g_p25" in self.ref and f"h_{rev['id']}_cagr" in self.ref:
            self.csec("9b. Base case vs history (do-nothing growth against the company's own record)")
            bt = self.rows["base_total"]
            self.put_scalar(f"Baseline revenue CAGR {self.base_year}-{self.years[-1]} (do nothing)", "%",
                            f"=({lc}{bt}/{bc}{bt})^(1/({lc}$4-{bc}$4))-1", "base_cagr")
            cg, p25, p75 = (self.ref[f"h_{rev['id']}_{k}"] for k in ("cagr", "g_p25", "g_p75"))
            tol = f"{I}$F${self.inrow['hist_tol']}"
            self.put_scalar("History band, low: lower of revenue CAGR and lower-percentile YoY growth, less tolerance",
                            "%", f"=MIN({cg},{p25})-{tol}", "hist_lo")
            self.put_scalar("History band, high: higher of revenue CAGR and upper-percentile YoY growth, plus tolerance",
                            "%", f"=MAX({cg},{p75})+{tol}", "hist_hi")
            br = self.meta.get("baseline_bridge")
            if br:
                self.blank()
                frm = br.get("from", "cagr")
                src = self.ref["h_" + rev["id"] + "_" + frm]
                start = self.put_scalar(f"Bridge to history: start - reported revenue {HIST_LABEL[frm]}", "%",
                                        f"={src}", "br_start")

                def rf(x):
                    if str(x).startswith("hist:"):
                        _, sid, stat = str(x).split(":")
                        return self.ref[f"h_{sid}_{stat}"]
                    return self.live[x]
                last = start
                for k, item in enumerate(br["items"]):
                    f = ("".join(f"+{rf(x)}" for x in item.get("plus") or []) +
                         "".join(f"-{rf(x)}" for x in item.get("minus") or [])).lstrip("+")
                    if item.get("weight"):   # a step on one stream counts at that stream's share of revenue
                        f = f"({f})*{rf(item['weight'])}"
                    last = self.put_scalar(f"Bridge to history: {item['name']}", "%", "=" + f, f"br_{k}")
                self.put_scalar("Bridged do-nothing growth (start + steps; additive approximation)", "%",
                                f"=SUM(E{start}:E{last})", "br_end")
                self.put_scalar("Bridge gap: bridged growth vs baseline revenue CAGR (absolute)", "%",
                                f"=ABS({self.local('br_end')}-{self.local('base_cagr')})", "br_gap")
            self.blank()
        bts = self.s.get("_backtest") or []
        if not bts or "backtest_tol" not in self.inrow:
            return
        drv = self.s["_drivers"]
        aby = {a["id"]: a for a in self.s["actuals"]}
        stream_act = {a["stream"]: a for a in self.s["actuals"] if a.get("stream")}
        self.csec("9c. Back-test: the driver tree run on reported driver history vs reported revenue")
        gaps, btrows = [], {}
        for sid in bts:
            st = next(x for x in self.s["streams"] if x["id"] == sid)
            yrs = set.intersection(*(set(self.hist_cols(aby[drv[x]])) for x in st["leaves"]))
            sm = f"*{self.live[st['scale']]}" if st.get("scale") else ""
            btrows[sid] = (self.put_row(
                f"Back-test: {st['name']} from driver history", self.cur,
                lambda c, p, y, st=st, yrs=yrs: ("=" + "*".join(f"Actuals!{c}{self.actrow[drv[x]]}"
                                                               for x in st["leaves"]) + sm) if y in yrs else None,
                cols="hist", ident=sid), yrs)
        comps = []   # (label, back-test rows, actual series, years)
        for sid, (r, yrs) in btrows.items():
            if sid in stream_act:
                a = stream_act[sid]
                comps.append((a["name"], [r], a, yrs & set(self.hist_cols(a))))
        if rev and len(bts) == len(self.s["streams"]):
            yrs = set.intersection(*(y for _, y in btrows.values())) & set(self.hist_cols(rev))
            comps.append((rev["name"], [r for r, _ in btrows.values()], rev, yrs))
        for name, rows, a, yrs in comps:
            if not yrs:
                continue
            ar = self.actrow[a["id"]]
            gaps.append(self.put_row(
                f"Back-test gap (%) vs {name}", "%",
                lambda c, p, y, rows=rows, ar=ar, yrs=yrs: f"=ABS(({'+'.join(f'{c}{x}' for x in rows)})/Actuals!{c}{ar}-1)"
                if y in yrs else None, cols="hist"))
        if gaps:
            h0, h1 = L(self.col(self.years[0])), L(self.base_col)
            self.put_scalar("Back-test: largest yearly gap vs reported", "%",
                            "=MAX(" + ",".join(f"{h0}{g}:{h1}{g}" for g in gaps) + ")", "bt_max")
        self.blank()

    def revenue_actual(self):
        acts = self.s.get("actuals", [])
        return next((a for a in acts if a.get("role") == "revenue"), None)

    def build_leaf(self, lf):
        cur = self.crow
        lid = lf["id"]
        info = {"uplift": {}}
        b = self.live[lf["base"]]
        g = self.live.get(lf.get("growth"))
        inf = self.live.get(lf.get("inflation"))
        mult = "".join(f"*(1+{x})" for x in (g, inf) if x)
        unit = lf.get("unit", "")
        info["base"] = self.put_row(f"Baseline: {lf['name']}", unit,
                                    lambda c, p, y: f"={b}" if y == self.base_year else f"={p}{self.crow}{mult}",
                                    cols="base+", ident=lid)
        if lf.get("ceiling"):
            cb = self.live[lf["ceiling"]]
            cg = self.live.get(lf.get("ceiling_growth"))
            cm = f"*(1+{cg})" if cg else ""
            info["ceiling"] = self.put_row(f"Ceiling: addressable {lf['name']} (market / capacity)", unit,
                                           lambda c, p, y: f"={cb}" if y == self.base_year else f"={p}{self.crow}{cm}",
                                           cols="base+")
        owners = [it for it in self.s["_initiatives"] if it["leaf"] == lid]
        for it in owners:
            ad, ret = self.ramp_rows(it)
            u = self.prod(it["uplift"])
            info["uplift"][it["id"]] = self.put_row(
                f"Uplift on {lf['name']} from {it['name']} (%)", "%",
                lambda c, p, y, u=u, ad=ad, ret=ret: f"={u}*{c}{ad}" + (f"*{c}{ret}" if ret else ""),
                ident=it["id"])
        if owners:
            ups = list(info["uplift"].values())
            info["total_uplift"] = self.put_row(f"Total uplift on {lf['name']} (%)", "%",
                                                lambda c, p, y: "=" + "+".join(f"{c}{r}" for r in ups))
        br = info["base"]

        def with_fn(c, p, y):
            if y == self.base_year or not owners:
                f = f"{c}{br}"
            else:
                f = f"{c}{br}*(1+{c}{info['total_uplift']})"
            if "ceiling" in info:
                f = f"MIN({f},{c}{info['ceiling']})"
            return "=" + f
        info["with"] = self.put_row(f"{lf['name']} with strategy", unit, with_fn, cols="base+", ident=lid)
        if "ceiling" in info:
            self.put_row(f"Baseline penetration of addressable market (% of ceiling): {lf['name']}", "%",
                         lambda c, p, y: f"={c}{br}/{c}{info['ceiling']}", cols="base+")
            info["pen"] = self.put_row(f"Penetration of addressable market (% of ceiling): {lf['name']}", "%",
                                       lambda c, p, y: f"={c}{info['with']}/{c}{info['ceiling']}", cols="base+")
        if owners:
            info["factor"] = self.put_row(f"Effective uplift factor after ceiling: {lf['name']}", "x",
                                          lambda c, p, y: f"={c}{info['with']}/{c}{br}", fmt="0.0000")
        self.leafrows[lid] = info
        self.blank()

    def ramp_rows(self, it):
        """Adoption (S-curve) row and, if a decay is defined, the retention row. Returns (ad, ret)."""
        ly = self.live[it["launch_year"]]
        t = f"({{c}}$4-{ly}+1)"
        rp = it["ramp"]
        if rp["type"] == "bass":
            p_, q_ = self.live[rp["p"]], self.live[rp["q"]]
            expr = (f"(1-EXP(-({p_}+{q_})*{t}))/(1+({q_}/{p_})*EXP(-({p_}+{q_})*{t}))")
        elif rp["type"] == "logistic":
            m_, k_ = self.live[rp["midpoint"]], self.live[rp["steepness"]]
            expr = f"1/(1+EXP(-{k_}*({t}-{m_})))"
        elif rp["type"] == "linear":
            n_ = self.live[rp["years_to_peak"]]
            expr = f"MIN(1,{t}/{n_})"
        else:
            expr = "1"
        self.adrow[it["id"]] = self.crow
        ad = self.put_row(f"Adoption ({rp['type']} curve): {it['name']}", "%",
                          lambda c, p, y, expr=expr, ly=ly: f"=IF({c}$4>={ly},{expr.format(c=c)},0)",
                          ident=it["id"])
        ret = None
        if it.get("decay"):
            d = self.live[it["decay"]]
            ret = self.put_row(f"Effect retention after decay: {it['name']}", "%",
                               lambda c, p, y, d=d, ly=ly: f"=IF({c}$4>={ly},(1-{d})^({c}$4-{ly}),0)",
                               ident=it["id"])
        return ad, ret

    def prod(self, ids):
        return "*".join(self.live[x] for x in ids_of(ids))

    def build_builds(self):
        """Initiatives that build their own volume x price line: a new revenue stream (e.g. entering a market
        with no baseline) or a cost avoided (e.g. solar kWh not bought from the grid)."""
        self.buildrows, self.shared, self.sharedpen = {}, {}, []
        builds = [it for it in self.s["_initiatives"] if it.get("model") == "build"]
        if not builds:
            return
        self.csec("4c. Build initiatives: volume (drivers x S-curve x decay, capped) x price")
        for it in builds:
            vol, pr = it["volume"], it["price"]
            info = {"kind": it["kind"]}
            ad, ret = self.ramp_rows(it)
            dr = self.prod(vol["drivers"])
            vg = self.live.get(vol.get("growth"))
            gm = f"*(1+{vg})^{{c}}$5" if vg else ""
            vu = vol.get("unit", "")
            info["peak"] = self.put_row(f"Volume at full adoption (drivers multiplied" + (", growing" if vg else "") +
                                        f"): {it['name']}", vu,
                                        lambda c, p, y, dr=dr, gm=gm: f"={dr}{gm.format(c=c)}", ident=it["id"])
            capof = vol.get("capacity_of")
            if capof:
                lr = self.leafrows[capof]
                prior = [self.buildrows[k]["vol"] for k in self.shared.get(capof, [])]
                info["spare"] = self.put_row(
                    f"Spare capacity left for {it['name']} (ceiling of {self.s['_leaves'][capof]['name']} - its "
                    f"volume with strategy - earlier builds)", vu,
                    lambda c, p, y, lr=lr, prior=prior: f"={c}{lr['ceiling']}-{c}{lr['with']}" +
                    "".join(f"-{c}{r}" for r in prior), ident=it["id"])
            if vol.get("ceiling"):
                cb = self.live[vol["ceiling"]]
                cg = self.live.get(vol.get("ceiling_growth"))
                cm = f"*(1+{cg})" if cg else ""
                info["ceiling"] = self.put_row(f"Ceiling: {vol.get('ceiling_name', 'addressable volume')} for {it['name']}",
                                               vu, lambda c, p, y, cb=cb, cm=cm: f"={cb}" if y == self.base_year
                                               else f"={p}{self.crow}{cm}", cols="base+", ident=it["id"])

            def vfn(c, p, y, info=info, ad=ad, ret=ret):
                f = f"{c}{info['peak']}*{c}{ad}" + (f"*{c}{ret}" if ret else "")
                caps = [f"{c}{info[k]}" for k in ("ceiling", "spare") if k in info]
                if caps:
                    f = f"MIN({f},{','.join(caps)})"
                return "=" + f
            info["vol"] = self.put_row(f"{vol.get('name', 'Volume')}: {it['name']}", vu, vfn, ident=it["id"])
            if "ceiling" in info:
                info["pen"] = self.put_row(f"Penetration of ceiling (% of ceiling): {it['name']}", "%",
                                           lambda c, p, y, info=info: f"={c}{info['vol']}/{c}{info['ceiling']}",
                                           ident=it["id"])
            pb = self.live[pr["base"]]
            pg = self.live.get(pr.get("growth"))
            pm = f"*(1+{pg})" if pg else ""
            info["price"] = self.put_row(f"{pr.get('name', 'Price per unit')}: {it['name']}", pr.get("unit", ""),
                                         lambda c, p, y, pb=pb, pm=pm: f"={pb}" if y == self.base_year
                                         else f"={p}{self.crow}{pm}", cols="base+", ident=it["id"])
            sc = self.prod(it.get("scale"))
            sm = f"*{sc}" if sc else ""
            what = "New revenue" if it["kind"] == "revenue" else "Cost saving"
            info["value"] = self.put_row(f"{what}: {it['name']}", self.cur,
                                         lambda c, p, y, info=info, sm=sm: f"={c}{info['vol']}*{c}{info['price']}{sm}",
                                         ident=it["id"])
            self.buildrows[it["id"]] = info
            if capof:
                self.shared.setdefault(capof, []).append(it["id"])
            self.blank()
        for lid, ks in self.shared.items():
            lr = self.leafrows[lid]
            vols = [self.buildrows[k]["vol"] for k in ks]
            self.sharedpen.append(self.put_row(
                f"Penetration of shared ceiling (% of ceiling): {self.s['_leaves'][lid]['name']} plus "
                + ", ".join(next(i["name"] for i in self.s["_initiatives"] if i["id"] == k) for k in ks), "%",
                lambda c, p, y, lr=lr, vols=vols: f"=({c}{lr['with']}" + "".join(f"+{c}{r}" for r in vols) +
                f")/{c}{lr['ceiling']}"))
            self.blank()

    def build_stream(self, st):
        info = {"via": {}}
        sc = self.live.get(st.get("scale"))
        sm = f"*{sc}" if sc else ""
        leaves = st["leaves"]
        base_prod = lambda c: "*".join(f"{c}{self.leafrows[x]['base']}" for x in leaves)
        with_prod = lambda c: "*".join(f"{c}{self.leafrows[x]['with']}" for x in leaves)
        info["base"] = self.put_row(f"Baseline revenue: {st['name']}", self.cur,
                                    lambda c, p, y: f"={base_prod(c)}{sm}", cols="base+", ident=st["id"])
        info["with"] = self.put_row(f"Revenue with strategy: {st['name']}", self.cur,
                                    lambda c, p, y: f"={with_prod(c)}{sm}", cols="base+", ident=st["id"])
        info["inc"] = self.put_row(f"Incremental revenue: {st['name']}", self.cur,
                                   lambda c, p, y: f"={c}{info['with']}-{c}{info['base']}", ident=st["id"])
        moved = [x for x in leaves if "factor" in self.leafrows[x]]
        if moved:
            info["den"] = self.put_row(f"Sum of log uplift factors: {st['name']}", "x",
                                       lambda c, p, y: "=" + "+".join(f"LN({c}{self.leafrows[x]['factor']})"
                                                                      for x in moved), fmt="0.0000")
            for x in moved:
                fr = self.leafrows[x]["factor"]
                sh = self.put_row(f"Log share of uplift: {self.s['_leaves'][x]['name']} in {st['name']}", "%",
                                  lambda c, p, y, fr=fr: f"=IF({c}{info['den']}=0,0,LN({c}{fr})/{c}{info['den']})")
                info["via"][x] = self.put_row(
                    f"Incremental revenue: {st['name']} via {self.s['_leaves'][x]['name']}", self.cur,
                    lambda c, p, y, sh=sh: f"={c}{info['inc']}*{c}{sh}")
        self.streamrows[st["id"]] = info
        self.blank()

    def build_totals_rows(self):
        """Baseline and target revenue totals (needed before outputs)."""
        sts = self.s["streams"]
        self.rows["base_total"] = self.put_row(
            "Baseline revenue (do nothing)", self.cur,
            lambda c, p, y: "=" + "+".join(f"{c}{self.streamrows[s['id']]['base']}" for s in sts), cols="base+")
        if self.meta.get("baseline_bridge"):
            self.sh["Calculations"].cell(self.rows["base_total"], 4, "FLAG: bridged to history (section 9b): " +
                                         self.meta.get("baseline_justification", ""))
        elif self.meta.get("baseline_justification"):
            self.sh["Calculations"].cell(self.rows["base_total"], 4, "Justification, not yet bridged to history: " +
                                         self.meta["baseline_justification"])
        newrev = [v["value"] for v in self.buildrows.values() if v["kind"] == "revenue"]
        nr = lambda c: "".join(f"+{c}{r}" for r in newrev)
        self.rows["target_total"] = self.put_row(
            "Target revenue (with strategy)", self.cur,
            lambda c, p, y: "=" + "+".join(f"{c}{self.streamrows[s['id']]['with']}" for s in sts)
            + ("" if y == self.base_year else nr(c)), cols="base+")
        self.rows["stream_inc"] = self.put_row(
            "Incremental revenue across streams (incl. new streams built by initiatives)", self.cur,
            lambda c, p, y: "=" + "+".join(f"{c}{self.streamrows[s['id']]['inc']}" for s in sts) + nr(c))
        rev = self.revenue_actual()
        if rev:
            ar = self.actrow[rev["id"]]
            self.rows["reported"] = self.put_row("Reported revenue (actuals)", self.cur,
                                                 lambda c, p, y: f"=Actuals!{c}{ar}", cols="hist")
        for a in self.s.get("actuals", []):
            if a is rev:
                continue
            ar = self.actrow[a["id"]]
            self.rows[f"act_{a['id']}"] = self.put_row(f"{a['name']} (actuals)", a.get("unit", self.cur),
                                                       lambda c, p, y, ar=ar: f"=Actuals!{c}{ar}", cols="hist")
        self.blank()

    def history_outputs(self):
        rev = self.revenue_actual()
        out = []
        if rev:
            ys = self.hist_cols(rev)
            r = rev["id"]
            out += [(f"Reported revenue CAGR {ys[0]}-{ys[-1]}" if ys else "Reported revenue CAGR", "%", f"h_{r}_cagr"),
                    ("Reported revenue: median YoY growth", "%", f"h_{r}_median_growth"),
                    ("Reported revenue: YoY growth, lower percentile", "%", f"h_{r}_g_p25"),
                    ("Reported revenue: YoY growth, upper percentile", "%", f"h_{r}_g_p75"),
                    (f"Baseline revenue CAGR {self.base_year}-{self.years[-1]} (do nothing)", "%", "base_cagr")]
        for a in self.s.get("actuals", []):
            if a.get("role") in MARGIN_ROLES:
                out.append((f"{a['name']}: median % of revenue", "%", f"h_{a['id']}_median_margin"))
        out.append(("Back-test: largest yearly gap vs reported", "%", "bt_max"))
        return out

    # ----------------------------------------------------------------------- Outputs
    def build_outputs(self):
        ws = self.sh["Outputs"]
        self.title(ws, f"Outputs - {self.meta['title']}", f"Currency: {self.cur}. All cells link to Calculations.")
        self.timeline_header(ws)
        ws.freeze_panes = ws.cell(5, FIRST_TL_COL)
        self.orow = 6
        S = "Outputs"
        self.csec("Headline results (base case unless Sensitivity switches are set)", S)
        for label, unit, key in (("NPV of strategy at WACC", self.cur, "npv"),
                                 ("  of which PV of terminal value", self.cur, "tv"),
                                 ("Horizon-only NPV (no terminal value)", self.cur, "npv_h"), ("IRR", "%", "irr"),
                                 ("Payback (years after base year)", "years", "payback"),
                                 ("WACC", "%", "wacc"),
                                 (f"Incremental revenue in {self.years[-1]}", self.cur, "rev_final"),
                                 (f"Incremental EBIT in {self.years[-1]}", self.cur, "ebit_final"),
                                 ("Cumulative incremental revenue", self.cur, "cum_rev"),
                                 (f"Baseline revenue in {self.years[-1]}", self.cur, "base_final"),
                                 (f"Target revenue in {self.years[-1]}", self.cur, "target_final")):
            r = self.put_scalar(label, unit, f"={self.ref[key]}", f"out_{key}", S)
            ws.cell(r, 5).fill = OUT_FILL
        self.blank(1, S)
        hist = [(lab, unit, k) for lab, unit, k in self.history_outputs() if k in self.ref]
        if hist:
            self.csec("Base case vs history", S)
            for label, unit, key in hist:
                self.put_scalar(label, unit, f"={self.ref[key]}", f"out_{key}", S)
            self.blank(1, S)
        self.csec("Revenue path", S)
        C = "Calculations!"
        o = {}
        if "reported" in self.rows:
            o["rep"] = self.put_row("Reported revenue (actuals)", self.cur,
                                    lambda c, p, y: f"={C}{c}{self.rows['reported']}", cols="hist", sheet=S)
        o["base"] = self.put_row("Baseline revenue (do nothing)", self.cur,
                                 lambda c, p, y: f"={C}{c}{self.rows['base_total']}", cols="base+", sheet=S)
        if self.meta.get("baseline_bridge"):
            ws.cell(o["base"], 4, "FLAG: bridged to history (Calculations 9b): " +
                    self.meta.get("baseline_justification", ""))
        elif self.meta.get("baseline_justification"):
            ws.cell(o["base"], 4, "Justification, not yet bridged to history: " + self.meta["baseline_justification"])
        o["target"] = self.put_row("Target revenue (with strategy)", self.cur,
                                   lambda c, p, y: f"={C}{c}{self.rows['target_total']}", cols="base+", sheet=S)
        self.blank(1, S)
        self.csec("Incremental revenue by pillar", S)
        prow = []
        for pl in self.s["pillars"]:
            prow.append(self.put_row(f"Incremental revenue: {pl['name']}", self.cur,
                                     lambda c, p, y, pl=pl: f"={C}{c}{self.pillrev[pl['id']]}", sheet=S))
        o["pill_total"] = self.put_row("Total incremental revenue", self.cur,
                                       lambda c, p, y: f"=SUM({c}{prow[0]}:{c}{prow[-1]})", sheet=S)
        o["prow"] = prow
        self.blank(1, S)
        self.csec("Incremental revenue by initiative", S)
        irow = []
        for it in self.s["_initiatives"]:
            if it["id"] in self.initrev:
                irow.append(self.put_row(f"Incremental revenue: {it['name']}", self.cur,
                                         lambda c, p, y, it=it: f"={C}{c}{self.initrev[it['id']]}", sheet=S))
        o["init_total"] = self.put_row("Total incremental revenue", self.cur,
                                       lambda c, p, y: f"=SUM({c}{irow[0]}:{c}{irow[-1]})" if irow else "=0", sheet=S)
        self.blank(1, S)
        self.csec("Incremental EBIT by pillar (revenue x margin + cost savings - opex)", S)
        perow = []
        for pl in self.s["pillars"]:
            perow.append(self.put_row(f"Incremental EBIT: {pl['name']}", self.cur,
                                      lambda c, p, y, pl=pl: f"={C}{c}{self.pillebit[pl['id']]}", sheet=S))
        self.put_row("Total incremental EBIT", self.cur, lambda c, p, y: f"=SUM({c}{perow[0]}:{c}{perow[-1]})", sheet=S)
        self.blank(1, S)
        self.csec("NPV by initiative (standalone; adds up to the strategy NPV)", S)
        for it in self.s["_initiatives"]:
            r_ = self.put_scalar(f"NPV: {it['name']}", self.cur, f"={self.ref['npv_' + it['id']]}",
                                 f"out_npv_{it['id']}", S)
        self.blank(1, S)
        self.csec("Incremental EBIT by initiative", S)
        ierow = []
        for it in self.s["_initiatives"]:
            ierow.append(self.put_row(f"Incremental EBIT: {it['name']}", self.cur,
                                      lambda c, p, y, it=it: f"={C}{c}{self.initebit[it['id']]}", sheet=S))
        self.put_row("Total incremental EBIT", self.cur, lambda c, p, y: f"=SUM({c}{ierow[0]}:{c}{ierow[-1]})", sheet=S)
        self.blank(1, S)
        self.csec("Incremental profit and cash", S)
        for key, label in (("ebit", "Incremental EBIT"), ("fcf", "Incremental free cash flow"),
                           ("cum", "Cumulative incremental free cash flow")):
            o[key] = self.put_row(label, self.cur, lambda c, p, y, key=key: f"={C}{c}{self.rows[key]}", sheet=S)
        self.out = o
        # charts
        fc, lc = self.col(self.years[0]), self.last_col
        from openpyxl.chart.series import SeriesLabel
        ch = LineChart()
        ch.title = f"Revenue: baseline vs with strategy ({self.cur})"
        ch.height, ch.width = 7.5, 16
        for key, nm in (("rep", "Reported"), ("base", "Baseline"), ("target", "With strategy")):
            if key in o:
                ch.add_data(Reference(ws, min_col=fc, max_col=lc, min_row=o[key]), from_rows=True,
                            titles_from_data=False)
                ch.series[-1].tx = SeriesLabel(v=nm)
        ch.set_categories(Reference(ws, min_col=fc, max_col=lc, min_row=4))
        ch.x_axis.delete = False
        ch.y_axis.delete = False
        ws.add_chart(ch, f"B{self.orow + 2}")
        bar = BarChart()
        bar.type, bar.grouping, bar.overlap = "col", "stacked", 100
        bar.title = f"Incremental revenue by pillar ({self.cur})"
        bar.height, bar.width = 7.5, 16
        for r in prow:
            bar.add_data(Reference(ws, min_col=fc, max_col=lc, min_row=r), from_rows=True, titles_from_data=False)
        for s_, pl in zip(bar.series, self.s["pillars"]):
            s_.tx = SeriesLabel(v=pl["name"])
        bar.set_categories(Reference(ws, min_col=fc, max_col=lc, min_row=4))
        bar.x_axis.delete = False
        bar.y_axis.delete = False
        ws.add_chart(bar, f"H{self.orow + 2}")

    # ----------------------------------------------------------------------- Bridge
    def build_bridge(self):
        ws = self.sh["Bridge"]
        fy = self.years[-1]
        self.title(ws, f"Revenue bridge {fy}: baseline to target ({self.cur})",
                   "One-slide impact bridge. Columns D-G drive the waterfall chart.")
        for j, h in enumerate(["Step", "Value", "Invisible base", "Increase", "Decrease", "Level", "Running total"], 2):
            c = ws.cell(4, j, h)
            c.font, c.fill = HEAD_FONT, HEAD_FILL
        for c in "CDEFGH":
            ws.column_dimensions[c].width = 14
        lc = L(self.last_col)
        r = 5
        ws.cell(r, 2, f"Baseline revenue {fy} (do nothing)")
        ws.cell(r, 3, f"=Calculations!{lc}{self.rows['base_total']}")
        ws.cell(r, 7, f"=C{r}")
        ws.cell(r, 8, f"=C{r}")
        first = r + 1
        r += 1
        for it in self.s["_initiatives"]:
            if it["id"] not in self.initrev:
                continue
            pl = next(p["name"] for p in self.s["pillars"] if p["id"] == it["pillar"])
            ws.cell(r, 2, f"{pl}: {it['name']}")
            ws.cell(r, 3, f"=Calculations!{lc}{self.initrev[it['id']]}")
            ws.cell(r, 4, f"=IF(C{r}>=0,H{r - 1},H{r - 1}+C{r})")
            ws.cell(r, 5, f"=MAX(C{r},0)")
            ws.cell(r, 6, f"=MAX(-C{r},0)")
            ws.cell(r, 8, f"=H{r - 1}+C{r}")
            r += 1
        last = r - 1
        ws.cell(r, 2, f"Target revenue {fy} (with strategy)")
        ws.cell(r, 3, f"=Calculations!{lc}{self.rows['target_total']}")
        ws.cell(r, 7, f"=C{r}")
        ws.cell(r, 8, f"=C{r}")
        self.bridge = {"base": r - (last - first + 2), "first": first, "last": last, "target": r}
        for rr in range(5, r + 1):
            for cc in range(3, 9):
                ws.cell(rr, cc).number_format = self.fmt_for(self.cur)
        bar = BarChart()
        bar.type, bar.grouping, bar.overlap = "col", "stacked", 100
        bar.title = f"Revenue bridge {fy} ({self.cur})"
        bar.height, bar.width = 9, 20
        data = Reference(ws, min_col=4, max_col=7, min_row=4, max_row=r)
        bar.add_data(data, titles_from_data=True)
        bar.set_categories(Reference(ws, min_col=2, min_row=5, max_row=r))
        bar.series[0].graphicalProperties.noFill = True
        bar.series[0].graphicalProperties.line.noFill = True
        for s_, colr in zip(bar.series[1:], ("70AD47", "C00000", "7F7F7F")):
            s_.graphicalProperties.solidFill = colr
        bar.legend = None
        bar.x_axis.delete = False
        bar.y_axis.delete = False
        ws.add_chart(bar, f"J4")
        # value bridge: incremental EBIT by initiative (shows cost savings and opex, which revenue cannot)
        r += 3
        ws.cell(r - 1, 2, f"Incremental EBIT bridge {fy} by initiative ({self.cur})").font = BOLD
        top = r
        ws.cell(r, 2, "Do nothing (no incremental EBIT)")
        ws.cell(r, 3, "=0")
        ws.cell(r, 7, f"=C{r}")
        ws.cell(r, 8, f"=C{r}")
        r += 1
        f2 = r
        for it in self.s["_initiatives"]:
            pl = next(p["name"] for p in self.s["pillars"] if p["id"] == it["pillar"])
            ws.cell(r, 2, f"{pl}: {it['name']}")
            ws.cell(r, 3, f"=Calculations!{lc}{self.initebit[it['id']]}")
            ws.cell(r, 4, f"=IF(C{r}>=0,H{r - 1},H{r - 1}+C{r})")
            ws.cell(r, 5, f"=MAX(C{r},0)")
            ws.cell(r, 6, f"=MAX(-C{r},0)")
            ws.cell(r, 8, f"=H{r - 1}+C{r}")
            r += 1
        ws.cell(r, 2, f"Incremental EBIT {fy} (with strategy)")
        ws.cell(r, 3, f"=Calculations!{lc}{self.rows['ebit']}")
        ws.cell(r, 7, f"=C{r}")
        ws.cell(r, 8, f"=C{r}")
        for rr in range(top, r + 1):
            for cc in range(3, 9):
                ws.cell(rr, cc).number_format = self.fmt_for(self.cur)
        self.bridge2 = {"base": top, "first": f2, "last": r - 1, "target": r}

    # ----------------------------------------------------------------------- Sensitivity
    def build_sensitivity(self):
        ws = self.sh["Sensitivity"]
        self.title(ws, "Sensitivity - scenarios, tornado, two-way tables",
                   "Live Excel Data Tables (Data > What-If). Leave the orange controls at Base / 0 / blank.")
        ws.column_dimensions["C"].width = 12
        for c in "DEFGHIJK":
            ws.column_dimensions[c].width = 14
        self.section(ws, 4, "Controls")
        ws["B5"], ws["E5"] = "Scenario switch (1 = Low, 2 = Base, 3 = High)", 2
        ws["B6"], ws["E6"] = "Tornado selector (0 = off; set by the tornado Data Tables)", 0
        names = [("ctl_scenario", "$E$5"), ("ctl_tornado", "$E$6")]
        for k, iid in enumerate(self.ovr):
            r = 7 + k
            ws.cell(r, 2, f"Override for two-way table: {self.by[iid]['name']} (leave blank)")
            names.append((f"ctl_ovr_{k + 1}", f"$E${r}"))
        for nm, a in names:
            ws[a.replace("$", "")].fill = CTL_FILL
            self.wb.defined_names[nm] = DefinedName(nm, attr_text=f"Sensitivity!{a}")
        npv = self.ref["npv"]
        r0 = 9 + len(self.ovr)
        # scenarios
        self.section(ws, r0, "Scenarios: Low / Base / High (every ranged input moves to its pessimistic / optimistic end)")
        heads = ["Scenario", "Switch", "NPV", f"Incr. revenue {self.years[-1]}", f"Incr. EBIT {self.years[-1]}",
                 "Cumulative incr. revenue"]
        for j, h in enumerate(heads, 2):
            c = ws.cell(r0 + 1, j, h)
            c.font, c.fill = HEAD_FONT, HEAD_FILL
        fr = r0 + 2
        ws.cell(fr, 2, "Live (current switches)")
        for j, key in enumerate(("npv", "rev_final", "ebit_final", "cum_rev"), 4):
            ws.cell(fr, j, f"={self.ref[key]}").number_format = self.fmt_for(self.cur)
        for k, nm in enumerate(("Low", "Base", "High"), 1):
            ws.cell(fr + k, 2, nm)
            ws.cell(fr + k, 3, k)
            for j in range(4, 8):
                ws.cell(fr + k, j).number_format = self.fmt_for(self.cur)
        ws.cell(fr + 1, 4).value = DataTableFormula(ref=f"D{fr + 1}:G{fr + 3}", dtr="0", r1="E5")
        self.scen = {"low": f"Sensitivity!$D${fr + 1}", "base": f"Sensitivity!$D${fr + 2}",
                     "high": f"Sensitivity!$D${fr + 3}", "base_code": f"Sensitivity!$C${fr + 2}"}
        # tornado
        tids = self.tornado_ids if self.tornado_ids is not None else \
            [i["id"] for i in self.s["inputs"] if i["type"] not in EXEMPT and i["low"] != i["high"]]
        t0 = fr + 6
        self.section(ws, t0, "Tornado: NPV with each input at its low and high value, others at base")
        th = ["Input", "Selector (low)", "NPV at low", "Selector (high)", "NPV at high", "Base NPV",
              "Delta at low", "Delta at high", "Swing", "Rank", "Sort key"]
        for j, h in enumerate(th, 2):
            c = ws.cell(t0 + 1, j, h)
            c.font, c.fill = HEAD_FONT, HEAD_FILL
        f_r = t0 + 2
        ws.cell(f_r, 2, "Live NPV")
        ws.cell(f_r, 4, f"={npv}")
        ws.cell(f_r, 6, f"={npv}")
        first = f_r + 1
        rows = []
        for k, iid in enumerate(tids):
            r = first + k
            ir = self.inrow[iid]
            ws.cell(r, 2, f"=Inputs!$B${ir}")
            ws.cell(r, 3, f"=-Inputs!$A${ir}")
            ws.cell(r, 5, f"=Inputs!$A${ir}")
            rows.append(r)
        last = first + len(tids) - 1
        if tids:
            ws.cell(first, 4).value = DataTableFormula(ref=f"D{first}:D{last}", dtr="0", r1="E6")
            ws.cell(first, 6).value = DataTableFormula(ref=f"F{first}:F{last}", dtr="0", r1="E6")
            for r in rows:
                ws.cell(r, 7, f"={npv}")
                ws.cell(r, 8, f"=D{r}-G{r}")
                ws.cell(r, 9, f"=F{r}-G{r}")
                ws.cell(r, 10, f"=ABS(F{r}-D{r})")
                ws.cell(r, 12, f"=J{r}+ROWS($J${first}:J{r})*Inputs!$F${self.inrow['tie_step']}")
                ws.cell(r, 11, f"=RANK(L{r},$L${first}:$L${last})")
                for cc in range(4, 11):
                    ws.cell(r, cc).number_format = self.fmt_for(self.cur)
        self.record["tornado"] = {"ids": tids, "first": first, "low_col": 4, "high_col": 6}
        # sorted top N for the chart
        topn = min(int(self.s.get("sensitivity", {}).get("tornado_top", 10)), len(tids))
        s0 = last + 3
        self.section(ws, s0, f"Tornado chart data: top {topn} inputs by swing")
        for j, h in enumerate(["Rank", "Input", "Delta at low", "Delta at high"], 2):
            c = ws.cell(s0 + 1, j, h)
            c.font, c.fill = HEAD_FONT, HEAD_FILL
        for k in range(topn):
            r = s0 + 2 + k
            m = f"MATCH(B{r},$K${first}:$K${last},0)"
            ws.cell(r, 2, f"=ROWS($B${s0 + 2}:B{r})")
            ws.cell(r, 3, f"=INDEX($B${first}:$B${last},{m})")
            ws.cell(r, 4, f"=INDEX($H${first}:$H${last},{m})").number_format = self.fmt_for(self.cur)
            ws.cell(r, 5, f"=INDEX($I${first}:$I${last},{m})").number_format = self.fmt_for(self.cur)
        if topn:
            ch = BarChart()
            ch.type, ch.grouping, ch.overlap = "bar", "clustered", 100
            ch.title = f"Tornado: change in NPV ({self.cur})"
            ch.height, ch.width = 9, 18
            ch.add_data(Reference(ws, min_col=4, max_col=5, min_row=s0 + 1, max_row=s0 + 1 + topn), titles_from_data=True)
            ch.set_categories(Reference(ws, min_col=3, min_row=s0 + 2, max_row=s0 + 1 + topn))
            ch.y_axis.delete = False
            ch.x_axis.delete = False
            ch.x_axis.scaling.orientation = "maxMin"
            ch.series[0].graphicalProperties.solidFill = "C00000"
            ch.series[1].graphicalProperties.solidFill = "70AD47"
            ws.add_chart(ch, f"M{t0}")
        # two-way tables
        w0 = s0 + 3 + topn
        steps = int(self.s.get("sensitivity", {}).get("steps", 5))
        self.record["two_way"] = []
        for n, (a, b) in enumerate(self.two_way or []):
            ra, rb = self.inrow[a], self.inrow[b]
            top = w0 + n * (steps + 5)
            self.section(ws, top, f"Two-way: NPV ({self.cur}) - rows: {self.by[a]['name']}; columns: {self.by[b]['name']}")
            hr = top + 2
            c1, c2 = 4, 4 + steps - 1
            ws.cell(hr, 3, f"={npv}").number_format = self.fmt_for(self.cur)
            ws.cell(hr - 1, 4, f"Columns: {self.by[b]['name']}")
            ws.cell(hr + 1, 2, f"Rows: {self.by[a]['name']}")
            for j in range(steps):
                cc = L(c1 + j)
                ws.cell(hr, c1 + j, f"=Inputs!$E${rb}+(Inputs!$G${rb}-Inputs!$E${rb})*"
                                    f"(COLUMNS($D${hr}:{cc}{hr})-1)/(COLUMNS($D${hr}:${L(c2)}${hr})-1)")
                ws.cell(hr, c1 + j).number_format = self.fmt_for(self.by[b]["unit"])
                ws.cell(hr, c1 + j).font = BOLD
            for i in range(steps):
                rr = hr + 1 + i
                ws.cell(rr, 3, f"=Inputs!$E${ra}+(Inputs!$G${ra}-Inputs!$E${ra})*"
                               f"(ROWS($C${hr + 1}:C{rr})-1)/(ROWS($C${hr + 1}:$C${hr + steps})-1)")
                ws.cell(rr, 3).number_format = self.fmt_for(self.by[a]["unit"])
                ws.cell(rr, 3).font = BOLD
                for j in range(steps):
                    ws.cell(rr, c1 + j).number_format = self.fmt_for(self.cur)
            oa, ob = self.ovr[a].split("!")[1].replace("$", ""), self.ovr[b].split("!")[1].replace("$", "")
            ws.cell(hr + 1, c1).value = DataTableFormula(ref=f"D{hr + 1}:{L(c2)}{hr + steps}", dt2D="1", dtr="1",
                                                         r1=ob, r2=oa)
            self.record["two_way"].append((a, b, f"Sensitivity!D{hr + 1}:{L(c2)}{hr + steps}"))

    # ----------------------------------------------------------------------- Checks
    def build_checks(self):
        ws = self.sh["Checks"]
        self.title(ws, "Checks", "Each check reads OK or ERROR. Total errors must be 0 before handover. "
                                 "Then run scripts/check_model.py (the independent judge).")
        self.timeline_header(ws)
        ws.column_dimensions["C"].width = 14
        ws.column_dimensions["D"].width = 14
        ws["B6"] = "Total errors (must be 0)"
        ws["B6"].font = BOLD
        ws["B8"], ws["C8"], ws["D8"], ws["E8"] = "Check", "Value", "Limit", "Result"
        for c in "BCDE":
            ws[f"{c}8"].font, ws[f"{c}8"].fill = HEAD_FONT, HEAD_FILL
        r = 9
        C = "Calculations!"
        I = "Inputs!"
        tol = f"{I}$F${self.inrow['sum_tol']}"
        fc, lc = L(self.col(self.proj_years[0])), L(self.last_col)
        results = []

        def row(label, value, limit, test):
            nonlocal r
            ws.cell(r, 2, label)
            ws.cell(r, 3, value)
            ws.cell(r, 4, limit)
            ws.cell(r, 5, f'=IF({test},"OK","ERROR")')
            results.append(r)
            r += 1

        def per_year(label, fn):
            nonlocal r
            ws.cell(r, 2, label)
            for y in self.proj_years:
                c = L(self.col(y))
                ws[f"{c}{r}"] = fn(c)
                ws[f"{c}{r}"].number_format = "0.0E+00"
            rr = r
            ws.cell(r, 3, f"=MAX({fc}{rr}:{lc}{rr})")
            ws.cell(r, 4, f"={tol}")
            ws.cell(r, 5, f'=IF(C{rr}<=D{rr},"OK","ERROR")')
            results.append(r)
            r += 1

        if "reported_base" in self.ref:
            row(f"Baseline calibration: model base-year revenue vs reported ({self.base_year})",
                f"=ABS({self.ref['model_base']}/{self.ref['reported_base']}-1)", f"={I}$F${self.inrow['calib_tol']}",
                f"C{r}<=D{r}")
        per_year("Initiatives add up to stream incremental revenue (abs. difference)",
                 lambda c: f"=ABS({C}{c}{self.rows['inc_total']}-{C}{c}{self.rows['stream_inc']})")
        per_year("Initiatives' EBIT adds up to incremental EBIT (abs. difference)",
                 lambda c: f"=ABS({C}{c}{self.rows['ebit_pillars']}-{C}{c}{self.rows['ebit']})")
        per_year("Pillars add up to initiatives (abs. difference)",
                 lambda c: f"=ABS({C}{c}{self.rows['pill_total']}-{C}{c}{self.rows['inc_total']})")
        per_year("Target = baseline + incremental (abs. difference)",
                 lambda c: f"=ABS({C}{c}{self.rows['target_total']}-{C}{c}{self.rows['base_total']}-"
                           f"{C}{c}{self.rows['stream_inc']})")
        if "base_cagr" in self.ref:
            just = self.meta.get("baseline_justification")
            inside = f"AND(C{r}>={self.ref['hist_lo']},C{r}<=D{r})"
            bridged = "br_gap" in self.ref
            row("Base case vs history: baseline revenue CAGR inside the historical band (Calculations 9b)" +
                (", or bridged to it within tolerance" if bridged else ""),
                f"={self.ref['base_cagr']}", f"={self.ref['hist_hi']}",
                f"OR({inside},{self.ref['br_gap']}<={I}$F${self.inrow['hist_tol']})" if bridged else inside)
            ws.cell(r - 1, 3).number_format = ws.cell(r - 1, 4).number_format = "0.0%"
            if just:
                ws.cell(r - 1, 6, f"Justification: {just}")
        if "bt_max" in self.ref:
            row("Back-test: driver tree on reported driver history reproduces reported revenue (largest yearly gap)",
                f"={self.ref['bt_max']}", f"={I}$F${self.inrow['backtest_tol']}", f"C{r}<=D{r}")
            ws.cell(r - 1, 3).number_format = ws.cell(r - 1, 4).number_format = "0.0%"
        b = self.bridge
        row("Bridge closes: baseline + steps = target (abs. difference)",
            f"=ABS(Bridge!C{b['base']}+SUM(Bridge!C{b['first']}:C{b['last']})-Bridge!C{b['target']})",
            f"={tol}", f"C{r}<=D{r}")
        b2 = self.bridge2
        row("EBIT bridge closes: steps = incremental EBIT (abs. difference)",
            f"=ABS(Bridge!C{b2['base']}+SUM(Bridge!C{b2['first']}:C{b2['last']})-Bridge!C{b2['target']})",
            f"={tol}", f"C{r}<=D{r}")
        self.objectives_checks(lambda a, b, c, t: row(a, b, c, t.replace("{r}", str(r))))
        pens = [self.leafrows[x]["pen"] for x in self.leafrows if "pen" in self.leafrows[x]]
        pens += [v["pen"] for v in self.buildrows.values() if "pen" in v] + self.sharedpen
        if pens:
            bc = L(self.base_col)
            mx = ",".join(f"{C}{bc}{p}:{lc}{p}" for p in pens)
            row("Penetration never exceeds 100% of the addressable market", f"=MAX({mx})", "<= 100%", f"C{r}<=1")
        withs = ",".join([f"{C}{L(self.base_col)}{v['with']}:{lc}{v['with']}" for v in self.leafrows.values()] +
                         [f"{C}{fc}{v['vol']}:{lc}{v['vol']}" for v in self.buildrows.values()])
        row("No driver goes negative", f"=MIN({withs})", ">= 0", f"C{r}>=0")
        lo, ba, hi = (f"{I}{c}6:{c}{self.inputs_last}" for c in "EFG")
        src, grd = (f"{I}{c}6:{c}{self.inputs_last}" for c in "IL")
        typ, bh, jus = (f"{I}{c}6:{c}{self.inputs_last}" for c in "DOQ")
        row("Register: low <= base <= high on every row (count of breaks)",
            f"=SUMPRODUCT(({lo}>{ba})*1)+SUMPRODUCT(({ba}>{hi})*1)", "0", f"C{r}=0")
        row("Register: every input has a source and a grade (count missing)",
            f'=SUMPRODUCT(({ba}<>"")*({src}=""))+SUMPRODUCT(({ba}<>"")*({grd}=""))', "0", f"C{r}=0")
        row("Register: uplifts above the reference-class upper quartile have a written justification (count missing)",
            f'=SUMPRODUCT(({typ}="uplift")*({ba}>{bh})*({jus}=""))', "0", f"C{r}=0")
        row("Initiative NPVs add up to the strategy NPV (abs. difference)",
            f"=ABS({self.ref['npv_sum']}-{self.ref['npv']})", f"={tol}", f"C{r}<=D{r}")
        row("WACC above terminal growth (Gordon model valid)", f"={self.ref['wacc']}-{C}{self.tv_g}", "> 0",
            f"C{r}>0")
        row("Scenarios ordered: NPV low <= base <= high", f"=IF(AND({self.scen['low']}<={self.scen['base']},"
            f"{self.scen['base']}<={self.scen['high']}),0,1)", "0", f"C{r}=0")
        ovr = "+".join(f"COUNT({v})" for v in self.ovr.values()) or "0"
        row("Controls neutral: scenario = Base, tornado off, no overrides (count set)",
            f"=(Sensitivity!$E$5<>{self.scen['base_code']})*1+(Sensitivity!$E$6<>0)*1+{ovr}", "0", f"C{r}=0")
        # growth of target revenue
        ws.cell(r, 2, "Max YoY growth of target revenue (must stay under the growth cap)")
        for y in self.proj_years:
            c, p = L(self.col(y)), L(self.col(y) - 1)
            ws[f"{c}{r}"] = f"={C}{c}{self.rows['target_total']}/{C}{p}{self.rows['target_total']}-1"
            ws[f"{c}{r}"].number_format = "0.0%"
        ws.cell(r, 3, f"=MAX({fc}{r}:{lc}{r})").number_format = "0.0%"
        ws.cell(r, 4, f"={I}$F${self.inrow['growth_cap']}").number_format = "0.0%"
        ws.cell(r, 5, f'=IF(C{r}<=D{r},"OK","ERROR")')
        results.append(r)
        r += 1
        ws["E6"] = f'=COUNTIF(E{results[0]}:E{results[-1]},"ERROR")'
        ws["E6"].font = Font(bold=True, color="C00000")
        ws["E6"].fill = OUT_FILL

    # ----------------------------------------------------------------------- Driver tree / data requests / cover
    def driver_tree_lines(self):
        s = self.s
        obj = self.meta.get("objective", "revenue")
        lines = [f"Target metric: incremental {obj} ({self.cur}), {self.proj_years[0]}-{self.years[-1]} vs do-nothing"]
        for st in s["streams"]:
            sc = f" x {self.nm(st['scale'])}" if st.get("scale") else ""
            lines.append(f"  Revenue stream: {st['name']} = " + " x ".join(s["_leaves"][x]["name"] for x in st["leaves"]) + sc)
            for x in st["leaves"]:
                lf = s["_leaves"][x]
                own = s["_owners"].get(x, [])
                names = [it["name"] for it in s["_initiatives"] if it["id"] in own]
                rule = s["split_rules"].get(x, "single owner" if own else "no initiative (baseline only)")
                src = (self.by.get(lf["base"]) or {}).get("source", "derived")
                lines.append(f"    Leaf: {lf['name']} [{lf.get('unit', '')}] - owner: "
                             f"{', '.join(names) or 'none'} - split: {rule} - measured by: {src}")
                for it in s["_initiatives"]:
                    if it["leaf"] == x:
                        lines.append(f"      Initiative: {it['name']} ({it['pillar']}) - {it.get('mechanism', '')}")
        for it in s["_initiatives"]:
            if it.get("model") == "build":
                v, pr = it["volume"], it["price"]
                what = "New revenue stream" if it["kind"] == "revenue" else "Cost saving"
                cap = f", capped by {self.nm(v['ceiling'])}" if v.get("ceiling") else ""
                if v.get("capacity_of"):
                    cap += f", within spare capacity of {self.s['_leaves'][v['capacity_of']]['name']}"
                lines.append(f"  {what}: {it['name']} ({it['pillar']}) = {v.get('name', 'volume')} x "
                             f"{pr.get('name', 'price')}")
                lines.append(f"    Volume = " + " x ".join(self.nm(x) for x in ids_of(v["drivers"])) +
                             f" x S-curve" + (" x retention" if it.get("decay") else "") + cap)
                lines.append(f"    Price: {self.nm(pr['base'])} - {(self.by.get(pr['base']) or {}).get('source', 'derived')}")
                lines.append(f"      Mechanism: {it.get('mechanism', '')}")
            elif not it.get("leaf"):
                lines.append(f"  Cost-only initiative: {it['name']} ({it['pillar']}) - {it.get('mechanism', '')}")
        lines.append("  Costs: initiative opex and capex (quantity x unit cost where given); incremental EBIT = "
                     "incremental revenue x margin + cost savings - opex")
        lines.append("  Cash: FCF = EBIT x (1 - tax) - capex, discounted at WACC = E/V x (rf + beta x ERP + CRP) + D/V x kd x (1 - tax)")
        return lines

    def build_tree(self):
        ws = self.sh["Driver tree"]
        self.title(ws, "Driver tree", "Every leaf is measurable and has one owning initiative or an explicit split rule.")
        ws.column_dimensions["B"].width = 140
        for k, line in enumerate(self.driver_tree_lines()):
            c = ws.cell(4 + k, 2, line.strip())
            c.alignment = Alignment(indent=(len(line) - len(line.lstrip())) // 2)
            if not line.startswith("    "):
                c.font = BOLD

    @staticmethod
    def _req_row(item, rq):
        db, f = rq.get("database", ""), rq.get("export_as", f"{item['id']}.xlsx")
        return ["Team pulls (IE library)", f"{item['id']} - {item['name']}", db, rq.get("table", ""),
                rq.get("filters", ""), rq.get("years", ""), f, rq.get("owner", ""), rq.get("status", "open"),
                rq.get("route", "Browser agent (supervised), else team manual export"),
                rq.get("fallback", f"Log in to {db} through the IE library portal, apply the filters, export as {f}, "
                                   f"upload it; until then the provisional value stays (see grade)")]

    def request_rows(self):
        out = []
        for i in self.s["inputs"]:
            if i.get("request"):
                out.append(self._req_row(i, i["request"]))
            fx = i.get("fetch")
            if fx:
                out.append(["Claude fetches (open source)", f"{i['id']} - {i['name']}", fx.get("source", ""),
                            fx.get("url", ""), fx.get("filters", ""), fx.get("years", ""), "-", "Claude",
                            fx.get("status", "open"), "Claude (web fetch)",
                            "Download from the URL by hand if network access is off"])
        for a in self.s.get("actuals", []):
            if a.get("request"):
                out.append(self._req_row(a, a["request"]))
        return out

    REQ_HEAD = ["Who", "Needed for (input)", "Database / source", "Exact table, report or URL", "Filters",
                "Years", "Export as (file name)", "Owner", "Status", "Route", "Manual fallback"]

    def build_requests(self, ws=None):
        ws = ws or self.sh["Data requests"]
        self.title(ws, "Data Request Sheet", "Pull each licensed item from the IE library, export with the exact file "
                                             "name, upload it, and Claude updates the Inputs row.")
        for j, (h, w) in enumerate(zip(self.REQ_HEAD, [24, 40, 26, 50, 40, 12, 30, 12, 9, 30, 60]), 1):
            c = ws.cell(4, j, h)
            c.font, c.fill = HEAD_FONT, HEAD_FILL
            ws.column_dimensions[L(j)].width = w
        for k, row in enumerate(self.request_rows()):
            for j, v in enumerate(row, 1):
                ws.cell(5 + k, j, v)

    def build_cover(self):
        ws = self.sh["Cover"]
        m = self.meta
        ws.column_dimensions["B"].width = 34
        ws.column_dimensions["C"].width = 110
        ws["B2"] = m["title"]
        ws["B2"].font = Font(bold=True, size=16)
        info = [("Company", m["company"]), ("Currency / units", self.cur),
                ("Timeline", f"{self.years[0]}-{self.base_year} actuals, {self.proj_years[0]}-{self.years[-1]} projection"),
                ("Objective", m.get("objective", "revenue")), ("Prepared by", m.get("prepared_by", "")),
                ("As of", m.get("as_of", "")), ("", ""),
                ("How to read", "Inputs -> Calculations -> Outputs / Bridge / Sensitivity; Checks must show 0 errors."),
                ("Colours", "Blue on yellow = typed input (Inputs, Actuals). Orange = Sensitivity control. "
                            "Green = headline output. Everything else is a formula."),
                ("Grades", "Source reliability A (reliable) - F (cannot judge); information credibility 1 (confirmed) - "
                           "6 (cannot judge). Every input carries one."),
                ("Scenarios", "Low/Base/High move every ranged input to its pessimistic/optimistic end (polarity aware)."),
                ("Recalculate", "Excel recalculates on open. Data Tables: Formulas > Calculation Options > Automatic."),
                ("Independent check", "python scripts/check_model.py <this file>  (must pass every check)"),
                ("History", self.history_note())]
        for k, (a, b) in enumerate(info):
            ws.cell(4 + k, 2, a).font = BOLD
            ws.cell(4 + k, 3, b)

    def build_cover_headline(self):
        """The case objectives first: base case vs with strategies, linked to the Objectives sheet."""
        ws = self.sh["Cover"]
        r = ws.max_row + 2
        ws.cell(r, 2, f"Headline: the case objectives in {self.years[-1]} (base case = do nothing)").font = BOLD
        r += 1
        for j, h in enumerate(["KPI", "Base case", "With strategies", "Change", "Better?"], 2):
            c = ws.cell(r, j, h)
            c.font, c.fill = HEAD_FONT, HEAD_FILL
        for k in self.kpis:
            r += 1
            o = self.objrow[k["id"]]
            ws.cell(r, 2, f"{k['name']} ({'pp' if k['ratio'] else k['unit']})")
            for j, col_ in ((3, "E"), (4, "F"), (5, "G"), (6, "J")):
                x = ws.cell(r, j, f"=Objectives!{col_}{o}")
                x.number_format = "0.0%" if k["unit"].startswith("%") else "#,##0.0"
        r += 1
        ws.cell(r, 2, "Valuation (secondary)")
        x = ws.cell(r, 3, f"={self.ref['npv']}")
        x.number_format = "#,##0.0"
        ws.cell(r, 4, f"NPV in {self.cur}; see Outputs. Breakdown through the consulting frameworks: Frameworks sheet.")

    def history_note(self):
        acts = self.s.get("actuals", [])
        linked = [i["id"] for i in self.s["inputs"] if i.get("_hist")]
        txt = (f"{len(acts)} reported series on Actuals; growth, CAGR, quartiles and margins in Calculations section 0. "
               f"{len(linked)} inputs take their value from history" + (f" ({', '.join(linked)})" if linked else "") +
               ". Baseline growth is checked against the historical band (Checks)")
        if self.s.get("_backtest"):
            txt += "; the driver tree is back-tested on reported driver history"
        just = self.meta.get("baseline_justification")
        return txt + "." + (f" Baseline justification: {just}" if just else "")

    # ----------------------------------------------------------------------- run
    def build(self):
        self.build_cover()
        self.build_inputs()
        self.build_actuals()
        self.build_calcs_all()
        self.link_history_inputs()
        self.build_outputs()
        self.build_objectives_sheet()
        self.build_frameworks_sheet()
        self.build_cover_headline()
        self.build_bridge()
        self.build_sensitivity()
        self.build_checks()
        self.build_tree()
        self.build_requests()
        self.wb.calculation.fullCalcOnLoad = True
        return self.wb

    def build_calcs_all(self):
        # Calculations needs the stream totals before the valuation block; build in order.
        orig = self.build_calcs
        self._orig_build_calcs = orig
        ws = self.sh["Calculations"]
        # patch: insert totals after streams by wrapping build_stream loop
        self._calcs_with_totals()

    def _calcs_with_totals(self):
        # same as build_calcs but adds the totals rows right after the streams block
        build_stream = self.build_stream
        streams = self.s["streams"]
        done = {"n": 0}

        def wrapped(st):
            build_stream(st)
            done["n"] += 1
            if done["n"] == len(streams):
                self.build_builds()
                self.csec("4b. Revenue totals and reported actuals")
                self.build_totals_rows()
        self.build_stream = wrapped
        self.build_calcs()
        self.build_stream = build_stream


# ----------------------------------------------------------------------------- two-pass build
def evaluate_tornado(path, rec, npv_ref):
    from xleval import Model, run_big_stack, is_num
    m = Model(path)

    def go():
        ctx = m.context()
        sh, a = npv_ref.split("!")
        base = ctx.value(sh, a.replace("$", ""))
        out = []
        for k, iid in enumerate(rec["ids"]):
            r = rec["first"] + k
            lo = ctx.get(("Sensitivity", r, rec["low_col"]))
            hi = ctx.get(("Sensitivity", r, rec["high_col"]))
            out.append((iid, lo, hi))
        return base, out
    return run_big_stack(go)


def build(spec_path_or_dict, out_path, requests_path=None, quiet=False):
    spec = load_spec(spec_path_or_dict) if isinstance(spec_path_or_dict, str) else spec_path_or_dict
    s, warns = normalise(spec)
    sens = s.get("sensitivity", {})
    # pass 1: every ranged input in the tornado; evaluate to rank them
    b1 = Builder(s, tornado_ids=None, two_way=[])
    b1.build()
    tmp = tempfile.mkdtemp(prefix="imbuild_")
    p1 = os.path.join(tmp, "pass1.xlsx")
    b1.wb.save(p1)
    base, rows = evaluate_tornado(p1, b1.record["tornado"], b1.ref["npv"])
    from xleval import is_num
    swings = [(iid, abs(hi - lo), lo, hi) for iid, lo, hi in rows if is_num(lo) and is_num(hi)]
    moving = [x for x in swings if x[1] > 1e-9 * max(1.0, abs(base))]
    dropped = [iid for iid, sw, _, _ in swings if sw <= 1e-9 * max(1.0, abs(base))]
    moving.sort(key=lambda x: -x[1])
    tornado_ids = [iid for iid, *_ in sorted(moving, key=lambda x: s["_by"][x[0]]["_n"])]
    two = sens.get("two_way")
    if not two:
        top = [iid for iid, *_ in moving]
        two = []
        if len(top) >= 2:
            two.append([top[0], top[1]])
        if len(top) >= 4:
            two.append([top[2], top[3]])
    for a, bb in two:
        if a not in tornado_ids or bb not in tornado_ids:
            raise SpecError(f"two-way pair ({a}, {bb}): both inputs need a range and must move NPV")
    s2, _ = normalise(spec)
    b2 = Builder(s2, tornado_ids=tornado_ids, two_way=two)
    b2.build()
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    b2.wb.save(out_path)
    if requests_path:
        wb = Workbook()
        ws = wb.active
        ws.title = "Data requests"
        b2.build_requests(ws)
        wb.save(requests_path)
    from xleval import Model, run_big_stack

    def per_init():
        ctx = Model(out_path).context()
        return {it["id"]: ctx.value("Calculations", b2.ref[f"npv_{it['id']}"].split("!")[1].replace("$", ""))
                for it in s2["_initiatives"]}
    by_init = run_big_stack(per_init)
    info = {"npv": base, "npv_by_initiative": by_init, "tornado": tornado_ids, "dropped_from_tornado": dropped, "two_way": two,
            "warnings": warns, "requests": len(b2.request_rows())}
    if not quiet:
        print(f"built {out_path}")
        print(f"  base-case NPV (built-in engine): {base:,.2f} {s['meta']['currency']}")
        print("  NPV by initiative: " + "; ".join(f"{it['id']} {by_init[it['id']]:,.2f}" for it in s2["_initiatives"]))
        print(f"  tornado inputs: {len(tornado_ids)}; two-way tables: {two}")
        if dropped:
            print(f"  not in tornado (range does not move NPV): {', '.join(dropped)}")
        print(f"  data requests: {info['requests']}" + (f" -> {requests_path}" if requests_path else ""))
        for w in warns:
            print(f"  WARNING {w}")
    return info


# ----------------------------------------------------------------------------- templates
def build_templates(outdir):
    here = os.path.dirname(os.path.abspath(__file__))
    spec = os.path.join(here, "..", "templates", "example_spec.json")
    os.makedirs(outdir, exist_ok=True)
    build(spec, os.path.join(outdir, "impact_model_template.xlsx"),
          os.path.join(outdir, "data_request_sheet_template.xlsx"), quiet=True)
    # stand-alone assumptions register template
    wb = Workbook()
    ws = wb.active
    ws.title = "Inputs"
    heads = REG_HEAD[:-2]
    ws["B1"] = "Assumptions register"
    ws["B1"].font = TITLE
    ws["B2"] = ("Type: actual | assumption | uplift | cost | plan | constant | conversion. Grade = source A-F + "
                "credibility 1-6. Uplifts need a reference class (p25 / median / p75 of what comparable initiatives "
                "achieved); a base above p75 needs a written justification.")
    for j, h in enumerate(heads, 1):
        c = ws.cell(5, j, h)
        c.font, c.fill = HEAD_FONT, HEAD_FILL
        ws.column_dimensions[L(j)].width = [6, 40, 10, 11, 10, 10, 10, 9, 40, 12, 36, 7, 11, 11, 11, 30, 36, 20][j - 1]
    ex = [1, "Peak uplift in visit frequency from loyalty app", "%", "uplift", 0.02, 0.04, 0.06, 1,
          "Factiva: case evidence of comparable loyalty launches; Scopus meta-analysis", "2026-10-01",
          "Scopus > TITLE-ABS-KEY(loyalty program AND purchase frequency)", "B3", 0.02, 0.04, 0.06,
          "Reference class: 7 grocery/convenience loyalty programmes (p25 / median / p75)", "", "initiative d1"]
    for j, v in enumerate(ex, 1):
        ws.cell(6, j, v)
    wb.save(os.path.join(outdir, "assumptions_register_template.xlsx"))
    print(f"templates written to {outdir}")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("spec", nargs="?")
    ap.add_argument("out", nargs="?")
    ap.add_argument("--requests", help="also write the Data Request Sheet as its own .xlsx")
    ap.add_argument("--templates", help="regenerate the blank templates into this folder")
    ap.add_argument("--validate", action="store_true", help="validate the spec only")
    a = ap.parse_args()
    if a.templates:
        build_templates(a.templates)
        return
    if not a.spec:
        ap.error("spec is required")
    try:
        if a.validate:
            _, warns = normalise(load_spec(a.spec))
            print("spec OK" + ("" if not warns else f" with {len(warns)} warnings"))
            for w in warns:
                print(f"  WARNING {w}")
            return
        if not a.out:
            ap.error("out path is required")
        build(a.spec, a.out, a.requests)
    except SpecError as e:
        sys.exit(f"SPEC ERROR: {e}")


if __name__ == "__main__":
    main()
