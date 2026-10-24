#!/usr/bin/env python3
"""Strategy revisions: propose on copies, let the team decide, apply only what was accepted.

    python scripts/revisions.py levers   Model_vA.xlsx spec_vA.json            # where impact could come from
    python scripts/revisions.py propose  spec_vA.json proposals.json --outdir revisions/
    python scripts/revisions.py apply    spec_vA.json proposals.json --out Model_vB.xlsx

`propose` validates each proposal, builds one copy of the model per proposal (version A is never
touched), and writes Revision_Proposals.xlsx / .md with the before/after impact. `apply` refuses to run
unless at least one proposal has decision.status = "accepted" with who decided and when; it writes
spec_vB.json + the version B model, a before/after bridge with the source behind each gain, and runs
the checker in revision mode. Proposal format: references/revisions.md.
Requires: openpyxl. Optional: matplotlib.
"""
from __future__ import annotations

import argparse
import copy
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

GRADE_RE = re.compile(r"^[A-F][1-6]$")
VALUE_FIELDS = ("low", "base", "high")
ESTIMATE_TYPES = ("uplift", "assumption", "cost")


class ProposalError(Exception):
    pass


# ----------------------------------------------------------------------------- patching
def _step(cur, seg, path):
    if isinstance(cur, dict):
        if seg not in cur:
            raise ProposalError(f"path '{path}': no key '{seg}'")
        return cur[seg]
    if isinstance(cur, list):
        for x in cur:
            if isinstance(x, dict) and x.get("id") == seg:
                return x
        if seg.isdigit() and int(seg) < len(cur):
            return cur[int(seg)]
        ids = [x.get("id") for x in cur if isinstance(x, dict)]
        raise ProposalError(f"path '{path}': no item with id '{seg}' (available: {', '.join(map(str, ids))})")
    raise ProposalError(f"path '{path}': cannot step into a {type(cur).__name__}")


def apply_op(spec, op):
    path = op.get("path", "")
    segs = [x for x in path.split("/") if x]
    kind = op.get("op")
    if kind not in ("set", "add", "remove") or not segs:
        raise ProposalError(f"op must be set/add/remove with a path, got {op}")
    parent = spec
    for seg in segs[:-1]:
        parent = _step(parent, seg, path)
    last = segs[-1]
    if kind == "add":
        target = _step(parent, last, path) if isinstance(parent, (dict, list)) else None
        if not isinstance(target, list):
            raise ProposalError(f"add needs a list at '{path}'")
        v = copy.deepcopy(op["value"])
        if isinstance(v, dict) and v.get("id") in {x.get("id") for x in target if isinstance(x, dict)}:
            raise ProposalError(f"add at '{path}': id '{v.get('id')}' already exists")
        target.append(v)
    elif kind == "set":
        if isinstance(parent, list):
            item = _step(parent, last, path)
            parent[parent.index(item)] = copy.deepcopy(op["value"])
        else:
            parent[last] = copy.deepcopy(op["value"])
    else:
        if isinstance(parent, list):
            parent.remove(_step(parent, last, path))
        else:
            parent.pop(last, None)


def touched_inputs(spec_before, ops):
    """Input ids whose register row a patch changes or adds."""
    out = []
    for op in ops:
        segs = [x for x in op["path"].split("/") if x]
        if segs[0] != "inputs":
            continue
        if len(segs) == 1 and op["op"] == "add":
            out.append(op["value"]["id"])
        elif len(segs) >= 2:
            out.append(segs[1])
    return list(dict.fromkeys(out))


def accepted(props):
    out = []
    for p in props["proposals"]:
        d = p.get("decision") or {}
        if str(d.get("status", "")).lower() == "accepted":
            if not d.get("by") or not d.get("on"):
                raise ProposalError(f"{p['id']}: accepted without 'by' and 'on' - record who decided and when")
            out.append(p)
    return out


def apply_proposals(spec_a, proposals, version="B"):
    """Spec of the revised version: spec A + the given proposals, inputs tagged with the revision id."""
    s = copy.deepcopy(spec_a)
    for p in proposals:
        before = copy.deepcopy(s)
        for op in p["patch"]:
            apply_op(s, op)
        by = {i["id"]: i for i in s["inputs"]}
        for iid in touched_inputs(before, p["patch"]):
            if iid in by:
                tags = [t for t in str(by[iid].get("revision", "")).split(", ") if t]
                by[iid]["revision"] = ", ".join(dict.fromkeys(tags + [p["id"]]))
    m = s.setdefault("meta", {})
    m["version"] = version
    m["revisions"] = [p["id"] for p in proposals]
    if proposals:
        m["title"] = f"{m.get('title', m.get('company', ''))} - version {version} (accepted: " \
                     f"{', '.join(p['id'] for p in proposals)})"
    return s


# ----------------------------------------------------------------------------- validation
def _initiatives(spec):
    return {it["id"]: it for p in spec.get("pillars", []) for it in p.get("initiatives", [])}


def validate(spec_a, prop):
    """Problems with one proposal (empty list = valid). Encodes: a revision changes the strategy, not just a
    number; every changed estimate has new evidence; uplifts stay within the reference-class p75 unless
    justified in writing."""
    from build_model import init_input_ids, normalise, SpecError
    probs = []
    pid = prop.get("id", "?")
    for k in ("id", "title", "change", "why", "risk", "patch", "evidence"):
        if not prop.get(k):
            probs.append(f"{pid}: missing '{k}'")
    for e in prop.get("evidence") or []:
        miss = [k for k in ("source", "link", "date", "grade") if not e.get(k)]
        if miss:
            probs.append(f"{pid}: evidence item without {', '.join(miss)}")
        elif not GRADE_RE.match(str(e["grade"]).upper()):
            probs.append(f"{pid}: evidence grade '{e['grade']}' is not A-F + 1-6")
    if probs and not prop.get("patch"):
        return probs
    try:
        spec_b = apply_proposals(spec_a, [prop])
        normalise(spec_b)
    except (ProposalError, SpecError, KeyError) as e:
        return probs + [f"{pid}: patch does not apply cleanly: {e}"]
    in_a = {i["id"]: i for i in spec_a["inputs"]}
    in_b = {i["id"]: i for i in spec_b["inputs"]}
    inits_a, inits_b = _initiatives(spec_a), _initiatives(spec_b)
    # 1. which initiatives does the strategy change touch?
    strat_inits = set()
    for op in prop["patch"]:
        segs = [x for x in op["path"].split("/") if x]
        if segs[0] == "pillars":
            if "initiatives" in segs:
                k = segs.index("initiatives")
                if len(segs) > k + 1:
                    strat_inits.add(segs[k + 1])
                elif op["op"] == "add":
                    strat_inits.add(op["value"]["id"])
            elif op["op"] == "add" and len(segs) == 1:
                strat_inits |= {it["id"] for it in op["value"].get("initiatives", [])}
        elif segs[0] in ("leaves", "streams", "split_rules"):
            strat_inits |= set(inits_b)  # structural change to the driver tree
    for iid in touched_inputs(spec_a, prop["patch"]):
        inp = in_b.get(iid) or in_a.get(iid)
        if inp and inp.get("type") == "plan":
            strat_inits |= {k for k, it in inits_b.items() if iid in init_input_ids(it)}
    if not strat_inits:
        probs.append(f"{pid}: number-only revision - it changes estimates but no strategy decision (initiative "
                     f"definition, launch/rollout, scale or other 'plan' input). Raising a number is not a revision.")
    derived_b = {d["id"]: d for d in spec_b.get("derived", [])}

    def expand(ids):
        out, todo = set(), list(ids)
        while todo:
            x = todo.pop()
            if x in out:
                continue
            out.add(x)
            todo += derived_b.get(x, {}).get("args", [])
        return out
    for iid in touched_inputs(spec_a, prop["patch"]):
        inp = in_b.get(iid) or in_a.get(iid)
        if inp and inp.get("type") == "plan":
            strat_inits |= {k for k, it in inits_b.items() if iid in expand(init_input_ids(it))}
    linked = set()
    for k in strat_inits:
        if k in inits_b:
            linked |= expand(init_input_ids(inits_b[k]))
    # 2. every changed estimate: linked to the strategy change, new evidence, inside the reference class
    for iid in touched_inputs(spec_a, prop["patch"]):
        b, a = in_b.get(iid), in_a.get(iid)
        if b is None:
            continue
        if b.get("type") in ESTIMATE_TYPES and strat_inits and iid not in linked:
            probs.append(f"{pid}: changes estimate '{iid}', which none of the changed initiatives uses")
        if a is not None and any(a.get(f) != b.get(f) for f in VALUE_FIELDS) and b.get("type") != "plan":
            if a.get("source") == b.get("source"):
                probs.append(f"{pid}: '{iid}' value changed but its source did not - cite the new evidence")
        if b.get("type") == "uplift":
            bh = b.get("bench_high")
            if bh is None:
                probs.append(f"{pid}: uplift '{iid}' has no reference class (bench_low / bench_median / bench_high)")
            elif b["base"] > bh and not b.get("justification"):
                probs.append(f"{pid}: uplift '{iid}' base {b['base']} exceeds the reference-class p75 {bh} with no "
                             f"written justification")
            if a is not None and (a.get("bench_high") != b.get("bench_high") or
                                  a.get("bench_low") != b.get("bench_low")) and \
                    a.get("bench_source") == b.get("bench_source"):
                probs.append(f"{pid}: '{iid}' reference class moved without a new bench_source")
    return probs


# ----------------------------------------------------------------------------- measuring
def metrics(xlsx, init_ids=None):
    from defence_pack import Pack, find_label
    from xleval import run_big_stack

    def go():
        spec_stub = {"meta": {"currency": ""}}
        P = Pack(xlsx, spec_stub)
        out = {"npv": P.calc("NPV of strategy"), "npv_h": P.calc("Horizon-only NPV"),
               "irr": P.calc("IRR of strategy"), "payback": P.calc("Payback (years")}
        ws = P.m.wb["Calculations"]
        lastc = max(c for (sh, r, c) in P.m.cells if sh == "Calculations" and r == 4)
        init = {}
        for (sh, r, c), v in sorted(P.m.cells.items(), key=lambda kv: kv[0][1]):
            if sh != "Calculations" or c != 2 or not isinstance(v, str):
                continue
            ident = ws.cell(r, 1).value
            if init_ids is not None and ident not in init_ids:
                continue
            for pre, key, col in (("Incremental EBIT: ", "ebit", lastc), ("NPV: ", "npv", 5),
                                  ("Incremental revenue: ", "rev", lastc)):
                if v.startswith(pre):
                    d = init.setdefault(ident, {"name": v[len(pre):]})
                    d[key] = P.v("Calculations", r, col)
        out["by"] = init
        r = next((r for (sh, r, c), v in P.m.cells.items() if sh == "Calculations" and c == 2 and v == "Incremental EBIT"),
                 None)
        out["ebit_final"] = P.v("Calculations", r, lastc) if r else None
        out["kpi"] = {}   # the case objectives: final-year KPI level with the strategies
        for (sh, r_, c), v in P.m.cells.items():
            if sh == "Calculations" and c == 2 and isinstance(v, str) and v.startswith("KPI ") and \
                    ", with strategies: " in v:
                out["kpi"][v.split(", with strategies: ", 1)[1]] = (P.v("Calculations", r_, 5),
                                                                     ws.cell(r_, 3).value or "")
        return out
    return run_big_stack(go)


def kpi_delta(base, m):
    """Change in each case-objective KPI vs version A: [(name, unit, delta)]."""
    out = []
    for name, (v, unit) in (m.get("kpi") or {}).items():
        b = (base.get("kpi") or {}).get(name, (None, ""))[0]
        if isinstance(v, (int, float)) and isinstance(b, (int, float)):
            out.append((name, unit, v - b))
    return out


def fmt_kpis(d):
    return "; ".join(f"{n} {x * 100:+.1f} pp" if u == "%" else f"{n} {x:+,.1f} {u}" for n, u, x in d
                     if abs(x) > 5e-4) or "no change"


def build_quiet(spec, out):
    from build_model import build
    return build(spec, out, quiet=True)


# ----------------------------------------------------------------------------- workbook comparison
def _norm(v):
    if v is None or v == "":
        return None
    if hasattr(v, "ref") and hasattr(v, "t"):     # DataTableFormula
        return ("DT",) + tuple(sorted((k, str(x)) for k, x in vars(v).items() if not k.startswith("_")))
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


def diff_workbooks(p1, p2, limit=50):
    """Cells whose formula or typed value differs (charts and styles ignored)."""
    from openpyxl import load_workbook
    a, b = load_workbook(p1), load_workbook(p2)
    out = []
    for name in sorted(set(a.sheetnames) | set(b.sheetnames)):
        if name not in a.sheetnames or name not in b.sheetnames:
            out.append((name, None, "sheet only in one file", ""))
            continue
        wa, wb_ = a[name], b[name]
        cells = {(c.row, c.column) for row in wa.iter_rows() for c in row if c.value is not None}
        cells |= {(c.row, c.column) for row in wb_.iter_rows() for c in row if c.value is not None}
        for r, c in sorted(cells):
            x, y = _norm(wa.cell(r, c).value), _norm(wb_.cell(r, c).value)
            if x != y:
                out.append((name, wa.cell(r, c).coordinate, x, y))
                if len(out) >= limit:
                    return out
    return out


# ----------------------------------------------------------------------------- reports
def _money(x, cur=""):
    return f"{x:,.1f}" if isinstance(x, (int, float)) else str(x)


def _ops_text(spec_a, prop):
    lines = []
    for op in prop["patch"]:
        before = "-"
        try:
            cur = spec_a
            for seg in [x for x in op["path"].split("/") if x]:
                cur = _step(cur, seg, op["path"])
            before = cur if not isinstance(cur, (dict, list)) else "(object)"
        except ProposalError:
            before = "(new)"
        after = op.get("value")
        if isinstance(after, dict):
            after = f"new {after.get('id', 'item')}: {after.get('name', '')}"
        lines.append((op["op"], op["path"], before, after if op["op"] != "remove" else "(removed)"))
    return lines


def write_proposals_xlsx(path, spec_a, props, results, base):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    cur = spec_a["meta"]["currency"]
    wb = Workbook()
    ws = wb.active
    ws.title = "Proposals"
    ws["B1"] = f"Revision proposals - {spec_a['meta'].get('company', '')} (version A untouched)"
    ws["B1"].font = Font(bold=True, size=14)
    ws["B2"] = ("The team decides. Set Decision to accepted / rejected, fill 'Decided by' and 'Date', and copy the "
                "decision into proposals.json; only accepted proposals are applied (revisions.py apply).")
    heads = ["ID", "Title", "What changes in the strategy", "Why (tornado / reference class)", "Evidence",
             "Reference-class check", f"NPV A ({cur})", f"NPV with proposal ({cur})", f"Change in NPV ({cur})",
             f"Change in final-year EBIT ({cur})", "Risk", "Validation", "Model copy", "Decision", "Decided by",
             "Date"]
    widths = [6, 30, 50, 50, 60, 36, 12, 14, 12, 14, 40, 40, 34, 11, 16, 12]
    hf, hfont = PatternFill("solid", fgColor="1F3864"), Font(bold=True, color="FFFFFF")
    for j, (h, w) in enumerate(zip(heads, widths), 1):
        c = ws.cell(4, j, h)
        c.fill, c.font = hf, hfont
        c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[chr(64 + j)].width = w
    for k, p in enumerate(props["proposals"]):
        r = 5 + k
        res = results[p["id"]]
        ev = "; ".join(f"{e.get('source')} ({e.get('grade')}, {e.get('date')}) {e.get('link')}"
                       for e in p.get("evidence", []))
        d = p.get("decision") or {}
        vals = [p["id"], p.get("title"), p.get("change"), p.get("why"), ev, res.get("refclass", ""),
                base["npv"], res.get("npv"), res.get("dnpv"), res.get("debit"), p.get("risk"),
                "OK" if not res["problems"] else " | ".join(res["problems"]), res.get("copy", ""),
                d.get("status", "pending"), d.get("by", ""), d.get("on", "")]
        for j, v in enumerate(vals, 1):
            c = ws.cell(r, j, v)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            if j in (7, 8, 9, 10):
                c.number_format = "#,##0.0;[Red]-#,##0.0"
    ws.freeze_panes = "C5"
    ch = wb.create_sheet("Changes")
    for j, h in enumerate(["Proposal", "Op", "Path in spec", "Version A", "Proposed"], 1):
        c = ch.cell(1, j, h)
        c.fill, c.font = hf, hfont
    r = 2
    for p in props["proposals"]:
        for op, where, before, after in _ops_text(spec_a, p):
            for j, v in enumerate([p["id"], op, where, str(before), str(after)], 1):
                ch.cell(r, j, v)
            r += 1
    for col, w in zip("ABCDE", (10, 8, 50, 40, 60)):
        ch.column_dimensions[col].width = w
    wb.save(path)


def refclass_note(spec_b, prop):
    notes = []
    by = {i["id"]: i for i in spec_b["inputs"]}
    for iid in touched_inputs(spec_b, prop["patch"]):
        i = by.get(iid)
        if i and i.get("type") == "uplift" and i.get("bench_high") is not None:
            pos = "<= median" if i.get("bench_median") is not None and i["base"] <= i["bench_median"] else \
                ("<= p75" if i["base"] <= i["bench_high"] else "ABOVE p75 (justified)" if i.get("justification")
                 else "ABOVE p75")
            notes.append(f"{iid}: {i['base']} {pos} (p75 {i['bench_high']})")
    return "; ".join(notes) or "no uplift changed"


def cmd_propose(a):
    spec_a, props = json.load(open(a.spec)), json.load(open(a.proposals))
    os.makedirs(os.path.join(a.outdir, "copies"), exist_ok=True)
    stem = re.sub(r"[^A-Za-z0-9]+", "_", spec_a["meta"].get("company", "Model")).strip("_")
    pa = os.path.join(a.outdir, "copies", f"{stem}_vA_reference.xlsx")
    build_quiet(spec_a, pa)
    base = metrics(pa)
    results, good = {}, []
    for p in props["proposals"]:
        probs = validate(spec_a, p)
        res = {"problems": probs}
        if not any("does not apply" in x for x in probs):
            sb = apply_proposals(spec_a, [p], version=f"A+{p['id']}")
            out = os.path.join(a.outdir, "copies", f"{stem}_{p['id']}.xlsx")
            build_quiet(sb, out)
            m = metrics(out)
            res.update(npv=m["npv"], dnpv=m["npv"] - base["npv"],
                       dnpv_h=(m["npv_h"] or 0) - (base["npv_h"] or 0),
                       dkpi=kpi_delta(base, m),
                       debit=(m["ebit_final"] or 0) - (base["ebit_final"] or 0), copy=os.path.relpath(out, a.outdir),
                       refclass=refclass_note(sb, p), by=m["by"])
            if not probs:
                good.append(p)
        results[p["id"]] = res
    combo = None
    if len(good) > 1:
        sb = apply_proposals(spec_a, good, version="A+all")
        out = os.path.join(a.outdir, "copies", f"{stem}_all_valid.xlsx")
        build_quiet(sb, out)
        combo = metrics(out)
    xl = os.path.join(a.outdir, "Revision_Proposals.xlsx")
    write_proposals_xlsx(xl, spec_a, props, results, base)
    cur = spec_a["meta"]["currency"]
    L = [f"# Revision proposals - {spec_a['meta'].get('company', '')}", "",
         f"Version A NPV: **{_money(base['npv'])} {cur}** (horizon-only {_money(base['npv_h'])}). Version A is "
         f"unchanged; each proposal was modelled in its own copy (`copies/`). **Nothing is applied until the team "
         f"decides.**", "",
         f"| ID | Proposal | Change in the case objectives ({spec_a['meta']['base_year'] + spec_a['meta']['horizon_years']}, "
         f"vs A) | Change in NPV ({cur}) | of which within the horizon | Reference class | Validation |",
         "|---|---|---|---:|---:|---|---|"]
    for p in props["proposals"]:
        r = results[p["id"]]
        L.append(f"| {p['id']} | {p.get('title')} | {fmt_kpis(r.get('dkpi') or [])} | {_money(r.get('dnpv'))} | "
                 f"{_money(r.get('dnpv_h'))} | {r.get('refclass', '-')} | "
                 f"{'OK' if not r['problems'] else 'NOT VALID'} |")
    if combo:
        L.append(f"\nAll valid proposals together: NPV {_money(combo['npv'])} {cur} "
                 f"({_money(combo['npv'] - base['npv'])} vs A; effects interact, so this is not the sum).")
    for p in props["proposals"]:
        r = results[p["id"]]
        L += ["", f"## {p['id']}. {p.get('title')}", f"**What changes:** {p.get('change')}", "",
              f"**Why:** {p.get('why')}", "", "**Evidence:**"]
        L += [f"- {e.get('finding', '')} - {e.get('source')} ({e.get('grade')}, {e.get('date')}) {e.get('link')}"
              for e in p.get("evidence", [])]
        L += ["", f"**Effect on the case objectives (vs version A):** {fmt_kpis(r.get('dkpi') or [])}."]
        L += ["", f"**Modelled impact:** NPV {_money(base['npv'])} -> {_money(r.get('npv'))} {cur} "
                  f"({_money(r.get('dnpv'))}, of which {_money(r.get('dnpv_h'))} within the horizon and the rest "
                  f"terminal value); change in final-year EBIT {_money(r.get('debit'))}.",
              f"**Reference class:** {r.get('refclass', '-')}", f"**Risk:** {p.get('risk')}"]
        dn, dh = r.get("dnpv"), r.get("dnpv_h")
        if isinstance(dn, (int, float)) and isinstance(dh, (int, float)) and dn > 0 and dh < 0.2 * dn:
            L.append(f"> **Mostly terminal value:** only {_money(dh)} of the {_money(dn)} gain falls within the "
                     f"horizon. Check that it is a real effect and not timing alone (e.g. opening faster).")
        if r["problems"]:
            L.append("**Not valid - fix before presenting:** " + "; ".join(r["problems"]))
        L.append("\n**Team decision:** ☐ accept ☐ reject  (who: ____  date: ____)")
    md = os.path.join(a.outdir, "Revision_Proposals.md")
    open(md, "w").write("\n".join(L) + "\n")
    print(f"version A NPV {base['npv']:,.2f} {cur}")
    for p in props["proposals"]:
        r = results[p["id"]]
        print(f"  {p['id']}: dNPV {_money(r.get('dnpv'))}  {'OK' if not r['problems'] else 'NOT VALID: ' + '; '.join(r['problems'])}")
    print(f"written {xl} and {md}")
    return 0 if all(not results[p["id"]]["problems"] for p in props["proposals"]) else 1


def draw_before_after(rows, cur, path, title):
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
    except ImportError:
        return False
    labels = [r[0] for r in rows]
    fig, ax = plt.subplots(figsize=(10, 4.6))
    xs = range(len(rows))
    ax.bar([x - 0.2 for x in xs], [r[1] for r in rows], width=0.4, color="#7F7F7F", label="Version A")
    ax.bar([x + 0.2 for x in xs], [r[2] for r in rows], width=0.4, color="#2E75B6", label="Version B")
    ax.axhline(0, color="#404040", lw=0.6)
    ax.set_xticks(list(xs))
    ax.set_xticklabels([l if len(l) < 28 else l[:26] + "..." for l in labels], fontsize=8, rotation=15)
    ax.set_ylabel(cur)
    ax.set_title(title, fontsize=12)
    ax.legend(frameon=False)
    ax.spines[["top", "right"]].set_visible(False)
    fig.tight_layout()
    fig.savefig(path, dpi=160)
    plt.close(fig)
    return True


def cmd_apply(a):
    spec_a, props = json.load(open(a.spec)), json.load(open(a.proposals))
    acc = accepted(props)
    if not acc:
        sys.exit("NOTHING APPLIED: no proposal has decision.status = 'accepted' (with 'by' and 'on'). "
                 "Present Revision_Proposals.md to the team and record their decision first.")
    for p in acc:
        probs = validate(spec_a, p)
        if probs:
            sys.exit("NOTHING APPLIED: accepted proposal is not valid: " + "; ".join(probs))
    out = os.path.abspath(a.out)
    va = props.get("version_a", {}).get("model")
    if va and os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(a.proposals)), va)) == out:
        sys.exit("refusing to overwrite version A - give version B its own file name")
    letter = a.version
    sb = apply_proposals(spec_a, acc, version=letter)
    spec_out = os.path.splitext(out)[0] + "_spec.json"
    json.dump(sb, open(spec_out, "w"), indent=1)
    build_quiet(sb, out)
    tmp_a = os.path.splitext(out)[0] + "_vA_reference.xlsx"
    build_quiet(spec_a, tmp_a)
    from build_model import init_input_ids
    ids_a, ids_b = set(_initiatives(spec_a)), set(_initiatives(sb))
    ma, mb = metrics(tmp_a, ids_a), metrics(out, ids_b)
    os.remove(tmp_a)
    cur = spec_a["meta"]["currency"]
    fy = spec_a["meta"]["base_year"] + spec_a["meta"]["horizon_years"]
    in_b = {i["id"]: i for i in sb["inputs"]}
    rows = []
    for k in list(dict.fromkeys(list(ma["by"]) + list(mb["by"]))):
        ea = (ma["by"].get(k) or {}).get("npv", 0) or 0
        eb = (mb["by"].get(k) or {}).get("npv", 0) or 0
        name = (mb["by"].get(k) or ma["by"].get(k))["name"]
        rows.append((name, ea, eb, k))
    inits_b = _initiatives(sb)
    L = [f"# Before / after: version A vs version {letter}", "",
         f"Accepted by the team: " + "; ".join(f"**{p['id']}** {p['title']} (decided by {p['decision']['by']} on "
                                               f"{p['decision']['on']})" for p in acc), "",
         f"NPV: {_money(ma['npv'])} -> **{_money(mb['npv'])} {cur}** ({_money(mb['npv'] - ma['npv'])}). "
         f"Horizon-only NPV: {_money(ma['npv_h'])} -> {_money(mb['npv_h'])}. "
         f"Incremental EBIT {fy}: {_money(ma['ebit_final'])} -> {_money(mb['ebit_final'])}.", ""]
    png = os.path.join(os.path.dirname(out), "before_after_bridge.png")
    if draw_before_after([(n, x, y) for n, x, y, _ in rows], cur, png, f"NPV by initiative: version A vs {letter} ({cur})"):
        L.append("![Before / after](before_after_bridge.png)\n")
    L += [f"| Initiative | NPV A | NPV {letter} | Change | Accepted change and the source behind it |",
          "|---|---:|---:|---:|---|"]
    for name, x, y, k in rows:
        src = []
        for p in acc:
            uses = set(init_input_ids(inits_b[k])) if k in inits_b else set()
            direct = any(f"/initiatives/{k}" in op["path"] or (isinstance(op.get("value"), dict) and
                                                                op["value"].get("id") == k) for op in p["patch"])
            hit = [iid for iid in touched_inputs(spec_a, p["patch"]) if iid in uses]
            if direct or hit:
                src.append(f"{p['id']} ({p['title']}): " + "; ".join(
                    [f"{e['source']} [{e['grade']}]" for e in p["evidence"]] +
                    [f"{in_b[i]['name']} = {in_b[i]['base']} ({in_b[i].get('source')}, {in_b[i].get('grade')})"
                     for i in hit if i in in_b]))
        why = " / ".join(src) if src else ("interaction: shares a revenue stream with a changed initiative (log-share split)" if abs(y - x) > 1e-9 else "-")
        L.append(f"| {name} | {_money(x)} | {_money(y)} | {_money(y - x)} | {why} |")
    L.append(f"| **Total (strategy NPV)** | **{_money(ma['npv'])}** | **{_money(mb['npv'])}** | "
             f"**{_money(mb['npv'] - ma['npv'])}** | |")
    md = os.path.join(os.path.dirname(out), f"Before_After_v{letter}.md")
    open(md, "w").write("\n".join(L) + "\n")
    print(f"version {letter} written: {out}\n  spec: {spec_out}\n  before/after: {md}")
    print(f"  NPV {ma['npv']:,.2f} -> {mb['npv']:,.2f} {cur}")
    print(f"next: python scripts/check_model.py {a.out} --version-a <Model_vA.xlsx> --spec-a {a.spec} "
          f"--proposals {a.proposals}")
    return 0


def cmd_levers(a):
    from defence_pack import Pack
    from xleval import run_big_stack
    from build_model import init_input_ids
    spec = json.load(open(a.spec))

    def go():
        P = Pack(a.xlsx, spec)
        base, torn = P.tornado()
        plan_ids = {i["id"] for i in spec["inputs"] if i.get("type") == "plan"}
        print(f"Base NPV {base:,.2f} {spec['meta']['currency']}. Top inputs by NPV swing:")
        print(f"{'input':<12} {'type':<10} {'swing':>10}  reference class (p25 / median / p75) and position")
        for t in torn[:a.top]:
            i, sp = P.inputs[t["key"]], next((x for x in spec["inputs"] if x["id"] == t["key"]), {})
            rc = ""
            if sp.get("type") == "uplift" and sp.get("bench_high") is not None:
                head = sp["bench_high"] - sp["base"]
                rc = (f"{sp.get('bench_low')} / {sp.get('bench_median')} / {sp['bench_high']}; base {sp['base']} "
                      f"-> headroom to p75 {head:+.4g}")
            print(f"{t['key']:<12} {str(i['type']):<10} {t['swing']:>10,.1f}  {rc}")
        print("\nStrategy decisions you can revise (plan inputs and initiative design):")
        for it in [x for p in spec["pillars"] for x in p.get("initiatives", [])]:
            dec = [x for x in init_input_ids(it) if x in plan_ids]
            ramp = (it.get("ramp") or {}).get("type", "step")
            print(f"  {it['id']}: {it['name']} - launch {it.get('launch_year')}, ramp {ramp}"
                  + (f", decisions: {', '.join(dec)}" if dec else ""))
        print("\nA revision must change one of these decisions (or add/replace an initiative) and cite evidence; "
              "estimates may move only as a consequence, inside the reference class.")
    run_big_stack(go)
    return 0


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("levers")
    p.add_argument("xlsx")
    p.add_argument("spec")
    p.add_argument("--top", type=int, default=8)
    p = sub.add_parser("propose")
    p.add_argument("spec")
    p.add_argument("proposals")
    p.add_argument("--outdir", default="revisions")
    p = sub.add_parser("apply")
    p.add_argument("spec")
    p.add_argument("proposals")
    p.add_argument("--out", required=True)
    p.add_argument("--version", default="B")
    a = ap.parse_args()
    try:
        sys.exit({"levers": cmd_levers, "propose": cmd_propose, "apply": cmd_apply}[a.cmd](a))
    except ProposalError as e:
        sys.exit(f"PROPOSAL ERROR: {e}")


if __name__ == "__main__":
    main()
