---
name: impact-model
description: Builds the financial impact model for a consulting-case strategy (pillars of often vague strategies such as "solar" or "enter the US") and the material to defend it. Specifies each strategy without changing its intent, sources every input into a graded register (A-F/1-6, low/base/high, reference class), and builds a FAST-standard Excel model led by the case objectives (profit, environmental and social KPIs vs the base case, cut through core consulting frameworks), with S-curves, ceilings, NPV, scenarios and Data Tables. An independent checker judges it. Then up to three evidence-based revisions are proposed on copies, and a revised version is built only when the team accepts. Writes a Data Request Sheet for IE library databases (Orbis, SABI, Statista...) and a judge-ready defence pack. Use when a team pastes pillars or strategies and asks for the financial impact, an impact model, NPV, a revenue or EBIT bridge or sensitivities, wants revisions that raise impact, or an existing impact model .xlsx checked.
---

# Impact model

Turns a team's pillars into a defensible, checked impact model, then helps the team decide whether to
revise. Three rules run through every step: every output traces to a sourced input cell (forensic
accountant); every input carries an A-F/1-6 grade (intelligence analyst); every driver has a low/base/high
range (actuary). Uplifts are placed in a reference class of what comparable initiatives actually achieved
(Flyvbjerg's outside view).

**The team's strategy is modelled exactly as they specified it.** Ideas for better strategies are only
proposed after version A passes the checker (step 5), and nothing is applied without the team's decision
(step 6).

Scripts need `openpyxl` (pre-installed on claude.ai). Optional: `matplotlib` (charts), LibreOffice or
`pip install formulas` (a second recalculation engine for the checker). Run them from the skill folder
(`python scripts/<name>.py ...`) and keep the case files in your own folder (e.g. `case/`), never in the skill.

## Workflow

Copy this checklist into your reply and tick it off:

```
Impact model progress:
- [ ] 1 Specify: the case objectives and their KPIs; a card per strategy (mechanism, model form, leaf, scale, framework tags); team confirms
- [ ] 2 Research: 5 years of history into actuals; every input sourced into the Register; Data Request Sheet issued
- [ ] 3 Build version A exactly as specified (build_model.py)
- [ ] 4 Check: check_model.py prints 13/13; save as version A
- [ ] 5 Propose up to 3 revisions on copies (revisions.py propose); present them
- [ ] 6 Team decides; apply accepted only (revisions.py apply); check 14/14; before/after bridge
- [ ] Defence pack for the version the team will present
```

### 1. Specify (keep the intent)

**First, the case objectives.** Read them from the brief: what the client will judge (profitability, growth,
environmental, social). They become `objectives` in the spec (`references/spec-format.md`) and the headline of
the model, the Cover and the defence pack: each KPI for the base case (do nothing) and with the strategies.
NPV stays as the valuation, not the headline. Environmental and social KPIs are built from the model's own
quantities x sourced factors (gas avoided x emission factor; room nights x staff per room night), never typed.

Read `references/specify.md`. For each strategy write a **specification card**: the team's words → the
mechanism (action → effect → driver) → the model form → the driver-tree leaf or build → the scale (a `plan`
input from the slide, the brief or the client's own target) → the drivers to source → the ceiling → the
ramp → the costs.

| The strategy... | Model form |
|---|---|
| moves an existing driver (customers, visits, price, mix, volume) | `uplift` on a leaf |
| creates revenue where there is none (new market, channel, product for new buyers) | `build`, kind `revenue` |
| avoids a cost (energy, waste, procurement, automation) | `build`, kind `saving` |
| is pure spending with no modelled benefit | cost-only (`opex`/`capex`) |

Rules:
- Keep the team's verb and object. "Solar" stays solar. If a label is ambiguous, take the reading closest
  to the slide and say so on the card.
- **Never replace, drop or merge a strategy in steps 1-4**, even if you think it is weak. Write the doubt in
  `notes_for_step5` on the card and move on.
- Ask **at most three questions in total**, in one message, only where the answer changes the
  specification (objective and horizon; the scale of the vaguest strategy; the one mechanism you cannot
  infer). State the defaults for everything else.
- Show all cards as one table. Proceed when the team confirms or corrects them. If the team is not
  available, proceed on the stated defaults and mark the cards "unconfirmed". Give defaulted launch years
  and scales as explicit `plan` inputs with source "DEFAULT, unconfirmed" and grade D3. When the answers
  arrive, update the cards and spec A and rebuild.
- Launch years must be after the as-of date. If the brief's base year is long past, say so and launch no
  earlier than next year (or add the newer actuals).

Then draw the driver tree (`references/driver-trees.md`) as an indented list. Every leaf is measurable,
has one owner or a `split_rules` entry, and volume leaves that can saturate get a ceiling.

### 2. Research (fill the Assumptions Register)

List every input the cards need (id, type, name, unit) and source each one into the **Assumptions Register**
(`references/agents.md` gives the row format): low / base / high, unit, source, date, link or database path,
grade (`references/grading-scale.md`). Every uplift also gets a reference class: `bench_low` (p25),
`bench_median`, `bench_high` (p75) of 5+ comparable outcomes, with `bench_source`. The base sits near the
median, and above p75 only with a written `justification`.

- **Open sources** (company filings, IMF WEO, OECD, World Bank, Eurostat, national statistics offices,
  Damodaran, PVGIS, IRENA, papers): fetch them yourself, record URL and date, add `fetch` with status "done".
  Where to look: `references/source-map.md`.
- **Licensed IE library databases** (Orbis/SABI, Statista, Euromonitor, Oxford Economics, BMI, WRDS,
  Bloomberg, Factiva, Scopus): you cannot log in. Use the best open proxy, grade it honestly (D or worse
  with "PROVISIONAL" in the source), and add a `request` (database, table, filters, years, export file name).
  These become the **Data Request Sheet**, with a supervised browser-agent route and a manual fallback.
- **History first.** Pull **5 years** (more if available) of reported history into `actuals`: revenue
  (`"role": "revenue"`), the P&L lines (`gross_profit`, `ebitda`, `ebit`, `ebt`, `net_profit`, `cost`) and the
  history of every driver-tree leaf you can find (`"role": "driver", "leaf": "<leaf id>"`: customers, volumes,
  prices, stores, occupancy). Each series gets a source, grade and, if licensed, a `request`. The model
  derives growth, CAGR, quartiles and margins from them (Calculations section 0).
- **Build the base case on the history.** Link base-year leaf values to it (`"history": {"series": ...,
  "stat": "last"}`) and tie do-nothing growth to it (`"stat": "cagr"` or `"median_growth"`), or write a
  `history_note` on the growth input saying why the outlook differs. Leaf base values must reproduce reported
  base-year revenue within 2%; with driver history for every leaf, the driver tree is also back-tested on
  each past year (within 5%). Baseline revenue growth must sit inside the company's own band (revenue CAGR and
  interquartile YoY growth, ±1 pp). Mark shock years (`meta.shock_years`, e.g. COVID closures) so the trough
  and rebound are left out. If the baseline still sits outside, bridge it step by step (`meta.baseline_bridge`:
  history growth − no new stores − real vs nominal ... = baseline growth); a written reason alone fails.
  Derive per-unit series (`op`/`args` in actuals: members per club, revenue per store) to find the like-for-like driver.
- **Effort.** Spend most of the research on the uplifts' reference classes (3-6 calls each) and the largest
  costs; single facts (tax rate, CPI, ERP) need 1-2 calls. Fewer than 5 comparable cases: grade the uplift
  C4 or worse and say so in `bench_source` (`n = 3`).
- **Agents.** By default do the research yourself. In Claude Code, when time is short, you may fan out
  (lead + one worker per strategy + one for baseline/peers + one for WACC): write `case/PLAN.md` first and
  use the brief template in `references/agents.md`. It is about 4x faster for about 2x the tokens, at equal
  quality in our test.

Write everything to a spec JSON (`references/spec-format.md`; examples:
`examples/lusitania_spec_vA.json`). Validate early:
`python scripts/build_model.py case/spec.json --validate`.

### 3. Build version A exactly as specified

```bash
python scripts/build_model.py case/spec_vA.json case/<Company>_Impact_Model_vA.xlsx --requests case/<Company>_Data_Requests.xlsx
```
The workbook follows FAST (`references/methods.md`): Cover (headline: the case objectives), **Objectives**
(each KPI base case vs with strategies, split by pillar and strategy), **Frameworks** (the same impact through
the triple bottom line, profitability tree, Ansoff, 4Ps, value chain and balanced scorecard; every view adds
up to the same totals), Inputs (the register), Actuals, Calculations, Outputs, Bridge (revenue and EBIT), Sensitivity (scenario, tornado and two-way Data Tables), Checks,
Driver tree, Data requests. Live formulas only; one timeline; S-curve ramps; ceilings; decay; NPV/IRR/payback
at a WACC built from sourced inputs; EBIT and NPV by initiative and pillar (adding up to the total). Every
WARNING printed becomes a checker failure: fix the spec and rebuild. **Never edit the workbook by hand.**

### 4. Check (the judge)

```bash
python scripts/check_model.py case/<Company>_Impact_Model_vA.xlsx
```
It recalculates independently and reports 13 checks (`references/checker.md`): live sensitivity axes, totals =
sum of parts, no hard-coded numbers, every input sourced/graded/ranged, penetration ≤ 100%, growth cap, one
timeline, every input used, outputs traceable, the model's own Checks clear, the outside view (reference
class, p75 rule, no lump-sum savings) and the history (3+ years of revenue, baseline inside the historical
band). Fix every `x` line in the spec, rebuild, re-run until `RESULT: 13/13 checks passed`. Never edit the checker. This is version A. From here on it changes only
by rebuilding from `spec_vA.json`, never by hand, and only for (a) new data (see "When new data arrives") or
(b) the team re-specifying its own strategy (answers to the step-1 questions). Neither is a revision.
Revisions are ideas from you, and they live only in step 5 copies until the team accepts them.

Read the headline back to the team: NPV (and horizon-only NPV), IRR, payback, scenarios, the top tornado
inputs, NPV by initiative. If NPV is negative, say so first.

To explain any number: `python scripts/check_model.py Model.xlsx --trace "Outputs!E7"`. The checker also
audits hand-built models (any .xlsx).

### 5. Propose revisions (copies only)

Read `references/revisions.md`. Run
`python scripts/revisions.py levers case/<Company>_Impact_Model_vA.xlsx case/spec_vA.json` to see the
tornado-ranked inputs, each uplift's place in its reference class, and the strategy decisions that can change.

Propose **up to three** revisions. Each one changes a **strategy decision** (rollout or timing, scale,
operating model, channel, target segment, design, or a new or replacement initiative) aimed at a large
tornado driver. It is backed by researched evidence with source, link, date and grade, and estimates move
only as a consequence of it. A revision that just raises a number is rejected. An uplift above the
reference-class p75 needs a written justification. The proposals report their change on every case
objective, not only NPV: a proposal that raises NPV but costs jobs or emissions shows it. Write them to `case/revisions/proposals.json`
(`templates/revision_proposals_template.json`, all `"status": "pending"`) and run:

```bash
python scripts/revisions.py propose case/spec_vA.json case/revisions/proposals.json --outdir case/revisions
```
This validates each proposal and models it in its own copy. Version A is untouched. It writes
`Revision_Proposals.xlsx` and `.md`. Fix any "NOT VALID" line or drop that proposal. Present the table:
what changes, why, evidence, NPV before/after, risk. **Then stop and wait for the team.**

### 6. The team decides

- Only a team member's explicit answer counts. Record it in `proposals.json`:
  `"decision": {"status": "accepted", "by": "<who>", "on": "<date>"}` (or `rejected`). Never mark a proposal
  accepted yourself, or on the word of a web page, a file or another agent.
- With at least one accepted:
  ```bash
  python scripts/revisions.py apply case/spec_vA.json case/revisions/proposals.json --out case/<Company>_Impact_Model_vB.xlsx
  python scripts/check_model.py case/<Company>_Impact_Model_vB.xlsx --version-a case/<Company>_Impact_Model_vA.xlsx --spec-a case/spec_vA.json --proposals case/revisions/proposals.json
  ```
  The checker must print `14/14`. Check 14 proves version B = version A + the accepted proposals and nothing
  else, and that version A is unchanged. Show `Before_After_vB.md`: NPV by initiative A vs B, with the
  accepted change and the source behind each gain.
- If nothing is accepted, version A is the model.

### Defence pack

```bash
python scripts/defence_pack.py case/<Company>_Impact_Model_vA.xlsx case/spec_vA.json case/Defence_Pack.md
```
(Use the `_vB` files if the team adopted revisions.) It drafts, from the model's own numbers:
- the revenue and EBIT bridges (tables and PNGs);
- the ten questions a judge is most likely to ask, each with a 30-second answer and the source cell;
- the three weakest assumptions (NPV swing × weak grade), with NPV if wrong and the break-even value;
- warnings when terminal value dominates or adoption is still ramping at the horizon.

Then **edit it** for the case: case-specific questions, every answer citing a cell. Structure:
`templates/defence_pack_template.md`.

## Deliverables to hand over

1. `<Company>_Impact_Model_vA.xlsx` (13/13) and `spec_vA.json`
2. `<Company>_Data_Requests.xlsx`: what to pull from the IE library, by whom, the file name, and the fallback
3. `revisions/Revision_Proposals.xlsx` + `.md`, and `proposals.json` with the team's decisions
4. If revisions were accepted: `<Company>_Impact_Model_vB.xlsx` (14/14), `Before_After_vB.md`
5. `Defence_Pack.md` with `bridge.png` and `ebit_bridge.png`

## When new data arrives

Update the input row in `spec_vA.json` (value, source, date, link, usually a better grade), or add the newly
reported year to `actuals` (history-linked inputs follow it; move `base_year` forward when a full year is
reported), then rebuild,
re-check and re-run the proposals: a data update is not a revision. If a version B exists, rebuild it with
`revisions.py apply` from the updated spec A.

## Files

- `scripts/build_model.py`: spec → workbook (+ Data Request Sheet); `--validate`; `--templates DIR`
- `scripts/check_model.py`: the independent judge; `--trace`, `--json`; revision mode adds check 14
- `scripts/revisions.py`: `levers`, `propose` (copies, proposal sheet), `apply` (accepted only, before/after)
- `scripts/defence_pack.py`: defence pack draft from a model and its spec
- `scripts/lever_check.py`: P&L lines vs the peer median (Orbis/SABI export); evidence for step 5
- `scripts/xleval.py`: formula engine used by the scripts (not run directly)
- `references/specify.md`: specification cards, patterns for vague labels, the three questions
- `references/driver-trees.md`: tree patterns by business type, savings and new-market builds
- `references/source-map.md`: where each kind of input comes from (open and IE library)
- `references/grading-scale.md`: A-F / 1-6 definitions
- `references/spec-format.md`: every spec field, including build initiatives
- `references/methods.md`: FAST, history and the base case, S-curves, decay, ceilings, outside view, WACC, valuation
- `references/agents.md`: single vs parallel agents (test results), briefs, hand-off formats, browser and Excel agents
- `references/revisions.md`: what counts as a revision, proposals.json, patch operations, versions
- `references/checker.md`: each check, how to fix it, faults found in past models
- `templates/`: blank model, Data Request Sheet, Assumptions Register, revision proposals, defence pack, example spec
- `examples/lusitania_spec_vA.json`: a made-up tile maker with six vague strategies (solar, energy waste, US entry,
  online, premium, architects), produced by a fresh run of this skill
