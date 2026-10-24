# Spec format (input to build_model.py)

## Contents
- Skeleton
- meta
- objectives (the headline: what the case is judged on)
- actuals
- inputs (the assumptions register)
- derived inputs
- leaves and streams (the driver tree)
- pillars and initiatives
- build initiatives (new revenue streams, cost savings)
- split_rules
- finance
- sensitivity
- Validation errors you may see

Complete examples: `examples/lusitania_spec_vA.json`, `examples/forza_spec_vA.json`. A neutral one: `templates/example_spec.json`.
Validate without building: `python scripts/build_model.py spec.json --validate`.

## Skeleton

```json
{
  "meta": {...},
  "actuals": [...],
  "inputs": [...],
  "leaves": [...],
  "streams": [...],
  "pillars": [{"id": "...", "name": "...", "initiatives": [...]}],
  "split_rules": {},
  "finance": {...},
  "sensitivity": {}
}
```

## meta

| Key | Required | Meaning |
|---|---|---|
| company, currency, base_year, horizon_years | yes | currency is the unit label everywhere, e.g. "EURm" |
| title, objective, prepared_by, as_of | no | objective: "revenue growth", "profitability", "cost", "cash" |
| first_hist_year | no | first actuals column (default base_year) |
| growth_cap | no | max YoY growth of target revenue, default 0.25 |
| calibration_tolerance | no | max gap model vs reported base-year revenue, default 0.02 |
| history_tolerance | no | baseline growth may sit this far outside the historical band, default 0.01 |
| backtest_tolerance | no | max yearly gap of the driver-tree back-test vs reported revenue, default 0.05 |
| shock_years, shock_note | no | years left out of the history statistics with the rebound after them, and what happened (e.g. `[2020, 2021]`, "COVID closures") |
| baseline_bridge | no | steps from historical growth to the do-nothing growth, required when it sits outside the band (below) |
| baseline_justification | no | the narrative behind the bridge (Checks, Outputs, defence pack); on its own it does not pass check 13 |

## objectives (the headline: what the case is judged on)

The headline of the model, the Cover and the defence pack is the case objectives, each shown for the base case
(do nothing) and with the strategies. NPV stays as the valuation. Take the objectives from the brief.

```json
"objectives": [
  {"id": "profit", "name": "Grow profit", "kind": "profitability"},
  {"id": "planet", "name": "Greener hotels", "kind": "environmental", "kpis": [
    {"id": "co2", "name": "CO2e from gas for hot water", "unit": "tCO2e", "direction": "lower",
     "baseline": ["gas_gwh", "ef_gas"],
     "effects": [{"initiative": "g1", "quantity": "volume", "factor": ["ef_gas"], "sign": -1}]}]},
  {"id": "people", "name": "Year-round jobs", "kind": "social", "kpis": [
    {"id": "jobs", "name": "Year-round jobs added", "unit": "FTE", "direction": "higher", "stock": true,
     "effects": [{"initiative": "e2", "quantity": "volume", "factor": ["fte_krn"], "per": ["per_k"]}]}]}]
```

| Field | Notes |
|---|---|
| kind | `profitability` (built in: EBITDA and EBITDA margin vs the base case when actuals have role `ebitda`; else the profit uplift), `growth` (built in: revenue), `environmental`, `social`, `other` (these need `kpis`) |
| kpis[].direction | `higher` or `lower` is better (Objectives shows "better" / "worse") |
| kpis[].baseline | input or derived ids multiplied together: the base-case level in the base year (e.g. gas used x emission factor). Omit it when only the change is counted (jobs added). `baseline_growth`: an input growing it per year |
| kpis[].effects | how each strategy moves the KPI: `quantity` of an initiative (`volume` of a build, `revenue`, `saving`, `opex`, `capex`, `ebit`) x `factor` ids / `per` ids, `sign` -1 for a reduction. Factors are sourced, graded inputs (emission factors: DESNZ/DEFRA, MITECO, IEA; jobs per unit: national statistics) |
| kpis[].stock | true for a level (jobs, members): the cumulative column is left empty |

Without `objectives` the headline defaults to profitability (with a warning).

**Framework tags** (on each initiative, for the Frameworks sheet):

```json
{"id": "e2", "name": "Open all year", ..., "frameworks": {"ansoff": "market_development", "four_p": "place",
 "value_chain": "operations"}}
```

| Tag | Values | Default when missing |
|---|---|---|
| ansoff | `penetration`, `product_development`, `market_development`, `diversification`, `efficiency` | uplift on a leaf: `penetration`; saving or cost-only: `efficiency`; a new revenue build: not tagged (warning) |
| four_p | `price`, `product`, `place`, `promotion`, `none` | uplift on a price leaf: `price`; saving: `none`; else not tagged |
| value_chain | `inbound_logistics`, `operations`, `outbound_logistics`, `marketing_sales`, `service`, `procurement`, `technology`, `hr`, `infrastructure` | not tagged (warning) |

## actuals

Reported history, one series per line: **5 years** if they exist (the checker needs 3+). Exactly one series has
`"role": "revenue"` (total reported revenue: calibration and the history band).

```json
{"id": "rev", "name": "Revenue (reported)", "unit": "EURm", "role": "revenue",
 "values": {"2020": 371.5, "2021": 389.0, "2022": 410.2, "2023": 431.0, "2024": 452.9},
 "source": "SABI > Cuentas modelo normal", "date": "2026-10-01", "grade": "A1",
 "request": {"database": "SABI", "table": "...", "filters": "...", "years": "2020-2024", "export_as": "sabi_x_pl.xlsx"}},
{"id": "cust_h", "name": "Customers (reported)", "unit": "m", "role": "driver", "leaf": "customers",
 "values": {"2020": 0.48, "2021": 0.486, "2022": 0.49, "2023": 0.495, "2024": 0.5}, "source": "...", "grade": "B2"},
{"id": "ebitda_h", "name": "EBITDA (reported)", "unit": "EURm", "role": "ebitda", "values": {...}, ...}
```

| Field | Notes |
|---|---|
| role | `revenue` (one), P&L lines `gross_profit`, `ebitda`, `ebit`, `ebt`, `net_profit`, `cost` (the model computes their % of revenue), `driver` (history of a driver-tree leaf: needs `leaf`), `other` (default) |
| leaf | for `driver`: the leaf whose level this is, in the leaf's units. When every leaf of a stream has one, the model back-tests the stream on each past year |
| stream | optional: this series is the reported revenue of one stream (e.g. room revenue); the back-test of that stream is compared with it |
| values | year -> value; years before `first_hist_year` (default: the earliest year given) or after `base_year` are ignored |
| op, args | instead of `values`: a derived series, a live formula on the Actuals sheet. `op` is `prod`, `sum`, `sub` or `div`; `args` are earlier actuals ids, e.g. `{"id": "mpc_h", "name": "Members per club (derived)", "op": "div", "args": ["mem_h", "clubs_h"]}`. Source, date and grade default to the arguments' |

The model derives, per series (Calculations section 0): YoY growth, CAGR, median and quartiles of growth, and
for P&L lines the % of revenue. Outputs show "Base case vs history".

**History-linked inputs.** Any input can take its value from a series instead of being typed:

```json
{"id": "cust0", "name": "Customers, base year", "unit": "m", "history": {"series": "cust_h", "stat": "last"}},
{"id": "cust_g", "name": "Baseline growth in customers", "unit": "%", "low": 0.0, "high": 0.02,
 "history": {"series": "cust_h", "stat": "cagr"}}
```

| stat | Base | Low / high (unless typed) | Type |
|---|---|---|---|
| `last` | last reported value | = base | actual |
| `cagr` | CAGR first-last year | lower / upper quartile of YoY growth | assumption |
| `median_growth` | median YoY growth | lower / upper quartile of YoY growth | assumption |
| `median_margin` | median % of revenue (P&L lines) | lower / upper quartile | assumption |

The Inputs cells become live links (black, not blue), so the base case moves when the Actuals are updated.
Source, date and grade default to the series'. Type `low`/`high` when the history range is narrow or the
outlook differs. A growth input that deliberately differs from history carries `history_note` (the reason;
the defence pack quotes it). An empty `source`/`date`/`grade` also takes the series' value.

**Bridge from history to the baseline** (needed when do-nothing growth sits outside the band):

```json
"baseline_bridge": {"from": "cagr", "items": [
  {"name": "no new clubs in the do-nothing case", "minus": ["hist:clubs_h:cagr"]},
  {"name": "members per club: outlook vs record", "plus": ["mem_g"], "minus": ["hist:mpc_h:cagr"]},
  {"name": "price at CPI, not its record (room share)", "plus": ["infl"], "minus": ["hist:adr_h:cagr"],
   "weight": "hist:room_share_h:last"}]}
```

`from`: `cagr` (default) or `median_growth` of reported revenue. Each step adds its `plus` refs and subtracts its
`minus` refs, times an optional `weight`. A ref is an input or derived id, or `hist:<actuals id>:<cagr |
median_growth | median_margin | last>`. The model writes the steps to Calculations 9b; the start plus the steps must
reach the baseline revenue CAGR within `history_tolerance` (the Checks sheet and checker check 13 both test it).

## inputs (the assumptions register)

```json
{"id": "d1_up", "name": "Peak uplift in visits from loyalty app", "unit": "%", "type": "uplift",
 "low": 0.02, "base": 0.04, "high": 0.06, "polarity": 1,
 "source": "...", "date": "2026-10-01", "link": "database path or URL", "grade": "B3",
 "bench_low": 0.02, "bench_median": 0.04, "bench_high": 0.06, "bench_source": "reference class: n cases ...",
 "justification": "",
 "request": {...}, "fetch": {"source": "IMF WEO", "url": "...", "status": "done"}}
```

| Field | Notes |
|---|---|
| type | `actual` (reported fact), `assumption`, `uplift` (initiative effect: needs a reference class), `cost` (polarity -1 by default; never negative: model savings as a build), `plan` (team decision: launch year, scale such as MWp or markets, rollout years), `constant`, `conversion` |
| bench_low / bench_median / bench_high | reference class of comparable outcomes: 25th percentile / median / 75th percentile. Base near the median; above p75 needs `justification` (check 12) |
| low / base / high | required for assumption, uplift, cost; `actual`/`plan`/`constant` may give base only. Give a `plan` input a range only when the team is still deciding it (it then enters the scenarios and tornado) |
| unit | "%" for rates (stored as decimals: 0.05), currency label for money, free text otherwise |
| polarity | +1 if higher is better for value, -1 if worse; drives which end is "low scenario" |
| request | adds a row to the Data Request Sheet (team pulls from the IE library) |
| fetch | adds an open-source row (Claude fetches); set status "done" once fetched |
| section | optional Inputs grouping; default inferred |

Launch years given as numbers in initiatives become `plan` inputs automatically.

## derived inputs

Arithmetic on inputs, so a derived value moves with its parents in every scenario (never type the result):
```json
"derived": [
  {"id": "spare_cap", "name": "Spare plant capacity", "unit": "m m2", "op": "sub", "args": ["capacity", "vol0"]},
  {"id": "prem_up", "name": "Price uplift from the premium mix shift", "unit": "%", "op": "prod", "args": ["prem_shift", "prem_gap"]}
]
```
`op` is `prod`, `sum`, `sub` (first minus the rest) or `div` (first divided by the rest). Args are input ids or
earlier derived ids. A derived id can be used anywhere an input id can. Derived values live in Calculations
section 1b, not in the register; the register holds their parents.

## leaves and streams (the driver tree)

```json
"leaves": [
  {"id": "customers", "name": "Active customers", "unit": "m", "base": "cust0",
   "growth": "cust_g", "ceiling": "market", "ceiling_growth": "market_g"},
  {"id": "spend", "name": "Spend per visit", "unit": "EUR", "base": "spend0", "inflation": "infl"}
],
"streams": [{"id": "retail", "name": "Retail revenue", "leaves": ["customers", "visits", "spend"], "scale": "to_m"}]
```

- Stream revenue = product of its leaves x optional `scale` (a `conversion` input, e.g. 0.000001 for EUR -> EURm).
- Baseline leaf path: base x (1 + growth) x (1 + inflation) per year. Ceiling path: ceiling x (1 + ceiling_growth).
- A leaf can appear in several streams (e.g. customers in fuel and shop streams).
- `kind`: `price` (money per unit) or `volume`. Default: `price` when the unit holds a currency, else
  `volume`. It feeds the profitability tree (more volume vs higher price).

## pillars and initiatives

```json
{"id": "acq", "name": "Acquire", "initiatives": [
  {"id": "a1", "name": "Corporate wellness partnerships", "mechanism": "Raise the number of members",
   "leaf": "members", "uplift": "a1_up", "launch_year": 2026,
   "ramp": {"type": "bass", "p": "bass_p", "q": "bass_q"},
   "decay": "a1_decay", "opex": "a1_opex", "capex": "a1_capex"}]}
```

| Field | Notes |
|---|---|
| leaf + uplift | uplift = % change of the leaf at full adoption: one id, or a list multiplied together (e.g. `["prem_shift", "prem_gap"]`). The outside-view check applies to inputs typed `uplift` |
| ramp | `bass` (p, q), `logistic` (midpoint, steepness), `linear` (years_to_peak), `step` |
| decay | annual erosion of the effect after launch (applied every year) |
| opex / capex | annual running cost from launch / capex in the launch year. One id, or a list of ids multiplied together (`["pv_mwp", "pv_capex_per_mwp"]` = quantity x unit cost). Never negative |
| opex_growth | annual escalation of opex from launch (e.g. wage inflation) |
| opex_ramp | `true`: opex scales with the initiative's adoption curve (variable running costs that grow with the rollout); default: full opex from launch |
| capex_years | spread capex evenly over n years from launch (phased rollout) |
| asset_life | straight-line depreciation of the capex over n years; gives the tax shield (tax is on EBIT less depreciation) |
| margin | contribution margin for this initiative's revenue (defaults to `finance.margin`): new markets (freight, duties), partner models, premium mix |
| model | `"build"` for an initiative with its own volume x price (next section); default: uplift on a leaf |
| cost-only initiative | omit leaf and uplift, give opex and/or capex (an enabler with no modelled benefit) |

## build initiatives (new revenue streams, cost savings)

For strategies with no baseline to uplift: entering a market, a new channel, or a cost avoided.

```json
{"id": "g1", "name": "Solar PV on roof and land", "model": "build", "kind": "saving",
 "mechanism": "On-site generation replaces grid electricity",
 "volume": {"name": "Solar electricity used on site", "unit": "GWh",
            "drivers": ["pv_mwp", "pv_yield", "pv_selfuse"], "ceiling": "elec_gwh", "ceiling_name": "electricity bought"},
 "price": {"name": "Avoided grid electricity price", "unit": "EURm/GWh", "base": "grid_price", "growth": "grid_price_g"},
 "launch_year": 2026, "ramp": {"type": "linear", "years_to_peak": "pv_build_years"}, "decay": "pv_degradation",
 "capex": ["pv_mwp", "pv_capex_mwp"], "asset_life": "pv_life", "opex": ["pv_mwp", "pv_om_mwp"]}
```
- `kind`: `revenue` (new stream: incremental revenue x its `margin`) or `saving` (straight to EBIT).
- Volume = product of `drivers` x (1 + `volume.growth`)^t x adoption (ramp) x retention (decay), capped by
  `ceiling` (grows at `ceiling_growth`). A penetration row is checked <= 100%.
- `volume.capacity_of: "<leaf id>"`: the build also fits inside what that leaf's ceiling leaves free (for example
  plant capacity shared with domestic sales and other builds). A combined penetration check covers the leaf
  plus every build that shares it.
- A percentage driver that is the initiative's effect (share of energy cut, share of online customers
  who are new) is typed `uplift` and carries a reference class like any uplift.
- Value = volume x price (price grows at `price.growth`) x optional `scale` (conversion ids).
- Units must multiply out to the model currency (e.g. GWh x EURm/GWh). Put any unit conversion in a
  `conversion` input listed in `scale`.

## split_rules

Needed only when two initiatives move the same leaf: `{"visits": "additive"}` (uplifts add;
incremental revenue split pro rata to each initiative's uplift). Without it the build stops.

## finance

```json
{"margin": "margin", "tax": "tax", "rf": "rf", "beta": "beta", "erp": "erp", "crp": "crp",
 "kd": "kd", "debt_weight": "wd", "terminal_growth": "tg", "terminal_switch": "tv_on"}
```
Each value is an input id. `crp` is optional. If `terminal_growth` / `terminal_switch` are omitted the builder
adds `tv_g` (fade -2%/yr, range -5% to +1%) and `tv_on` (1 = include terminal value); set
`meta.terminal_value: 0` for horizon-only valuation.

## sensitivity

```json
{"two_way": [["margin", "bass_p"], ["d1_up", "dv_up"]], "tornado_top": 10, "steps": 5}
```
Without `two_way` the builder pairs the top four tornado inputs. Both inputs of a pair need a range that moves NPV.

## Validation errors you may see

| Message | Meaning |
|---|---|
| `leaf 'x' is moved by [a, b]: add split_rules` | two owners for one leaf: add the rule or merge initiatives |
| `input 'x' is not defined in inputs[]` | a reference to a missing input id |
| `Bass p must be > 0` | p appears in a denominator |
| `two-way pair (a, b): both inputs need a range and must move NPV` | pick tornado inputs |
| WARNING `... checker will fail` | the model builds, but a register row is incomplete: fix before handover |
