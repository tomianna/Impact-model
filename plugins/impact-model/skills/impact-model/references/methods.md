# Methods notes

## Contents
- Case objectives and the consulting frameworks
- History and the base case
- FAST Standard as applied in the generated workbook
- Ramp-up: Bass diffusion and logistic S-curves
- Decay (the old "flatten rate")
- Ceilings and penetration
- Reference-class forecasting (outside view)
- Starting benchmarks (verify before citing)
- Attribution: splitting incremental revenue without double counting
- WACC from sourced inputs
- Scenarios, tornado, two-way tables
- NPV, IRR, payback conventions

## Case objectives and the consulting frameworks

The model answers the case question, so its headline is the case objectives (profitability, growth,
environmental, social), each in the base case (do nothing) and with the strategies. NPV is the valuation that
supports it. The impact is then cut through the frameworks consultants and judges use. Each view splits the
same totals, so they reconcile (model Checks), and the team can say which lever, growth vector or activity
the impact comes from.

| # | Framework | What it answers | In the model |
|---|---|---|---|
| 1 | Issue tree / MECE (objective -> pillars -> strategies -> drivers) | Does every strategy serve an objective, with nothing counted twice? | Objectives sheet; driver tree; one owner per leaf (`split_rules`) |
| 2 | Profitability tree (profit = volume x price x margin - costs) | Where does the extra profit come from? | Frameworks 2: volume, price, new streams, savings, added costs |
| 3 | Triple bottom line (Elkington: profit, planet, people) | How does each pillar move each objective? | Frameworks 1 and the Objectives sheet |
| 4 | Ansoff matrix | Which growth vector, and how risky? | Frameworks 3 (`frameworks.ansoff`) |
| 5 | Marketing mix (4Ps) | Which commercial lever? | Frameworks 4 (`frameworks.four_p`) |
| 6 | Value chain (Porter) | Which activity changes? | Frameworks 5 (`frameworks.value_chain`) |
| 7 | Balanced scorecard (Kaplan & Norton; outcomes widened to the TBL) | Outcomes, customers, processes, enablers | Frameworks 6 |
| 8 | 3Cs (company, customers, competitors) | Can we win? | Customers: the leaves moved; competitors: each uplift's reference class; company: capex, opex and grades |
| 9 | Porter's five forces | Will the effect last? | Decay (competitors copy, buyers bargain) and the terminal fade; say which force on the card |
| 10 | McKinsey 7S | Can the company deliver it? | Ramp-up, opex for skills and systems; the defence pack's weakest assumptions |

Frameworks 1-7 are computed; 8-10 are checks the team argues from the model (cards, reference classes, decay
inputs). Sources: case-interview guides agree on profitability, market entry, 3Cs, 4Ps, five forces, Ansoff,
value chain, BCG and 7S as the core set (e.g. roadtooffer.com, consultport.com; profitability is ~40% of
cases); TBL from Elkington (1994), with Kaplan's TBL scorecard. Do not force a framework that does not fit
the case: leave a tag out and the view shows it as not tagged.

Environmental and social KPIs follow the same three rules as money: each is a modelled quantity x a sourced,
graded factor (low/base/high), so it moves with the scenarios. State the boundary (e.g. "CO2e from gas for hot
water, scope 1") and the factor's vintage; a judge will ask.

## History and the base case

The do-nothing baseline is the yardstick for every uplift, so it is built on the company's own record:
1. **Collect** 5 years of reported revenue, the P&L lines and the history of each driver leaf (`actuals`).
2. **Process** (Calculations section 0, live formulas): YoY growth, CAGR, median and quartiles of growth, and
   each P&L line's % of revenue. Per-unit series (members per club, spend per visit, revenue per store) are
   derived on the Actuals sheet (`op`/`args`), because like-for-like growth is usually the real driver.
3. **Anchor**: base-year leaf values link to the last reported value (`stat: last`); the base year must
   reproduce reported revenue within 2% (calibration).
4. **Shocks**: list shock years in `meta.shock_years` (e.g. COVID closures 2020-21, an energy spike) with a
   `shock_note`. Those years and the rebound year after each are left out of the medians, quartiles and the
   band, and the CAGR runs between the first and last non-shock years. For a trough and rebound, the
   median of all YoY changes is *more* distorted than the CAGR (the rebound years dominate the sample), so
   do not switch to the median: exclude the shock years instead.
5. **Grow**: do-nothing growth comes from history (`cagr` or `median_growth`) or from an outlook with a
   `history_note`.
6. **Test**: baseline revenue CAGR must sit inside [min(CAGR, p25 YoY) - 1 pp, max(CAGR, p75 YoY) + 1 pp]
   (the interquartile band, like the reference class for uplifts). If it does not, a written reason is
   not enough: **bridge** it (`meta.baseline_bridge`). Start from the historical CAGR (or median), then add
   one step per difference, each built from inputs or history statistics, e.g. "no new stores in the
   do-nothing case" = minus the store-count CAGR, "price at CPI instead of its record" = CPI minus the
   price CAGR (times the stream's revenue share when it moves one stream only). The steps must add up to the
   baseline CAGR within 1 pp. The checker recomputes the start and the sum itself. With driver history for
   every leaf, the driver tree is back-tested on each past year (gap <= 5%). A driver series derived from the
   revenue it is tested against is left out, since it would match by construction.
7. **Margins**: compare the contribution margin on incremental revenue with the company's EBITDA margin
   history. It is normally higher (fixed costs do not grow with the extra sales); say why in the defence pack.

Watch nominal vs real: reported revenue includes inflation, so a baseline without price inflation will sit
below the band. Add CPI to the price leaf, or bridge it (a "real vs nominal" step = minus CPI).

## FAST Standard as applied in the generated workbook

FAST = Flexible, Appropriate, Structured, Transparent (FAST Standard Organisation, fast-standard.org).
The rules the builder enforces and the checker verifies:

| FAST rule | In the workbook | Checker |
|---|---|---|
| Separate inputs, calculations, outputs, checks | Inputs + Actuals / Calculations / Outputs, Bridge, Sensitivity / Checks | role of each sheet |
| One consistent timeline | Columns G onward = same year on every time-based sheet; rows 3-4 hold the period labels | timeline |
| No hard-coded numbers in formulas | Only 0 and 1 may appear inside formulas; everything else is an Inputs cell | hardcodes |
| Inputs entered once | Calculations section 1 links every input once ("live inputs") | trace, unused |
| Flow left-to-right, top-to-bottom | Each calculation block reads only from blocks above it | - |
| Live formulas | Sensitivities are Excel Data Tables, not pasted grids | sensitivity |
| Checks in the model | Checks sheet, "Total errors" must be 0 | checks_sheet |

Formatting: blue on yellow = typed input; orange = Sensitivity control; green = headline output.

## Ramp-up: Bass diffusion and logistic S-curves

Bass cumulative adoption share after t years from launch (t = year - launch + 1):

    F(t) = (1 - exp(-(p+q) t)) / (1 + (q/p) exp(-(p+q) t))

- p = innovation (external influence: advertising, launch push); q = imitation (word of mouth).
- Effect in year t = peak uplift x F(t) x retention(t).
- Peak timing: t* = ln(q/p) / (p+q). With p = 0.03, q = 0.38: t* = 6.2 years, F(6) = 44%.
- **Choose p, q from analogous launches**, not from habit. The durable-goods meta-analysis averages
  (p = 0.03, q = 0.38) describe market-wide diffusion of new products over a decade. An initiative rolled out
  to an existing customer base usually adopts faster: calibrate p, q to a comparable launch (Factiva: members
  of a competitor's app after 1, 2, 3 years) and record that in Source. If no analogue exists, keep the
  meta-analysis values, widen the range, and grade `C4`.
- Logistic alternative: F(t) = 1 / (1 + exp(-k (t - m))), m = midpoint year, k = steepness. Use when the
  team has a target "50% adoption by year m".
- Linear: F(t) = min(1, t / n). Step: F(t) = 1 from launch (only for switch-on effects like a price change).

## Decay (the old "flatten rate")

Effects erode as competitors copy and novelty wears off: retention(t) = (1 - d)^(t - 1).
Every initiative with a defined decay input has it applied in its uplift row; the checker fails any
decay/flatten rate that is defined but unused (a common fault in hand-built models).
Old-model conversion: "growth g in year 1, flattening x f per year" is a cumulative effect of g / (1 - f)
(e.g. 1.5% / (1 - 0.7) = 5%). Use that as the peak uplift, then apply the S-curve and decay.

## Ceilings and penetration

- With strategy = MIN(baseline x (1 + uplifts), ceiling). The ceiling grows at its own sourced rate.
- Penetration row = leaf / ceiling, checked <= 100% in the model and by the checker.
- When the ceiling binds, the effective uplift (with / baseline) is below the planned uplift: the
  attribution uses the effective factor, so capped revenue is never claimed.

## Reference-class forecasting (outside view)

Kahneman & Lovallo (1993, Management Science, "Timid choices and bold forecasts") and Flyvbjerg
(2006, Project Management Journal, "From Nobel Prize to project management: getting risks right"):
planners overestimate benefits and underestimate time and cost because they reason from the inside.
The correction, as an action for every uplift:

1. Define the reference class: comparable initiatives (same mechanism, similar industry and market maturity).
2. Collect outcomes: 5+ cases from Factiva / ABI/INFORM, or effect sizes from a meta-analysis (Scopus).
3. Record the distribution: `bench_low` = 25th percentile, `bench_median`, `bench_high` = 75th percentile
   (upper quartile), with `bench_source` (n cases, what each measured, converted to the leaf's units).
4. Place the team's uplift: the base sits near the median; anything up to p75 is acceptable; above p75 needs
   a written `justification` with evidence of why this case beats three in four comparables. Otherwise pull
   it back (check 12 enforces this).
5. Time and cost: delay launches and inflate costs to the reference class too (Flyvbjerg's uplift
   on budgets); use the high end of cost ranges in the low scenario (polarity -1 does this).
6. Revisions (step 5) use the same class: a revision may move the strategy toward what the better
   comparables did (rollout, operating model, channel), never the number past p75 without justification.

## Starting benchmarks (verify before citing)

These are starting points for the reference class, not substitutes for it. Check the primary source in
Scopus before quoting a number to a panel.

| Effect | Benchmark | Source | Suggested grade |
|---|---|---|---|
| Bass p, q (consumer durables) | mean p = 0.03, q = 0.38 | Sultan, Farley & Lehmann (1990), JMR 27(1) | B2 |
| Price elasticity of demand (brands) | mean about -2.6 | Bijmolt, van Heerde & Pieters (2005), JMR 42(2) | B2 |
| Short-term advertising elasticity | mean about 0.12 | Sethuraman, Tellis & Briesch (2011), JMR 48(3) | B2 |
| Loyalty reward programme sales lift | about 6% during programme | Taylor & Neslin (2005), J. Retailing 81(4) (secondary summary) | C3 |
| Frequency programmes raise visit frequency, strongest for light buyers | direction, not size | Liu (2007), J. Marketing 71(4); Lal & Bell (2003) | C3 |
| Price vs profit leverage | 1% price ~ 11% operating profit (1990s S&P sample) | Marn & Rosiello (1992), HBR | C3 |

## Attribution: splitting incremental revenue without double counting

Within a stream R = L1 x L2 x ... x Ln. With strategy each leaf moves by factor f_i = with_i / baseline_i.
Incremental revenue dR is split by log shares:

    share_i = ln(f_i) / sum_j ln(f_j)          (shares add to exactly 100%)

Then each leaf's share is split across its initiatives pro rata to their uplifts (split rule "additive").
This makes a common hand-built method exact and live. Checks prove initiatives and pillars add back to the
stream total in every year.

## WACC from sourced inputs

    ke   = rf + beta x ERP (+ CRP for emerging markets)
    kd   = pre-tax cost of debt x (1 - tax)
    WACC = (1 - D/V) x ke + D/V x kd

- rf in local currency = 10y government yield - sovereign default spread (Damodaran).
- beta: peer median, relevered if the target's leverage differs: beta_L = beta_U x (1 + (1 - t) D/E).
- Each component is its own Inputs row with source, date and grade; Q7 of the defence pack reads them.

## Scenarios, tornado, two-way tables

- Every non-exempt input has low / base / high. Polarity (+1: higher is better, -1: higher is worse, e.g.
  costs) maps low/high to pessimistic/optimistic. The scenario switch moves all of them at once.
- Tornado: each input to its low and its high with all others at base (two 1-way Data Tables on a
  selector); sorted live by swing with RANK + COUNTIF; inputs whose range does not move NPV are left out.
- Two-way tables: rows and columns are the input's own low-high range in 5 steps, fed through override
  cells. The checker proves both axes feed NPV and no row or column repeats.

## NPV, IRR, payback conventions

- Incremental cash flows only (strategy minus do-nothing), end-of-year discounting, t = year - base year.
- Terminal value: Gordon growth on a normalised final-year incremental FCF (final-year growth capex added
  back, replacement capex = depreciation taken off, so a capex tranche in the last year is not capitalised
  forever). The build warns when capex runs past the horizon (not counted) and when capex has no
  `asset_life` (then no replacement capex is assumed), included by default with a sourced fade
  (`tv_g`: -2%/yr, range -5% to +1%) and a switch (`tv_on`). Horizon-only NPV is always reported next to it,
  and the defence pack states the terminal-value share. Without it, initiatives with lasting effects
  (new stores, maturing apps) look value-destroying over a 5-year horizon.
- If initiatives are still ramping in the final year (adoption well below 100%), the terminal value
  capitalises a level that is still growing (understates) or one never reached (if the decay is high). The
  defence pack flags final-year adoption below 50%: lengthen the horizon or say so.
- Capex with an `asset_life` is depreciated straight-line. Tax is charged on EBIT less depreciation, so FCF =
  EBIT x (1 - t) + t x depreciation - capex. Without `asset_life` there is no tax shield (conservative).
- IRR uses horizon cash flows only. Some engines (LibreOffice) return n/a for strongly negative IRRs; Excel
  and the built-in engine solve them.
- IRR shows "n/a" when cash flows never change sign. Payback = years after the base year until cumulative
  FCF turns non-negative; "beyond horizon" if it never does; "no investment phase" when cash flow is never
  negative (then quote the horizon-only NPV, not payback, to the judges).
- Tax on negative incremental EBIT is a credit (assumes group profits absorb it): say so if asked.
