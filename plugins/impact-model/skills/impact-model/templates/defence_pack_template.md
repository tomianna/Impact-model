# Defence pack - [Company] [case name]

_Model: `[file].xlsx` (version [A/B], checker [13/13 | 14/14] on [date]). Currency: [unit]. Base year [YYYY], horizon to [YYYY]._

## 1. Impact on the case objectives (base case = do nothing)

| Objective | KPI | Base case | With strategies | Change | Change % | Better? |
|---|---|---:|---:|---:|---:|---|
| Profit | EBITDA | | | | | |
| Planet | [e.g. CO2e, tCO2e] | | | | | |
| People | [e.g. jobs, FTE] | | | | | |

How each strategy moves each objective (table from the Objectives sheet), then one line per framework view
(triple bottom line, profitability tree, Ansoff, 4Ps, value chain) from the Frameworks sheet. Explain any KPI
that moves the wrong way.

## 1b. Revenue and EBIT bridge

![Revenue bridge](bridge.png)

| Step | [YYYY] revenue | Cell |
|---|---:|---|
| Baseline (do nothing) | [x] | `Bridge!C5` |
| [Pillar]: [initiative] | +[x] | `Bridge!C6` |
| ... | | |
| Target (with strategy) | [x] | `Bridge![C]` |

**Headline (one sentence, numbers from the model):** the strategy adds [x] of revenue in [YYYY]
(+[x]% vs do-nothing); NPV [x] at a [x]% WACC; IRR [x]; payback [x] years.
Low / base / high NPV: [x] / [x] / [x].

**Value bridge (EBIT [YYYY] by initiative):** ![EBIT bridge](ebit_bridge.png): cost savings and running
costs show here, not in revenue. NPV by initiative (adds up to the total): [table from Outputs].

**If revisions were adopted (version B):** ![Before / after](before_after_bridge.png): NPV by initiative,
A vs B, with the accepted change and the source behind each gain (`Before_After_vB.md`). Say that the team
chose it and why.

**Lever check:** against [3-5 peers] the largest gap to the peer median is [line] ([x]); the strategy
targets [line] - [consistent / why we still chose it].

## 2. Ten questions a judge is most likely to ask

Each answer: 30 seconds out loud (about 70 words), one number, one source, one cell.

**Q1. Where does your baseline come from?**
[Reported accounts, source, grade; driver tree reproduces it within x%.]
_Source cell: `Actuals!...`_

**Q2. Why does your baseline grow x% a year when the company grew y%?**
[Company record: CAGR, interquartile YoY growth over 5 years; baseline inside the band (or the written reason);
inputs linked to history; back-test gap; EBITDA margin history vs the contribution margin used.]
_Source cell: `Calculations` sections 0 and 9b; `Checks` 'Base case vs history'_

**Q3. Why is [largest uplift] credible?**
[Inside the reference class x-y% from n comparable cases / meta-analysis; NPV if at the low end.]
_Source cell: `Inputs!F..`_

**Q4. How fast does it ramp - aren't you assuming day-one impact?**
[S-curve, p and q from analogous launch, % of peak reached by year n.]

**Q5. Aren't you double counting between pillars?**
[One owner per leaf, log-share attribution, live checks.]

**Q6. What if it works less well than you hope?**
[Low scenario NPV, which input matters most.]

**Q7. How did you get your discount rate?**
[rf, beta, ERP, CRP, kd, D/V - each sourced.]

**Q8. What does it cost and when does it pay back?**

**Q9. Is the market big enough?**
[Ceiling, penetration reached, source.]

**Q10. Why this lever and not another?**
[Peer benchmark result.]

## 3. Three weakest assumptions

| # | Assumption | Base | Grade | NPV if at the pessimistic end | Break-even | How we will strengthen it |
|---|---|---|---|---|---|---|
| 1 | | | | | | [database / request row] |
| 2 | | | | | | |
| 3 | | | | | | |

For each: one sentence on what changes if it is wrong, and what evidence would move its grade up.
