# check_model.py: what each check means and how to fix it

Run: `python scripts/check_model.py Model.xlsx` (add `--verbose` for info lines, `--json report.json`,
`--trace "Outputs!E7"` to print the source chain of one cell, `--growth-cap 0.30` to change the cap).
A revised model (version B, C...) is checked in revision mode, which adds check 14:
`python scripts/check_model.py Model_vB.xlsx --version-a Model_vA.xlsx --spec-a spec_vA.json --proposals proposals.json`.
Pass = `RESULT: 13/13` for version A, `14/14` for a revised version. Objectives and Frameworks sheets are output
sheets: their totals (check 3), traces (check 10) and the model's own reconciliation rows (check 11) are tested.

Recalculation: LibreOffice headless if installed, else the `formulas` library, else the built-in
engine (`xleval.py`). The built-in engine always re-runs Data Tables by substitution and is
cross-checked against the main engine (differences appear as a warning under check 1).

| # | Check | Fails when | Typical fix |
|---|---|---|---|
| 1 | Recalculated; no formula errors | any #DIV/0!, #REF!, #NAME?, #VALUE! | fix the formula; never wrap errors in IFERROR to hide them |
| 2 | Sensitivity tables | an axis input cell does not feed the output; two rows/columns identical (recalculated or as saved); axis cell depends on its own input cell; axis not ordered; grid is pasted numbers | point the Data Table at the input the model actually uses; never put `=input` in an axis cell; rebuild pasted grids as Data Tables |
| 3 | Totals = sum of parts | a row labelled Total differs from the rows above it, or sums a % row | totals are plain SUMs of the block above; keep % rows out of the block |
| 4 | No hard-coded numbers | a formula on a calc/output/check sheet contains any number other than 0 or 1 | move the number to Inputs with a source (unit conversions too) |
| 5 | Sources and grades | an input row lacks source, date, A-F/1-6 grade, or low/base/high; numbers typed on calc sheets | complete the register row |
| 6 | Penetration and ceilings | a penetration/share row exceeds 100%; a volume row compounds above the cap in a projection year with no MIN against a ceiling (reported history, Actuals mirrors and back-test rows are not tested) | add a `ceiling` to the leaf |
| 7 | Growth cap | a level row (with an actual or base value) grows faster than the cap (default 25%/yr) in a projection year and is not flagged | justify with "FLAG: reason" in the row, or temper the driver |
| 8 | Timeline | repeated periods, uneven steps, or the same year in different columns on different sheets | one timeline, same columns everywhere |
| 9 | Inputs used | a labelled input (e.g. a decay/flatten rate) is never referenced | wire it in or delete it |
| 10 | Trace | an output depends on a typed number that has no sourced, graded register row | move it to Inputs |
| 11 | Checks sheet | no Checks sheet, or a check reads ERROR / FALSE / non-zero errors | fix the model until Total errors = 0 |
| 12 | Outside view | an uplift has no reference class (p25 and p75, with source); an uplift's base is above the reference-class upper quartile (p75) with no written justification; a saving is typed as a lump-sum negative cost | add the reference class; pull the base back to the median or justify it in writing; rebuild the saving as a `build` of kind `saving` (volume x price) |
| 13 | History | no revenue row with 3+ reported years (Actuals sheet); no baseline (do-nothing) revenue row; baseline revenue CAGR outside the band [min(revenue CAGR, p25 YoY growth) - 1 pp, max(CAGR, p75) + 1 pp] unless "Bridge to history:" rows start at the historical CAGR or median and add up to the baseline CAGR within 1 pp (a written justification alone fails). Years flagged 1 on a "Shock year" row of Actuals are left out, with the rebound year after each. Warns: fewer than 5 years; no register input linked to the actuals | pull 5 years into `actuals`; link base-year values and growth to history (`history` on the input); mark `meta.shock_years`; fix the growth, or add `meta.baseline_bridge`. The model's own Checks sheet also back-tests the driver tree on driver history (check 11) |
| 14 | Revision (revised models only) | no proposal has a recorded team decision (`accepted`, by, on); an accepted proposal fails validation (number-only, no new source, above p75); the revised workbook differs from "spec A + accepted proposals" (hand edits or unaccepted proposals); version A was edited after its build | never edit a workbook by hand; apply only what the team accepted with `revisions.py apply`; rebuild version A from its spec |

## What it catches in hand-built models (regression tests)

| Fault | Check |
|---|---|
| a sensitivity axis that changes nothing (input feeds no output) | 2 |
| a volume compounding 30%+/yr with no capacity ceiling | 6, 7 |
| a decay or flatten rate defined but never applied | 9 |
| a "total" that is not the sum of its lines | 3 |
| half-year and annual headers mixed on one timeline | 8 |
| repeated rows or columns pasted into a sensitivity table | 2 |
| uplifts typed with no source | 5, 10, 12 |
| no reported revenue history, or no do-nothing projection | 13 |
| savings typed as lump-sum negative opex | 12 |
| a revision that only raises a number; a hand-edited version B | `revisions.py propose`; 14 |
