# impact-model

A Claude skill that turns a case team's pillars and strategies into a checked, sourced Excel impact model. It
leads with the case objectives (profit, environmental and social KPIs vs the base case), cuts the impact
through the core consulting frameworks, proposes evidence-based revisions, and writes a judge-ready defence
pack.

## claude.ai (web, desktop, mobile)
Download [`impact-model.zip`](impact-model.zip). In claude.ai, turn on **Settings > Capabilities > Code execution
and file creation**, then upload the zip under **Skills**. Start a new chat, attach the case pack and paste the
case objective plus your pillars and strategies, ending with "Build the impact model". Re-upload the zip when it
is updated (delete the old skill first).

## Claude Code (updates with one command)
```
/plugin marketplace add tomianna/impact-model
/plugin install impact-model@impact-model
```
Update: `/plugin marketplace update impact-model`, or turn on auto-update for the marketplace in `/plugin`.
Python 3.9+ with `openpyxl` is required.

## What you get
`<Company>_Impact_Model_vA.xlsx` (Cover, Objectives, Frameworks, Inputs, Actuals, Calculations, Outputs, Bridge,
Sensitivity, Checks), a Data Request Sheet for licensed databases, up to three revision proposals (applied only
when you accept them), and `Defence_Pack.md`. An independent checker must pass every check before handover.
