# impact-model

A Claude skill for case-competition teams. Paste your pillars and strategies, even vague ones like "solar" or
"enter the US", and Claude builds a sourced, checked Excel impact model and the material to defend it.

- **Leads with the case objectives:** profit, environmental and social KPIs, each for the base case (do
  nothing) and with your strategies, split by pillar and strategy.
- **Cuts the impact through the core consulting frameworks:** triple bottom line, profitability tree,
  Ansoff, 4Ps, value chain, balanced scorecard. Every view adds up to the same totals.
- **Sources every input:** low / base / high, source, date, link and an A-F / 1-6 reliability grade. Each
  uplift is compared with what similar initiatives actually achieved.
- **Builds the base case on 5 years of reported history** and back-tests the driver tree on it.
- **Is checked by an independent checker** (13 checks: live formulas, totals, ceilings, sources, history...).
- **Proposes up to three revisions** to raise impact, on copies, and applies only the ones your team accepts.
- **Writes a defence pack:** the headline, the ten questions a judge is most likely to ask, and the three
  weakest assumptions.

## Install

### claude.ai (web, desktop or mobile)
1. Download **[impact-model.zip](https://github.com/tomianna/impact-model/raw/main/impact-model.zip)**.
2. In claude.ai, open **Settings > Capabilities** and turn on **Code execution and file creation**.
3. Under **Skills**, click **Upload skill** and choose the zip. Make sure it is switched on.

### Claude Code
```
/plugin marketplace add tomianna/impact-model
/plugin install impact-model@impact-model
```
Needs Python 3.9+ with `openpyxl` (`pip install openpyxl`). Optional: `matplotlib` for the bridge charts, and
LibreOffice or `pip install formulas` for a second recalculation engine in the checker.

## Use it

Start a **new chat**, attach the case pack (PDF or Excel) if you have one, and paste:

> We're a case team. Client: **[company]**, **[country]**. The case asks us to **[objective, e.g. grow profit
> and cut environmental impact]**. Our pillars:
> 1. **[Pillar]**: [strategy A], [strategy B]
> 2. **[Pillar]**: [strategy C], [strategy D]
>
> Base year **[2025]**, horizon **5 years**. Build the impact model.

Claude then works through six steps:

| Step | What happens | What you do |
|---|---|---|
| 1. Specify | The case objectives and their KPIs, plus one card per strategy (mechanism, model form, scale). Your strategies are kept exactly as you wrote them | Answer at most 3 questions; confirm or correct the cards |
| 2. Research | 5 years of history and every input sourced and graded; a Data Request Sheet for licensed databases | Pull the files on the Data Request Sheet from your library (optional, improves grades) |
| 3. Build | Version A of the Excel model, exactly as specified | - |
| 4. Check | The independent checker must print `13/13` | - |
| 5. Propose | Up to 3 evidence-based revisions, each modelled on a copy, with its effect on every objective | Read them |
| 6. Decide | Nothing changes until you decide | Reply e.g. "accept R1, reject R2 and R3"; version B is built and checked (`14/14`) |

Ask for the defence pack at any point, e.g. "write the defence pack for version A".

## What you get

| File | Contents |
|---|---|
| `<Company>_Impact_Model_vA.xlsx` | Cover, **Objectives**, **Frameworks**, Inputs (the assumptions register), Actuals, Calculations, Outputs, Bridge, Sensitivity (scenarios, tornado, two-way tables), Checks, Driver tree, Data requests |
| `<Company>_Data_Requests.xlsx` | What to pull from Orbis, SABI, Statista, Euromonitor, Factiva, etc., with the exact table and file name |
| `revisions/Revision_Proposals.xlsx` / `.md` | The proposals, their evidence and their effect on each objective and on NPV |
| `Defence_Pack.md` | Headline on the case objectives, framework views, judge Q&A, weakest assumptions |
| `<Company>_Impact_Model_vB.xlsx` | Only after you accept revisions |

## Built-in protections (Claude Code plugin)

The plugin ships two hooks that run before every file edit and shell command:
- **Version A is locked.** A built `*_vA.xlsx` cannot be edited, overwritten, moved or deleted, by hand or by a
  shell command; change the spec and rebuild instead. Changing `spec_vA.json` or rebuilding version A asks you
  first (fine for new data, never for a revision).
- **Your team decides.** Claude must ask you before it marks a revision accepted or rejected, or runs
  `revisions.py apply`.

On claude.ai, hooks don't run; the skill follows the same rules from its instructions.

## Tips

- **Open the workbook in Excel** (not a preview) so the Data Tables recalculate. Change inputs only on the
  Inputs sheet, or better, ask Claude to change the spec and rebuild. Don't edit the formulas.
- **Licensed databases:** Claude cannot log in to Orbis, SABI or Statista. Those inputs are marked
  PROVISIONAL with a data request. Pull the file, upload it, and ask Claude to update the model.
- **Long runs:** research and building take a while. If Claude stops partway, say "continue".
- **New data:** "Here is the 2025 P&L, update the model" re-runs the build and the checks; it is not a revision.
- **Check an existing model:** upload any impact model `.xlsx` and ask "check this model".

## Updates

- **claude.ai:** download the zip again, delete the old skill under Settings > Skills, and upload the new one.
- **Claude Code:** `/plugin marketplace update impact-model`, or turn on auto-update for this marketplace in
  `/plugin` > Marketplaces.

## Examples inside the skill

`examples/lusitania_spec_vA.json` (ceramics manufacturer, six vague strategies) and
`examples/forza_spec_vA.json` (gym chain, COVID shock in the history). Both companies are made up; the market,
macro and benchmark data come from real sources.
