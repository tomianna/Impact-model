# Grading scale (intelligence-analyst / Admiralty code)

Every input gets two characters: source reliability (letter) + information credibility (number).
Write it in the `Grade` column of Inputs and Actuals, e.g. `B2`. The checker rejects anything
that is not `[A-F][1-6]`.

## Source reliability (who says it)

| Grade | Meaning | Typical case-competition sources |
|---|---|---|
| A | Completely reliable | Audited statutory accounts (Orbis/SABI, company filings), national statistics office, Eurostat, IMF/OECD/World Bank, tax law |
| B | Usually reliable | Damodaran datasets, Bloomberg/WRDS data, Statista Market Insights, Euromonitor, Oxford Economics, BMI, peer-reviewed meta-analyses |
| C | Fairly reliable | Single peer-reviewed study, industry association reports, the case pack itself, management interviews, quality press (FT, Expansión) |
| D | Not usually reliable | Team estimate built from analogues, consultant blog posts, vendor case studies |
| E | Unreliable | Anonymous web pages, marketing claims, LLM recollection without a source |
| F | Reliability cannot be judged | Source unknown (inherited number, "from last year's deck") |

## Information credibility (how well the number itself holds up)

| Grade | Meaning | Test |
|---|---|---|
| 1 | Confirmed by other sources | Two independent sources agree within the low-high range |
| 2 | Probably true | Consistent with other evidence, one source |
| 3 | Possibly true | Plausible, not contradicted, not corroborated |
| 4 | Doubtful | Sits outside the reference-class range or relies on optimistic logic |
| 5 | Improbable | Contradicted by other evidence |
| 6 | Truth cannot be judged | No way to check yet |

## Rules of thumb

- Reported base-year financials: `A1`-`A2`. A press summary of the same numbers is `C2` - replace it.
- A competition case pack: the client's own reported figures in it are `B2` (reliable source, not
  independently confirmed); the case writer's estimates or round numbers are `C3`.
- A number inherited from an old model with no source is `F6` until traced. Trace it or replace it.
- Uplifts: rarely better than `B3`. If the uplift is above the reference-class high, credibility is `4`
  unless the written justification cites evidence that the case is genuinely different.
- Derived numbers (e.g. spend per visit = revenue / (customers x visits)) take the weakest grade of
  their parents, and say "Derived: ..." in Source.
- Grades drive the defence pack: the three weakest assumptions are ranked by NPV swing x weakness,
  where weakness = (letter index + number index) / 10 + 0.1 (A1 = 0.1, F6 = 1.1).
- Upgrading a grade is the fastest way to strengthen the defence: pull the licensed source named in
  the Data Request Sheet, then update Source, Date, Link and Grade on the same row.
