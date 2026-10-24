# Specify: from a vague strategy label to a modellable specification

## Contents
- The rule: keep the intent
- The specification card (one per strategy)
- Choosing the model form: uplift, build or cost-only
- Patterns for common vague labels
- Sizing the scale without inventing it
- The three questions
- Worked example: Lusitânia (6 vague strategies)

## The rule: keep the intent

Step 1 makes each strategy concrete enough to model. It never changes what the team chose.
- Keep the team's verb and object. "Solar" stays solar: not "energy efficiency", and not "buy green power".
- If a label could mean two things, take the reading closest to the slide and say so on the card.
- You may think a strategy is weak. Write that in `notes_for_step5` and say nothing more until step 5.
  Replacing, dropping or merging a strategy at steps 1-4 is a failure, even if it looks better.
- The team confirms the cards (one message) before research starts.

## The specification card (one per strategy)

| Field | Content |
|---|---|
| Strategy (team's words) | "Solar" |
| Pillar | Go green |
| Mechanism (action -> effect -> driver) | Install PV on roof and land -> generate electricity on site -> buy fewer MWh from the grid |
| Model form | `build`, kind `saving` (or `uplift` on a leaf, or cost-only) |
| Driver-tree link | Energy cost = electricity bought x price; solar reduces electricity bought |
| Scale (team decision, type `plan`) | 12 MWp (the brief's maximum), launch 2026 |
| Drivers to source | yield MWh/MWp (PVGIS), share used on site, grid price (Eurostat), capex/MWp (IRENA), O&M, degradation |
| Ceiling | on-site use <= electricity bought (95 GWh) |
| Ramp | step or a 1-2 year linear build (rollout plan) |
| Costs | capex = MWp x EUR/MWp, `asset_life` 25; opex = MWp x O&M |
| What would make it fail | low self-consumption, falling grid prices |
| Case objectives it moves | e.g. profit (cost saving), planet (gas avoided x emission factor) |
| Framework tags | Ansoff vector, 4P, value-chain activity (`frameworks` on the initiative) |
| notes_for_step5 | (ideas to raise impact: do not act on them yet) |

## Choosing the model form

| The strategy... | Form | Spec |
|---|---|---|
| moves an existing revenue driver (more customers, visits, price, mix) | uplift | `leaf` + `uplift` (+ reference class) |
| creates revenue where today there is none (new market, new channel, new product line sold to new buyers) | build, kind revenue | `volume.drivers` (market x share), `volume.ceiling` (market size or spare capacity), `price`, own `margin` |
| avoids a cost (energy, waste, procurement, automation) | build, kind saving | volume avoided x price of what is avoided |
| is pure spending with no modelled benefit (an enabler) | cost-only | `opex` / `capex`, positive numbers |

Two strategies that move the same leaf need `split_rules: {"leaf": "additive"}`.

## Patterns for common vague labels

| Label | Usual mechanism | Form | Scale (plan) | Key drivers and sources |
|---|---|---|---|---|
| Solar / renewables | on-site generation replaces grid power | build saving | MWp | PVGIS yield, Eurostat nrg_pc_205 price, IRENA capex |
| Cut energy waste / efficiency | less energy per unit produced | build saving | share of plant retrofitted | energy use (case), % cut (audits, BREF, peer retrofits), price |
| Electrify / heat pumps | gas replaced by electricity | build saving (gas avoided) + opex (electricity added) | units installed | COP, gas and power prices, carbon price |
| Enter market X | sell to new buyers in X | build revenue | target share by year 5 | market size (Statista/Euromonitor/trade body), local price, freight + duties -> margin |
| Sell online / D2C / e-commerce | reach new buyers, or capture the intermediary margin | build revenue (new buyers) or uplift on price (margin capture) + cannibalisation | share of sales online | online share of the category (Eurostat isoc_ec_eseln2, Statista), channel margin |
| Premium line / go premium | mix shift to higher price and margin | uplift on price (mix) with own margin, or build revenue | premium share of volume | price premium (case pack), segment size |
| Become the X's choice (architects, doctors...) | specification / recommendation wins projects | uplift on volume | number of reps, accounts | win rates, specified-project share (trade press, Factiva) |
| Loyalty / app | more visits per customer | uplift on frequency | members targeted | loyalty meta-analyses (Scopus) |
| Partnerships | access to a partner's customers | uplift on customers, or build revenue | partners signed | partner reach, conversion |
| Digitalise / AI / automation | fewer hours or errors | build saving | processes in scope | hours x cost (case), % saved (studies) |
| Sustainability label / ESG | price premium or retention | uplift on price or churn | share of range certified | willingness-to-pay studies |

## Sizing the scale without inventing it

Order of preference: (1) the team's slide or the brief (12 MWp, 3 markets); (2) the client's own target in its
annual report; (3) the reference class of what comparable firms did (median); (4) a physical limit
(roof area, capacity, licences). Record it as a `plan` input with that source. Every scale gets a
range through the drivers that depend on it, never a hidden typed result.

## The three questions

Ask at most three, in one message, only where the answer changes the specification:
1. Objective and horizon (profit, revenue, cash; base year = last reported).
2. The scale of the vaguest strategy, if the brief gives none ("How many markets, by when?").
3. The one mechanism you cannot infer ("Online: new installers, or the ones distributors serve today?").
Default everything else and state the defaults on the cards.

## Worked example: Lusitânia (6 vague strategies)

| Strategy | Mechanism | Form | Scale |
|---|---|---|---|
| Solar | PV replaces grid electricity | build saving | 12 MWp (brief) |
| Cut energy waste | kiln heat recovery and burner upgrades cut gas per m2 | build saving | % of gas cut at full rollout, retrofit over 3 years |
| Enter the US | sell tiles to US distributors | build revenue, own margin, ceiling = spare capacity | target US share |
| Sell online to installers | B2B portal reaches installers that distributors don't serve | build revenue, own margin | online volume |
| Premium design line | mix shift to large-format/design at 1.6x price | uplift on price, own margin | premium share of volume |
| Become the architects' choice | specification sales win projects | uplift on volume | specified share |

The resulting spec: `examples/lusitania_spec_vA.json`.

## Patterns for things the format does not do directly

- **Cannibalisation** (online takes sales from distributors): put a net-new share in the build's volume
  drivers (`["onl_customers", "onl_new_share"]`, `onl_new_share` typed `uplift` with a reference class).
- **A scale decision that sets an uplift** (premium share 8% -> 15% sets the price uplift): make the uplift
  a `derived` input (`prod` of the shift and the price premium) or list both ids in `uplift`, so they move
  together in scenarios.
- **Several initiatives on one capacity** (domestic volume + US + online on one plant): give the leaf its
  capacity `ceiling` and each build `"capacity_of": "<leaf>"`. The builds share what the leaf leaves free,
  and a combined penetration check covers them all.
- **Arithmetic on inputs** (spare capacity = capacity - volume; gas price + CO2 x ETS): `derived` inputs, not
  typed results.
- **Time-varying tax rates:** not supported. Use the average over the horizon with a range.
- **A segment with its own economics** (winter season, a new customer type): a `build` with its own volume,
  price and margin, rather than an uplift on a shared leaf.
- **Interactions between initiatives** (direct booking applies to the extra nights that loyalty creates):
  not modelled. Each build's drivers are inputs. Say so in the defence pack, and size the overlap in
  `notes_for_step5`.
