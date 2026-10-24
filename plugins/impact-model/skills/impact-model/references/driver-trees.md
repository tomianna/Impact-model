# Driver-tree patterns by business type

## Contents
- Rules for every tree
- How a tree maps to the spec
- Retail, fuel, restaurants, grocery
- Subscription, SaaS, gyms, telecom
- Sports clubs, venues, events, museums
- Tourism, hotels, airlines
- B2B, tenders, professional services
- Manufacturing, consumer goods
- Marketplaces and platforms
- Cost and profitability programmes
- From revenue to EBIT and free cash flow
- Double-counting traps

## Rules for every tree

1. Decompose the target metric until each leaf is something you can **measure from a named source** (see source-map.md). "Brand strength" is not a leaf; "repeat-purchase rate (loyalty data)" is.
2. Revenue = product of leaves within a stream; streams add. Keep leaves multiplicative inside a stream so uplifts compound correctly.
3. Map each initiative to **one leaf** it moves. If two initiatives move the same leaf, write a split rule (`"additive"`: uplifts add, incremental revenue split pro rata to each uplift) - or merge them.
4. Every volume leaf that can saturate gets a **ceiling** (market size, capacity, population). Penetration = leaf / ceiling and must stay <= 100%.
5. Restate every initiative as *action -> mechanism -> leaf*: "Launch tiered rewards -> members visit more often -> visits per customer".

## How a tree maps to the spec

```
stream "Retail revenue" = customers x visits x spend            -> streams[].leaves
leaf customers: base 1.40m, growth 0%, ceiling 8.45m cars        -> leaves[] (+ inputs)
initiative "Acquire new drivers": +5% peak on customers, Bass    -> pillars[].initiatives[]
```

Uplifts are % changes on the leaf at full adoption. Absolute effects ("+1 visit") are converted: +1 / 22 = +4.5%.

## Retail, fuel, restaurants, grocery

```
Revenue = customers x visits per customer x spend per visit
          spend per visit = items per basket x price per item (split if a pricing initiative exists)
Ceiling: customers <= addressable population (drivers, households in catchment)
```
Typical initiatives: loyalty (visits), basket/upsell (spend), new stores (customers via catchment), price architecture (price per item).
Peer lever check: gross margin %, staff cost % of sales, sales per m2 (Orbis/SABI, Euromonitor).

## Subscription, SaaS, gyms, telecom

```
Revenue = average members x ARPU x 12
          members_t = members_{t-1} x (1 - churn) + gross adds
Simplified leaf form for the builder: members (level) x ARPU (level)
Ceiling: members <= addressable users x achievable share
```
Model churn reduction as an uplift on the members leaf: retention gain of d points raises steady-state members by roughly d / churn. Show the conversion in the Source column.

## Sports clubs, venues, events, museums

```
Matchday revenue = games x attendance per game x ticket yield (+ F&B spend per head)
Commercial revenue = sponsorship + merchandising + media (each a level leaf)
Ceiling: attendance per game <= stadium capacity x max occupancy (e.g. 95%)
```
A common failure: attendance compounding at 32%/yr with no capacity ceiling. Always add `ceiling`.

## Tourism, hotels, airlines

```
Revenue = capacity (rooms x nights | seats x flights) x occupancy (load factor) x average rate
Ceiling: occupancy <= 100% (practically 90-95%)
```
Occupancy is a share leaf: the builder's penetration check covers it when you set capacity as the ceiling of the volume leaf.

## B2B, tenders, professional services

```
Revenue = bids submitted x win rate x average contract value
Ceiling: bids <= addressable tenders in market (public procurement portals)
```
A common failure: mixing half-year and annual columns and subtracted base growth from a "total": keep one annual timeline and let totals be plain sums.

## Manufacturing, consumer goods

```
Revenue = volume sold (units, m2, tonnes) x net price per unit
          or distribution points x rate of sale per point x net price
EBIT    = revenue x contribution margin - fixed costs
Ceiling: volume <= plant capacity (or the universe of outlets: Euromonitor / Nielsen)
Energy cost = electricity bought (MWh) x grid price + gas used (MWh) x gas price (+ CO2 x ETS price)
```
- New export market: a `build` of kind `revenue`. Volume = target market size x target share, capped by
  market size or spare capacity (capacity - baseline volume). Price = local net price. Give it its own
  `margin` (freight, duties, distributor margin).
- Mix shift to premium: uplift on net price per unit, with its own `margin` when premium carries a higher margin.
  Or a `build` of kind `revenue` for the premium line when it is sold to new customers.

## Marketplaces and platforms

```
Revenue = active buyers x orders per buyer x average order value x take rate
Ceiling: active buyers <= online population x category penetration
```

## Cost and profitability programmes

```
EBIT = revenue - variable cost per unit x volume - fixed costs
Saving = volume avoided (MWh, kg, hours, km) x price of what is avoided x S-curve x retention
```
Never type a saving as a lump-sum negative cost: the checker fails it (check 12) because its drivers would
be invisible. Model it as a `build` of kind `saving`:
- **Solar / own generation:** volume = installed MWp (`plan`) x yield (MWh per MWp per year, PVGIS) x share
  used on site, capped by electricity bought; price = avoided grid price (Eurostat nrg_pc_205);
  capex = MWp x capex per MWp (IRENA), `asset_life` 25; opex = MWp x O&M per MWp per year; decay = panel degradation.
- **Energy efficiency:** volume = energy use (MWh, case pack) x % cut at full rollout (reference class:
  audits, BREF, comparable retrofits); price = avoided energy price; capex phased with `capex_years`.
- **Procurement / waste / automation:** volume = spend or hours in scope x % reduction; price = 1 (spend) or
  the hourly cost.
Margin-mix initiatives: an initiative-level `margin` input (not a change to the global margin).

## From revenue to EBIT and free cash flow

The builder carries every initiative through:

```
Incremental EBIT = incremental revenue x contribution margin - initiative opex
Incremental FCF  = incremental EBIT x (1 - tax) - initiative capex
NPV = sum of FCF x 1/(1 + WACC)^t  (+ optional terminal value switch)
```
Contribution margin = (revenue - variable costs) / revenue from the base-year P&L, not the EBIT margin: an extra euro of sales carries variable cost, not fixed cost.

## Double-counting traps

- Two initiatives both claim "more customers": one owner or a split rule.
- New-store revenue counted in both "customers" and "visits": new stores move the customers leaf only.
- Price increase and "premium range" both on spend per visit: split spend into items x price, or split rule.
- Market growth counted as strategy impact: baseline growth goes in the leaf's `growth`, not in uplifts.
- Inflation inside an uplift: keep inflation on the price leaf only.
