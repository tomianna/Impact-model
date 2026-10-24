# Source map: driver type -> best source

## Contents
- How to use this map
- Company financials and peers
- Market size, penetration, prices, volumes
- Macro: inflation, GDP, FX, population
- Cost of capital (WACC inputs)
- Initiative effect sizes (reference classes)
- Startups and new business models
- Open sources Claude can fetch directly
- Data Request Sheet: how to write a good row

## How to use this map

1. For every leaf of the driver tree and every assumption, pick the first source in the row that the team can access.
2. Licensed IE library databases: Claude cannot log in. Write a Data Request Sheet row (`request` in the spec) with database, exact table/report, filters, years and the export file name. The team pulls it and uploads the file; Claude then updates Source / Date / Link / Grade on Inputs.
3. Open sources: Claude fetches them directly (`fetch` in the spec, status "done" once fetched) and records the URL and retrieval date.
4. Never take a base-year number from a press summary when the filing exists.

## Company financials and peers

| Need | Best source (IE library) | Where exactly | Grade |
|---|---|---|---|
| Base-year P&L, balance sheet (private EU firms) | **Orbis** (Bureau van Dijk) | Company > Financial profile > Global standard format; "Unconsolidated" for the operating entity | A1-A2 |
| Spanish / Portuguese companies | **SABI** | Company report > Cuentas (modelo normal / abreviado); export to Excel | A1-A2 |
| Listed company financials | Company annual report / 20-F / CNMV, SEC EDGAR filings | Income statement, segment note | A1 |
| 3-5 peers for the lever challenge | Orbis / SABI peer search | Search by NACE/CNAE code + country + size band; export cost lines as % of revenue | A2 |
| Peer margins, consensus forecasts | **Compustat Global**, **I/B/E/S** via WRDS | Compustat: fundq/funda (SALE, COGS, XSGA, OIADP); I/B/E/S: consensus EPS/revenue | A2-B2 |
| Industry financial benchmarks | **First Research**, **BMI Industry Reports** | First Research: "Industry financial ratios"; BMI: industry forecast tables | B2 |

## Market size, penetration, prices, volumes

| Need | Best source | Where exactly | Grade |
|---|---|---|---|
| Market size and growth by product/country | **Statista Market Insights** (sector modules) | Market Insights > sector > country > Revenue, Users, Penetration, ARPU | B2-B3 |
| Penetration, unit prices, channel shares | **Euromonitor Passport** | Passport > Industries > category > Market sizes / Brand shares / Pricing | B2 |
| Country and industry forecasts | **Oxford Economics** | Global Industry Service / Global Economic Model databank | B2 |
| Capacity ceilings (stadium seats, stores, beds) | Company filings, league/regulator data, national statistics | e.g. LaLiga stadium capacities, INE hotel occupancy survey | A2 |
| Addressable population | National statistics office, Eurostat | Population by age/region; households; vehicle stock (road_eqs_carpda) | A1 |
| Fitness and leisure (members, penetration, churn) | **Deloitte / EuropeActive European Health & Fitness Market Report**; HFA (ex-IHRSA) Global Report; operator filings (Basic-Fit, The Gym Group, PureGym bond reports) | EHFMR country tables (members, penetration, revenue per member); annual reports: members per club, churn, club capex | B2 (filings A2) |

## Macro: inflation, GDP, FX, population

| Need | Open source (Claude fetches) | Series | Grade |
|---|---|---|---|
| CPI forecast | **IMF WEO** database | PCPIPCH (avg consumer prices, % change) | A2 |
| GDP growth | IMF WEO; **OECD.Stat** Economic Outlook | NGDP_RPCH; OECD EO real GDP | A2 |
| Historical CPI, HICP | **Eurostat** | prc_hicp_aind | A1 |
| Population, demography | **World Bank** WDI; Eurostat; INE (Spain), INS (Romania), INEGI (Mexico) | SP.POP.TOTL; demo_pjan | A1 |
| FX | ECB reference rates; Bloomberg | EXR dataset | A1 |

## Cost of capital (WACC inputs)

| Input | Best source | Notes | Grade |
|---|---|---|---|
| Risk-free rate | **Bloomberg / Workspace** (10y government yield, local currency); open: central bank, Trading Economics | For emerging markets use local yield minus the sovereign default spread (Damodaran) | A2-B2 |
| Beta | **Bloomberg** BETA (2y weekly vs local/regional index); WRDS (CRSP + Compustat); open: Damodaran "Betas by sector" | Use peer median, relever to target D/E if leverage differs | B2 |
| Equity risk premium, country risk premium | Damodaran "Country Default Spreads and Risk Premiums" (open) | Mature ERP = total ERP - CRP | A2 |
| Cost of debt | Orbis (interest paid / average debt); Bloomberg corporate yields | Pre-tax; apply tax in the WACC | B2-C3 |
| Tax rate | National tax law; OECD corporate tax database | Statutory rate; effective if losses carried forward | A1 |
| Capital structure | Damodaran industry D/E; Orbis balance sheet | Target or peer-median D/V | B2 |

## Initiative effect sizes (reference classes)

| Need | Best source | Search recipe | Grade |
|---|---|---|---|
| What similar initiatives achieved (case evidence) | **Factiva**, **ABI/INFORM** | "(company OR competitor) AND (initiative keyword) AND (increase OR grew OR uplift) AND %"; 5+ cases; record each in a reference-class table | C2-C3 |
| Academic effect sizes (elasticities, loyalty, pricing, advertising) | **Scopus**, **Web of Science** | TITLE-ABS-KEY("meta-analysis" AND elasticity AND category); prefer meta-analyses | B2 |
| Diffusion parameters (Bass p, q) | Scopus: Sultan, Farley & Lehmann (1990); Van den Bulte & Stremersch (2004); category studies | Match the launch type (durable, service, app) | B2-B3 |
| Startup comparables (growth, funding, take-rates) | **Crunchbase** | Organization search: category + region + founded year; export | C3 |

## Startups and new business models

Use Crunchbase for comparables, Statista for market size, Factiva for launch case evidence. Grade
start-up self-reported metrics no better than `C3`.

## Open sources Claude can fetch directly

- IMF WEO: https://www.imf.org/en/Publications/WEO/weo-database
- OECD.Stat / OECD Data Explorer: https://data-explorer.oecd.org
- World Bank WDI API: https://api.worldbank.org/v2/country/{iso}/indicator/{code}?format=json
- Eurostat API: https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/{dataset}?geo={XX}
- Damodaran data: https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datacurrent.html
- National statistics: INE (ine.es), INS (insse.ro), INEGI (inegi.org.mx), ISTAT, INSEE, ONS
- Company filings: investor-relations pages, SEC EDGAR, CNMV, national business registers

Record for each: URL, retrieval date, series code, and grade.

## Data Request Sheet: how to write a good row

A good row lets a teammate export the file in two minutes without asking a question:

| Field | Good example | Bad example |
|---|---|---|
| Database | Orbis | "a database" |
| Exact table / report | <Company> SRL > Financial profile > Global standard format (P&L) | "financials" |
| Filters | Unconsolidated, RON, thousands | - |
| Years | 2019-2024 | "recent" |
| Export as | orbis_mol_ro_pl_2019_2024.xlsx | "export" |

File name convention: `<database>_<entity>_<content>_<years>.<ext>`, lowercase, no spaces.
