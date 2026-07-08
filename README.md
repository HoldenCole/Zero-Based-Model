# Zero-Based US Naphtha Model

A bottom-up ("zero-based") naphtha supply model for a naphtha trading desk.
Every refinery is built from its individual process units — nothing is
inherited from history. Each refinery has its own **box** (mirroring the
whiteboard design): units, capacities, utilization, naphtha yields per cut,
total naphtha, and a forward weekly production strip. Boxes roll up to
**PADD** and **US** balances, and outages (planned turnarounds first,
unplanned scenarios second) flow straight through to the forward balance so
the desk can see the **marginal barrel** and the **balance at risk**.

Scope today: **PADD 3** first, then the rest of the US, then Europe and
Asia-Pacific (the schema already carries a `region` field so expansion is
additive, not a rebuild).

## Download the Excel model

A pre-built copy lives in the repo at **[`dist/US_Naphtha_Model.xlsx`](dist/US_Naphtha_Model.xlsx)**.
On GitHub, open that file and click **Download** (or the raw/download button)
to pull the latest workbook without running anything.

To refresh it after editing any `data/` file:

```bash
./build_workbook.sh          # simple 3-sheet model
./build_workbook.sh --full   # extended workbook (balances, dashboard, checks)
```

then commit `dist/US_Naphtha_Model.xlsx`.

---

## How the model computes naphtha (the "model cut")

For each refinery, each day:

```
for each unit:
    effective_capacity = nameplate_kbd x (1 - offline_fraction_from_outages)
    throughput         = effective_capacity x utilization
    for each cut (LVN, HVN, ...):
        production[cut] += throughput x yield(unit_type, cut)
```

- Yields are **signed**: producers are positive (CDU, VDU, FCC, coker...),
  consumers are negative (reformer and isom eat naphtha as feed). The sum is
  therefore **net naphtha available to the market** — the number a trader
  cares about.
- `Total naphtha = CDU_throughput x yield + VDU_throughput x yield + ...`
  exactly as sketched on the whiteboard (`CDU x 1% + VDU x 0.5% ...`).

## Assumption hierarchy (general → specific)

Every number resolves through this chain, and the reports show *which layer*
supplied it, so you always know what is a house assumption vs. manual intel:

1. **Manual override** — refinery/unit-specific, dated, with a source note
   (`data/overrides/refinery_overrides.csv`). This is the knob you turn when
   intel lands ("X refinery is down and has to do this, which means this").
2. **PADD assumption** — per-PADD defaults (`data/assumptions/padd_overrides.yaml`).
3. **Global default** — across-the-board yields/utilization
   (`data/assumptions/global.yaml`).

## Outages and balance at risk

- `data/outages/planned_outages.csv` — turnaround schedule (planned first).
- `data/outages/unplanned_outages.csv` — live unplanned events. When an
  outage hits, add one row (refinery, unit, dates, % offline, source,
  confidence) and every downstream number — refinery box, PADD balance, US
  balance — updates immediately.
- **Balance at risk** = base-case production (no outages) minus
  outage-adjusted production, per refinery / PADD / week.
- `data/scenarios/` — hypothetical outage/flow scenarios layered on top of
  the base case without touching the live data.

## Flows, demand, and intel

- `data/flows/trade_flows.csv` — imports / exports / inter-PADD and
  intra-company transfers (so we know whose barrels move where, who to go to,
  and how to trade around it). Ship-tracking fixtures land here.
- `data/demand/demand.csv` — naphtha dispositions per PADD (cracker feed,
  gasoline blending, diluent, exports base-load).
- `data/intel/market_intel.csv` — a dated log of market intel, each row
  optionally linked to an outage id and carrying an estimated kbd impact, so
  the narrative ("refinery X is down → must buy Y → marginal barrel is Z")
  stays attached to the numbers.

## Repository layout

```
naphtha_model/          the engine (pure Python + pandas)
  config.py             PADDs, unit types, constants
  schema.py             Refinery / Unit / Outage / Override / Flow dataclasses
  loaders.py            CSV + YAML ingestion & validation
  assumptions.py        the override -> PADD -> global resolution chain
  engine.py             production math, outage application, weekly axis
  balance.py            PADD / US supply-demand balance, balance at risk
  scenario.py           scenario layering on top of the base case
  report.py             refinery "box" rendering, balance tables, Excel export
  cli.py                command line entry points
data/
  reference/            refineries.csv, units.csv  (registry — fill from the
                        capacity sheet; current rows are PLACEHOLDERS)
  assumptions/          global.yaml, padd_overrides.yaml
  overrides/            refinery_overrides.csv  (manual adjustment layer)
  outages/              planned_outages.csv, unplanned_outages.csv
  flows/                trade_flows.csv
  demand/               demand.csv
  intel/                market_intel.csv
  scenarios/            example_scenario.yaml
docs/data_intake.md     column spec for the refinery capacity sheet
tests/                  pytest suite for the math and the hierarchy
```

## Using it

```bash
pip install -r requirements.txt

python -m naphtha_model list                      # refinery registry
python -m naphtha_model box MOTIVA_PAR            # one refinery's box
python -m naphtha_model boxes --padd 3            # all boxes in a PADD
python -m naphtha_model balance --padd 3          # forward weekly PADD balance
python -m naphtha_model risk --padd 3             # balance at risk from outages
python -m naphtha_model scenario data/scenarios/example_scenario.yaml
python -m naphtha_model export --out naphtha_model.xlsx   # THE desk workbook
python -m naphtha_model export --dump --out flat.xlsx     # flat value dump
python -m naphtha_model ingest-yields data/raw/estimated_refinery_outputs_2024.xlsx
python -m naphtha_model calibrate --padd 3                # model vs 2024 actuals

pytest                                            # run the test suite
```

## The desk workbook (primary output)

`export` builds a **live, formula-driven Excel model** — not a value dump.
Once generated, the desk drives it entirely inside Excel; Python only
regenerates it when the underlying data/ files change.

**The default is the simple three-sheet model** (current desk preference):

- **Boxes** — one box per refinery: units, capacity, utilization, signed
  yields, orange manual-override cells, bold net-naphtha total. Yield-mode
  refineries read their crude capacity and 2024 net yield live from Data.
- **Assumptions** — per-PADD utilization/yield inputs plus the yield-mode
  cut split.
- **Data** — the imported registry + 2024 yields. Type a crude capacity and
  that refinery's box lights up.

`export --full` builds the extended workbook described below (forward
weekly strips, PADD balances, outage/scenario machinery, dashboard charts,
data checks) — same data, more surface:

- **Boxes** — every refinery's whiteboard box: unit rows with capacity
  (blue input), utilization and per-cut yields (grey formulas that read the
  Assumptions tab), orange **manual override** cells that beat the
  assumptions (clear the cell to revert), and a forward weekly net-naphtha
  strip that prorates outages by overlap days automatically.
- **Assumptions** — per-PADD utilization and yield matrices (blue inputs).
  Edit a yield and every box, balance and chart moves.
- **Outages** — one table for planned / unplanned / scenario rows. Add a row
  the moment an outage hits; the weekly strips and balance update instantly.
  `category = scenario` rows only count when the **scenario toggle**
  (Model!B5) is YES — flip it to layer the what-if on and off.
- **Flows** — ship-tracking cargo table (imports/exports/transfers, vessel,
  counterparty). Cargoes hit the balance in the week of their date; scenario
  cargoes obey the same toggle.
- **Balance** — per-PADD weekly supply / flows / demand / balance / at-risk,
  all SUMIFS over the input tabs.
- **Dashboard** — live charts: supply forecast (base vs outage-adjusted vs
  demand), balance, supply-at-risk bars, cargo-flow bars, and per-refinery
  forecast lines. They move when any input or the scenario toggle changes.
- **Checks** — 11 automated data-quality checks (dangling IDs, impossible
  yields/utilizations/dates, invalid flows, negative supply) with PASS/FAIL
  status and an overall ALL CHECKS PASS cell.
- **Calibration** — model-implied net naphtha yield vs 2024 actuals per
  refinery (deltas over ±2 pts flag red) plus PADD averages. This is how the
  desk tunes the unit yield assumptions.
- **Yields_2024** — the full 2024 estimated-output reference table.

The Excel formulas are verified in the test suite by recalculating the
workbook headlessly (LibreOffice) and comparing every balance number against
the Python engine.

All commands accept `--start YYYY-MM-DD` and `--weeks N` to move the forward
window.

## Methodology sequence (desk plan)

1. **Nameplate capacity** — stated unit capacities (REM 2026 + EA site
   nameplate). DONE, in the boxes.
2. **Effective capacity** — nameplate is not how units actually run (India
   tends to exceed nameplate; the US usually runs below it). Effective =
   max demonstrated annual throughput 2017–2024 excl. 2020 (RDT actuals),
   per unit, in `data/reference/effective_capacity.csv`, the boxes' "eff
   cap" column, and `python -m naphtha_model capacity`. DONE — US CDUs:
   17.7M b/d nameplate, 16.7M effective, ~15.5M running in 2024.
3. **Crude slate** — actual crude diet per refinery (REM, 2010→present, in
   `data/reference/crude_slate.csv`; purchased feedstocks incl. merchant
   naphtha/reformate buying in `feedstock_slate.csv`).
   `python -m naphtha_model slate <refinery_id>`. Feeds the naphtha-yield
   discussion (light vs heavy slate). DATA IN — quality/API typing per
   crude stream is the next desk decision.
4. **The kit, unit by unit** — CDU → overhead distillation → naphtha
   hydrofiner (NHT, now in the boxes) → reformer → mogas pool. This is the
   yield-tuning walk: set per-unit naphtha cuts and reformer/isom pull so
   each refinery's implied net naphtha matches its 2024 actual
   (Calibration view).
5. **Blend economics** — max-heavy vs max-light crude scenarios, yield
   impact, logistics costs, and where blends clear vs market prices.
   Needs: crude/product price feeds and freight assumptions.

## Status / roadmap

- [x] Foundation: schema, assumption hierarchy, engine, balances, boxes,
      scenarios, CLI, tests
- [x] Desk workbook: live formula-driven Excel model with charts, scenario
      toggle, manual overrides, and data checks
- [x] 2024 yields ingested: all 123 US refineries in the registry with
      actual net naphtha yields, yield-mode fallback, calibration view
- [x] EA site-level nameplate capacities ingested (monthly 2023-2026,
      `ingest-capacity`): all 123 refineries carry real crude capacity,
      shut sites stamped, per-PADD sums reconciled to the EA PADD series
- [ ] Add the unmatched splitter/asphalt sites (Channelview, Magellan &
      Kinder splitters, ~120 kbd) as SPLITTER units with the desk
- [ ] Replace the placeholder unit-level detail for Motiva PAR / Baytown /
      Galveston Bay with real unit capacities and tune yields via the
      calibration view
- [ ] Populate the PADD 3 turnaround schedule
- [ ] Wire in ship-tracking / fixture feed into `data/flows/`
- [ ] Refinery margin lens ("how does this refinery think about margins")
- [ ] Europe, then Asia-Pacific regions

## Data provenance

- **Real:** refinery registry & 2024 net product yields (desk's Estimated
  Refinery Outputs file); site-level monthly nameplate capacities (desk's EA
  file, stamped 2026-07); **unit-by-unit capacities for 111 refineries**
  (desk's REM export, 2026 column — `ingest-units`); **actual 2024 per-unit
  utilizations for 515 units** (desk's RefineryDataTool export, applied as
  unit-level overrides); crude slate by refinery, 2021 full product yields,
  and the EA US/Europe/Asia monthly naphtha balance in `data/reference/`;
  EIA Refinery Capacity Report 2026 kept in `data/raw/` as a cross-check.
- **Assumptions (desk-tunable, clearly labeled):** per-unit-type naphtha
  yields in `data/assumptions/` — the one remaining invented layer, and the
  reason unit-detail refineries' Calibration deltas are non-zero. Tune them
  against the 2024 actuals in the Calibration view.
- **Yield-mode (no honest unit split):** 12 refineries where REM only
  carries combined entities (PBF Delaware City + Paulsboro, Marathon LA
  Carson + Wilmington) or no units — they run capacity x utilization x
  2024 actual yield.
- **Empty until sourced:** outages, flows, demand, intel.
