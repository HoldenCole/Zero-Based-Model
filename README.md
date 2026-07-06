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
python -m naphtha_model export --out naphtha_model.xlsx   # Excel workbook

pytest                                            # run the test suite
```

All commands accept `--start YYYY-MM-DD` and `--weeks N` to move the forward
window.

## Status / roadmap

- [x] Foundation: schema, assumption hierarchy, engine, balances, boxes,
      scenarios, CLI, tests (this commit)
- [ ] Load the real US refinery capacity sheet (replaces placeholder rows in
      `data/reference/`) — see `docs/data_intake.md` for the expected columns
- [ ] Populate the PADD 3 turnaround schedule
- [ ] Wire in ship-tracking / fixture feed into `data/flows/`
- [ ] Refinery margin lens ("how does this refinery think about margins")
- [ ] Europe, then Asia-Pacific regions

> **Placeholder data warning:** every capacity, yield, and utilization number
> currently in `data/` is an illustrative placeholder so the model runs
> end-to-end. Do not trade off these numbers until the real capacity sheet
> and desk assumptions are loaded.
