# Data intake: the refinery capacity sheet

When the US refinery capacity sheet arrives, it replaces the PLACEHOLDER rows
in `data/reference/refineries.csv` and `data/reference/units.csv`. Any format
works (Excel/CSV) — it just needs to map onto these columns. If the sheet has
different headers, keep the sheet as-is in `data/raw/` and we'll write a
one-time converter.

## refineries.csv — one row per refinery

| column               | required | example              | notes                                  |
|----------------------|----------|----------------------|----------------------------------------|
| refinery_id          | yes      | MOTIVA_PAR           | short unique key, no spaces            |
| name                 | yes      | Motiva Port Arthur   |                                        |
| owner                | no       | Saudi Aramco         | who to go to / who trades the barrels  |
| city, state          | no       | Port Arthur, TX      |                                        |
| padd                 | yes      | 3                    | 1–5                                    |
| region               | no       | US                   | US / EUROPE / APAC (defaults to US)    |
| crude_capacity_kbd   | no       | 626                  | headline crude capacity                |
| status               | no       | operating            | operating / shut / mothballed          |
| notes                | no       |                      |                                        |

## units.csv — one row per process unit per refinery

| column        | required | example | notes                                                |
|---------------|----------|---------|------------------------------------------------------|
| refinery_id   | yes      | MOTIVA_PAR | must match refineries.csv                         |
| unit_id       | yes      | CDU-1   | unique within the refinery                           |
| unit_type     | yes      | CDU     | CDU, VDU, FCC, COKER, HYDROCRACKER, REFORMER, ISOM, ALKY, SPLITTER |
| capacity_kbd  | yes      | 350     | nameplate, thousand barrels/day                      |
| notes         | no       |         |                                                      |

The model needs at minimum the naphtha-relevant units: **CDUs** (and
splitters) for straight-run make, **FCC/coker/hydrocracker** if their naphtha
reaches the market pool, and **reformers/isom** because they consume naphtha
and set the net number.

## What the desk sets afterwards (no code involved)

- `data/assumptions/global.yaml` — house yields per unit type, default utilization
- `data/assumptions/padd_overrides.yaml` — per-PADD deviations
- `data/overrides/refinery_overrides.csv` — dated, sourced, refinery/unit-level
  manual tweaks (the intel knob)
- `data/outages/planned_outages.csv` — the turnaround schedule
- `data/outages/unplanned_outages.csv` — live outages as they hit
- `data/flows/trade_flows.csv`, `data/demand/demand.csv`, `data/intel/market_intel.csv`
