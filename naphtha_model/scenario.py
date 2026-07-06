"""Scenario layering: hypothetical outages / flow shifts on top of the base case.

A scenario YAML looks like:

    name: MOTIVA_CDU_TRIP
    description: What if Motiva PAR loses its big CDU for 3 weeks?
    outages:
      - outage_id: SCN-MOTIVA-CDU
        refinery_id: MOTIVA_PAR
        unit_id: CDU-1
        start: 2026-07-13
        end: 2026-08-02
        offline_pct: 100
        outage_type: unplanned
    flows:
      - date: 2026-07-20
        padd: 3
        direction: import
        volume_kbd: 10
        counterparty: NWE
        notes: replacement cargo pulled transatlantic

The scenario is evaluated against the live data without modifying it.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

import yaml

from .balance import PaddWeek, us_balance
from .loaders import ModelData
from .schema import Outage, TradeFlow


@dataclass
class Scenario:
    name: str
    description: str = ""
    outages: list[Outage] = field(default_factory=list)
    flows: list[TradeFlow] = field(default_factory=list)


def _as_date(value) -> date:
    if isinstance(value, date):
        return value
    return datetime.strptime(str(value), "%Y-%m-%d").date()


def load_scenario(path: Path) -> Scenario:
    with Path(path).open(encoding="utf-8") as fh:
        raw = yaml.safe_load(fh)
    outages = [
        Outage(
            outage_id=o["outage_id"],
            refinery_id=o["refinery_id"],
            unit_id=o.get("unit_id", ""),
            start=_as_date(o["start"]),
            end=_as_date(o["end"]),
            offline_pct=float(o["offline_pct"]),
            outage_type=o.get("outage_type", "unplanned"),
            source=f"scenario:{raw['name']}",
            notes=o.get("notes", ""),
        )
        for o in raw.get("outages", [])
    ]
    flows = [
        TradeFlow(
            flow_date=_as_date(f["date"]),
            padd=int(f["padd"]),
            direction=f["direction"],
            volume_kbd=float(f["volume_kbd"]),
            cut=f.get("cut", "TOTAL"),
            counterparty=f.get("counterparty", ""),
            source=f"scenario:{raw['name']}",
            notes=f.get("notes", ""),
        )
        for f in raw.get("flows", [])
    ]
    return Scenario(
        name=raw["name"],
        description=raw.get("description", ""),
        outages=outages,
        flows=flows,
    )


@dataclass
class ScenarioResult:
    scenario: Scenario
    base: dict[int, list[PaddWeek]]
    case: dict[int, list[PaddWeek]]

    def delta_balance(self, padd: int) -> dict[date, float]:
        """Scenario balance minus base balance, per week (kbd)."""
        base_weeks = {w.week_start: w.balance_kbd for w in self.base.get(padd, [])}
        return {
            w.week_start: w.balance_kbd - base_weeks.get(w.week_start, 0.0)
            for w in self.case.get(padd, [])
        }


def run_scenario(scenario: Scenario, data: ModelData, axis: list[date]) -> ScenarioResult:
    base = us_balance(data.refineries, axis, data.book, data.outages, data.flows, data.demand)
    case = us_balance(
        data.refineries,
        axis,
        data.book,
        data.outages + scenario.outages,
        data.flows + scenario.flows,
        data.demand,
    )
    return ScenarioResult(scenario=scenario, base=base, case=case)
