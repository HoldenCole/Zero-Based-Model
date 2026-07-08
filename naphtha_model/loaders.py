"""Load and validate the CSV / YAML data files into model objects."""

from __future__ import annotations

import csv
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

import yaml

from .assumptions import AssumptionBook
from .config import DATA_DIR
from .schema import (
    DemandItem,
    IntelNote,
    Outage,
    Override,
    ProcessUnit,
    Refinery,
    TradeFlow,
)


def _parse_date(value: str, context: str) -> date:
    try:
        return datetime.strptime(value.strip(), "%Y-%m-%d").date()
    except ValueError as exc:
        raise ValueError(f"{context}: bad date {value!r}, expected YYYY-MM-DD") from exc


def _parse_optional_date(value: str, context: str) -> date | None:
    value = (value or "").strip()
    return _parse_date(value, context) if value else None


def _read_csv(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        return []
    with path.open(newline="", encoding="utf-8") as fh:
        return [
            {k.strip(): (v or "").strip() for k, v in row.items()}
            for row in csv.DictReader(fh)
            # skip blank lines and full-line comments
            if any((v or "").strip() for v in row.values())
            and not next(iter(row.values()), "").startswith("#")
        ]


@dataclass
class ModelData:
    refineries: list[Refinery] = field(default_factory=list)
    book: AssumptionBook | None = None
    outages: list[Outage] = field(default_factory=list)
    flows: list[TradeFlow] = field(default_factory=list)
    demand: list[DemandItem] = field(default_factory=list)
    intel: list[IntelNote] = field(default_factory=list)

    def refinery(self, refinery_id: str) -> Refinery:
        for r in self.refineries:
            if r.refinery_id == refinery_id:
                return r
        raise KeyError(
            f"unknown refinery {refinery_id!r}; known: "
            f"{[r.refinery_id for r in self.refineries]}"
        )


def load_refineries(data_dir: Path = DATA_DIR) -> list[Refinery]:
    ref_rows = _read_csv(data_dir / "reference" / "refineries.csv")
    unit_rows = _read_csv(data_dir / "reference" / "units.csv")

    refineries: dict[str, Refinery] = {}
    for row in ref_rows:
        r = Refinery(
            refinery_id=row["refinery_id"],
            name=row["name"],
            owner=row.get("owner", ""),
            city=row.get("city", ""),
            state=row.get("state", ""),
            padd=int(row["padd"]),
            region=row.get("region", "US") or "US",
            crude_capacity_kbd=float(row.get("crude_capacity_kbd") or 0),
            status=row.get("status", "operating") or "operating",
            notes=row.get("notes", ""),
        )
        if r.refinery_id in refineries:
            raise ValueError(f"duplicate refinery_id {r.refinery_id!r}")
        refineries[r.refinery_id] = r

    for row in unit_rows:
        rid = row["refinery_id"]
        if rid not in refineries:
            raise ValueError(f"units.csv references unknown refinery {rid!r}")
        refineries[rid].units.append(
            ProcessUnit(
                refinery_id=rid,
                unit_id=row["unit_id"],
                unit_type=row["unit_type"].upper(),
                capacity_kbd=float(row["capacity_kbd"]),
                notes=row.get("notes", ""),
            )
        )
    return list(refineries.values())


def _read_global_cfg(data_dir: Path) -> dict:
    with (data_dir / "assumptions" / "global.yaml").open(encoding="utf-8") as fh:
        return yaml.safe_load(fh)


def load_assumptions(
    data_dir: Path = DATA_DIR, extra_overrides: list[Override] | None = None
) -> AssumptionBook:
    global_cfg = _read_global_cfg(data_dir)

    padd_path = data_dir / "assumptions" / "padd_overrides.yaml"
    padd_cfg: dict = {}
    if padd_path.exists():
        with padd_path.open(encoding="utf-8") as fh:
            padd_cfg = yaml.safe_load(fh) or {}

    overrides = [
        Override(
            refinery_id=row["refinery_id"],
            unit_id=row.get("unit_id", ""),
            field_name=row["field"],
            cut=row.get("cut", ""),
            value=float(row["value"]),
            start=_parse_optional_date(row.get("start", ""), "override.start"),
            end=_parse_optional_date(row.get("end", ""), "override.end"),
            source=row.get("source", ""),
            notes=row.get("notes", ""),
        )
        for row in _read_csv(data_dir / "overrides" / "refinery_overrides.csv")
    ]
    return AssumptionBook(global_cfg, padd_cfg, overrides + list(extra_overrides or []))


def load_outages(data_dir: Path = DATA_DIR) -> list[Outage]:
    outages: list[Outage] = []
    for fname, default_type in (
        ("planned_outages.csv", "planned"),
        ("unplanned_outages.csv", "unplanned"),
    ):
        for row in _read_csv(data_dir / "outages" / fname):
            outages.append(
                Outage(
                    outage_id=row["outage_id"],
                    refinery_id=row["refinery_id"],
                    unit_id=row.get("unit_id", ""),
                    start=_parse_date(row["start"], f"{fname}:{row['outage_id']}"),
                    end=_parse_date(row["end"], f"{fname}:{row['outage_id']}"),
                    offline_pct=float(row["offline_pct"]),
                    outage_type=row.get("outage_type", "") or default_type,
                    source=row.get("source", ""),
                    confidence=row.get("confidence", ""),
                    notes=row.get("notes", ""),
                )
            )
    ids = [o.outage_id for o in outages]
    dupes = {i for i in ids if ids.count(i) > 1}
    if dupes:
        raise ValueError(f"duplicate outage_id(s): {sorted(dupes)}")
    return outages


def load_flows(data_dir: Path = DATA_DIR) -> list[TradeFlow]:
    return [
        TradeFlow(
            flow_date=_parse_date(row["date"], "trade_flows.csv"),
            padd=int(row["padd"]),
            direction=row["direction"],
            volume_kbd=float(row["volume_kbd"]),
            cut=row.get("cut", "TOTAL") or "TOTAL",
            counterparty=row.get("counterparty", ""),
            vessel=row.get("vessel", ""),
            source=row.get("source", ""),
            notes=row.get("notes", ""),
        )
        for row in _read_csv(data_dir / "flows" / "trade_flows.csv")
    ]


def load_demand(data_dir: Path = DATA_DIR) -> list[DemandItem]:
    return [
        DemandItem(
            padd=int(row["padd"]),
            sector=row["sector"],
            volume_kbd=float(row["volume_kbd"]),
            start=_parse_optional_date(row.get("start", ""), "demand.start"),
            end=_parse_optional_date(row.get("end", ""), "demand.end"),
            notes=row.get("notes", ""),
        )
        for row in _read_csv(data_dir / "demand" / "demand.csv")
    ]


def load_intel(data_dir: Path = DATA_DIR) -> list[IntelNote]:
    return [
        IntelNote(
            note_date=_parse_date(row["date"], "market_intel.csv"),
            headline=row["headline"],
            refinery_id=row.get("refinery_id", ""),
            padd=int(row["padd"]) if row.get("padd") else None,
            impact_kbd=float(row.get("impact_kbd") or 0),
            linked_outage_id=row.get("linked_outage_id", ""),
            source=row.get("source", ""),
            confidence=row.get("confidence", ""),
            notes=row.get("notes", ""),
        )
        for row in _read_csv(data_dir / "intel" / "market_intel.csv")
    ]


def load_yields_2024(data_dir: Path = DATA_DIR) -> dict[str, dict]:
    """2024 actual product yields keyed by refinery_id (empty if not ingested)."""
    rows = _read_csv(data_dir / "reference" / "refinery_yields_2024.csv")
    return {row["refinery_id"]: row for row in rows}


def _apply_yields_2024(
    refineries: list[Refinery], yields24: dict[str, dict], global_cfg: dict
) -> list[Override]:
    """Attach 2024 net naphtha yields; give unit-less refineries a synthetic
    CRUDE-EST unit (yield-mode: crude x utilization x 2024 net yield, split
    across cuts by global.yaml yield_mode.cut_shares). Returns the synthetic
    yield overrides for the AssumptionBook."""
    shares: dict = (global_cfg.get("yield_mode") or {}).get("cut_shares") or {}
    cuts = list(global_cfg.get("cuts", []))
    if shares:
        if set(shares) != set(cuts):
            raise ValueError(
                f"yield_mode.cut_shares must cover every cut {cuts}, got {list(shares)}"
            )
        if abs(sum(shares.values()) - 1.0) > 1e-6:
            raise ValueError("yield_mode.cut_shares must sum to 1")

    synth: list[Override] = []
    for r in refineries:
        y = yields24.get(r.refinery_id)
        if y is None:
            continue
        r.naphtha_yield_pct = float(y["naphtha_pct"])
        if r.units or not shares:
            continue
        r.units.append(
            ProcessUnit(
                refinery_id=r.refinery_id,
                unit_id="CRUDE-EST",
                unit_type="CDU",
                capacity_kbd=r.crude_capacity_kbd,
                notes="yield-mode: 2024 net naphtha yield applied to crude runs",
            )
        )
        for cut, share in shares.items():
            synth.append(
                Override(
                    refinery_id=r.refinery_id,
                    unit_id="CRUDE-EST",
                    field_name="yield",
                    cut=cut,
                    value=float(share) * r.naphtha_yield_pct / 100.0,
                    source="2024 actual net yield",
                )
            )
    return synth


def load_unit_utilization(data_dir: Path = DATA_DIR) -> list[Override]:
    """Actual per-unit utilization (RefineryDataTool) as utilization
    overrides — the most specific layer, beating PADD/global defaults."""
    return [
        Override(
            refinery_id=row["refinery_id"],
            unit_id=row["unit_id"],
            field_name="utilization",
            value=float(row["utilization"]),
            source=f"RDT {row['year']} actual",
        )
        for row in _read_csv(data_dir / "reference" / "unit_utilization.csv")
    ]


def load_all(data_dir: Path = DATA_DIR) -> ModelData:
    refineries = load_refineries(data_dir)
    global_cfg = _read_global_cfg(data_dir)
    synth = _apply_yields_2024(refineries, load_yields_2024(data_dir), global_cfg)
    synth += load_unit_utilization(data_dir)
    return ModelData(
        refineries=refineries,
        book=load_assumptions(data_dir, extra_overrides=synth),
        outages=load_outages(data_dir),
        flows=load_flows(data_dir),
        demand=load_demand(data_dir),
        intel=load_intel(data_dir),
    )
