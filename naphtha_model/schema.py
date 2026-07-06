"""Core dataclasses: the building blocks of the zero-based model."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date

from .config import CONSUMER_UNIT_TYPES, FLOW_DIRECTIONS, OUTAGE_TYPES, PADDS, UNIT_TYPES


@dataclass(frozen=True)
class ProcessUnit:
    refinery_id: str
    unit_id: str          # unique within the refinery, e.g. "CDU-1"
    unit_type: str        # one of config.UNIT_TYPES
    capacity_kbd: float   # nameplate, thousand barrels per day
    notes: str = ""

    def __post_init__(self) -> None:
        if self.unit_type not in UNIT_TYPES:
            raise ValueError(
                f"{self.refinery_id}/{self.unit_id}: unknown unit_type "
                f"{self.unit_type!r} (expected one of {UNIT_TYPES})"
            )
        if self.capacity_kbd < 0:
            raise ValueError(f"{self.refinery_id}/{self.unit_id}: negative capacity")

    @property
    def is_consumer(self) -> bool:
        return self.unit_type in CONSUMER_UNIT_TYPES


@dataclass
class Refinery:
    refinery_id: str
    name: str
    owner: str
    city: str
    state: str
    padd: int
    region: str = "US"
    crude_capacity_kbd: float = 0.0
    status: str = "operating"
    notes: str = ""
    units: list[ProcessUnit] = field(default_factory=list)

    def __post_init__(self) -> None:
        if self.region == "US" and self.padd not in PADDS:
            raise ValueError(f"{self.refinery_id}: invalid PADD {self.padd}")

    def unit(self, unit_id: str) -> ProcessUnit:
        for u in self.units:
            if u.unit_id == unit_id:
                return u
        raise KeyError(f"{self.refinery_id}: no unit {unit_id!r}")


@dataclass(frozen=True)
class Outage:
    outage_id: str
    refinery_id: str
    unit_id: str            # blank/"" means the whole refinery
    start: date
    end: date               # inclusive
    offline_pct: float      # 0-100, share of the unit's capacity offline
    outage_type: str        # planned | unplanned
    source: str = ""
    confidence: str = ""    # e.g. high / medium / low
    notes: str = ""

    def __post_init__(self) -> None:
        if self.outage_type not in OUTAGE_TYPES:
            raise ValueError(f"{self.outage_id}: outage_type must be one of {OUTAGE_TYPES}")
        if not (0 <= self.offline_pct <= 100):
            raise ValueError(f"{self.outage_id}: offline_pct must be 0-100")
        if self.end < self.start:
            raise ValueError(f"{self.outage_id}: end before start")

    def active_on(self, day: date) -> bool:
        return self.start <= day <= self.end

    @property
    def offline_fraction(self) -> float:
        return self.offline_pct / 100.0


@dataclass(frozen=True)
class Override:
    """A dated manual adjustment — the trader's knob.

    field == "utilization": value replaces the utilization fraction (0-1).
    field == "yield":       value replaces the yield fraction for `cut`
                            (signed; consumers negative). Requires `cut`.
    Targets a single unit if unit_id is set, otherwise every unit in the
    refinery (for "yield", every unit — use unit_id for unit-type precision).
    """

    refinery_id: str
    field_name: str         # "utilization" | "yield"
    value: float
    unit_id: str = ""       # blank = whole refinery
    cut: str = ""           # required when field_name == "yield"
    start: date | None = None
    end: date | None = None
    source: str = ""
    notes: str = ""

    def __post_init__(self) -> None:
        if self.field_name not in ("utilization", "yield"):
            raise ValueError(f"override field must be utilization|yield, got {self.field_name!r}")
        if self.field_name == "yield" and not self.cut:
            raise ValueError("yield override requires a cut (e.g. LVN, HVN)")

    def active_on(self, day: date) -> bool:
        if self.start and day < self.start:
            return False
        if self.end and day > self.end:
            return False
        return True

    def targets(self, refinery_id: str, unit_id: str) -> bool:
        if self.refinery_id != refinery_id:
            return False
        return self.unit_id in ("", unit_id)


@dataclass(frozen=True)
class TradeFlow:
    flow_date: date
    padd: int
    direction: str          # import | export | transfer_in | transfer_out
    volume_kbd: float       # average kbd over the covered week(s)
    cut: str = "TOTAL"
    counterparty: str = ""  # region/company on the other side
    vessel: str = ""
    source: str = ""
    notes: str = ""

    def __post_init__(self) -> None:
        if self.direction not in FLOW_DIRECTIONS:
            raise ValueError(f"flow direction must be one of {FLOW_DIRECTIONS}")

    @property
    def signed_kbd(self) -> float:
        """Positive adds supply to the PADD, negative removes it."""
        return self.volume_kbd if self.direction in ("import", "transfer_in") else -self.volume_kbd


@dataclass(frozen=True)
class DemandItem:
    padd: int
    sector: str             # petchem_cracker | gasoline_blending | diluent | other
    volume_kbd: float
    start: date | None = None
    end: date | None = None
    notes: str = ""

    def active_on(self, day: date) -> bool:
        if self.start and day < self.start:
            return False
        if self.end and day > self.end:
            return False
        return True


@dataclass(frozen=True)
class IntelNote:
    note_date: date
    headline: str
    refinery_id: str = ""
    padd: int | None = None
    impact_kbd: float = 0.0     # signed: + adds naphtha supply, - removes it
    linked_outage_id: str = ""
    source: str = ""
    confidence: str = ""
    notes: str = ""
