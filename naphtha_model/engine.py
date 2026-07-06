"""Production math: unit throughput, outage application, forward weekly axis.

All volumes are kbd (thousand barrels per day). Weekly numbers are the
average of the 7 daily values in the week, so an outage that starts mid-week
shows a partial-week impact instead of an all-or-nothing step.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, timedelta

from .assumptions import AssumptionBook
from .config import DAYS_PER_WEEK, DEFAULT_FORWARD_WEEKS
from .schema import Outage, ProcessUnit, Refinery


def build_axis(start: date, weeks: int = DEFAULT_FORWARD_WEEKS) -> list[date]:
    """Week-start dates for the forward window."""
    return [start + timedelta(weeks=w) for w in range(weeks)]


def offline_fraction(unit: ProcessUnit, day: date, outages: list[Outage]) -> float:
    """Fraction of the unit offline on `day`.

    Overlapping outages are combined by taking the maximum (an outage row for
    the whole refinery, unit_id blank, applies to every unit). Capped at 1.
    """
    frac = 0.0
    for o in outages:
        if o.refinery_id != unit.refinery_id:
            continue
        if o.unit_id not in ("", unit.unit_id):
            continue
        if o.active_on(day):
            frac = max(frac, o.offline_fraction)
    return min(frac, 1.0)


@dataclass
class UnitDay:
    """One unit's contribution on one day — the row inside the refinery box."""

    unit: ProcessUnit
    effective_capacity_kbd: float
    utilization: float
    utilization_source: str
    throughput_kbd: float
    yields: dict[str, float]           # cut -> fraction (signed)
    yield_sources: dict[str, str]      # cut -> where the yield came from
    production_kbd: dict[str, float]   # cut -> kbd (signed)
    offline_frac: float


@dataclass
class RefineryDay:
    refinery: Refinery
    day: date
    unit_days: list[UnitDay] = field(default_factory=list)

    @property
    def production_kbd(self) -> dict[str, float]:
        """Signed production per cut, summed over units."""
        totals: dict[str, float] = {}
        for ud in self.unit_days:
            for cut, kbd in ud.production_kbd.items():
                totals[cut] = totals.get(cut, 0.0) + kbd
        return totals

    @property
    def gross_kbd(self) -> float:
        """Naphtha made (positive contributions only)."""
        return sum(
            kbd for ud in self.unit_days for kbd in ud.production_kbd.values() if kbd > 0
        )

    @property
    def consumed_kbd(self) -> float:
        """Naphtha eaten internally (reformer/isom feed), as a positive number."""
        return -sum(
            kbd for ud in self.unit_days for kbd in ud.production_kbd.values() if kbd < 0
        )

    @property
    def net_kbd(self) -> float:
        """Net naphtha available to the market."""
        return self.gross_kbd - self.consumed_kbd


def refinery_day(
    refinery: Refinery,
    day: date,
    book: AssumptionBook,
    outages: list[Outage],
    include_outages: bool = True,
) -> RefineryDay:
    result = RefineryDay(refinery=refinery, day=day)
    for unit in refinery.units:
        off = offline_fraction(unit, day, outages) if include_outages else 0.0
        eff_cap = unit.capacity_kbd * (1.0 - off)
        util = book.utilization(refinery, unit, day)
        throughput = eff_cap * util.value
        resolved_yields = book.yields(refinery, unit, day)
        production = {cut: throughput * r.value for cut, r in resolved_yields.items()}
        result.unit_days.append(
            UnitDay(
                unit=unit,
                effective_capacity_kbd=eff_cap,
                utilization=util.value,
                utilization_source=util.source,
                throughput_kbd=throughput,
                yields={c: r.value for c, r in resolved_yields.items()},
                yield_sources={c: r.source for c, r in resolved_yields.items()},
                production_kbd=production,
                offline_frac=off,
            )
        )
    return result


def weekly_net_kbd(
    refinery: Refinery,
    axis: list[date],
    book: AssumptionBook,
    outages: list[Outage],
    include_outages: bool = True,
) -> dict[date, float]:
    """Average net naphtha kbd for each week on the axis."""
    out: dict[date, float] = {}
    for week_start in axis:
        days = [week_start + timedelta(days=d) for d in range(DAYS_PER_WEEK)]
        daily = [
            refinery_day(refinery, d, book, outages, include_outages).net_kbd for d in days
        ]
        out[week_start] = sum(daily) / DAYS_PER_WEEK
    return out


def balance_at_risk_kbd(
    refinery: Refinery,
    axis: list[date],
    book: AssumptionBook,
    outages: list[Outage],
) -> dict[date, float]:
    """Weekly naphtha lost to outages: base case minus outage case (>= 0)."""
    base = weekly_net_kbd(refinery, axis, book, outages, include_outages=False)
    hit = weekly_net_kbd(refinery, axis, book, outages, include_outages=True)
    return {w: base[w] - hit[w] for w in axis}
