"""PADD and US supply/demand balances built from the refinery boxes.

Balance per PADD per week (all kbd):

    refinery net naphtha supply
  + imports + transfers in
  - exports - transfers out
  - demand (cracker feed, gasoline blending, diluent, ...)
  = net balance  (positive = length / barrels looking for a home,
                  negative = short / must pull barrels in)
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, timedelta

from .assumptions import AssumptionBook
from .config import DAYS_PER_WEEK
from .engine import balance_at_risk_kbd, weekly_net_kbd
from .schema import DemandItem, Outage, Refinery, TradeFlow


@dataclass
class PaddWeek:
    padd: int
    week_start: date
    supply_kbd: float
    flows_kbd: float      # signed net trade/transfer
    demand_kbd: float
    at_risk_kbd: float    # supply lost to outages this week

    @property
    def balance_kbd(self) -> float:
        return self.supply_kbd + self.flows_kbd - self.demand_kbd


def _week_days(week_start: date) -> list[date]:
    return [week_start + timedelta(days=d) for d in range(DAYS_PER_WEEK)]


def padd_balance(
    padd: int,
    refineries: list[Refinery],
    axis: list[date],
    book: AssumptionBook,
    outages: list[Outage],
    flows: list[TradeFlow],
    demand: list[DemandItem],
) -> list[PaddWeek]:
    padd_refs = [r for r in refineries if r.padd == padd and r.region == "US"]
    supply: dict[date, float] = {w: 0.0 for w in axis}
    at_risk: dict[date, float] = {w: 0.0 for w in axis}
    for r in padd_refs:
        for w, kbd in weekly_net_kbd(r, axis, book, outages).items():
            supply[w] += kbd
        for w, kbd in balance_at_risk_kbd(r, axis, book, outages).items():
            at_risk[w] += kbd

    weeks: list[PaddWeek] = []
    for w in axis:
        days = _week_days(w)
        flow_kbd = sum(
            f.signed_kbd for f in flows if f.padd == padd and days[0] <= f.flow_date <= days[-1]
        )
        # Demand items are steady-state kbd rates; average their active days.
        demand_kbd = sum(
            d.volume_kbd * sum(1 for day in days if d.active_on(day)) / DAYS_PER_WEEK
            for d in demand
            if d.padd == padd
        )
        weeks.append(
            PaddWeek(
                padd=padd,
                week_start=w,
                supply_kbd=supply[w],
                flows_kbd=flow_kbd,
                demand_kbd=demand_kbd,
                at_risk_kbd=at_risk[w],
            )
        )
    return weeks


def us_balance(
    refineries: list[Refinery],
    axis: list[date],
    book: AssumptionBook,
    outages: list[Outage],
    flows: list[TradeFlow],
    demand: list[DemandItem],
) -> dict[int, list[PaddWeek]]:
    """Balance for every PADD that has refineries, demand, or flows loaded."""
    padds = sorted(
        {r.padd for r in refineries if r.region == "US"}
        | {d.padd for d in demand}
        | {f.padd for f in flows}
    )
    return {
        p: padd_balance(p, refineries, axis, book, outages, flows, demand) for p in padds
    }
