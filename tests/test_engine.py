from datetime import date

import pytest

from conftest import make_outage
from naphtha_model.engine import (
    balance_at_risk_kbd,
    build_axis,
    offline_fraction,
    refinery_day,
    weekly_net_kbd,
)

DAY = date(2026, 7, 6)  # before any test outage


def test_production_math(refinery, book):
    """CDU 200 kbd x 0.95 util x yields; reformer 50 kbd consumes HVN 1:1."""
    snap = refinery_day(refinery, DAY, book, outages=[])
    prod = snap.production_kbd
    cdu_throughput = 200 * 0.95
    ref_throughput = 50 * 0.95
    assert prod["LVN"] == pytest.approx(cdu_throughput * 0.05)
    assert prod["HVN"] == pytest.approx(cdu_throughput * 0.12 - ref_throughput)
    assert snap.gross_kbd == pytest.approx(cdu_throughput * (0.05 + 0.12))
    assert snap.consumed_kbd == pytest.approx(ref_throughput)
    assert snap.net_kbd == pytest.approx(snap.gross_kbd - snap.consumed_kbd)


def test_offline_fraction_overlap_takes_max(refinery):
    cdu = refinery.unit("CDU-1")
    outages = [
        make_outage(outage_id="A", offline_pct=40),
        make_outage(outage_id="B", offline_pct=70),
    ]
    assert offline_fraction(cdu, date(2026, 7, 15), outages) == 0.70
    assert offline_fraction(cdu, date(2026, 7, 25), outages) == 0.0


def test_whole_refinery_outage_hits_every_unit(refinery):
    outages = [make_outage(unit_id="", offline_pct=100)]
    for unit in refinery.units:
        assert offline_fraction(unit, date(2026, 7, 15), outages) == 1.0


def test_outage_reduces_production(refinery, book):
    outages = [make_outage(offline_pct=100)]  # CDU-1 down 7/13-7/19
    hit = refinery_day(refinery, date(2026, 7, 15), book, outages)
    base = refinery_day(refinery, date(2026, 7, 15), book, [])
    # CDU makes nothing; reformer still pulls feed, so net drops below base
    assert hit.gross_kbd == 0.0
    assert hit.net_kbd < base.net_kbd


def test_weekly_average_partial_outage(refinery, book):
    """Outage covering 7/13-7/19 wipes exactly the second week of the axis."""
    axis = build_axis(date(2026, 7, 6), weeks=3)
    outages = [make_outage(offline_pct=100)]
    weekly = weekly_net_kbd(refinery, axis, book, outages)
    base = weekly_net_kbd(refinery, axis, book, outages, include_outages=False)
    assert weekly[date(2026, 7, 6)] == pytest.approx(base[date(2026, 7, 6)])
    assert weekly[date(2026, 7, 13)] < base[date(2026, 7, 13)]
    assert weekly[date(2026, 7, 20)] == pytest.approx(base[date(2026, 7, 20)])


def test_balance_at_risk_is_base_minus_hit(refinery, book):
    axis = build_axis(date(2026, 7, 6), weeks=3)
    outages = [make_outage(offline_pct=100)]
    risk = balance_at_risk_kbd(refinery, axis, book, outages)
    assert risk[date(2026, 7, 6)] == pytest.approx(0.0)
    cdu_weekly_net = 200 * 0.95 * (0.05 + 0.12)
    assert risk[date(2026, 7, 13)] == pytest.approx(cdu_weekly_net)
