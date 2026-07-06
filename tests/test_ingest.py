"""Ingestion of the 2024 yields workbook + yield-mode loading."""

from datetime import date

import pytest

from naphtha_model.config import DATA_DIR
from naphtha_model.engine import build_axis, refinery_day
from naphtha_model.ingest import parse_yields_workbook
from naphtha_model.loaders import load_all, load_yields_2024

RAW = DATA_DIR / "raw" / "estimated_refinery_outputs_2024.xlsx"


@pytest.mark.skipif(not RAW.exists(), reason="raw 2024 yields workbook not present")
def test_parse_yields_workbook():
    records = parse_yields_workbook(RAW)
    assert len(records) == 123
    # every PADD represented; same-city refineries stay distinct
    assert {r["padd"] for r in records} == {1, 2, 3, 4, 5}
    port_arthur = [r for r in records if r["city"] == "Port Arthur"]
    assert len(port_arthur) == 4
    # JV owners folded into the operator row
    wood_river = next(r for r in records if r["city"] == "Wood River")
    assert len(wood_river["owners"]) == 2


def test_yields_csv_and_registry_consistent():
    yields = load_yields_2024()
    data = load_all()
    assert len(yields) == 123
    ids = {r.refinery_id for r in data.refineries}
    assert set(yields) <= ids
    # the original trio kept their ids and unit detail
    for rid in ("MOTIVA_PAR", "XOM_BAYTOWN", "MPC_GALV_BAY"):
        r = data.refinery(rid)
        assert r.naphtha_yield_pct is not None
        assert all(u.unit_id != "CRUDE-EST" for u in r.units)


def test_yield_mode_math():
    """Yield-mode refinery: net = crude x util x 2024 yield, split across cuts."""
    data = load_all()
    r = data.refinery("EXXONMOBIL_BATON_ROUGE")
    assert [u.unit_id for u in r.units] == ["CRUDE-EST"]
    # no capacity yet -> zero production but the machinery is live
    day = date(2026, 7, 6)
    snap = refinery_day(r, day, data.book, [])
    assert snap.net_kbd == 0.0
    # simulate the capacity sheet landing
    object.__setattr__(r.units[0], "capacity_kbd", 500.0)
    snap = refinery_day(r, day, data.book, [])
    util = data.book.utilization(r, r.units[0], day).value
    expected = 500.0 * util * r.naphtha_yield_pct / 100.0
    assert snap.net_kbd == pytest.approx(expected)
    # cut split follows global.yaml yield_mode.cut_shares
    prod = snap.production_kbd
    assert prod["LVN"] == pytest.approx(expected * 0.35)
    assert prod["HVN"] == pytest.approx(expected * 0.65)


def test_calibration_inputs_available():
    data = load_all()
    with_yield = [r for r in data.refineries if r.naphtha_yield_pct is not None]
    assert len(with_yield) == 123
    axis = build_axis(date(2026, 7, 6), 2)
    # model still runs end-to-end over the full registry
    total = sum(
        refinery_day(r, axis[0], data.book, data.outages).net_kbd
        for r in data.refineries
    )
    assert total > 0
