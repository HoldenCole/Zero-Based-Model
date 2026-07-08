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
    """Yield-mode refinery: net = crude x util x 2024 yield, split across cuts.
    PBF Delaware City stays yield-mode because REM only carries the combined
    PBF East Coast entity, which can't be split honestly."""
    data = load_all()
    r = data.refinery("PBF_DELAWARE_CITY")
    assert [u.unit_id for u in r.units] == ["CRUDE-EST"]
    assert r.crude_capacity_kbd > 0  # EA nameplate ingested
    day = date(2026, 7, 6)
    snap = refinery_day(r, day, data.book, [])
    util = data.book.utilization(r, r.units[0], day).value
    expected = r.crude_capacity_kbd * util * r.naphtha_yield_pct / 100.0
    assert snap.net_kbd == pytest.approx(expected)
    # cut split follows global.yaml yield_mode.cut_shares
    prod = snap.production_kbd
    assert prod["LVN"] == pytest.approx(expected * 0.35)
    assert prod["HVN"] == pytest.approx(expected * 0.65)


def test_unit_ingestion_state():
    """After REM/RDT ingestion: most refineries run unit-detail with real
    capacities and actual 2024 unit utilizations."""
    data = load_all()
    day = date(2026, 7, 6)
    unit_detail = [r for r in data.refineries
                   if r.units and r.units[0].unit_id != "CRUDE-EST"]
    assert len(unit_detail) >= 100
    motiva = data.refinery("MOTIVA_PAR")
    types = {u.unit_type for u in motiva.units}
    assert {"CDU", "VDU", "FCC", "COKER", "REFORMER", "ISOM"} <= types
    cdu = motiva.unit("CDU")
    res = data.book.utilization(motiva, cdu, day)
    assert "RDT" in res.source          # actual 2024 utilization applied
    assert 0.5 < res.value <= 1.1
    # splitters came through where REM has them
    splitters = [u for r in data.refineries for u in r.units
                 if u.unit_type == "SPLITTER"]
    assert splitters


def test_capacity_ingestion_state():
    """Registry state after EA capacity ingestion: every refinery stamped,
    shut sites at zero, per-PADD sums close to EA PADD nameplate."""
    import csv

    from naphtha_model.ingest import COMPONENT_NOTES

    data = load_all()
    stamped = [r for r in data.refineries
               if r.refinery_id not in COMPONENT_NOTES]
    operating = [r for r in stamped if r.crude_capacity_kbd > 0]
    assert len(operating) >= 100
    for rid in ("MOTIVA_PAR", "XOM_BAYTOWN", "MPC_GALV_BAY"):
        assert data.refinery(rid).crude_capacity_kbd > 500
    assert data.refinery("SHELL_CONVENT").crude_capacity_kbd == 0
    assert data.refinery("SHELL_CONVENT").status == "shut"

    # monthly series present for every site
    monthly = DATA_DIR / "reference" / "site_capacity_monthly.csv"
    rows = list(csv.DictReader(monthly.open()))
    assert len({r["site"] for r in rows}) == 150

    # reconcile vs the EA PADD nameplate series (gap = known unmatched
    # splitter/asphalt sites, < 200 kbd nationally)
    padd_raw = DATA_DIR / "raw" / "ea_padd_capacity_2023_2026.csv"
    if padd_raw.exists():
        ea = {}
        for r in csv.DictReader(padd_raw.open(encoding="utf-8-sig")):
            if (r["source"] == "Energy Aspects" and r["date"] == "2026-07-01"
                    and "nameplate capacity for PADD" in r["series_label"]):
                ea[int(r["sub_region"].split("_")[1])] = float(r["value"])
        reg_total = sum(r.crude_capacity_kbd for r in data.refineries)
        assert sum(ea.values()) - reg_total == pytest.approx(174.0, abs=1.0)


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
