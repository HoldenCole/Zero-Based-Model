"""The three-sheet simple workbook (current default export)."""

import shutil
import subprocess
from datetime import date
from pathlib import Path

import openpyxl
import pytest

from naphtha_model.engine import build_axis, refinery_day
from naphtha_model.loaders import load_all
from naphtha_model.workbook import build_simple_workbook

AXIS = build_axis(date(2026, 7, 6), 8)


@pytest.fixture(scope="module")
def built(tmp_path_factory):
    data = load_all()
    out = tmp_path_factory.mktemp("wb") / "simple.xlsx"
    build_simple_workbook(data, AXIS, out)
    return data, out


def _totals(path) -> dict[str, float]:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Boxes"]
    return {
        ws.cell(row=r, column=2).value: ws.cell(row=r, column=14).value
        for r in range(3, 900)
        if ws.cell(row=r, column=1).value == "TOTAL"
    }


def test_exactly_three_sheets(built):
    _, out = built
    wb = openpyxl.load_workbook(out)
    assert wb.sheetnames == ["Boxes", "Assumptions", "Data"]


def test_every_refinery_has_a_box(built):
    data, out = built
    wb = openpyxl.load_workbook(out)
    ws = wb["Boxes"]
    total_rows = [r for r in range(3, 900) if ws.cell(row=r, column=1).value == "TOTAL"]
    assert len(total_rows) == len(data.refineries)
    # yield-mode capacity and yields read the Data sheet live
    unit_row = next(
        r for r in range(3, 900)
        if ws.cell(row=r, column=1).value == "UNIT"
        and ws.cell(row=r, column=4).value == "CRUDE-EST"
    )
    assert "Data!" in str(ws.cell(row=unit_row, column=6).value)
    assert "Data!" in str(ws.cell(row=unit_row, column=12).value)


@pytest.mark.skipif(shutil.which("soffice") is None, reason="LibreOffice not installed")
def test_boxes_match_engine_and_capacity_lights_up(built, tmp_path):
    data, out = built
    # simulate the capacity sheet: type 500 kbd into a yield-mode refinery
    wb = openpyxl.load_workbook(out)
    ws = wb["Data"]
    row = next(r for r in range(3, 300)
               if ws.cell(row=r, column=1).value == "EXXONMOBIL_BATON_ROUGE")
    ws.cell(row=row, column=7, value=500)
    capped = tmp_path / "capped.xlsx"
    wb.save(capped)

    profile = tmp_path / "loprofile"
    for f in (out, capped):
        subprocess.run(
            ["soffice", f"-env:UserInstallation=file://{profile}", "--headless",
             "--convert-to", "xlsx", "--outdir", str(tmp_path / "recalc"), str(f)],
            check=True, capture_output=True, timeout=300,
        )
    base = _totals(tmp_path / "recalc" / Path(out).name)
    cap = _totals(tmp_path / "recalc" / "capped.xlsx")

    day = AXIS[0]
    for rid in ("MOTIVA_PAR", "XOM_BAYTOWN", "MPC_GALV_BAY"):
        py = refinery_day(data.refinery(rid), day, data.book, [], include_outages=False).net_kbd
        assert base[rid] == pytest.approx(py, abs=0.05)

    r = data.refinery("EXXONMOBIL_BATON_ROUGE")
    util = data.book.utilization(r, r.units[0], day).value
    assert base["EXXONMOBIL_BATON_ROUGE"] == pytest.approx(
        r.crude_capacity_kbd * util * r.naphtha_yield_pct / 100, abs=0.01
    )
    assert cap["EXXONMOBIL_BATON_ROUGE"] == pytest.approx(
        500 * util * r.naphtha_yield_pct / 100, abs=0.01
    )
