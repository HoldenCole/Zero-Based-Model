"""End-to-end checks against the repository's live data/ directory, so a bad
edit to any CSV/YAML breaks the build instead of silently corrupting numbers.
"""

from datetime import date

from naphtha_model.engine import build_axis
from naphtha_model.balance import us_balance
from naphtha_model.loaders import load_all
from naphtha_model.report import render_padd_balance, render_refinery_box
from naphtha_model.scenario import load_scenario, run_scenario
from naphtha_model.config import DATA_DIR


def test_data_directory_loads():
    data = load_all()
    assert data.refineries, "refinery registry is empty"
    assert data.book is not None
    for r in data.refineries:
        assert r.units, f"{r.refinery_id} has no units"


def test_boxes_and_balance_render():
    data = load_all()
    axis = build_axis(date(2026, 7, 6), weeks=4)
    for r in data.refineries:
        box = render_refinery_box(r, data, axis)
        assert r.name in box
        assert "NET NAPHTHA" in box
    balances = us_balance(
        data.refineries, axis, data.book, data.outages, data.flows, data.demand
    )
    assert 3 in balances
    text = render_padd_balance(3, balances[3])
    assert "PADD 3" in text


def test_example_scenario_runs():
    data = load_all()
    axis = build_axis(date(2026, 7, 6), weeks=4)
    scn = load_scenario(DATA_DIR / "scenarios" / "example_scenario.yaml")
    result = run_scenario(scn, data, axis)
    delta = result.delta_balance(3)
    # losing a CDU must make PADD 3 shorter in the affected weeks
    assert min(delta.values()) < 0
