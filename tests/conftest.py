from datetime import date

import pytest

from naphtha_model.assumptions import AssumptionBook
from naphtha_model.schema import Outage, Override, ProcessUnit, Refinery

GLOBAL_CFG = {
    "cuts": ["LVN", "HVN"],
    "utilization": {"default": 0.90},
    "yields": {
        "CDU": {"LVN": 0.05, "HVN": 0.10},
        "REFORMER": {"LVN": 0.0, "HVN": -1.0},
    },
}

PADD_CFG = {3: {"utilization": 0.95, "yields": {"CDU": {"HVN": 0.12}}}}


@pytest.fixture
def refinery() -> Refinery:
    r = Refinery(
        refinery_id="TEST_REF",
        name="Test Refinery",
        owner="Test Co",
        city="Houston",
        state="TX",
        padd=3,
        crude_capacity_kbd=300,
    )
    r.units = [
        ProcessUnit("TEST_REF", "CDU-1", "CDU", 200.0),
        ProcessUnit("TEST_REF", "REF-1", "REFORMER", 50.0),
    ]
    return r


@pytest.fixture
def padd1_refinery() -> Refinery:
    r = Refinery(
        refinery_id="P1_REF",
        name="East Coast Refinery",
        owner="Test Co",
        city="Philadelphia",
        state="PA",
        padd=1,
        crude_capacity_kbd=100,
    )
    r.units = [ProcessUnit("P1_REF", "CDU-1", "CDU", 100.0)]
    return r


@pytest.fixture
def book() -> AssumptionBook:
    return AssumptionBook(GLOBAL_CFG, PADD_CFG, [])


def make_book(overrides: list[Override]) -> AssumptionBook:
    return AssumptionBook(GLOBAL_CFG, PADD_CFG, overrides)


def make_outage(**kw) -> Outage:
    defaults = dict(
        outage_id="O1",
        refinery_id="TEST_REF",
        unit_id="CDU-1",
        start=date(2026, 7, 13),
        end=date(2026, 7, 19),
        offline_pct=100,
        outage_type="planned",
    )
    defaults.update(kw)
    return Outage(**defaults)
