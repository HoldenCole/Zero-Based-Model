from datetime import date

import pytest

from conftest import make_outage
from naphtha_model.balance import padd_balance, us_balance
from naphtha_model.engine import build_axis
from naphtha_model.schema import DemandItem, TradeFlow

AXIS = build_axis(date(2026, 7, 6), weeks=2)


def test_padd_balance_composition(refinery, book):
    flows = [
        TradeFlow(flow_date=date(2026, 7, 8), padd=3, direction="export", volume_kbd=10),
        TradeFlow(flow_date=date(2026, 7, 9), padd=3, direction="import", volume_kbd=4),
    ]
    demand = [DemandItem(padd=3, sector="petchem_cracker", volume_kbd=15)]
    weeks = padd_balance(3, [refinery], AXIS, book, [], flows, demand)

    w1 = weeks[0]
    cdu_net = 200 * 0.95 * (0.05 + 0.12)
    reformer_pull = 50 * 0.95
    assert w1.supply_kbd == pytest.approx(cdu_net - reformer_pull)
    assert w1.flows_kbd == pytest.approx(-6)  # -10 export + 4 import
    assert w1.demand_kbd == pytest.approx(15)
    assert w1.balance_kbd == pytest.approx(w1.supply_kbd - 6 - 15)
    # flows dated in week 1 do not leak into week 2
    assert weeks[1].flows_kbd == 0.0


def test_at_risk_flows_into_padd(refinery, book):
    outages = [make_outage(offline_pct=100)]  # second week
    weeks = padd_balance(3, [refinery], AXIS, book, outages, [], [])
    assert weeks[0].at_risk_kbd == pytest.approx(0.0)
    assert weeks[1].at_risk_kbd > 0


def test_us_balance_groups_by_padd(refinery, padd1_refinery, book):
    balances = us_balance([refinery, padd1_refinery], AXIS, book, [], [], [])
    assert set(balances) == {1, 3}
    assert balances[1][0].supply_kbd == pytest.approx(100 * 0.90 * 0.15)


def test_demand_only_padd_still_appears(refinery, book):
    demand = [DemandItem(padd=2, sector="gasoline_blending", volume_kbd=30)]
    balances = us_balance([refinery], AXIS, book, [], [], demand)
    assert 2 in balances
    assert balances[2][0].balance_kbd == pytest.approx(-30)
