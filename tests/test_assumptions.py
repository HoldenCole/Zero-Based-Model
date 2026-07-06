from datetime import date

from conftest import make_book
from naphtha_model.schema import Override

DAY = date(2026, 7, 6)


def test_padd_beats_global(refinery, padd1_refinery, book):
    cdu = refinery.unit("CDU-1")
    # PADD 3 has overrides
    assert book.utilization(refinery, cdu, DAY).value == 0.95
    assert book.yield_for(refinery, cdu, "HVN", DAY).value == 0.12
    # PADD 3 has no LVN CDU override -> falls to global
    assert book.yield_for(refinery, cdu, "LVN", DAY).value == 0.05
    # PADD 1 has nothing -> global everywhere
    p1_cdu = padd1_refinery.unit("CDU-1")
    assert book.utilization(padd1_refinery, p1_cdu, DAY).value == 0.90
    assert book.yield_for(padd1_refinery, p1_cdu, "HVN", DAY).value == 0.10


def test_override_beats_padd(refinery):
    book = make_book(
        [Override(refinery_id="TEST_REF", field_name="utilization", value=0.5)]
    )
    cdu = refinery.unit("CDU-1")
    res = book.utilization(refinery, cdu, DAY)
    assert res.value == 0.5
    assert res.source.startswith("override")


def test_unit_override_beats_refinery_override(refinery):
    book = make_book(
        [
            Override(refinery_id="TEST_REF", field_name="utilization", value=0.5),
            Override(
                refinery_id="TEST_REF", unit_id="CDU-1", field_name="utilization", value=0.7
            ),
        ]
    )
    assert book.utilization(refinery, refinery.unit("CDU-1"), DAY).value == 0.7
    # the reformer only matches the refinery-wide row
    assert book.utilization(refinery, refinery.unit("REF-1"), DAY).value == 0.5


def test_dated_override_window(refinery):
    book = make_book(
        [
            Override(
                refinery_id="TEST_REF",
                field_name="yield",
                cut="HVN",
                value=0.2,
                start=date(2026, 7, 10),
                end=date(2026, 7, 20),
            )
        ]
    )
    cdu = refinery.unit("CDU-1")
    assert book.yield_for(refinery, cdu, "HVN", date(2026, 7, 9)).value == 0.12
    assert book.yield_for(refinery, cdu, "HVN", date(2026, 7, 15)).value == 0.2
    assert book.yield_for(refinery, cdu, "HVN", date(2026, 7, 21)).value == 0.12


def test_sources_are_reported(refinery, book):
    cdu = refinery.unit("CDU-1")
    assert book.utilization(refinery, cdu, DAY).source == "PADD 3 default"
    assert book.yield_for(refinery, cdu, "LVN", DAY).source == "global default"
