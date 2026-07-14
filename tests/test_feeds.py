"""Feed pulls: URL building, payload parsing, pagination (network mocked)."""

import csv
import json

from naphtha_model.feeds import (EIA_SERIES, ea_pop_url, eia_url, parse_eia,
                                 pull_ea, pull_eia, pull_iir)


def test_eia_url_shape():
    url = eia_url("PET.MNFUPUS2.M", "MYKEY")
    assert url.startswith("https://api.eia.gov/v2/seriesid/PET.MNFUPUS2.M?")
    assert "api_key=MYKEY" in url
    assert "sort[0][direction]=desc" in url
    # the workbook variant asks for XML so FILTERXML can parse it
    assert eia_url("PET.WPULEUS3.W", "K", out_xml=True).endswith("&out=xml")


def test_parse_and_pull_eia(tmp_path):
    canned = {"response": {"data": [
        {"period": "2026-03", "value": "148", "units": "MBBL/D"},
        {"period": "2026-02", "value": "160", "units": "MBBL/D"},
    ]}}
    rows = parse_eia(canned)
    assert rows[0] == {"period": "2026-03", "value": 148.0, "units": "MBBL/D"}

    out, all_rows = pull_eia("K", series=["PET.MNFUPUS2.M"],
                             out_csv=tmp_path / "feed.csv",
                             fetch=lambda url: canned)
    assert out.exists()
    text = out.read_text()
    assert "us_petchem_naphtha_product_supplied" in text
    assert len(all_rows) == 2


def test_pull_ea_paginates_and_filters(tmp_path):
    assert "marginal_range=1900-01-01," in ea_pop_url("K")
    page1 = {"data": [
        {"ReferenceDate": "2026-06-01", "FlowBreakdown": "TOTDEMO",
         "ObservedValue": 139.0, "GeneralizedSource": "OilX",
         "During": "[2026-07-01,9999-12-31)"},
        {"ReferenceDate": "2022-01-01", "FlowBreakdown": "TOTDEMO",
         "ObservedValue": 999.0, "GeneralizedSource": "Actual",
         "During": "x"},                       # before `since` -> dropped
    ], "token": "NEXT"}
    page2 = {"data": [
        {"ReferenceDate": "2026-05-01", "FlowBreakdown": "BALANCE",
         "ObservedValue": -12.0, "GeneralizedSource": "Actual",
         "During": "x"},
    ]}

    def fetch(url):
        return page2 if "token=NEXT" in url else page1

    out, summary = pull_ea("K", out_csv=tmp_path / "bal.csv", fetch=fetch)
    assert summary["pages"] == 2 and summary["rows"] == 2
    rows = list(csv.DictReader(out.open()))
    assert {(r["month"], r["flow"]) for r in rows} == {
        ("2026-06", "TOTDEMO"), ("2026-05", "BALANCE")}
    assert rows[1]["source"] == "OilX"


def test_pull_iir_paginates_no_csv(tmp_path):
    ev = {"offlineEventKey": "1**2*", "eventId": 1, "eventType": "Planned",
          "eventStartDate": "2026-09-01T05:00:00Z[UTC]",
          "eventEndDate": "2026-10-01T05:00:00Z[UTC]",
          "plantName": "X", "unitTypeDesc": "CDU",
          "offlineCapacity": {"unitCapacity": 100, "capacityOffline": 50},
          "plantPhysicalAddress": {"countryName": "U.S.A."}}
    calls = []

    def fetch(url, headers=None, method="GET"):
        calls.append((url, method))
        assert headers["Authorization"].startswith("Bearer ")
        page = {"totalCount": 3, "offlineEvents": [ev] * (2 if "offset=0" in url else 1)}
        return json.dumps(page).encode()

    out, summary = pull_iir("TOK", out_json=tmp_path / "iir.json",
                            page_size=2, fetch=fetch, refresh_csvs=False)
    assert out.exists()
    assert summary["records"] == 3 and summary["pages"] == 2
    assert summary["is_oev"]
    assert all(m == "POST" for _, m in calls)
    assert "physicalAddressCountryName=U.S.A." in calls[0][0]


def test_workbook_and_module_series_agree():
    from naphtha_model.workbook import SimpleWorkbook

    tab_ids = {sid for _, sid, _ in SimpleWorkbook.EIA_FEED_SERIES}
    assert tab_ids == set(EIA_SERIES)
