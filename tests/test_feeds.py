"""Feed pulls: URL building and payload parsing (network mocked)."""

import json

from naphtha_model.feeds import (EIA_SERIES, eia_url, parse_eia, pull_eia,
                                 pull_iir)


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


def test_pull_iir_detects_oev_shape(tmp_path):
    records = [{"offlineEventKey": 1, "eventId": 2, "eventType": "Turnaround",
                "eventStartDate": "2026-09-01", "eventEndDate": "2026-10-01",
                "plantName": "X", "unitTypeDesc": "CDU"}]
    raw = json.dumps(records).encode()
    out, summary = pull_iir("https://x/?q=1", "TOK",
                            out_json=tmp_path / "iir.json",
                            fetch=lambda url, headers=None: raw)
    assert out.exists() and summary["records"] == 1 and summary["is_oev"]


def test_workbook_and_module_series_agree():
    from naphtha_model.workbook import SimpleWorkbook

    tab_ids = {sid for _, sid, _ in SimpleWorkbook.EIA_FEED_SERIES}
    assert tab_ids == set(EIA_SERIES)
