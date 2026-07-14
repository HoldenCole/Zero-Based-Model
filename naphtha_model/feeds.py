"""Live data feeds — the scriptable twin of the workbook's Live Feeds tab.

The workbook pulls EIA/IIR straight into cells via WEBSERVICE(); these
functions are the fallback that refresh the repo CSVs the build ingests
(useful headless, on a Mac, or when WEBSERVICE is blocked).

EIA: API v2 with a free key (eia.gov/opendata/register.php). Series IDs
below were verified against eia.gov dnav on 2026-07-14.

IIR: the desk supplies its working REST query URL + a 30-day token (the
source offline-events workbook refreshed the same way via Power Query).
We save the raw response; when it parses to the known OEV record shape we
report the fields so it can be wired into ingest_outages.
"""
from __future__ import annotations

import csv
import json
import urllib.parse
import urllib.request
from pathlib import Path

from .config import DATA_DIR

# series id -> (slug, unit) — same list the Live Feeds tab ships with
EIA_SERIES: dict[str, tuple[str, str]] = {
    "PET.MNFUPUS2.M": ("us_petchem_naphtha_product_supplied", "kbd"),
    "PET.WGIRIUS2.W": ("us_gross_inputs_to_refineries", "kbd"),
    "PET.WCRRIUS2.W": ("us_refinery_crude_runs", "kbd"),
    "PET.WOCLEUS2.W": ("us_operable_refining_capacity", "kbd"),
    "PET.WPULEUS3.W": ("us_refinery_utilization", "pct"),
}

# fields we expect in an IIR offline-event record (from the source DB)
IIR_OEV_FIELDS = {
    "offlineEventKey", "eventId", "eventType", "eventStartDate",
    "eventEndDate", "plantName", "unitTypeDesc",
}


def eia_url(series_id: str, api_key: str, length: int = 520,
            out_xml: bool = False) -> str:
    """The exact request the Live Feeds tab builds (JSON by default here)."""
    return (
        "https://api.eia.gov/v2/seriesid/" + urllib.parse.quote(series_id)
        + "?api_key=" + urllib.parse.quote(api_key)
        + "&sort[0][column]=period&sort[0][direction]=desc"
        + f"&length={length}"
        + ("&out=xml" if out_xml else "")
    )


def parse_eia(payload: dict) -> list[dict]:
    """v2 JSON -> [{period, value, units}] (newest first as requested)."""
    rows = (payload.get("response") or {}).get("data") or []
    out = []
    for row in rows:
        val = row.get("value")
        out.append({
            "period": row.get("period"),
            "value": float(val) if val not in (None, "") else None,
            "units": row.get("units") or row.get("unit") or "",
        })
    return out


def _http_json(url: str) -> dict:
    with urllib.request.urlopen(url, timeout=60) as resp:
        return json.load(resp)


def pull_eia(api_key: str, series: list[str] | None = None,
             out_csv: Path | None = None, fetch=None) -> tuple[Path, list[dict]]:
    """Fetch the series and land them in data/reference/eia_feed.csv
    (series, name, period, value, units) — newest first per series."""
    fetch = fetch or _http_json
    series = series or list(EIA_SERIES)
    out_csv = Path(out_csv or DATA_DIR / "reference" / "eia_feed.csv")
    all_rows: list[dict] = []
    for sid in series:
        slug = EIA_SERIES.get(sid, (sid.lower().replace(".", "_"), ""))[0]
        for row in parse_eia(fetch(eia_url(sid, api_key))):
            all_rows.append({"series": sid, "name": slug, **row})
    out_csv.parent.mkdir(parents=True, exist_ok=True)
    with out_csv.open("w", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=["series", "name", "period",
                                           "value", "units"])
        w.writeheader()
        w.writerows(all_rows)
    return out_csv, all_rows


def _http_raw(url: str, headers: dict | None = None) -> bytes:
    req = urllib.request.Request(url, headers=headers or {})
    with urllib.request.urlopen(req, timeout=120) as resp:
        return resp.read()


def pull_iir(url: str, token: str, out_json: Path | None = None,
             auth: str = "query", fetch=None) -> tuple[Path, dict]:
    """Fetch the desk's IIR query and save the raw response.

    auth='query'  -> appends &token=... (what WEBSERVICE on the tab does)
    auth='bearer' -> sends Authorization: Bearer <token> instead
    Returns (path, summary) where summary says whether the payload parsed
    and whether it looks like OEV records (so it can feed ingest_outages).
    """
    fetch = fetch or _http_raw
    if auth == "query":
        full = url + ("&" if "?" in url else "?") + "token=" + \
            urllib.parse.quote(token)
        raw = fetch(full)
    else:
        raw = fetch(url, {"Authorization": f"Bearer {token}"})
    out_json = Path(out_json or DATA_DIR.parent / "data" / "raw"
                    / "iir_pull.json")
    out_json.parent.mkdir(parents=True, exist_ok=True)
    out_json.write_bytes(raw)

    summary: dict = {"bytes": len(raw), "records": 0, "is_oev": False,
                     "fields": []}
    try:
        payload = json.loads(raw)
    except (ValueError, UnicodeDecodeError):
        return out_json, summary
    records = payload if isinstance(payload, list) else next(
        (v for v in payload.values() if isinstance(v, list)), []) \
        if isinstance(payload, dict) else []
    summary["records"] = len(records)
    if records and isinstance(records[0], dict):
        fields = set(records[0])
        summary["fields"] = sorted(fields)[:40]
        summary["is_oev"] = len(IIR_OEV_FIELDS & fields) >= 4
    return out_json, summary
