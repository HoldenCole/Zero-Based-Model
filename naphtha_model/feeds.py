"""Live data feeds — the scriptable twin of the workbook's Live Feeds tab.

The workbook pulls EIA straight into cells via WEBSERVICE(); these
functions are the fallback that refresh the repo CSVs the build ingests
(useful headless, on a Mac, or when WEBSERVICE is blocked) — and the ONLY
path for IIR (POST-only API) and EA (JSON-only responses).

EIA: API v2 with a free key (eia.gov/opendata/register.php). Series IDs
below were verified against eia.gov dnav on 2026-07-14.

EA:  OilX/Energy Aspects REST API (developer.energyaspects.com), UUID key
in the api_key query param. /balances/country/pop carries the US NAPHTHA
monthly balance (same FlowBreakdown codes the workbook already uses).

IIR: api.industrialinfo.com/idb/v2.6 — Bearer token (30-day life),
POST-only. Endpoints recovered from the desk workbook's Power Query:
offlineevents/summary, units/summary, units/detail.

Keys live in the gitignored secrets.yaml at the repo root; every pull
falls back to it when no key is passed explicitly.
"""
from __future__ import annotations

import csv
import json
import urllib.parse
import urllib.request
from pathlib import Path

from .config import DATA_DIR


def load_secrets(path: Path | None = None) -> dict:
    """secrets.yaml at the repo root (gitignored — never committed)."""
    import yaml

    p = Path(path or DATA_DIR.parent / "secrets.yaml")
    if p.exists():
        return yaml.safe_load(p.open()) or {}
    return {}

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


def _http_raw(url: str, headers: dict | None = None,
              method: str = "GET") -> bytes:
    hdrs = dict(headers or {})
    if method == "POST":
        # IIR 500s on urllib's default form-encoded type for empty bodies
        hdrs.setdefault("Content-Type", "application/json")
    req = urllib.request.Request(
        url, data=b"" if method == "POST" else None,
        headers=hdrs, method=method)
    with urllib.request.urlopen(req, timeout=180) as resp:
        return resp.read()


# ------------------------------------------------------------------ EA

EA_BASE = "https://api.energyaspects.com/oilx/v2"
# columns of data/reference/us_naphtha_balance_monthly.csv
EA_BAL_FIELDS = ["month", "flow", "kbd", "source"]


def ea_pop_url(api_key: str, product: str = "NAPHTHA", country: str = "US",
               token: str | None = None) -> str:
    """Country-level product balance, latest snapshot of every series
    (marginal_range from the beginning of data to now)."""
    url = (f"{EA_BASE}/balances/country/pop?api_key={urllib.parse.quote(api_key)}"
           f"&product={product}&country={country}&marginal_range=1900-01-01,")
    if token:
        url += "&token=" + urllib.parse.quote(token)
    return url


def pull_ea(api_key: str | None = None, out_csv: Path | None = None,
            since: str = "2023-01", through: str | None = None,
            product: str = "NAPHTHA", country: str = "US",
            fetch=None) -> tuple[Path, dict]:
    """US naphtha monthly balance from EA -> the exact CSV the workbook's
    US Balance tab is built from (month, flow, kbd, source).

    `through` caps the horizon (default: keep every month EA returns,
    including OilX nowcast months — the source column says which is which).
    """
    api_key = api_key or load_secrets().get("ea_api_key")
    if not api_key:
        raise ValueError("no EA api key: pass api_key or set ea_api_key "
                         "in secrets.yaml")
    fetch = fetch or _http_json
    out_csv = Path(out_csv or DATA_DIR / "reference"
                   / "us_naphtha_balance_monthly.csv")
    latest: dict[tuple[str, str], dict] = {}
    token, pages = None, 0
    while True:
        payload = fetch(ea_pop_url(api_key, product, country, token))
        pages += 1
        for row in payload.get("data") or []:
            month = str(row.get("ReferenceDate", ""))[:7]
            flow = row.get("FlowBreakdown", "")
            if not month or not flow or month < since:
                continue
            if through and month > through:
                continue
            # marginal_range should yield one row per series point; if the
            # API ever returns several validity intervals, latest During wins
            key = (month, flow)
            if key not in latest or str(row.get("During", "")) >= str(
                    latest[key].get("During", "")):
                latest[key] = row
        token = payload.get("token")
        if not token or pages > 200:
            break
    rows = [{"month": m, "flow": f,
             "kbd": latest[(m, f)].get("ObservedValue"),
             "source": latest[(m, f)].get("GeneralizedSource", "")}
            for m, f in sorted(latest)]
    out_csv.parent.mkdir(parents=True, exist_ok=True)
    with out_csv.open("w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=EA_BAL_FIELDS)
        w.writeheader()
        w.writerows(rows)
    months = sorted({m for m, _ in latest})
    return out_csv, {"rows": len(rows), "pages": pages,
                     "months": f"{months[0]} -> {months[-1]}" if months else "",
                     "flows": len({f for _, f in latest})}


# ----------------------------------------------------------------- IIR

IIR_BASE = "https://api.industrialinfo.com/idb/v2.6"
IIR_OEV_QUERY = (
    f"{IIR_BASE}/offlineevents/summary"
    "?eventStatusDesc=Ongoing&eventStatusDesc=Future&eventStatusDesc=Past"
    "&industryCodeDesc=Petroleum+Refining+%28HPI%29"
)


def _flatten_oev(rec: dict) -> dict:
    """Nest -> dotted keys, matching the desk xlsx export's column names."""
    out = dict(rec)
    for parent in ("offlineCapacity", "plantPhysicalAddress"):
        sub = out.pop(parent, None) or {}
        for k, v in sub.items():
            out[f"{parent}.{k}"] = v
    return out


def pull_iir(token: str | None = None, url: str = IIR_OEV_QUERY,
             country: str = "U.S.A.", out_json: Path | None = None,
             as_of: str | None = None, page_size: int = 1000,
             fetch=None, refresh_csvs: bool = True) -> tuple[Path, dict]:
    """Offline events from the IIR API (POST-only, Bearer token) -> raw
    JSON + refreshed outage_events.csv / current_outages.csv.

    Defaults reproduce the desk workbook's Power Query (all statuses,
    refining industry) but filtered server-side to the US."""
    token = token or load_secrets().get("iir_token")
    if not token:
        raise ValueError("no IIR token: pass token or set iir_token in "
                         "secrets.yaml (tokens expire every 30 days)")
    fetch = fetch or _http_raw
    hdrs = {"Authorization": f"Bearer {token}"}
    q = url
    if country:
        q += "&physicalAddressCountryName=" + urllib.parse.quote(country)
    records, offset, pages = [], 0, 0
    while True:
        raw = fetch(f"{q}&limit={page_size}&offset={offset}", hdrs, "POST")
        payload = json.loads(raw)
        batch = payload.get("offlineEvents") or []
        records.extend(batch)
        pages += 1
        offset += page_size
        if offset >= int(payload.get("totalCount") or 0) or not batch \
                or pages > 200:
            break
    out_json = Path(out_json or DATA_DIR.parent / "data" / "raw"
                    / "iir_pull.json")
    out_json.parent.mkdir(parents=True, exist_ok=True)
    out_json.write_text(json.dumps(records, indent=1))

    summary: dict = {"records": len(records), "pages": pages,
                     "is_oev": bool(records) and
                     len(IIR_OEV_FIELDS & set(records[0])) >= 4}
    if refresh_csvs and summary["is_oev"]:
        from datetime import date as _date

        from .ingest import ingest_outage_records

        flat = [_flatten_oev(r) for r in records]
        summary["ingest"] = ingest_outage_records(
            flat, as_of=as_of or _date.today().isoformat())
    return out_json, summary


# --------------------------------------------------------------- Kpler

KPLER_BASE = "https://api.kpler.com/v1"


def pull_kpler(key: str | None = None, endpoint: str = "trades",
               params: str = "products=naphtha&size=100",
               out_json: Path | None = None, fetch=None) -> tuple[Path, dict]:
    """Kpler liquids API (Basic auth with the console-issued base64 key).
    Lands the raw response for the Kpler Flows tab pipeline. NOTE: the key
    the desk supplied on 2026-07-14 fails auth (401) — likely mistyped in
    transit; paste a fresh copy into secrets.yaml to activate this."""
    key = key or load_secrets().get("kpler_key")
    if not key:
        raise ValueError("no Kpler key: set kpler_key in secrets.yaml "
                         "(copy the exact base64 string from the console)")
    fetch = fetch or _http_raw
    raw = fetch(f"{KPLER_BASE}/{endpoint}?{params}",
                {"Authorization": "Basic " + key})
    out_json = Path(out_json or DATA_DIR.parent / "data" / "raw"
                    / f"kpler_{endpoint}.json")
    out_json.parent.mkdir(parents=True, exist_ok=True)
    out_json.write_bytes(raw)
    try:
        payload = json.loads(raw)
        n = len(payload) if isinstance(payload, list) else \
            len(payload.get("data") or []) if isinstance(payload, dict) else 0
    except ValueError:
        n = None                      # some endpoints return CSV-ish text
    return out_json, {"bytes": len(raw), "records": n}
