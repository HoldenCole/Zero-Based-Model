"""Ingest the desk's 'Estimated Refinery Outputs' workbook (2024 yields).

    python -m naphtha_model ingest-yields <file.xlsx>

Does two things:
1. Writes data/reference/refinery_yields_2024.csv — per-refinery 2024 product
   yields (% of crude), naphtha included. This is NET merchant naphtha
   (after reformer/isom pull), matching the model's net-naphtha concept.
2. Merges every refinery into data/reference/refineries.csv. Existing rows
   (matched by city + PADD + operator token) keep their id, name, capacity,
   units and notes; new rows arrive with crude_capacity_kbd=0 ("yield-mode")
   until the capacity sheet lands.

Refineries that have no unit detail get a synthetic CRUDE-EST unit at load
time (see loaders.py): net naphtha = crude capacity x utilization x 2024
yield, split across cuts by assumptions/global.yaml yield_mode.cut_shares.
"""

from __future__ import annotations

import csv
import re
from pathlib import Path

import openpyxl

from .config import DATA_DIR

ROMAN_PADD = {"PADD I": 1, "PADD II": 2, "PADD III": 3, "PADD IV": 4, "PADD V": 5}

YIELD_COLS = [
    ("gasoil_diesel_pct", 8),
    ("gasoline_pct", 9),
    ("hfo_pct", 10),
    ("kero_jet_pct", 11),
    ("lpg_pct", 12),
    ("naphtha_pct", 13),
    ("total_pct", 14),
]


def _blank(v) -> bool:
    return v is None or str(v).strip() == ""


def _slug(operator: str, city: str) -> str:
    op_token = re.sub(r"[^A-Za-z0-9]", "", operator.split()[0]).upper()
    city_token = re.sub(r"[^A-Za-z0-9]+", "_", city.strip()).upper().strip("_")
    return f"{op_token}_{city_token}"


def parse_yields_workbook(path: Path) -> list[dict]:
    """One dict per refinery. Handles the pivot-style layout: region/country/
    PADD/state/city forward-filled, extra owner rows (blank operator) folded
    into the owner string of the refinery above."""
    ws = openpyxl.load_workbook(path, data_only=True).active
    refineries: list[dict] = []
    cur = [None] * 5  # region, country, padd, state, city
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] == "Region" or all(_blank(v) for v in row):
            continue
        for i in range(5):
            if not _blank(row[i]):
                cur[i] = str(row[i]).strip()
        operator = row[5]
        if _blank(operator):
            # extra owner row for the refinery above (JV)
            if not _blank(row[6]) and refineries and str(row[6]).strip() != "Owner":
                refineries[-1]["owners"].append((str(row[6]).strip(), row[7]))
            continue
        if cur[2] not in ROMAN_PADD:
            raise ValueError(f"unrecognised PADD {cur[2]!r} for {operator}")
        rec = {
            "padd": ROMAN_PADD[cur[2]],
            "state": cur[3] or "",
            "city": cur[4] or "",
            "operator": str(operator).strip(),
            "owners": [(str(row[6]).strip(), row[7])] if not _blank(row[6]) else [],
        }
        for name, col in YIELD_COLS:
            rec[name] = round(float(row[col] or 0.0), 4)
        refineries.append(rec)
    return refineries


def _owner_string(owners: list[tuple[str, object]]) -> str:
    if not owners:
        return ""
    if len(owners) == 1 and float(owners[0][1] or 100) == 100:
        return owners[0][0]
    return " / ".join(f"{name} ({float(share or 0):.0f}%)" for name, share in owners)


def _read_registry(path: Path) -> list[dict]:
    if not path.exists():
        return []
    with path.open(newline="", encoding="utf-8") as fh:
        return list(csv.DictReader(fh))


def _match_existing(rec: dict, registry: list[dict]) -> dict | None:
    """Same city + PADD, and the operator's first token appears in the
    existing name or owner — keeps hand-curated ids like MOTIVA_PAR."""
    token = rec["operator"].split()[0].lower()
    for row in registry:
        if (
            row["city"].strip().lower() == rec["city"].strip().lower()
            and int(row["padd"]) == rec["padd"]
            and (token in row["name"].lower() or token in row["owner"].lower())
        ):
            return row
    return None


def ingest_yields(xlsx_path: Path, data_dir: Path = DATA_DIR) -> dict:
    records = parse_yields_workbook(Path(xlsx_path))
    registry_path = data_dir / "reference" / "refineries.csv"
    registry = _read_registry(registry_path)
    # match only against rows that existed before this ingestion, and let
    # each pre-existing row be claimed at most once
    preexisting = list(registry)
    claimed: set[str] = set()
    matched = added = 0
    used_ids = {row["refinery_id"] for row in registry}
    yields_rows = []

    for rec in records:
        existing = _match_existing(rec, preexisting)
        if existing is not None and existing["refinery_id"] in claimed:
            existing = None
        if existing is not None:
            claimed.add(existing["refinery_id"])
            rid = existing["refinery_id"]
            matched += 1
        else:
            rid = _slug(rec["operator"], rec["city"])
            base, n = rid, 2
            while rid in used_ids:
                rid = f"{base}_{n}"
                n += 1
            used_ids.add(rid)
            registry.append(
                {
                    "refinery_id": rid,
                    "name": f"{rec['operator']} {rec['city']}",
                    "owner": _owner_string(rec["owners"]) or rec["operator"],
                    "city": rec["city"],
                    "state": rec["state"],
                    "padd": str(rec["padd"]),
                    "region": "US",
                    "crude_capacity_kbd": "0",
                    "status": "operating",
                    "notes": "yield-mode; awaiting capacity sheet",
                }
            )
            added += 1
        yields_rows.append(
            {
                "refinery_id": rid,
                "padd": rec["padd"],
                "state": rec["state"],
                "city": rec["city"],
                "operator": rec["operator"],
                "owners": _owner_string(rec["owners"]),
                **{name: rec[name] for name, _ in YIELD_COLS},
            }
        )

    reg_fields = ["refinery_id", "name", "owner", "city", "state", "padd",
                  "region", "crude_capacity_kbd", "status", "notes"]
    with registry_path.open("w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=reg_fields)
        w.writeheader()
        for row in registry:
            w.writerow({k: row.get(k, "") for k in reg_fields})

    yields_path = data_dir / "reference" / "refinery_yields_2024.csv"
    y_fields = ["refinery_id", "padd", "state", "city", "operator", "owners"] + [
        name for name, _ in YIELD_COLS
    ]
    with yields_path.open("w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=y_fields)
        w.writeheader()
        w.writerows(yields_rows)

    return {"refineries": len(records), "matched_existing": matched,
            "added_to_registry": added, "yields_csv": str(yields_path),
            "registry_csv": str(registry_path)}


# ------------------------------------------------------------------ capacity
#
# EA site-level monthly nameplate capacity (kb/d), 2023-2026. Site names are
# city-based with parenthetical disambiguators; the alias map below resolves
# every site that automatic matching (exact city, then operator hint, then
# name-contains) cannot. None = site is not in the registry (shut before the
# 2024 yields vintage, or a splitter/topping asset not yet modelled) — its
# monthly series is still written to site_capacity_monthly.csv for reference.

SITE_ALIASES: dict[str, str | None] = {
    "Anacortes": "MARATHON_ANACORTES",            # ex-Tesoro/Andeavor plant
    "Anacortes (Puget Sound)": "HF_ANACORTES",    # ex-Shell Puget Sound
    "Bakersfield": None,                          # ex-Big West, converted/shut
    "Bakersfield (SJR)": "SAN_BAKERSFIELD",       # San Joaquin Refining
    "Bayway": "PHILLIPS_LINDEN",                  # P66 Bayway, Linden NJ
    "Billings": "PAR_BILLINGS",                   # ex-ExxonMobil Billings
    "Cherry Point": "BP_FERNDALE",                # BP Cherry Point
    "Corpus Christi (Bill Greehey)": "VALERO_CORPUS_CHRISTI",
    "Corpus Christi (Trigeant)": None,            # asphalt, not in registry
    "Corpus Christi Splitter (Magellan)": None,   # splitter, not in registry yet
    "Channelview Splitter": None,                 # splitter, not in registry yet
    "Eagle Springs (Ely)": "FORELAND_TONOPAH",    # Foreland, NV
    "Ferndale": "PHILLIPS_FERNDALE",
    "Galena Park": "KINDER_HOUSTON",              # Kinder Morgan splitters
    "Galveston Bay": "MPC_GALV_BAY",
    "Houston Splitter": None,                     # splitter, not in registry yet
    "Kapolei (Ewa Beach -East Plant)": "PAR_EWA_BEACH",
    "Kenai": "MARATHON_NIKISKI",
    "Kuparuk Topping Unit": "CONOCOPHILLIPS_ANCHORAGE",
    "Lake Charles (Pelican)": None,
    "Lake Charles (Westlake - P66)": "PHILLIPS_WESTLAKE",
    "Los Angeles (Wilmington and Carson)": "PHILLIPS_LOS_ANGELES",  # shut 2025
    "McKee (Sunray)": "VALERO_SUNRAY",
    "Navajo (Artesia)": "HF_ARTESIA",
    "North Pole (FHR)": None,                     # FHR North Pole, shut
    "North Salt Lake": "BIG_NORTH_SLATERVILLE_CANAL",  # Big West Oil
    "Paulsboro": None,                            # asphalt plant, not PBF
    "Paulsboro (PBF)": "PBF_PAULSBORO",
    "Pine Bend (Rosemount)": "FLINT_ROSEMOUNT",
    "Port Arthur (Total)": "TOTAL_PORT_ARTHUR",
    "Saraland": "VERTEX_MOBILE",                  # Vertex Saraland (Mobile AL)
    "Sinclair, WY": "SINCLAIR_SINCLAIR",
    "Wilmington (Marathon)": "MARATHON_LOS_ANGELES",
    "Wilmington (Valero)": "VALERO_LOS_ANGELES",
    "Wilmington Asphalt Refinery": None,
    "Woods Cross": "HF_WOODS_CROSS",
    # shut sites with no registry row (pre-2024 closures):
    "Alliance (Belle Chasse)": None, "Bloomfield": None, "Brownsville": None,
    "Cheyenne": None, "Davis": None, "Dickinson": None,
    "Gallup": None, "Golden Eagle (Martinez)": None, "Paramount": None,
    "Par West (Barbers Point, ex-Island Energy)": None, "Perth Amboy": None,
    "Philadelphia Complex": None, "Rodeo": None, "Santa Maria": None,
    "Santa Maria (Arroyo Grande)": None, "Savannah": None, "Southland": None,
}

# registry rows that are duplicate/component entities in the EA yields data;
# their capacity is carried on the primary row
COMPONENT_NOTES = {
    "MARATHON_TEXAS_CITY_GALVESTON_BAY":
        "EA yields component of the Galveston Bay complex; "
        "capacity carried on MPC_GALV_BAY",
}

_STATE_SUFFIX = re.compile(
    r",\s*(mt|wv|tx|la|ca|wy|ut|nd|mn|ks|ok|pa|nj|wa|ak|hi|nm|co|nv|mi|oh|il|in|"
    r"ky|tn|wi|al|ar|ms|de|ga|sc|va|md|mo)$"
)

_HINT_ALIASES = {"p66": "phillips", "cvx": "chevron", "mpc": "marathon",
                 "fhr": "flint", "basf total": "basf"}


def _norm(s: str) -> str:
    s = _STATE_SUFFIX.sub("", s.lower().strip())
    s = re.sub(r"[^a-z0-9 ]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def match_site(site: str, registry: list[dict]) -> str | None:
    """Resolve an EA site name to a registry refinery_id (None = no match)."""
    if site in SITE_ALIASES:
        return SITE_ALIASES[site]
    m = re.match(r"^(.*?)\s*(?:\((.*)\))?$", site)
    base, hint = _norm(m.group(1)), _norm(m.group(2) or "")
    hint = _HINT_ALIASES.get(hint, hint)

    exact = [r for r in registry if _norm(r["city"]) == base]
    if len(exact) == 1:
        return exact[0]["refinery_id"]
    cands = exact or [
        r for r in registry
        if base in _norm(r["name"]) or (_norm(r["city"]) and _norm(r["city"]) in base)
    ]
    if len(cands) > 1 and hint:
        hinted = [r for r in cands
                  if hint.split()[0] in _norm(r["name"] + " " + r["owner"])]
        if len(hinted) == 1:
            return hinted[0]["refinery_id"]
        cands = hinted or cands
    if len(cands) == 1:
        return cands[0]["refinery_id"]
    return None


def ingest_capacity(
    csv_path: Path, data_dir: Path = DATA_DIR, as_of: str = "2026-07"
) -> dict:
    """Load EA site nameplate capacities: stamp the as-of month into the
    registry (multi-site refineries are summed), write the full monthly
    series to data/reference/site_capacity_monthly.csv."""
    registry_path = data_dir / "reference" / "refineries.csv"
    registry = _read_registry(registry_path)

    with Path(csv_path).open(newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))

    sites = sorted({r["refinery"] for r in rows})
    site_to_rid = {s: match_site(s, registry) for s in sites}

    monthly_path = data_dir / "reference" / "site_capacity_monthly.csv"
    with monthly_path.open("w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["site", "refinery_id", "month", "capacity_kbd"])
        for r in sorted(rows, key=lambda r: (r["refinery"], r["date"])):
            w.writerow([r["refinery"], site_to_rid[r["refinery"]] or "",
                        r["date"][:7], r["value"]])

    # as-of capacity per refinery (sum of its matched sites)
    per_rid: dict[str, float] = {}
    for r in rows:
        if r["date"][:7] != as_of:
            continue
        rid = site_to_rid[r["refinery"]]
        if rid:
            per_rid[rid] = per_rid.get(rid, 0.0) + float(r["value"])

    stamped = 0
    for row in registry:
        rid = row["refinery_id"]
        if rid in per_rid:
            row["crude_capacity_kbd"] = f"{per_rid[rid]:g}"
            row["status"] = "operating" if per_rid[rid] > 0 else "shut"
            row["notes"] = f"capacity: EA nameplate {as_of}"
            stamped += 1
        if rid in COMPONENT_NOTES:
            row["notes"] = COMPONENT_NOTES[rid]

    reg_fields = ["refinery_id", "name", "owner", "city", "state", "padd",
                  "region", "crude_capacity_kbd", "status", "notes"]
    with registry_path.open("w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=reg_fields)
        w.writeheader()
        for row in registry:
            w.writerow({k: row.get(k, "") for k in reg_fields})

    unmatched = {s: v for s, v in (
        (s, sum(float(r["value"]) for r in rows
                if r["refinery"] == s and r["date"][:7] == as_of))
        for s in sites if site_to_rid[s] is None
    ) if v > 0}
    unstamped = [row["refinery_id"] for row in registry
                 if row["refinery_id"] not in per_rid
                 and row["refinery_id"] not in COMPONENT_NOTES]
    return {
        "sites": len(sites),
        "matched_sites": sum(1 for v in site_to_rid.values() if v),
        "refineries_stamped": stamped,
        "us_total_kbd": round(sum(per_rid.values()), 1),
        "unmatched_active_sites": unmatched,
        "registry_without_capacity": unstamped,
        "monthly_csv": str(monthly_path),
        "as_of": as_of,
    }
