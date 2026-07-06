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
