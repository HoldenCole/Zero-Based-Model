"""Model-wide constants and reference vocabularies."""

from __future__ import annotations

from pathlib import Path

# Repository root (this file lives in <root>/naphtha_model/).
ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"

PADDS: dict[int, str] = {
    1: "East Coast",
    2: "Midwest",
    3: "Gulf Coast",
    4: "Rocky Mountains",
    5: "West Coast",
}

REGIONS = ["US", "EUROPE", "APAC"]

# Process units the model understands. Yields for each are set in
# data/assumptions/global.yaml (and can be overridden per PADD / refinery).
UNIT_TYPES = [
    "CDU",          # crude distillation — primary naphtha source
    "VDU",          # vacuum distillation
    "FCC",          # fluid catalytic cracker — cat naphtha
    "COKER",        # delayed coker — coker naphtha
    "HYDROCRACKER", # hydrocracker — HC naphtha
    "REFORMER",     # catalytic reformer — CONSUMES heavy naphtha
    "ISOM",         # isomerization — CONSUMES light naphtha
    "ALKY",         # alkylation (no straight naphtha make; kept for the box)
    "SPLITTER",     # condensate splitter — heavy naphtha maker
    "NHT",          # naphtha hydrofiner/hydrotreater — pass-through in the
                    # naphtha path (CDU -> OH dist -> NHT -> reformer/mogas)
]

# Units whose yields are expected to be negative (naphtha consumers).
CONSUMER_UNIT_TYPES = {"REFORMER", "ISOM"}

OUTAGE_TYPES = ["planned", "unplanned"]
FLOW_DIRECTIONS = ["import", "export", "transfer_in", "transfer_out"]

DAYS_PER_WEEK = 7
DEFAULT_FORWARD_WEEKS = 8
