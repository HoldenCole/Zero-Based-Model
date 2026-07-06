"""Assumption resolution: manual override -> PADD default -> global default.

Every lookup also returns *where* the number came from, so reports can show
which figures are house defaults and which are trader-set intel.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date

from .schema import Override, ProcessUnit, Refinery

GLOBAL_SOURCE = "global default"


@dataclass(frozen=True)
class Resolved:
    value: float
    source: str  # "override (<note/source>)" | "PADD <n> default" | "global default"


class AssumptionBook:
    """Resolves utilization and yields for any refinery unit on any date.

    global_cfg: parsed data/assumptions/global.yaml
        {"cuts": [...], "utilization": {"default": f},
         "yields": {UNIT_TYPE: {CUT: fraction}}}
    padd_cfg: parsed data/assumptions/padd_overrides.yaml
        {padd:int: {"utilization": f, "yields": {UNIT_TYPE: {CUT: f}}}}
    overrides: list[Override] from data/overrides/refinery_overrides.csv
    """

    def __init__(self, global_cfg: dict, padd_cfg: dict, overrides: list[Override]):
        self.global_cfg = global_cfg
        self.padd_cfg = padd_cfg or {}
        self.overrides = overrides or []
        self.cuts: list[str] = list(global_cfg.get("cuts", []))
        if not self.cuts:
            raise ValueError("global assumptions must define at least one cut")

    # -- utilization ---------------------------------------------------

    def utilization(self, refinery: Refinery, unit: ProcessUnit, day: date) -> Resolved:
        ov = self._find_override(refinery, unit, day, "utilization")
        if ov is not None:
            return Resolved(ov.value, _override_label(ov))
        padd = self.padd_cfg.get(refinery.padd, {})
        if "utilization" in padd:
            return Resolved(float(padd["utilization"]), f"PADD {refinery.padd} default")
        return Resolved(float(self.global_cfg["utilization"]["default"]), GLOBAL_SOURCE)

    # -- yields ----------------------------------------------------------

    def yield_for(self, refinery: Refinery, unit: ProcessUnit, cut: str, day: date) -> Resolved:
        ov = self._find_override(refinery, unit, day, "yield", cut=cut)
        if ov is not None:
            return Resolved(ov.value, _override_label(ov))
        padd_yields = self.padd_cfg.get(refinery.padd, {}).get("yields", {})
        if cut in padd_yields.get(unit.unit_type, {}):
            return Resolved(
                float(padd_yields[unit.unit_type][cut]), f"PADD {refinery.padd} default"
            )
        global_yields = self.global_cfg.get("yields", {})
        value = float(global_yields.get(unit.unit_type, {}).get(cut, 0.0))
        return Resolved(value, GLOBAL_SOURCE)

    def yields(self, refinery: Refinery, unit: ProcessUnit, day: date) -> dict[str, Resolved]:
        return {cut: self.yield_for(refinery, unit, cut, day) for cut in self.cuts}

    # -- internals ---------------------------------------------------------

    def _find_override(
        self,
        refinery: Refinery,
        unit: ProcessUnit,
        day: date,
        field_name: str,
        cut: str = "",
    ) -> Override | None:
        """Most specific active override wins: unit-level beats refinery-level;
        among equals, the last row in the file wins (latest entry)."""
        best: Override | None = None
        best_rank = -1
        for i, ov in enumerate(self.overrides):
            if ov.field_name != field_name:
                continue
            if field_name == "yield" and ov.cut != cut:
                continue
            if not ov.targets(refinery.refinery_id, unit.unit_id):
                continue
            if not ov.active_on(day):
                continue
            rank = (2 if ov.unit_id else 1) * 1_000_000 + i
            if rank > best_rank:
                best, best_rank = ov, rank
        return best


def _override_label(ov: Override) -> str:
    detail = ov.source or ov.notes or "manual"
    return f"override ({detail})"
