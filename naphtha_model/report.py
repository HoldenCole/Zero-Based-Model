"""Rendering: the refinery "box", balance tables, and Excel export."""

from __future__ import annotations

from datetime import date
from pathlib import Path

import pandas as pd

from .balance import PaddWeek
from .engine import balance_at_risk_kbd, refinery_day, weekly_net_kbd
from .loaders import ModelData
from .schema import Outage, Refinery


def _fmt_week(d: date) -> str:
    return f"{d.month}/{d.day}"


def render_refinery_box(
    refinery: Refinery,
    data: ModelData,
    axis: list[date],
) -> str:
    """The whiteboard box: units, capacities, utilization, yields, production,
    total naphtha, model cut, forward weekly strip, and active outages."""
    book = data.book
    today = axis[0]
    snap = refinery_day(refinery, today, book, data.outages)
    cuts = book.cuts

    lines: list[str] = []
    title = (
        f"{refinery.name} [{refinery.refinery_id}]  —  PADD {refinery.padd}, "
        f"{refinery.city}, {refinery.state}  —  {refinery.owner}"
    )
    width = max(96, len(title) + 4)
    bar = "=" * width
    lines += [bar, title, bar]
    lines.append(
        f"Crude capacity: {refinery.crude_capacity_kbd:,.0f} kbd   "
        f"Status: {refinery.status}   As of: {today.isoformat()}"
    )
    if refinery.notes:
        lines.append(f"Notes: {refinery.notes}")
    lines.append("")

    # Unit table
    cut_hdr = "  ".join(f"{c:>7}" for c in cuts)
    lines.append(
        f"{'Unit':<10} {'Type':<12} {'Cap kbd':>8} {'Offline':>8} {'Util':>6} "
        f"{'Thruput':>8}  yields: {cut_hdr}   source"
    )
    lines.append("-" * width)
    for ud in snap.unit_days:
        yields_str = "  ".join(f"{ud.yields[c] * 100:>6.1f}%" for c in cuts)
        srcs = {ud.utilization_source, *ud.yield_sources.values()}
        src = "MIXED" if len(srcs) > 1 else next(iter(srcs))
        lines.append(
            f"{ud.unit.unit_id:<10} {ud.unit.unit_type:<12} "
            f"{ud.unit.capacity_kbd:>8.0f} {ud.offline_frac * 100:>7.0f}% "
            f"{ud.utilization * 100:>5.0f}% {ud.throughput_kbd:>8.1f}  "
            f"         {yields_str}   {src}"
        )
    lines.append("-" * width)

    # Production by cut (the "Produce" column on the whiteboard)
    prod = snap.production_kbd
    prod_str = "   ".join(f"{c}: {prod.get(c, 0.0):+.1f}" for c in cuts)
    lines.append(f"Production kbd (signed, consumers negative):   {prod_str}")
    lines.append(
        f"Gross naphtha: {snap.gross_kbd:.1f} kbd   "
        f"Consumed internally (reformer/isom feed): {snap.consumed_kbd:.1f} kbd   "
        f"=> NET NAPHTHA: {snap.net_kbd:.1f} kbd"
    )
    lines.append("Model cut: sum over units of throughput x yield (per cut, signed)")
    lines.append("")

    # Forward weekly strip
    weekly = weekly_net_kbd(refinery, axis, book, data.outages)
    at_risk = balance_at_risk_kbd(refinery, axis, book, data.outages)
    strip = " | ".join(
        f"{_fmt_week(w)} {weekly[w]:.1f}{'*' if at_risk[w] > 0.05 else ''}" for w in axis
    )
    lines.append(f"Forward net naphtha (kbd, weekly avg):  {strip}")
    lines.append("(* = week impacted by an outage)")

    # Outages touching this refinery inside the window
    ref_outages = [
        o
        for o in data.outages
        if o.refinery_id == refinery.refinery_id and o.end >= axis[0]
    ]
    if ref_outages:
        lines.append("")
        lines.append("Outages:")
        for o in sorted(ref_outages, key=lambda o: o.start):
            unit = o.unit_id or "WHOLE REFINERY"
            lines.append(
                f"  - [{o.outage_type}] {unit} {o.start.isoformat()} -> {o.end.isoformat()} "
                f"({o.offline_pct:.0f}% offline)"
                + (f"  conf: {o.confidence}" if o.confidence else "")
                + (f"  src: {o.source}" if o.source else "")
                + (f"  — {o.notes}" if o.notes else "")
            )

    # Linked market intel
    notes = [n for n in data.intel if n.refinery_id == refinery.refinery_id]
    if notes:
        lines.append("")
        lines.append("Market intel:")
        for n in sorted(notes, key=lambda n: n.note_date, reverse=True):
            impact = f" ({n.impact_kbd:+.0f} kbd)" if n.impact_kbd else ""
            lines.append(f"  - {n.note_date.isoformat()}: {n.headline}{impact}")

    lines.append(bar)
    return "\n".join(lines)


def render_padd_balance(padd: int, weeks: list[PaddWeek]) -> str:
    lines = [
        f"PADD {padd} naphtha balance (kbd, weekly averages)",
        f"{'Week':<8} {'Supply':>9} {'Flows':>8} {'Demand':>9} {'Balance':>9} {'At risk':>9}",
        "-" * 56,
    ]
    for w in weeks:
        lines.append(
            f"{_fmt_week(w.week_start):<8} {w.supply_kbd:>9.1f} {w.flows_kbd:>8.1f} "
            f"{w.demand_kbd:>9.1f} {w.balance_kbd:>+9.1f} {w.at_risk_kbd:>9.1f}"
        )
    lines.append("-" * 56)
    lines.append("Balance: + = length (barrels looking for a home), - = short (must pull barrels)")
    return "\n".join(lines)


# ---------------------------------------------------------------- Excel export


def export_workbook(
    data: ModelData,
    axis: list[date],
    balances: dict[int, list[PaddWeek]],
    out_path: Path,
) -> Path:
    """Dump the whole model state to a multi-tab Excel workbook."""
    book = data.book

    refinery_rows = [
        {
            "refinery_id": r.refinery_id,
            "name": r.name,
            "owner": r.owner,
            "city": r.city,
            "state": r.state,
            "padd": r.padd,
            "crude_capacity_kbd": r.crude_capacity_kbd,
            "status": r.status,
            "notes": r.notes,
        }
        for r in data.refineries
    ]

    unit_rows = []
    for r in data.refineries:
        snap = refinery_day(r, axis[0], book, data.outages)
        for ud in snap.unit_days:
            row = {
                "refinery_id": r.refinery_id,
                "unit_id": ud.unit.unit_id,
                "unit_type": ud.unit.unit_type,
                "capacity_kbd": ud.unit.capacity_kbd,
                "offline_pct": round(ud.offline_frac * 100, 1),
                "utilization": ud.utilization,
                "utilization_source": ud.utilization_source,
                "throughput_kbd": round(ud.throughput_kbd, 2),
            }
            for cut in book.cuts:
                row[f"yield_{cut}"] = ud.yields[cut]
                row[f"prod_{cut}_kbd"] = round(ud.production_kbd[cut], 2)
            unit_rows.append(row)

    forward_rows = []
    for r in data.refineries:
        weekly = weekly_net_kbd(r, axis, book, data.outages)
        at_risk = balance_at_risk_kbd(r, axis, book, data.outages)
        row: dict = {"refinery_id": r.refinery_id, "padd": r.padd}
        for w in axis:
            row[w.isoformat()] = round(weekly[w], 2)
            row[f"{w.isoformat()}_at_risk"] = round(at_risk[w], 2)
        forward_rows.append(row)

    balance_rows = [
        {
            "padd": p,
            "week_start": w.week_start.isoformat(),
            "supply_kbd": round(w.supply_kbd, 2),
            "flows_kbd": round(w.flows_kbd, 2),
            "demand_kbd": round(w.demand_kbd, 2),
            "balance_kbd": round(w.balance_kbd, 2),
            "at_risk_kbd": round(w.at_risk_kbd, 2),
        }
        for p, weeks in balances.items()
        for w in weeks
    ]

    outage_rows = [
        {
            "outage_id": o.outage_id,
            "refinery_id": o.refinery_id,
            "unit_id": o.unit_id,
            "start": o.start.isoformat(),
            "end": o.end.isoformat(),
            "offline_pct": o.offline_pct,
            "type": o.outage_type,
            "source": o.source,
            "confidence": o.confidence,
            "notes": o.notes,
        }
        for o in data.outages
    ]

    out_path = Path(out_path)
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        pd.DataFrame(refinery_rows).to_excel(writer, sheet_name="Refineries", index=False)
        pd.DataFrame(unit_rows).to_excel(writer, sheet_name="Units", index=False)
        pd.DataFrame(forward_rows).to_excel(writer, sheet_name="Forward_Net_Naphtha", index=False)
        pd.DataFrame(balance_rows).to_excel(writer, sheet_name="PADD_Balance", index=False)
        pd.DataFrame(outage_rows).to_excel(writer, sheet_name="Outages", index=False)
    return out_path
