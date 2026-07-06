"""Command line interface: python -m naphtha_model <command> ..."""

from __future__ import annotations

import argparse
from datetime import date, datetime
from pathlib import Path

from .balance import padd_balance, us_balance
from .config import DEFAULT_FORWARD_WEEKS
from .engine import build_axis
from .loaders import load_all
from .report import export_workbook, render_padd_balance, render_refinery_box
from .scenario import load_scenario, run_scenario


def _parse_start(value: str | None) -> date:
    if value:
        return datetime.strptime(value, "%Y-%m-%d").date()
    return date.today()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="naphtha_model", description="Zero-based US naphtha model"
    )
    parser.add_argument("--start", help="forward window start (YYYY-MM-DD), default today")
    parser.add_argument(
        "--weeks", type=int, default=DEFAULT_FORWARD_WEEKS, help="forward window length"
    )
    sub = parser.add_subparsers(dest="command", required=True)

    sub.add_parser("list", help="list the refinery registry")

    p_box = sub.add_parser("box", help="render one refinery's box")
    p_box.add_argument("refinery_id")

    p_boxes = sub.add_parser("boxes", help="render every box (optionally one PADD)")
    p_boxes.add_argument("--padd", type=int)

    p_bal = sub.add_parser("balance", help="forward weekly balance")
    p_bal.add_argument("--padd", type=int, help="one PADD (default: all loaded)")

    p_risk = sub.add_parser("risk", help="balance at risk from outages")
    p_risk.add_argument("--padd", type=int)

    p_scn = sub.add_parser("scenario", help="run a scenario YAML against the base case")
    p_scn.add_argument("path")

    p_ing = sub.add_parser(
        "ingest-yields", help="ingest an 'Estimated Refinery Outputs' yields workbook"
    )
    p_ing.add_argument("path")

    p_cal = sub.add_parser(
        "calibrate", help="model-implied net naphtha yield vs 2024 actuals"
    )
    p_cal.add_argument("--padd", type=int)

    p_exp = sub.add_parser("export", help="build the live desk Excel workbook")
    p_exp.add_argument("--out", default="naphtha_model.xlsx")
    p_exp.add_argument(
        "--full", action="store_true",
        help="full workbook (balance, dashboard, outages, scenario toggle, checks) "
             "instead of the simple three-sheet model",
    )
    p_exp.add_argument(
        "--dump", action="store_true",
        help="flat value dump instead of the formula-driven desk workbook",
    )

    args = parser.parse_args(argv)

    if args.command == "ingest-yields":
        from .ingest import ingest_yields

        result = ingest_yields(Path(args.path))
        for k, v in result.items():
            print(f"{k}: {v}")
        return 0

    data = load_all()
    axis = build_axis(_parse_start(args.start), args.weeks)

    if args.command == "list":
        print(f"{'refinery_id':<16} {'PADD':<5} {'crude kbd':>10}  name / owner")
        print("-" * 78)
        for r in sorted(data.refineries, key=lambda r: (r.padd, -r.crude_capacity_kbd)):
            print(
                f"{r.refinery_id:<16} {r.padd:<5} {r.crude_capacity_kbd:>10,.0f}  "
                f"{r.name} — {r.owner}"
            )

    elif args.command == "box":
        print(render_refinery_box(data.refinery(args.refinery_id), data, axis))

    elif args.command == "boxes":
        refs = [r for r in data.refineries if args.padd is None or r.padd == args.padd]
        for r in sorted(refs, key=lambda r: -r.crude_capacity_kbd):
            print(render_refinery_box(r, data, axis))
            print()

    elif args.command == "balance":
        balances = us_balance(data.refineries, axis, data.book, data.outages, data.flows, data.demand)
        padds = [args.padd] if args.padd is not None else sorted(balances)
        for p in padds:
            weeks = balances.get(p) or padd_balance(
                p, data.refineries, axis, data.book, data.outages, data.flows, data.demand
            )
            print(render_padd_balance(p, weeks))
            print()

    elif args.command == "risk":
        balances = us_balance(data.refineries, axis, data.book, data.outages, data.flows, data.demand)
        padds = [args.padd] if args.padd is not None else sorted(balances)
        print(f"{'PADD':<6}" + "".join(f"{w.month}/{w.day:>2}".rjust(9) for w in axis))
        for p in padds:
            weeks = balances.get(p, [])
            print(f"{p:<6}" + "".join(f"{w.at_risk_kbd:>9.1f}" for w in weeks))
        print("\nkbd of naphtha supply lost to outages (base case minus outage case)")

    elif args.command == "scenario":
        scn = load_scenario(Path(args.path))
        result = run_scenario(scn, data, axis)
        print(f"Scenario: {scn.name} — {scn.description}\n")
        for p in sorted(result.case):
            delta = result.delta_balance(p)
            if all(abs(v) < 0.05 for v in delta.values()):
                continue
            print(f"PADD {p} balance delta vs base (kbd):")
            print("  " + " | ".join(f"{w.month}/{w.day} {v:+.1f}" for w, v in delta.items()))
            print(render_padd_balance(p, result.case[p]))
            print()

    elif args.command == "calibrate":
        from .engine import refinery_day

        day = axis[0]
        rows = []
        for r in data.refineries:
            if args.padd is not None and r.padd != args.padd:
                continue
            if r.naphtha_yield_pct is None:
                continue
            mode = "yield-mode" if any(u.unit_id == "CRUDE-EST" for u in r.units) else "unit-detail"
            if r.crude_capacity_kbd > 0 and r.units:
                snap = refinery_day(r, day, data.book, [], include_outages=False)
                util = data.book.utilization(r, r.units[0], day).value
                implied = snap.net_kbd / (r.crude_capacity_kbd * util) * 100
                delta = implied - r.naphtha_yield_pct
            else:
                implied = delta = None
            rows.append((r, mode, implied, delta))

        rows.sort(key=lambda x: -abs(x[3]) if x[3] is not None else 1)
        print(f"{'refinery':<22} {'PADD':<5} {'crude':>7} {'mode':<12} "
              f"{'2024 act%':>9} {'model%':>8} {'delta':>7}")
        print("-" * 78)
        for r, mode, implied, delta in rows:
            imp = f"{implied:8.2f}" if implied is not None else "     n/a"
            dl = f"{delta:+7.2f}" if delta is not None else "    n/a"
            print(f"{r.refinery_id:<22} {r.padd:<5} {r.crude_capacity_kbd:>7,.0f} "
                  f"{mode:<12} {r.naphtha_yield_pct:>9.2f} {imp} {dl}")
        print("\nPADD simple-average 2024 net naphtha yield (% of crude):")
        for p in sorted({r.padd for r, *_ in rows}):
            ys = [r.naphtha_yield_pct for r, *_ in rows if r.padd == p]
            print(f"  PADD {p}: {sum(ys) / len(ys):.2f}%  ({len(ys)} refineries)")

    elif args.command == "export":
        if args.dump:
            balances = us_balance(
                data.refineries, axis, data.book, data.outages, data.flows, data.demand
            )
            path = export_workbook(data, axis, balances, Path(args.out))
        elif args.full:
            from .config import DATA_DIR
            from .workbook import build_desk_workbook

            scenario_dir = DATA_DIR / "scenarios"
            scenarios = [
                load_scenario(p) for p in sorted(scenario_dir.glob("*.yaml"))
            ] if scenario_dir.exists() else []
            path = build_desk_workbook(data, axis, Path(args.out), scenarios=scenarios)
        else:
            from .workbook import build_simple_workbook

            path = build_simple_workbook(data, axis, Path(args.out))
        print(f"wrote {path}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
