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

    p_cap = sub.add_parser(
        "ingest-capacity", help="ingest EA site-level monthly nameplate capacities"
    )
    p_cap.add_argument("path")
    p_cap.add_argument("--as-of", default="2026-07", help="month to stamp (YYYY-MM)")

    p_units = sub.add_parser(
        "ingest-units", help="ingest REM unit capacities (+ optional RDT utilizations)"
    )
    p_units.add_argument("rem_csv")
    p_units.add_argument("--utilization", help="RefineryDataTool throughput CSV")
    p_units.add_argument("--year", default="2026", help="REM capacity year column")
    p_units.add_argument("--util-year", default="2024", help="utilization year")

    sub.add_parser(
        "ingest-reference",
        help="extract crude slate / 2021 yields / US naphtha balance from data/raw",
    )

    p_oev = sub.add_parser(
        "ingest-outages", help="ingest the desk offline-events export (OEVs)"
    )
    p_oev.add_argument("path")
    p_oev.add_argument("--as-of", default="2026-07-09",
                       help="snapshot date for current_outages.csv")

    p_out = sub.add_parser("outages", help="ongoing + upcoming outage events")
    p_out.add_argument("--days", type=int, default=90)
    p_out.add_argument("--padd", type=int)

    p_cal = sub.add_parser(
        "calibrate", help="model-implied net naphtha yield vs 2024 actuals"
    )
    p_cal.add_argument("--padd", type=int)

    p_capr = sub.add_parser(
        "capacity", help="nameplate vs effective (demonstrated) capacity vs current rate"
    )
    p_capr.add_argument("--padd", type=int)
    p_capr.add_argument("--unit", default="CDU", help="unit id to report (default CDU)")

    p_slate = sub.add_parser(
        "slate", help="crude / feedstock slate for one refinery (API, light vs heavy)"
    )
    p_slate.add_argument("refinery_id")

    p_peia = sub.add_parser(
        "pull-eia",
        help="pull EIA weekly/monthly series into data/reference/eia_feed.csv "
             "(same series the workbook's Live Feeds tab pulls)")
    p_peia.add_argument("--api-key", required=True,
                        help="free key from eia.gov/opendata/register.php")
    p_peia.add_argument("--series", nargs="*", default=None,
                        help="EIA series ids (default: the Live Feeds set)")

    p_piir = sub.add_parser(
        "pull-iir",
        help="pull US refinery offline events from the IIR API and refresh "
             "the outage CSVs (token from secrets.yaml if not passed)")
    p_piir.add_argument("--token", default=None,
                        help="IIR Bearer token (expires every 30 days); "
                             "default: iir_token in secrets.yaml")
    p_piir.add_argument("--country", default="U.S.A.",
                        help="IIR country filter (default U.S.A.; '' = world)")
    p_piir.add_argument("--as-of", default=None,
                        help="snapshot date for current_outages.csv "
                             "(default: today)")
    p_piir.add_argument("--no-csv", action="store_true",
                        help="save raw JSON only, don't refresh outage CSVs")
    p_piir.add_argument("--out", default=None, help="output json path")

    p_pea = sub.add_parser(
        "pull-ea",
        help="pull the EA US naphtha monthly balance into "
             "data/reference/us_naphtha_balance_monthly.csv "
             "(key from secrets.yaml if not passed)")
    p_pea.add_argument("--api-key", default=None,
                       help="EA/OilX api key; default: ea_api_key in secrets.yaml")
    p_pea.add_argument("--since", default="2023-01",
                       help="first month kept (default 2023-01)")
    p_pea.add_argument("--product", default="NAPHTHA")
    p_pea.add_argument("--country", default="US")

    p_pk = sub.add_parser(
        "pull-kpler",
        help="pull a Kpler liquids endpoint to data/raw/ (key from "
             "secrets.yaml)")
    p_pk.add_argument("--key", default=None,
                      help="base64 Kpler key; default: kpler_key in secrets.yaml")
    p_pk.add_argument("--endpoint", default="trades")
    p_pk.add_argument("--params", default="products=naphtha&size=100")

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

    if args.command == "pull-eia":
        from .feeds import pull_eia

        path, rows = pull_eia(args.api_key, series=args.series or None)
        print(f"wrote {path} ({len(rows)} rows)")
        seen = set()
        for row in rows:                      # newest-first per series
            if row["series"] not in seen:
                seen.add(row["series"])
                print(f"    {row['series']}: {row['value']} "
                      f"{row['units']} ({row['period']})")
        return 0

    if args.command == "pull-iir":
        from .feeds import pull_iir

        path, summary = pull_iir(
            token=args.token, country=args.country, as_of=args.as_of,
            refresh_csvs=not args.no_csv,
            out_json=Path(args.out) if args.out else None)
        print(f"wrote {path} ({summary['records']} records, "
              f"{summary['pages']} pages)")
        for k, v in (summary.get("ingest") or {}).items():
            if isinstance(v, list):
                print(f"{k} ({len(v)}):")
                for x in v:
                    print(f"    {x}")
            else:
                print(f"{k}: {v}")
        return 0

    if args.command == "pull-ea":
        from .feeds import pull_ea

        path, summary = pull_ea(api_key=args.api_key, since=args.since,
                                product=args.product, country=args.country)
        print(f"wrote {path}")
        for k, v in summary.items():
            print(f"{k}: {v}")
        return 0

    if args.command == "pull-kpler":
        from .feeds import pull_kpler

        path, summary = pull_kpler(key=args.key, endpoint=args.endpoint,
                                   params=args.params)
        print(f"wrote {path}: {summary}")
        return 0

    if args.command == "ingest-yields":
        from .ingest import ingest_yields

        result = ingest_yields(Path(args.path))
        for k, v in result.items():
            print(f"{k}: {v}")
        return 0

    if args.command == "ingest-units":
        from .ingest import ingest_units

        result = ingest_units(
            Path(args.rem_csv),
            rdt_csv=Path(args.utilization) if args.utilization else None,
            year=args.year,
            util_year=args.util_year,
        )
        for k, v in result.items():
            print(f"{k}: {v}")
        return 0

    if args.command == "ingest-reference":
        from .ingest import ingest_reference

        for k, v in ingest_reference().items():
            print(f"{k}: {v}")
        return 0

    if args.command == "ingest-outages":
        from .ingest import ingest_outages

        result = ingest_outages(Path(args.path), as_of=args.as_of)
        for k, v in result.items():
            if isinstance(v, list):
                print(f"{k} ({len(v)}):")
                for x in v:
                    print(f"    {x}")
            else:
                print(f"{k}: {v}")
        return 0

    if args.command == "ingest-capacity":
        from .ingest import ingest_capacity

        result = ingest_capacity(Path(args.path), as_of=args.as_of)
        for k, v in result.items():
            if isinstance(v, dict):
                print(f"{k}:")
                for kk, vv in sorted(v.items(), key=lambda x: -x[1]):
                    print(f"    {vv:8.1f}  {kk}")
            elif isinstance(v, list):
                print(f"{k} ({len(v)}): {', '.join(v)}")
            else:
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

    elif args.command == "capacity":
        import csv as _csv

        from .config import DATA_DIR

        eff = {}
        with (DATA_DIR / "reference" / "effective_capacity.csv").open() as fh:
            for row in _csv.DictReader(fh):
                eff[(row["refinery_id"], row["unit_id"])] = row
        day = axis[0]
        print(f"{'refinery':<26} {'PADD':<5} {'plate':>7} {'effective':>10} "
              f"{'eff yr':>7} {'eff/plate':>10} {'run 2024':>9}")
        print("-" * 82)
        padd_tot: dict[int, list[float]] = {}
        for r in sorted(data.refineries, key=lambda x: (x.padd, -x.crude_capacity_kbd)):
            if args.padd is not None and r.padd != args.padd:
                continue
            row = eff.get((r.refinery_id, args.unit))
            unit = next((u for u in r.units if u.unit_id == args.unit), None)
            if unit is None:
                continue
            util = data.book.utilization(r, unit, day).value
            plate = unit.capacity_kbd
            t = padd_tot.setdefault(r.padd, [0.0, 0.0, 0.0])
            t[0] += plate
            t[2] += plate * util
            if row:
                e = float(row["effective_kbd"])
                t[1] += e
                print(f"{r.refinery_id:<26} {r.padd:<5} {plate:>7,.0f} {e:>10,.1f} "
                      f"{row['effective_year']:>7} {e / plate:>9.0%} {util:>8.0%}")
            else:
                t[1] += plate * util  # no history: assume current rate
                print(f"{r.refinery_id:<26} {r.padd:<5} {plate:>7,.0f} {'n/a':>10} "
                      f"{'':>7} {'':>10} {util:>8.0%}")
        print("-" * 82)
        print(f"{'PADD':<6}{'plate':>10}{'effective':>12}{'running':>10}  ({args.unit}, kbd)")
        for p in sorted(padd_tot):
            plate, e, run = padd_tot[p]
            print(f"{p:<6}{plate:>10,.0f}{e:>12,.0f}{run:>10,.0f}")
        tp = [sum(v[i] for v in padd_tot.values()) for i in range(3)]
        print(f"{'US':<6}{tp[0]:>10,.0f}{tp[1]:>12,.0f}{tp[2]:>10,.0f}")
        print("\neffective = max demonstrated annual throughput 2017-2024 excl. 2020 (RDT)")

    elif args.command == "slate":
        import csv as _csv

        from .config import DATA_DIR

        rid = args.refinery_id
        r = data.refinery(rid)
        print(f"{r.name} [{rid}] — PADD {r.padd} — crude {r.crude_capacity_kbd:,.0f} kbd\n")

        with (DATA_DIR / "reference" / "crude_slate.csv").open() as fh:
            slate = [row for row in _csv.DictReader(fh) if row["refinery_id"] == rid]
        if slate:
            latest = max(row["year"] for row in slate)
            rows = sorted((row for row in slate if row["year"] == latest),
                          key=lambda x: -float(x["slate_pct"] or 0))
            print(f"crude slate, {latest} (REM):")
            for row in rows[:12]:
                print(f"  {float(row['slate_pct']):>5.1f}%  {row['crude_stream']}"
                      f"  ({row['source_country']})")

        with (DATA_DIR / "reference" / "feedstock_slate.csv").open() as fh:
            feeds = [row for row in _csv.DictReader(fh)
                     if row["refinery_id"] == rid and row["year"] == "2024"
                     and float(row["kbd"]) > 0]
        if feeds:
            print("\npurchased supplementary feedstocks, 2024 (RefineryDataTool):")
            print(f"  {'feedstock':<34} {'kbd':>7} {'to unit':<14} {'API':>6}")
            for f in sorted(feeds, key=lambda x: -float(x["kbd"])):
                api = f"{float(f['api']):.1f}" if f["api"] else ""
                print(f"  {f['feedstock']:<34} {float(f['kbd']):>7.1f} "
                      f"{f['to_unit']:<14} {api:>6}")
            naph = [f for f in feeds
                    if f["feedstock"].lower() in ("naphtha", "reformate")]
            if naph:
                kbd = sum(float(f["kbd"]) for f in naph)
                print(f"\n  >> buys {kbd:.1f} kbd of naphtha/reformate — "
                      f"a merchant naphtha BUYER")

    elif args.command == "outages":
        import csv as _csv
        from datetime import timedelta

        from .config import DATA_DIR

        day = axis[0]
        horizon = (day + timedelta(days=args.days)).isoformat()
        padd_of = {r.refinery_id: r.padd for r in data.refineries}
        with (DATA_DIR / "reference" / "outage_events.csv").open() as fh:
            evs = [e for e in _csv.DictReader(fh)
                   if e["refinery_id"] and e["model_unit"]
                   and e["end"] >= day.isoformat() and e["start"] <= horizon
                   and (args.padd is None or padd_of.get(e["refinery_id"]) == args.padd)]
        print(f"naphtha-relevant outage events, {day} -> {horizon} "
              f"({'PADD ' + str(args.padd) if args.padd else 'all PADDs'})\n")
        print(f"{'start':<11} {'end':<11} {'refinery':<26} {'unit':<7} "
              f"{'type':<10} {'offline kbd':>11}  cause")
        print("-" * 100)
        for e in sorted(evs, key=lambda x: (x["start"], x["refinery_id"])):
            off = float(e["capacity_offline"] or 0) / 1000
            live = "LIVE " if e["start"] <= day.isoformat() else ""
            print(f"{e['start']:<11} {e['end']:<11} {e['refinery_id']:<26} "
                  f"{e['model_unit']:<7} {live}{e['event_type']:<9} {off:>11,.1f}  "
                  f"{(e['cause'] or '')[:30]}")
        print(f"\n{len(evs)} events. Live events prefill the boxes' OFFLINE % "
              f"column on the next export.")

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
