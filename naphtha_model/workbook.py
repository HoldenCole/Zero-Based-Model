"""Desk workbook builder.

Generates a live, formula-driven Excel model: traders edit blue/orange input
cells (assumptions, outages, flows, scenario toggle) directly in Excel and
every downstream number — refinery boxes, PADD balance, charts, checks —
recalculates without touching Python. Python is the *generator/refresher*:
it seeds the workbook from the data/ directory and wires the formulas.

Tabs
----
README       how to drive the model, color legend
Model        settings: forward start date, scenario toggle
Dashboard    line/bar charts: supply forecast vs demand, balance, outage
             at-risk, cargo flows, per-refinery forecast
Boxes        one whiteboard-style box per refinery (units, capacity,
             utilization, yields, weekly forward strip) — all live formulas
Balance      per-PADD weekly supply/flows/demand/balance/at-risk
Assumptions  per-PADD utilization + yield matrices (inputs)
Outages      planned / unplanned / scenario rows (inputs)
Flows        ship-tracking / cargo table (inputs)
Demand       demand by PADD & sector (inputs)
Refineries   registry (reference)
Intel        market intel log (reference)
Checks       automated data-quality checks with PASS/FAIL status
"""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference, Series
from openpyxl.chart.marker import Marker
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

from .config import DATA_DIR, PADDS, UNIT_TYPES
from .loaders import ModelData
from .schema import ProcessUnit, Refinery

# ------------------------------------------------------------------ styling

NAVY = "1F3864"
FILL_BANNER = PatternFill("solid", fgColor=NAVY)
FILL_INPUT = PatternFill("solid", fgColor="DDEBF7")     # blue: editable input
FILL_OVERRIDE = PatternFill("solid", fgColor="FCE4D6")  # orange: manual override
FILL_CALC = PatternFill("solid", fgColor="F2F2F2")      # grey: formula, don't type
FILL_TOTAL = PatternFill("solid", fgColor="FFF2CC")     # yellow: totals
FILL_PASS = PatternFill("solid", fgColor="C6EFCE")
FILL_FAIL = PatternFill("solid", fgColor="FFC7CE")

FONT_BANNER = Font(bold=True, color="FFFFFF", size=12)
FONT_HDR = Font(bold=True, size=9)
FONT_SMALL = Font(size=9)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# ------------------------------------------------------- scan-range budgets
# Formulas scan fixed ranges so traders can append rows without re-wiring.

OUT_LO, OUT_HI = 2, 201        # Outages data rows
FLOW_LO, FLOW_HI = 2, 300      # Flows data rows
DEM_LO, DEM_HI = 2, 100        # Demand data rows
REF_LO, REF_HI = 2, 300        # Refineries registry rows
BOX_LO, BOX_HI = 4, 2600       # Boxes data rows


def _banner(ws, row: int, text: str, span: int) -> None:
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    c = ws.cell(row=row, column=1, value=text)
    c.fill = FILL_BANNER
    c.font = FONT_BANNER
    for col in range(2, span + 1):
        ws.cell(row=row, column=col).fill = FILL_BANNER


def _hdr(ws, row: int, col: int, text: str) -> None:
    c = ws.cell(row=row, column=col, value=text)
    c.font = FONT_HDR
    c.fill = FILL_CALC
    c.border = BORDER
    c.alignment = Alignment(wrap_text=True, vertical="center")


def _style(cell, fill=None, fmt=None, bold=False, border=True):
    if fill is not None:
        cell.fill = fill
    if fmt is not None:
        cell.number_format = fmt
    if bold:
        cell.font = Font(bold=True, size=9)
    else:
        cell.font = FONT_SMALL
    if border:
        cell.border = BORDER
    return cell


class DeskWorkbook:
    """Builds the workbook. One instance per export."""

    def __init__(self, data: ModelData, axis: list[date], scenarios=None):
        self.scenarios = scenarios or []
        self.data = data
        self.book = data.book
        self.axis = axis
        self.weeks = len(axis)
        self.cuts = list(self.book.cuts)
        self.padds = sorted({r.padd for r in data.refineries if r.region == "US"})
        self.wb = Workbook()

        # Boxes column map (1-indexed), dynamic in the number of cuts:
        # A type | B refinery | C padd | D unit | E unit_type | F cap |
        # G eff cap | H util_ovr | I util | (ovr, yield) per cut | ysum |
        # base | net wk1..wkW | off% wk1..wkW
        self.c_type, self.c_rid, self.c_padd, self.c_uid, self.c_utype = 1, 2, 3, 4, 5
        self.c_cap, self.c_effcap, self.c_utilov, self.c_util = 6, 7, 8, 9
        self.c_cut0 = 10                               # first cut override col
        self.c_ysum = self.c_cut0 + 2 * len(self.cuts)
        self.c_base = self.c_ysum + 1
        self.c_net0 = self.c_base + 1                  # first weekly net col
        self.c_off0 = self.c_net0 + self.weeks         # first weekly off% col
        self.c_last = self.c_off0 + self.weeks - 1

        # effective capacity (demonstrated max annual throughput) if ingested
        self.eff_caps: dict[tuple[str, str], float] = {}
        eff_path = DATA_DIR / "reference" / "effective_capacity.csv"
        if eff_path.exists():
            import csv as _csv

            with eff_path.open() as fh:
                for row in _csv.DictReader(fh):
                    self.eff_caps[(row["refinery_id"], row["unit_id"])] = float(
                        row["effective_kbd"]
                    )

        # populated as sheets are built
        self.assump_refs: dict = {}
        self.balance_rows: dict[int, dict[str, int]] = {}
        self.balance_week_row: dict[int, int] = {}
        self.refinery_total_rows: dict[str, int] = {}

    # -------------------------------------------------------------- helpers

    def _cut_ovr_col(self, i: int) -> int:
        return self.c_cut0 + 2 * i

    def _cut_yld_col(self, i: int) -> int:
        return self.c_cut0 + 2 * i + 1

    def _resolved_padd(self, padd: int):
        """Resolved (util, {unit_type: {cut: yield}}) for a PADD on day one,
        ignoring refinery-specific overrides (dummy refinery)."""
        dummy = Refinery(
            refinery_id="__PADD__", name="", owner="", city="", state="",
            padd=padd, crude_capacity_kbd=0,
        )
        day = self.axis[0]
        yields = {}
        util = None
        for ut in UNIT_TYPES:
            unit = ProcessUnit("__PADD__", "__U__", ut, 0.0)
            if util is None:
                util = self.book.utilization(dummy, unit, day).value
            yields[ut] = {
                cut: self.book.yield_for(dummy, unit, cut, day).value
                for cut in self.cuts
            }
        return util, yields

    # ---------------------------------------------------------------- build

    def build(self, out_path: Path) -> Path:
        self._sheet_model()
        self._sheet_assumptions()
        self._sheet_refineries()
        self._sheet_outages()
        self._sheet_flows()
        self._sheet_demand()
        self._sheet_intel()
        self._sheet_boxes()
        self._sheet_balance()
        self._sheet_calibration()
        self._sheet_yields_2024()
        self._sheet_dashboard()
        self._sheet_checks()
        self._sheet_readme()

        if "Sheet" in self.wb.sheetnames:
            del self.wb["Sheet"]
        order = [
            "README", "Dashboard", "Boxes", "Balance", "Calibration",
            "Assumptions", "Outages", "Flows", "Demand", "Refineries",
            "Yields_2024", "Intel", "Checks", "Model",
        ]
        self.wb._sheets = [self.wb[name] for name in order]
        out_path = Path(out_path)
        self.wb.save(out_path)
        return out_path

    # ---------------------------------------------------------------- Model

    def _sheet_model(self) -> None:
        ws = self.wb.create_sheet("Model")
        _banner(ws, 1, "Model settings", 4)
        rows = [
            ("Forward window start (edit to shift every date in the book)", self.axis[0], "yyyy-mm-dd", FILL_INPUT),
            ("Weeks in window (fixed at generation; regenerate to change)", self.weeks, "0", FILL_CALC),
            ("Include SCENARIO outage rows? (YES/NO)", "NO", None, FILL_INPUT),
            ("Generated", datetime.now().strftime("%Y-%m-%d %H:%M"), None, FILL_CALC),
            ("Regenerate with", "python -m naphtha_model export --out <file>", None, FILL_CALC),
        ]
        for i, (label, value, fmt, fill) in enumerate(rows, start=3):
            ws.cell(row=i, column=1, value=label).font = FONT_SMALL
            _style(ws.cell(row=i, column=2, value=value), fill=fill, fmt=fmt)
        dv = DataValidation(type="list", formula1='"YES,NO"', allow_blank=False)
        ws.add_data_validation(dv)
        dv.add("B5")
        ws.column_dimensions["A"].width = 55
        ws.column_dimensions["B"].width = 40
        # cell used by formulas elsewhere
        self.start_ref = "Model!$B$3"
        self.toggle_ref = "Model!$B$5"

    # ---------------------------------------------------------- Assumptions

    def _sheet_assumptions(self) -> None:
        ws = self.wb.create_sheet("Assumptions")
        span = 6 + len(self.cuts)
        _banner(
            ws, 1,
            "Assumptions — blue cells are inputs. Yields are % of unit throughput; "
            "consumers (reformer/isom) are NEGATIVE. Edits flow straight into Boxes/Balance.",
            span + 8,
        )
        # long lookup tables (right-hand side) that Boxes formulas read
        UL_PADD, UL_VAL = span + 2, span + 3            # util long: padd, value
        YL_PADD, YL_UT, YL_CUT, YL_VAL = span + 5, span + 6, span + 7, span + 8
        for col, label in [
            (UL_PADD, "padd"), (UL_VAL, "util"),
            (YL_PADD, "padd"), (YL_UT, "unit_type"), (YL_CUT, "cut"), (YL_VAL, "yield"),
        ]:
            _hdr(ws, 3, col, label)

        ul_row, yl_row = 4, 4
        r = 3
        for padd in self.padds:
            util, yields = self._resolved_padd(padd)
            _banner(ws, r, f"PADD {padd} — {PADDS.get(padd, '')}", span)
            ws.cell(row=r + 1, column=1, value="Utilization").font = FONT_HDR
            ucell = ws.cell(row=r + 1, column=2, value=util)
            _style(ucell, fill=FILL_INPUT, fmt="0.0%")
            # util long row points at the matrix cell
            ws.cell(row=ul_row, column=UL_PADD, value=padd)
            ws.cell(row=ul_row, column=UL_VAL, value=f"={ucell.coordinate}")
            ul_row += 1

            _hdr(ws, r + 3, 1, "Unit type")
            for j, cut in enumerate(self.cuts):
                _hdr(ws, r + 3, 2 + j, cut)
            for i, ut in enumerate(UNIT_TYPES):
                row = r + 4 + i
                ws.cell(row=row, column=1, value=ut).font = FONT_SMALL
                for j, cut in enumerate(self.cuts):
                    cell = ws.cell(row=row, column=2 + j, value=yields[ut][cut])
                    _style(cell, fill=FILL_INPUT, fmt="0.0%")
                    ws.cell(row=yl_row, column=YL_PADD, value=padd)
                    ws.cell(row=yl_row, column=YL_UT, value=ut)
                    ws.cell(row=yl_row, column=YL_CUT, value=cut)
                    ws.cell(row=yl_row, column=YL_VAL, value=f"={cell.coordinate}")
                    yl_row += 1
            r += 5 + len(UNIT_TYPES) + 1

        ul_col_p, ul_col_v = get_column_letter(UL_PADD), get_column_letter(UL_VAL)
        yl_p, yl_ut = get_column_letter(YL_PADD), get_column_letter(YL_UT)
        yl_c, yl_v = get_column_letter(YL_CUT), get_column_letter(YL_VAL)
        self.assump_refs = {
            "util": (
                f"Assumptions!${ul_col_v}$4:${ul_col_v}${ul_row - 1}",
                f"Assumptions!${ul_col_p}$4:${ul_col_p}${ul_row - 1}",
            ),
            "yield": (
                f"Assumptions!${yl_v}$4:${yl_v}${yl_row - 1}",
                f"Assumptions!${yl_p}$4:${yl_p}${yl_row - 1}",
                f"Assumptions!${yl_ut}$4:${yl_ut}${yl_row - 1}",
                f"Assumptions!${yl_c}$4:${yl_c}${yl_row - 1}",
            ),
        }
        ws.column_dimensions["A"].width = 16
        ws.freeze_panes = "A2"

    # ------------------------------------------------------------ Refineries

    def _sheet_refineries(self) -> None:
        ws = self.wb.create_sheet("Refineries")
        _banner(ws, 1, "Refinery registry (reference — edit via data/reference and regenerate)", 9)
        headers = ["refinery_id", "name", "owner", "city", "state", "padd",
                   "crude_kbd", "status", "notes"]
        for c, h in enumerate(headers, start=1):
            _hdr(ws, 2, c, h)
        for i, r in enumerate(self.data.refineries, start=3):
            vals = [r.refinery_id, r.name, r.owner, r.city, r.state, r.padd,
                    r.crude_capacity_kbd, r.status, r.notes]
            for c, v in enumerate(vals, start=1):
                _style(ws.cell(row=i, column=c, value=v))
        widths = [16, 24, 20, 14, 6, 6, 10, 10, 44]
        for c, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(c)].width = w

    # -------------------------------------------------------------- Outages

    def _sheet_outages(self) -> None:
        ws = self.wb.create_sheet("Outages")
        headers = ["category", "refinery_id", "unit_id", "outage_id", "start",
                   "end", "offline_pct", "active", "source", "confidence", "notes"]
        for c, h in enumerate(headers, start=1):
            _hdr(ws, 1, c, h)
        ws.cell(row=1, column=12, value=(
            "Add a row the moment an outage hits — Boxes/Balance/Dashboard update "
            "instantly. category=scenario rows only count when Model!B5 = YES."
        )).font = FONT_SMALL

        scenario_outages = [
            (s.name, o) for s in self.scenarios for o in s.outages
        ]
        row = OUT_LO
        for o in self.data.outages:
            vals = [o.outage_type, o.refinery_id, o.unit_id, o.outage_id,
                    o.start, o.end, o.offline_pct, None, o.source, o.confidence, o.notes]
            for c, v in enumerate(vals, start=1):
                cell = ws.cell(row=row, column=c, value=v)
                fill = FILL_CALC if c == 8 else FILL_INPUT
                fmt = "yyyy-mm-dd" if c in (5, 6) else ("0" if c == 7 else None)
                _style(cell, fill=fill, fmt=fmt)
            row += 1
        for scn_name, o in scenario_outages:
            vals = ["scenario", o.refinery_id, o.unit_id, o.outage_id,
                    o.start, o.end, o.offline_pct, None, scn_name, "", o.notes]
            for c, v in enumerate(vals, start=1):
                cell = ws.cell(row=row, column=c, value=v)
                fill = FILL_CALC if c == 8 else FILL_INPUT
                fmt = "yyyy-mm-dd" if c in (5, 6) else ("0" if c == 7 else None)
                _style(cell, fill=fill, fmt=fmt)
            row += 1
        # blank input rows + the active formula down the whole scan range
        for rr in range(OUT_LO, OUT_HI + 1):
            cell = ws.cell(
                row=rr, column=8,
                value=f'=IF($B{rr}="","",IF($A{rr}="scenario",{self.toggle_ref},"YES"))',
            )
            _style(cell, fill=FILL_CALC)
            if rr >= row:
                for c in (1, 2, 3, 4, 5, 6, 7, 9, 10, 11):
                    fmt = "yyyy-mm-dd" if c in (5, 6) else None
                    _style(ws.cell(row=rr, column=c), fill=FILL_INPUT, fmt=fmt)

        dv = DataValidation(type="list", formula1='"planned,unplanned,scenario"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"A{OUT_LO}:A{OUT_HI}")
        widths = [11, 14, 9, 14, 11, 11, 11, 8, 16, 11, 40]
        for c, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(c)].width = w
        ws.freeze_panes = "A2"

    # ---------------------------------------------------------------- Flows

    def _sheet_flows(self) -> None:
        ws = self.wb.create_sheet("Flows")
        headers = ["date", "padd", "direction", "volume_kbd", "cut",
                   "counterparty", "vessel", "source", "notes", "category", "signed_kbd"]
        for c, h in enumerate(headers, start=1):
            _hdr(ws, 1, c, h)
        ws.cell(row=1, column=12, value=(
            "Ship-tracking / cargo table. import & transfer_in ADD to the PADD, "
            "export & transfer_out REMOVE. category=scenario rows only count when "
            "Model!B5 = YES. signed_kbd is automatic."
        )).font = FONT_SMALL

        rows = [(f, "") for f in self.data.flows] + [
            (f, "scenario") for s in self.scenarios for f in s.flows
        ]
        row = FLOW_LO
        for f, category in rows:
            vals = [f.flow_date, f.padd, f.direction, f.volume_kbd, f.cut,
                    f.counterparty, f.vessel, f.source, f.notes, category]
            for c, v in enumerate(vals, start=1):
                _style(ws.cell(row=row, column=c, value=v),
                       fill=FILL_INPUT, fmt="yyyy-mm-dd" if c == 1 else None)
            row += 1
        for rr in range(FLOW_LO, FLOW_HI + 1):
            cell = ws.cell(
                row=rr, column=11,
                value=(f'=IF($C{rr}="","",'
                       f'IF(AND($J{rr}="scenario",{self.toggle_ref}<>"YES"),0,'
                       f'IF(OR($C{rr}="import",$C{rr}="transfer_in"),$D{rr},-$D{rr})))'),
            )
            _style(cell, fill=FILL_CALC, fmt="0.0")
            if rr >= row:
                for c in range(1, 11):
                    _style(ws.cell(row=rr, column=c), fill=FILL_INPUT,
                           fmt="yyyy-mm-dd" if c == 1 else None)

        dv = DataValidation(type="list",
                            formula1='"import,export,transfer_in,transfer_out"',
                            allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"C{FLOW_LO}:C{FLOW_HI}")
        dv2 = DataValidation(type="list", formula1='"scenario"', allow_blank=True)
        ws.add_data_validation(dv2)
        dv2.add(f"J{FLOW_LO}:J{FLOW_HI}")
        widths = [11, 6, 12, 11, 8, 16, 16, 14, 30, 10, 11]
        for c, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(c)].width = w
        ws.freeze_panes = "A2"

    # --------------------------------------------------------------- Demand

    def _sheet_demand(self) -> None:
        ws = self.wb.create_sheet("Demand")
        headers = ["padd", "sector", "volume_kbd", "start", "end", "notes"]
        for c, h in enumerate(headers, start=1):
            _hdr(ws, 1, c, h)
        ws.cell(row=1, column=7, value=(
            "Steady-state naphtha dispositions per PADD. The workbook treats every "
            "row as always-on; dated demand is applied by the Python engine."
        )).font = FONT_SMALL
        for i, d in enumerate(self.data.demand, start=DEM_LO):
            vals = [d.padd, d.sector, d.volume_kbd, d.start, d.end, d.notes]
            for c, v in enumerate(vals, start=1):
                _style(ws.cell(row=i, column=c, value=v),
                       fill=FILL_INPUT, fmt="yyyy-mm-dd" if c in (4, 5) else None)
        widths = [6, 20, 11, 11, 11, 44]
        for c, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(c)].width = w

    # ---------------------------------------------------------------- Intel

    def _sheet_intel(self) -> None:
        ws = self.wb.create_sheet("Intel")
        headers = ["date", "refinery_id", "padd", "headline", "impact_kbd",
                   "linked_outage_id", "source", "confidence", "notes"]
        for c, h in enumerate(headers, start=1):
            _hdr(ws, 1, c, h)
        for i, n in enumerate(self.data.intel, start=2):
            vals = [n.note_date, n.refinery_id, n.padd, n.headline, n.impact_kbd,
                    n.linked_outage_id, n.source, n.confidence, n.notes]
            for c, v in enumerate(vals, start=1):
                _style(ws.cell(row=i, column=c, value=v),
                       fill=FILL_INPUT, fmt="yyyy-mm-dd" if c == 1 else None)
        widths = [11, 14, 6, 60, 11, 15, 12, 11, 40]
        for c, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(c)].width = w

    # ---------------------------------------------------------------- Boxes

    def _offline_formula(self, row: int, week_idx: int, ws_name: str = "Boxes") -> str:
        """Prorated offline fraction for one unit-week: sum over matching active
        outage rows of offline% x overlap_days/7, capped at 1. Elementwise
        min/max built from ABS so it works inside SUMPRODUCT."""
        net_col = get_column_letter(self.c_net0 + week_idx)
        ws_ = f"${net_col}$2"                     # week start date (header)
        we = f"(${net_col}$2+6)"                  # week end date
        S = f"Outages!$E${OUT_LO}:$E${OUT_HI}"
        E = f"Outages!$F${OUT_LO}:$F${OUT_HI}"
        min_e_we = f"(({E}+{we}-ABS({E}-{we}))/2)"
        max_s_ws = f"(({S}+{ws_}+ABS({S}-{ws_}))/2)"
        od = f"({min_e_we}-{max_s_ws}+1)"
        clip = f"(({od}+ABS({od}))/2)"
        flags = (
            f"(Outages!$B${OUT_LO}:$B${OUT_HI}=$B{row})"
            f"*((Outages!$C${OUT_LO}:$C${OUT_HI}=$D{row})"
            f"+(Outages!$C${OUT_LO}:$C${OUT_HI}=\"\"))"
            f"*(Outages!$H${OUT_LO}:$H${OUT_HI}=\"YES\")"
        )
        return (
            f"=MIN(1,SUMPRODUCT({flags}"
            f"*(Outages!$G${OUT_LO}:$G${OUT_HI}/100)*{clip}/7))"
        )

    def _sheet_boxes(self) -> None:
        ws = self.wb.create_sheet("Boxes")
        _banner(
            ws, 1,
            "Refinery boxes — every refinery, unit by unit. Blue = capacity input, "
            "orange = manual overrides (clear the cell to fall back to PADD/global "
            "assumptions), grey = live formulas. Weekly strip nets outages automatically.",
            self.c_last,
        )
        # week date headers (row 2) over both the net strip and the off% strip
        for k in range(self.weeks):
            for c0 in (self.c_net0, self.c_off0):
                cell = ws.cell(row=2, column=c0 + k, value=f"={self.start_ref}+{k}*7")
                _style(cell, fill=FILL_CALC, fmt="m/d", bold=True)
        # static headers (row 3)
        labels = {
            self.c_type: "row", self.c_rid: "refinery_id", self.c_padd: "padd",
            self.c_uid: "unit", self.c_utype: "type", self.c_cap: "cap kbd",
            self.c_effcap: "eff cap", self.c_utilov: "util ovr", self.c_util: "util",
            self.c_ysum: "Σ yield", self.c_base: "base net kbd",
        }
        for i, cut in enumerate(self.cuts):
            labels[self._cut_ovr_col(i)] = f"{cut} ovr"
            labels[self._cut_yld_col(i)] = f"{cut} yield"
        for col, text in labels.items():
            _hdr(ws, 3, col, text)
        for k in range(self.weeks):
            _hdr(ws, 3, self.c_net0 + k, "net kbd")
            _hdr(ws, 3, self.c_off0 + k, "off%")

        util_val, util_padd = self.assump_refs["util"]
        y_val, y_padd, y_ut, y_cut = self.assump_refs["yield"]
        row = BOX_LO
        day1 = self.axis[0]

        for ref in sorted(self.data.refineries, key=lambda r: (r.padd, -r.crude_capacity_kbd)):
            _banner(
                ws, row,
                f"  {ref.name}  [{ref.refinery_id}]   —   {ref.owner}   —   PADD {ref.padd}"
                f"   —   crude {ref.crude_capacity_kbd:,.0f} kbd",
                self.c_last,
            )
            row += 1
            first_unit = row
            ovl = get_column_letter(self.c_utilov)
            for unit in ref.units:
                ws.cell(row=row, column=self.c_type, value="UNIT").font = FONT_SMALL
                _style(ws.cell(row=row, column=self.c_rid, value=ref.refinery_id))
                _style(ws.cell(row=row, column=self.c_padd, value=ref.padd))
                _style(ws.cell(row=row, column=self.c_uid, value=unit.unit_id))
                _style(ws.cell(row=row, column=self.c_utype, value=unit.unit_type))
                _style(ws.cell(row=row, column=self.c_cap, value=unit.capacity_kbd),
                       fill=FILL_INPUT, fmt="0")
                _style(ws.cell(row=row, column=self.c_effcap,
                               value=self.eff_caps.get((ref.refinery_id, unit.unit_id))),
                       fill=FILL_CALC, fmt="0")

                # manual override cells, prefilled from data/overrides when active
                ov_util = self.book._find_override(ref, unit, day1, "utilization")
                _style(ws.cell(row=row, column=self.c_utilov,
                               value=ov_util.value if ov_util else None),
                       fill=FILL_OVERRIDE, fmt="0.0%")
                _style(ws.cell(
                    row=row, column=self.c_util,
                    value=(f'=IF(${ovl}{row}<>"",${ovl}{row},'
                           f"SUMIFS({util_val},{util_padd},$C{row}))"),
                ), fill=FILL_CALC, fmt="0.0%")

                for i, cut in enumerate(self.cuts):
                    oc = get_column_letter(self._cut_ovr_col(i))
                    ov = self.book._find_override(ref, unit, day1, "yield", cut=cut)
                    _style(ws.cell(row=row, column=self._cut_ovr_col(i),
                                   value=ov.value if ov else None),
                           fill=FILL_OVERRIDE, fmt="0.0%")
                    _style(ws.cell(
                        row=row, column=self._cut_yld_col(i),
                        value=(f'=IF(${oc}{row}<>"",${oc}{row},'
                               f"SUMIFS({y_val},{y_padd},$C{row},{y_ut},$E{row},"
                               f'{y_cut},"{cut}"))'),
                    ), fill=FILL_CALC, fmt="0.0%")

                ysum = "+".join(
                    f"{get_column_letter(self._cut_yld_col(i))}{row}"
                    for i in range(len(self.cuts))
                )
                _style(ws.cell(row=row, column=self.c_ysum, value=f"={ysum}"),
                       fill=FILL_CALC, fmt="0.0%")
                cap_l, util_l = get_column_letter(self.c_cap), get_column_letter(self.c_util)
                ysum_l = get_column_letter(self.c_ysum)
                _style(ws.cell(row=row, column=self.c_base,
                               value=f"=${cap_l}{row}*${util_l}{row}*${ysum_l}{row}"),
                       fill=FILL_CALC, fmt="0.0")

                for k in range(self.weeks):
                    off_l = get_column_letter(self.c_off0 + k)
                    _style(ws.cell(row=row, column=self.c_off0 + k,
                                   value=self._offline_formula(row, k)),
                           fill=FILL_CALC, fmt="0%")
                    _style(ws.cell(
                        row=row, column=self.c_net0 + k,
                        value=(f"=${cap_l}{row}*(1-${off_l}{row})"
                               f"*${util_l}{row}*${ysum_l}{row}"),
                    ), fill=FILL_CALC, fmt="0.0")
                row += 1

            # TOTAL row
            ws.cell(row=row, column=self.c_type, value="TOTAL").font = FONT_HDR
            _style(ws.cell(row=row, column=self.c_rid, value=ref.refinery_id),
                   fill=FILL_TOTAL, bold=True)
            _style(ws.cell(row=row, column=self.c_padd, value=ref.padd),
                   fill=FILL_TOTAL, bold=True)
            _style(ws.cell(row=row, column=self.c_uid, value="NET NAPHTHA"),
                   fill=FILL_TOTAL, bold=True)
            base_l = get_column_letter(self.c_base)
            _style(ws.cell(row=row, column=self.c_base,
                           value=f"=SUM({base_l}{first_unit}:{base_l}{row - 1})"),
                   fill=FILL_TOTAL, fmt="0.0", bold=True)
            for k in range(self.weeks):
                nl = get_column_letter(self.c_net0 + k)
                _style(ws.cell(row=row, column=self.c_net0 + k,
                               value=f"=SUM({nl}{first_unit}:{nl}{row - 1})"),
                       fill=FILL_TOTAL, fmt="0.0", bold=True)
            self.refinery_total_rows[ref.refinery_id] = row
            row += 2

        widths = {1: 7, 2: 14, 3: 5, 4: 12, 5: 12, 6: 8, 7: 8, 8: 8, 9: 7}
        for i in range(len(self.cuts)):
            widths[self._cut_ovr_col(i)] = 8
            widths[self._cut_yld_col(i)] = 8
        widths[self.c_ysum] = 8
        widths[self.c_base] = 10
        for k in range(self.weeks):
            widths[self.c_net0 + k] = 7
            widths[self.c_off0 + k] = 6
        for col, w in widths.items():
            ws.column_dimensions[get_column_letter(col)].width = w
        ws.freeze_panes = ws.cell(row=BOX_LO, column=self.c_utilov).coordinate

    # -------------------------------------------------------------- Balance

    def _sheet_balance(self) -> None:
        ws = self.wb.create_sheet("Balance")
        span = 1 + self.weeks
        _banner(ws, 1, "PADD naphtha balance — kbd, weekly. + = length, - = short. "
                       "All rows are live formulas over Boxes / Flows / Demand.", span + 2)
        r = 3
        for padd in self.padds:
            _banner(ws, r, f"PADD {padd} — {PADDS.get(padd, '')}", span)
            week_row = r + 1
            _hdr(ws, week_row, 1, "Week starting")
            for k in range(self.weeks):
                cell = ws.cell(row=week_row, column=2 + k, value=f"={self.start_ref}+{k}*7")
                _style(cell, fill=FILL_CALC, fmt="m/d", bold=True)

            rows = {
                "supply": "Refinery net supply (w/ outages)",
                "base": "Refinery net supply (base, no outages)",
                "risk": "At risk — supply lost to outages",
                "flows": "Net trade flows (ships in - out)",
                "demand": "Demand (crackers, blending, diluent)",
                "balance": "BALANCE (supply + flows - demand)",
            }
            idx = {name: week_row + 1 + i for i, name in enumerate(rows)}
            for name, label in rows.items():
                bold = name == "balance"
                _style(ws.cell(row=idx[name], column=1, value=label),
                       fill=FILL_TOTAL if bold else FILL_CALC, bold=True)

            box_type = f"Boxes!$A${BOX_LO}:$A${BOX_HI}"
            box_padd = f"Boxes!$C${BOX_LO}:$C${BOX_HI}"
            base_l = get_column_letter(self.c_base)
            for k in range(self.weeks):
                col = get_column_letter(2 + k)
                wk = f"{col}${week_row}"
                net_l = get_column_letter(self.c_net0 + k)
                cur = (f'=SUMIFS(Boxes!${net_l}${BOX_LO}:${net_l}${BOX_HI},'
                       f'{box_type},"TOTAL",{box_padd},{padd})')
                base = (f'=SUMIFS(Boxes!${base_l}${BOX_LO}:${base_l}${BOX_HI},'
                        f'{box_type},"TOTAL",{box_padd},{padd})')
                flows = (f"=SUMIFS(Flows!$K${FLOW_LO}:$K${FLOW_HI},"
                         f"Flows!$B${FLOW_LO}:$B${FLOW_HI},{padd},"
                         f'Flows!$A${FLOW_LO}:$A${FLOW_HI},">="&{wk},'
                         f'Flows!$A${FLOW_LO}:$A${FLOW_HI},"<="&({wk}+6))')
                demand = (f"=SUMIFS(Demand!$C${DEM_LO}:$C${DEM_HI},"
                          f"Demand!$A${DEM_LO}:$A${DEM_HI},{padd})")
                values = {
                    "supply": cur,
                    "base": base,
                    "risk": f"={col}{idx['base']}-{col}{idx['supply']}",
                    "flows": flows,
                    "demand": demand,
                    "balance": f"={col}{idx['supply']}+{col}{idx['flows']}-{col}{idx['demand']}",
                }
                for name, formula in values.items():
                    bold = name == "balance"
                    _style(ws.cell(row=idx[name], column=2 + k, value=formula),
                           fill=FILL_TOTAL if bold else FILL_CALC, fmt="0.0", bold=bold)

            self.balance_rows[padd] = idx
            self.balance_week_row[padd] = week_row
            r = idx["balance"] + 3

        ws.column_dimensions["A"].width = 38
        for k in range(self.weeks):
            ws.column_dimensions[get_column_letter(2 + k)].width = 9
        ws.freeze_panes = "B2"

    # ---------------------------------------------------------- Calibration

    def _sheet_calibration(self) -> None:
        """Model-implied net naphtha yield vs 2024 actual, live formulas."""
        ws = self.wb.create_sheet("Calibration")
        _banner(
            ws, 1,
            "Calibration — model-implied net naphtha yield vs 2024 actuals. Big deltas "
            "mean the unit yield assumptions need tuning (or the 2024 number reflects a "
            "different naphtha definition). Yield-mode refineries match by construction.",
            10,
        )
        headers = ["refinery_id", "name", "padd", "crude kbd", "PADD util",
                   "model net kbd", "implied yield", "2024 actual", "delta (pts)", "mode"]
        for c, h in enumerate(headers, start=1):
            _hdr(ws, 3, c, h)

        util_val, util_padd = self.assump_refs["util"]
        box_type = f"Boxes!$A${BOX_LO}:$A${BOX_HI}"
        box_rid = f"Boxes!$B${BOX_LO}:$B${BOX_HI}"
        base_l = get_column_letter(self.c_base)
        r = 4
        for ref in sorted(self.data.refineries, key=lambda x: (x.padd, -x.crude_capacity_kbd)):
            if ref.naphtha_yield_pct is None:
                continue
            mode = ("yield-mode" if any(u.unit_id == "CRUDE-EST" for u in ref.units)
                    else "unit-detail")
            _style(ws.cell(row=r, column=1, value=ref.refinery_id))
            _style(ws.cell(row=r, column=2, value=ref.name))
            _style(ws.cell(row=r, column=3, value=ref.padd))
            _style(ws.cell(row=r, column=4, value=ref.crude_capacity_kbd), fmt="0")
            _style(ws.cell(row=r, column=5,
                           value=f"=SUMIFS({util_val},{util_padd},$C{r})"),
                   fill=FILL_CALC, fmt="0.0%")
            _style(ws.cell(
                row=r, column=6,
                value=(f'=SUMIFS(Boxes!${base_l}${BOX_LO}:${base_l}${BOX_HI},'
                       f'{box_type},"TOTAL",{box_rid},$A{r})'),
            ), fill=FILL_CALC, fmt="0.0")
            _style(ws.cell(row=r, column=7,
                           value=f'=IF($D{r}*$E{r}=0,"",$F{r}/($D{r}*$E{r}))'),
                   fill=FILL_CALC, fmt="0.00%")
            _style(ws.cell(row=r, column=8, value=ref.naphtha_yield_pct / 100.0),
                   fmt="0.00%")
            _style(ws.cell(row=r, column=9, value=f'=IF($G{r}="","",$G{r}-$H{r})'),
                   fill=FILL_CALC, fmt="+0.00%;-0.00%")
            _style(ws.cell(row=r, column=10, value=mode))
            r += 1
        last = r - 1

        # flag deltas larger than +/-2 points
        ws.conditional_formatting.add(
            f"I4:I{last}",
            CellIsRule(operator="greaterThan", formula=["0.02"], fill=FILL_FAIL))
        ws.conditional_formatting.add(
            f"I4:I{last}",
            CellIsRule(operator="lessThan", formula=["-0.02"], fill=FILL_FAIL))

        # PADD averages of 2024 actual net naphtha yield
        s = last + 3
        _banner(ws, s, "PADD simple-average 2024 net naphtha yield", 3)
        _hdr(ws, s + 1, 1, "PADD")
        _hdr(ws, s + 1, 2, "avg 2024 yield")
        _hdr(ws, s + 1, 3, "refineries")
        for i, padd in enumerate(self.padds):
            row = s + 2 + i
            _style(ws.cell(row=row, column=1, value=padd))
            _style(ws.cell(row=row, column=2,
                           value=f"=AVERAGEIF($C$4:$C${last},$A{row},$H$4:$H${last})"),
                   fill=FILL_CALC, fmt="0.00%")
            _style(ws.cell(row=row, column=3,
                           value=f"=COUNTIF($C$4:$C${last},$A{row})"),
                   fill=FILL_CALC, fmt="0")

        widths = [22, 34, 6, 10, 9, 12, 12, 11, 11, 12]
        for c, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(c)].width = w
        ws.freeze_panes = "A4"

    # ----------------------------------------------------------- Yields_2024

    def _sheet_yields_2024(self) -> None:
        """Full 2024 estimated-output reference table (all products)."""
        ws = self.wb.create_sheet("Yields_2024")
        _banner(ws, 1, "2024 estimated refinery yields (% of crude) — reference; "
                       "re-ingest via: python -m naphtha_model ingest-yields <file>", 13)
        from .loaders import load_yields_2024

        rows = load_yields_2024()
        headers = ["refinery_id", "padd", "state", "city", "operator", "owners",
                   "gasoil_diesel", "gasoline", "hfo", "kero_jet", "lpg",
                   "naphtha", "total"]
        keys = ["refinery_id", "padd", "state", "city", "operator", "owners",
                "gasoil_diesel_pct", "gasoline_pct", "hfo_pct", "kero_jet_pct",
                "lpg_pct", "naphtha_pct", "total_pct"]
        for c, h in enumerate(headers, start=1):
            _hdr(ws, 2, c, h)
        for i, row in enumerate(rows.values(), start=3):
            for c, key in enumerate(keys, start=1):
                v = row.get(key, "")
                if key.endswith("_pct"):
                    v = float(v) / 100.0 if v not in ("", None) else None
                elif key == "padd":
                    v = int(v)
                _style(ws.cell(row=i, column=c, value=v),
                       fmt="0.00%" if key.endswith("_pct") else None)
        widths = [22, 6, 14, 16, 26, 30, 12, 10, 8, 10, 8, 10, 8]
        for c, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(c)].width = w
        ws.freeze_panes = "A3"

    # ------------------------------------------------------------ Dashboard

    def _line(self, title: str, y_title: str) -> LineChart:
        ch = LineChart()
        ch.title = title
        ch.style = 2
        ch.y_axis.title = y_title
        ch.x_axis.title = "Week starting"
        ch.x_axis.number_format = "m/d"
        ch.height, ch.width = 8.5, 17
        return ch

    def _series(self, ws, row: int, title: str) -> Series:
        ref = Reference(ws, min_col=2, max_col=1 + self.weeks, min_row=row, max_row=row)
        s = Series(ref, title=title)
        s.marker = Marker(symbol="circle", size=5)
        s.smooth = False
        return s

    def _sheet_dashboard(self) -> None:
        ws = self.wb.create_sheet("Dashboard")
        _banner(ws, 1, "Dashboard — every chart is live: edit Assumptions / Outages / "
                       "Flows or flip the scenario toggle (Model!B5) and these move.", 20)
        bal = self.wb["Balance"]
        boxes = self.wb["Boxes"]
        padd = 3 if 3 in self.balance_rows else self.padds[0]
        idx = self.balance_rows[padd]
        cats = Reference(bal, min_col=2, max_col=1 + self.weeks,
                         min_row=self.balance_week_row[padd],
                         max_row=self.balance_week_row[padd])

        ch1 = self._line(f"PADD {padd} net naphtha supply forecast (kbd)", "kbd")
        for row_name, label in [("supply", "Supply w/ outages & scenario"),
                                ("base", "Supply base (no outages)"),
                                ("demand", "Demand")]:
            ch1.series.append(self._series(bal, idx[row_name], label))
        ch1.set_categories(cats)
        ws.add_chart(ch1, "B3")

        ch2 = self._line(f"PADD {padd} balance (+ length / - short, kbd)", "kbd")
        ch2.series.append(self._series(bal, idx["balance"], "Balance"))
        ch2.set_categories(cats)
        ws.add_chart(ch2, "L3")

        ch3 = BarChart()
        ch3.title = f"PADD {padd} supply at risk from outages (kbd)"
        ch3.y_axis.title = "kbd"
        ch3.height, ch3.width = 8.5, 17
        ch3.add_data(Reference(bal, min_col=2, max_col=1 + self.weeks,
                               min_row=idx["risk"], max_row=idx["risk"]),
                     from_rows=True, titles_from_data=False)
        ch3.series[0].tx = None
        ch3.set_categories(cats)
        ws.add_chart(ch3, "B21")

        ch4 = BarChart()
        ch4.title = f"PADD {padd} net cargo flows — ship tracking (kbd)"
        ch4.y_axis.title = "kbd (+in / -out)"
        ch4.height, ch4.width = 8.5, 17
        ch4.add_data(Reference(bal, min_col=2, max_col=1 + self.weeks,
                               min_row=idx["flows"], max_row=idx["flows"]),
                     from_rows=True, titles_from_data=False)
        ch4.series[0].tx = None
        ch4.set_categories(cats)
        ws.add_chart(ch4, "L21")

        ch5 = self._line("Net naphtha by refinery — top 10 by crude capacity (kbd/week)", "kbd")
        box_cats = Reference(boxes, min_col=self.c_net0,
                             max_col=self.c_net0 + self.weeks - 1, min_row=2, max_row=2)
        top = sorted(
            (r for r in self.data.refineries if r.crude_capacity_kbd > 0),
            key=lambda r: -r.crude_capacity_kbd,
        )[:10]
        for ref in top:
            total_row = self.refinery_total_rows[ref.refinery_id]
            ref_vals = Reference(boxes, min_col=self.c_net0,
                                 max_col=self.c_net0 + self.weeks - 1,
                                 min_row=total_row, max_row=total_row)
            s = Series(ref_vals, title=ref.name)
            s.marker = Marker(symbol="circle", size=5)
            ch5.series.append(s)
        ch5.set_categories(box_cats)
        ws.add_chart(ch5, "B39")

    # --------------------------------------------------------------- Checks

    def _sheet_checks(self) -> None:
        ws = self.wb.create_sheet("Checks")
        _banner(ws, 1, "Data checks — every check must PASS before numbers are trusted.", 6)
        _hdr(ws, 3, 1, "Check")
        _hdr(ws, 3, 2, "Failures")
        _hdr(ws, 3, 3, "Status")

        # valid-direction list used by check formulas
        for i, d in enumerate(["import", "export", "transfer_in", "transfer_out"]):
            ws.cell(row=2 + i, column=8, value=d).font = FONT_SMALL
        ws.column_dimensions["H"].hidden = True

        ref_ids = f"Refineries!$A${REF_LO + 1}:$A${REF_HI}"
        yield_checks = "+".join(
            f'COUNTIF(Boxes!${get_column_letter(self._cut_yld_col(i))}${BOX_LO}:'
            f'${get_column_letter(self._cut_yld_col(i))}${BOX_HI},">1.2")'
            f'+COUNTIF(Boxes!${get_column_letter(self._cut_yld_col(i))}${BOX_LO}:'
            f'${get_column_letter(self._cut_yld_col(i))}${BOX_HI},"<-1.2")'
            for i in range(len(self.cuts))
        )
        supply_neg = "+".join(
            f'COUNTIF(Balance!$B${idx["supply"]}:'
            f'${get_column_letter(1 + self.weeks)}${idx["supply"]},"<0")'
            for idx in self.balance_rows.values()
        )
        checks = [
            ("Duplicate refinery IDs in registry",
             f"=SUMPRODUCT((COUNTIF({ref_ids},{ref_ids})>1)*({ref_ids}<>\"\"))"),
            ("Boxes units pointing at unknown refinery IDs",
             f"=SUMPRODUCT((Boxes!$A${BOX_LO}:$A${BOX_HI}=\"UNIT\")"
             f"*(COUNTIF({ref_ids},Boxes!$B${BOX_LO}:$B${BOX_HI})=0))"),
            ("Utilization outside 0-110%",
             f'=COUNTIF(Boxes!$H${BOX_LO}:$H${BOX_HI},">1.1")'
             f'+COUNTIF(Boxes!$H${BOX_LO}:$H${BOX_HI},"<0")'),
            ("Yields outside +/-120% of throughput", f"={yield_checks}"),
            ("Negative unit capacities", f'=COUNTIF(Boxes!$F${BOX_LO}:$F${BOX_HI},"<0")'),
            ("Outages ending before they start",
             f"=SUMPRODUCT((Outages!$B${OUT_LO}:$B${OUT_HI}<>\"\")"
             f"*(Outages!$F${OUT_LO}:$F${OUT_HI}<Outages!$E${OUT_LO}:$E${OUT_HI}))"),
            ("Outage offline % outside 0-100",
             f'=COUNTIF(Outages!$G${OUT_LO}:$G${OUT_HI},">100")'
             f'+COUNTIF(Outages!$G${OUT_LO}:$G${OUT_HI},"<0")'),
            ("Outages pointing at unknown refinery IDs",
             f"=SUMPRODUCT((Outages!$B${OUT_LO}:$B${OUT_HI}<>\"\")"
             f"*(COUNTIF({ref_ids},Outages!$B${OUT_LO}:$B${OUT_HI})=0))"),
            ("Flows with invalid direction",
             f"=SUMPRODUCT((Flows!$C${FLOW_LO}:$C${FLOW_HI}<>\"\")"
             f"*(COUNTIF($H$2:$H$5,Flows!$C${FLOW_LO}:$C${FLOW_HI})=0))"),
            ("Flows missing a PADD",
             f"=SUMPRODUCT((Flows!$A${FLOW_LO}:$A${FLOW_HI}<>\"\")"
             f"*(Flows!$B${FLOW_LO}:$B${FLOW_HI}=\"\"))"),
            ("Weeks where a PADD's net refinery supply is negative "
             "(reformer pull exceeds naphtha make — check yields)",
             f"={supply_neg}"),
        ]
        r = 4
        for label, formula in checks:
            ws.cell(row=r, column=1, value=label).font = FONT_SMALL
            _style(ws.cell(row=r, column=2, value=formula), fill=FILL_CALC, fmt="0")
            _style(ws.cell(row=r, column=3, value=f'=IF(B{r}=0,"PASS","FAIL")'),
                   fill=FILL_CALC, bold=True)
            r += 1
        last = r - 1
        ws.cell(row=r + 1, column=1, value="OVERALL").font = Font(bold=True, size=12)
        master = ws.cell(row=r + 1, column=3,
                         value=f'=IF(SUM(B4:B{last})=0,"ALL CHECKS PASS","REVIEW FAILURES")')
        master.font = Font(bold=True, size=12)

        rng = f"C4:C{r + 1}"
        ws.conditional_formatting.add(
            rng, CellIsRule(operator="equal", formula=['"FAIL"'], fill=FILL_FAIL))
        ws.conditional_formatting.add(
            rng, CellIsRule(operator="equal", formula=['"REVIEW FAILURES"'], fill=FILL_FAIL))
        ws.conditional_formatting.add(
            rng, CellIsRule(operator="equal", formula=['"PASS"'], fill=FILL_PASS))
        ws.conditional_formatting.add(
            rng, CellIsRule(operator="equal", formula=['"ALL CHECKS PASS"'], fill=FILL_PASS))
        ws.column_dimensions["A"].width = 72
        ws.column_dimensions["B"].width = 10
        ws.column_dimensions["C"].width = 16

    # ---------------------------------------------------------------- README

    def _sheet_readme(self) -> None:
        ws = self.wb.create_sheet("README")
        _banner(ws, 1, "Zero-Based US Naphtha Model — desk workbook", 10)
        lines = [
            "",
            "WHAT THIS IS",
            "A bottom-up naphtha model: every refinery is built unit-by-unit (its 'box' on the",
            "Boxes tab), rolled up to PADD balances, with outages and ship-tracking flowing",
            "straight into the forward weekly forecast on the Dashboard.",
            "",
            "COLOR CODE",
            "", "", "",
            "HOW TO DRIVE IT",
            "1. Assumptions tab: PADD-level utilization & yields per unit type (blue).",
            "2. Boxes tab: per-unit orange override cells beat the PADD assumptions;",
            "   clear the cell to fall back. Capacity cells are blue inputs.",
            "3. Outages tab: add a row the moment an outage hits (planned or unplanned).",
            "   The weekly strip prorates partial weeks automatically.",
            "4. Scenario analysis: add rows with category = scenario, then flip",
            "   Model!B5 to YES to layer them on. Dashboard shows base vs case.",
            "5. Flows tab: ship-tracking cargoes (imports/exports/transfers) hit the",
            "   balance in the week of their date.",
            "6. Checks tab: must read ALL CHECKS PASS before trusting the numbers.",
            "",
            "MODEL CUT",
            "Net naphtha = sum over units of capacity x (1 - offline%) x utilization x yield.",
            "Yields are signed: reformers/isom consume naphtha (negative), so the total is",
            "net naphtha available to the market.",
            "",
            "PLACEHOLDER WARNING",
            "Registry, capacities, yields and demand are illustrative placeholders until the",
            "real capacity sheet and desk assumptions are loaded. Do not trade off them yet.",
            "",
            "Regenerate from the data/ directory:  python -m naphtha_model export --out <file>",
        ]
        for i, text in enumerate(lines, start=2):
            cell = ws.cell(row=i, column=1, value=text)
            cell.font = FONT_HDR if text.isupper() and text else FONT_SMALL
        # color legend swatches
        legend = [(FILL_INPUT, "BLUE — input: type here"),
                  (FILL_OVERRIDE, "ORANGE — manual override: beats assumptions; clear to revert"),
                  (FILL_CALC, "GREY — formula: do not type")]
        for j, (fill, label) in enumerate(legend):
            row = 9 + j
            ws.cell(row=row, column=2).fill = fill
            ws.cell(row=row, column=2).border = BORDER
            ws.cell(row=row, column=3, value=label).font = FONT_SMALL
        ws.column_dimensions["A"].width = 90


class SimpleWorkbook(DeskWorkbook):
    """The pared-down three-sheet model: Boxes, Assumptions, Data.

    - Boxes: one box per refinery; net naphtha = capacity x utilization x
      signed yields, live against the Assumptions sheet. No forward strip,
      no outage math — that stays in the Python engine for now.
    - Assumptions: per-PADD utilization/yield inputs plus the yield-mode
      cut split.
    - Data: the imported registry + 2024 yields. Typing a crude capacity
      here lights up that refinery's box (yield-mode rows read it live).
    """

    def build(self, out_path: Path) -> Path:
        self._sheet_assumptions()
        self._simple_share_block()
        self._sheet_data()
        self._simple_boxes()
        if "Sheet" in self.wb.sheetnames:
            del self.wb["Sheet"]
        self.wb._sheets = [self.wb[n] for n in ("Boxes", "Assumptions", "Data")]
        out_path = Path(out_path)
        self.wb.save(out_path)
        return out_path

    def _simple_share_block(self) -> None:
        """Yield-mode cut split inputs, placed right of the lookup tables."""
        ws = self.wb["Assumptions"]
        ws["A1"] = (
            "Assumptions — blue cells are inputs. Yields are % of unit throughput; "
            "consumers (reformer/isom) are NEGATIVE. Edits flow straight into the Boxes sheet."
        )
        col = 6 + len(self.cuts) + 10
        _hdr(ws, 3, col, "Yield-mode cut split")
        _hdr(ws, 3, col + 1, "share")
        shares = (self.book.global_cfg.get("yield_mode") or {}).get("cut_shares") or {}
        self.share_refs: dict[str, str] = {}
        for i, cut in enumerate(self.cuts):
            row = 4 + i
            ws.cell(row=row, column=col, value=cut).font = FONT_SMALL
            cell = ws.cell(row=row, column=col + 1, value=float(shares.get(cut, 0.0)))
            _style(cell, fill=FILL_INPUT, fmt="0%")
            self.share_refs[cut] = f"Assumptions!${get_column_letter(col + 1)}${row}"
        note = ws.cell(row=4 + len(self.cuts) + 1, column=col,
                       value="how a yield-mode refinery's 2024 net yield splits across cuts")
        note.font = FONT_SMALL

    def _sheet_data(self) -> None:
        from .loaders import load_yields_2024

        ws = self.wb.create_sheet("Data")
        _banner(
            ws, 1,
            "Imported data — refinery registry + 2024 net yields (% of crude). "
            "Type a crude capacity (blue) and that refinery's box lights up. "
            "Re-import: python -m naphtha_model ingest-yields <file>",
            15,
        )
        headers = ["refinery_id", "name", "owner", "padd", "state", "city",
                   "crude_capacity_kbd", "status", "gasoil_diesel", "gasoline",
                   "hfo", "kero_jet", "lpg", "naphtha", "total"]
        for c, h in enumerate(headers, start=1):
            _hdr(ws, 2, c, h)
        yields = load_yields_2024()
        y_keys = ["gasoil_diesel_pct", "gasoline_pct", "hfo_pct", "kero_jet_pct",
                  "lpg_pct", "naphtha_pct", "total_pct"]
        r = 3
        for ref in sorted(self.data.refineries,
                          key=lambda x: (x.padd, -x.crude_capacity_kbd, x.name)):
            y = yields.get(ref.refinery_id, {})
            _style(ws.cell(row=r, column=1, value=ref.refinery_id))
            _style(ws.cell(row=r, column=2, value=ref.name))
            _style(ws.cell(row=r, column=3, value=ref.owner))
            _style(ws.cell(row=r, column=4, value=ref.padd))
            _style(ws.cell(row=r, column=5, value=ref.state))
            _style(ws.cell(row=r, column=6, value=ref.city))
            _style(ws.cell(row=r, column=7, value=ref.crude_capacity_kbd),
                   fill=FILL_INPUT, fmt="0")
            _style(ws.cell(row=r, column=8, value=ref.status))
            for j, key in enumerate(y_keys):
                v = y.get(key)
                _style(ws.cell(row=r, column=9 + j,
                               value=float(v) / 100.0 if v not in (None, "") else None),
                       fmt="0.00%")
            r += 1
        self.data_last_row = r - 1
        widths = [22, 30, 30, 6, 13, 15, 10, 9, 12, 9, 8, 9, 8, 9, 8]
        for c, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(c)].width = w
        ws.freeze_panes = "A3"

    def _simple_boxes(self) -> None:
        ws = self.wb.create_sheet("Boxes")
        last_col = self.c_base
        _banner(
            ws, 1,
            "Refinery boxes — net naphtha = capacity x utilization x signed yields. "
            "Blue = input, orange = manual override (clear to fall back to the "
            "Assumptions sheet), grey = formula. Yield-mode rows read the Data sheet.",
            last_col,
        )
        labels = {
            self.c_type: "row", self.c_rid: "refinery_id", self.c_padd: "padd",
            self.c_uid: "unit", self.c_utype: "type", self.c_cap: "cap kbd",
            self.c_effcap: "eff cap", self.c_utilov: "util ovr", self.c_util: "util",
            self.c_ysum: "Σ yield", self.c_base: "net naphtha kbd",
        }
        for i, cut in enumerate(self.cuts):
            labels[self._cut_ovr_col(i)] = f"{cut} ovr"
            labels[self._cut_yld_col(i)] = f"{cut} yield"
        for col, text in labels.items():
            _hdr(ws, 2, col, text)

        util_val, util_padd = self.assump_refs["util"]
        y_val, y_padd, y_ut, y_cut = self.assump_refs["yield"]
        data_id = f"Data!$A$3:$A${self.data_last_row}"
        data_cap = f"Data!$G$3:$G${self.data_last_row}"
        data_naph = f"Data!$N$3:$N${self.data_last_row}"
        day1 = self.axis[0]
        row = 3
        for ref in sorted(self.data.refineries,
                          key=lambda x: (x.padd, -x.crude_capacity_kbd, x.name)):
            crude = (f"crude {ref.crude_capacity_kbd:,.0f} kbd"
                     if ref.crude_capacity_kbd else "crude: pending capacity sheet")
            _banner(ws, row,
                    f"  {ref.name}  [{ref.refinery_id}]   —   {ref.owner}   —   "
                    f"PADD {ref.padd}   —   {crude}", last_col)
            row += 1
            first_unit = row
            ovl = get_column_letter(self.c_utilov)
            for unit in ref.units:
                est = unit.unit_id == "CRUDE-EST"
                ws.cell(row=row, column=self.c_type, value="UNIT").font = FONT_SMALL
                _style(ws.cell(row=row, column=self.c_rid, value=ref.refinery_id))
                _style(ws.cell(row=row, column=self.c_padd, value=ref.padd))
                _style(ws.cell(row=row, column=self.c_uid, value=unit.unit_id))
                _style(ws.cell(row=row, column=self.c_utype,
                               value="est. yield" if est else unit.unit_type))
                if est:
                    _style(ws.cell(row=row, column=self.c_cap,
                                   value=f"=SUMIFS({data_cap},{data_id},$B{row})"),
                           fill=FILL_CALC, fmt="0")
                else:
                    _style(ws.cell(row=row, column=self.c_cap, value=unit.capacity_kbd),
                           fill=FILL_INPUT, fmt="0")
                _style(ws.cell(row=row, column=self.c_effcap,
                               value=self.eff_caps.get((ref.refinery_id, unit.unit_id))),
                       fill=FILL_CALC, fmt="0")

                ov_util = self.book._find_override(ref, unit, day1, "utilization")
                _style(ws.cell(row=row, column=self.c_utilov,
                               value=ov_util.value if ov_util else None),
                       fill=FILL_OVERRIDE, fmt="0.0%")
                _style(ws.cell(
                    row=row, column=self.c_util,
                    value=(f'=IF(${ovl}{row}<>"",${ovl}{row},'
                           f"SUMIFS({util_val},{util_padd},$C{row}))"),
                ), fill=FILL_CALC, fmt="0.0%")

                for i, cut in enumerate(self.cuts):
                    oc = get_column_letter(self._cut_ovr_col(i))
                    if est:
                        # 2024 net yield (Data sheet) split by the cut shares
                        _style(ws.cell(row=row, column=self._cut_ovr_col(i)),
                               fill=FILL_OVERRIDE, fmt="0.0%")
                        formula = (
                            f'=IF(${oc}{row}<>"",${oc}{row},'
                            f"SUMIFS({data_naph},{data_id},$B{row})"
                            f"*{self.share_refs[cut]})"
                        )
                    else:
                        ov = self.book._find_override(ref, unit, day1, "yield", cut=cut)
                        _style(ws.cell(row=row, column=self._cut_ovr_col(i),
                                       value=ov.value if ov else None),
                               fill=FILL_OVERRIDE, fmt="0.0%")
                        formula = (
                            f'=IF(${oc}{row}<>"",${oc}{row},'
                            f"SUMIFS({y_val},{y_padd},$C{row},{y_ut},$E{row},"
                            f'{y_cut},"{cut}"))'
                        )
                    _style(ws.cell(row=row, column=self._cut_yld_col(i), value=formula),
                           fill=FILL_CALC, fmt="0.00%")

                ysum = "+".join(
                    f"{get_column_letter(self._cut_yld_col(i))}{row}"
                    for i in range(len(self.cuts))
                )
                _style(ws.cell(row=row, column=self.c_ysum, value=f"={ysum}"),
                       fill=FILL_CALC, fmt="0.00%")
                cap_l = get_column_letter(self.c_cap)
                util_l = get_column_letter(self.c_util)
                ysum_l = get_column_letter(self.c_ysum)
                _style(ws.cell(row=row, column=self.c_base,
                               value=f"=${cap_l}{row}*${util_l}{row}*${ysum_l}{row}"),
                       fill=FILL_CALC, fmt="0.0")
                row += 1

            ws.cell(row=row, column=self.c_type, value="TOTAL").font = FONT_HDR
            _style(ws.cell(row=row, column=self.c_rid, value=ref.refinery_id),
                   fill=FILL_TOTAL, bold=True)
            _style(ws.cell(row=row, column=self.c_padd, value=ref.padd),
                   fill=FILL_TOTAL, bold=True)
            _style(ws.cell(row=row, column=self.c_uid, value="NET NAPHTHA"),
                   fill=FILL_TOTAL, bold=True)
            base_l = get_column_letter(self.c_base)
            _style(ws.cell(row=row, column=self.c_base,
                           value=f"=SUM({base_l}{first_unit}:{base_l}{row - 1})"),
                   fill=FILL_TOTAL, fmt="0.0", bold=True)
            self.refinery_total_rows[ref.refinery_id] = row
            row += 2

        widths = {1: 7, 2: 22, 3: 5, 4: 12, 5: 12, 6: 8, 7: 8, 8: 8, 9: 7}
        for i in range(len(self.cuts)):
            widths[self._cut_ovr_col(i)] = 8
            widths[self._cut_yld_col(i)] = 9
        widths[self.c_ysum] = 9
        widths[self.c_base] = 14
        for col, w in widths.items():
            ws.column_dimensions[get_column_letter(col)].width = w
        ws.freeze_panes = "A3"


def build_desk_workbook(
    data: ModelData, axis: list[date], out_path: Path, scenarios=None
) -> Path:
    return DeskWorkbook(data, axis, scenarios=scenarios).build(out_path)


def build_simple_workbook(data: ModelData, axis: list[date], out_path: Path) -> Path:
    return SimpleWorkbook(data, axis).build(out_path)
