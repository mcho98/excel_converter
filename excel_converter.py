#!/usr/bin/env python3
"""Convert a time clock export into the monthly client claim workbook.

Replaces the old VBA macros (tt_1_15 / tt_16_31 / reset).

    python excel_converter.py TIMECLOCK.xlsx --chart CLAIM_OR_CHART.xlsx
    python excel_converter.py                      # no arguments: file pickers

What it does
1. Reads the time clock export (one header row per client, shift rows below it).
2. Reads the "Family# Client# Chart" sheet (family -> clients, AHS IDs, monthly hours).
3. Writes a NEW workbook with one 26-row claim block per client, a Summary sheet,
   an Input sheet (every shift that was read) and an Issues sheet.

Nothing is modified in place and there is no state to reset: every run builds the
output from scratch from the two input files.
"""
from __future__ import annotations

import argparse
import calendar
import collections
import datetime as dt
import json
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as col

# --- layout of the claim sheet (matches the old workbook) -------------------
FIRST_BLOCK_ROW = 15      # row of the first block's family/client numbers
BLOCK_PITCH = 26          # rows from one block to the next
FIRST_DAY_COL = 5         # column E = day 1
LAST_DAY_COL = 35         # column AI = day 31
SUM_COL = 36              # AJ: row total
GRAND_COL = 37            # AK: all-category total
DEFAULT_RATE = 34.03      # $/hr, same as the old template
DEFAULT_CHART_SHEET = "Family# Client# Chart"
# config/ folder beside the script (or beside the .exe when packaged with PyInstaller)
APP_DIR = Path(sys.executable if getattr(sys, "frozen", False) else __file__).resolve().parent
if APP_DIR.parts[-3:-1] == ("excel_converter.app", "Contents"):      # macOS .app: use the folder holding the .app
    APP_DIR = APP_DIR.parents[2]
CONFIG_FILE = APP_DIR / "config" / "excel_converter.json"
OUTPUT_DIR = APP_DIR / "output"

BOLD = Font(bold=True)
WARN_FILL = PatternFill("solid", fgColor="FFF2CC")
WARN_FONT = Font(bold=True, color="C00000")
HEAD_FILL = PatternFill("solid", fgColor="D9E1F2")

# Claim sheet colors, copied from the old workbook (theme colors resolved to RGB).
YELLOW, GREY = "FFFF00", "E7E6E6"
GREEN_L, GREEN_M, GREEN_D = "E2F0D9", "C5E0B4", "A9D18E"
ORANGE_L, ORANGE_M = "FBE5D6", "F8CBAD"
BLUE_L, BLUE_M = "DEEBF7", "BDD7EE"
# (row offset in block, first column, last column, color); later entries paint over earlier ones
BLOCK_FILLS = [
    (0, 1, 2, YELLOW),                                               # family / client
    (1, 1, 2, GREEN_L),
    (2, 1, 2, ORANGE_L), (2, 4, 4, BLUE_L),                          # client name
    (3, 1, 2, ORANGE_M), (3, 3, 4, BLUE_L),
    (4, 3, 4, BLUE_L),                                               # rate
    (8, GRAND_COL, GRAND_COL, GREEN_M),
    *[(o, 3, 4, BLUE_M) for o in range(9, 19)],                      # AHS#, claim/day, $ labels
    (10, FIRST_DAY_COL, LAST_DAY_COL, BLUE_L), (10, SUM_COL, SUM_COL, YELLOW),   # daily total
    (12, 4, 4, GREEN_D), (12, FIRST_DAY_COL, LAST_DAY_COL, BLUE_L),             # claim/day
    (12, SUM_COL, GRAND_COL, GREEN_M),
]


class InputError(Exception):
    """A problem with an input file that the user needs to fix."""


# --- data -------------------------------------------------------------------
@dataclass
class Shift:
    row: int
    start: dt.date | None
    end: dt.date | None
    hours: float


@dataclass
class ClientEntry:
    row: int
    header: str
    family: int | None
    total: float | None
    shifts: list[Shift] = field(default_factory=list)


@dataclass
class ChartRecord:
    row: int
    family: int
    client: int
    name: str
    ahs_id: object
    homemaking: float | None
    personal_care: float | None
    respite: float | None


@dataclass
class Block:
    family: int | None
    client: int | None
    name: str
    ahs_id: object
    monthly: tuple
    daily: dict          # day-of-month -> hours (only on the client that receives the hours)
    note: str = ""


@dataclass
class Issue:
    level: str           # "ERROR" (data is missing/wrong in the output) or "WARN"
    where: str
    message: str


# --- small helpers ----------------------------------------------------------
def to_int(v):
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, int):
        return v
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, str) and v.strip().isdigit():
        return int(v.strip())
    return None


def to_float(v):
    if isinstance(v, bool) or v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        try:
            return float(v.strip())
        except ValueError:
            return None
    return None


def to_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, str):
        for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y"):
            try:
                return dt.datetime.strptime(v.strip(), fmt).date()
            except ValueError:
                pass
    return None


def clean(text) -> str:
    """Collapse whitespace (incl. non-breaking spaces) in a name."""
    return re.sub(r"\s+", " ", str(text).replace("\xa0", " ")).strip()


def norm(text) -> str:
    """Header text for matching: lower case, no spaces ('AHS ID #' == 'ahs id#')."""
    return re.sub(r"\s+", "", clean(text)).lower()


def find_columns(header_row, wanted: dict, where: str) -> dict:
    """Map logical names to column indexes by header text, not by fixed letters."""
    labels = {norm(v): i for i, v in enumerate(header_row) if v not in (None, "")}
    found, missing = {}, []
    for key, label in wanted.items():
        if norm(label) in labels:
            found[key] = labels[norm(label)]
        else:
            missing.append(label)
    if missing:
        raise InputError(
            f"{where}: expected column(s) {missing} but found {sorted(labels)}. "
            "Is this the right file?"
        )
    return found


# --- readers ----------------------------------------------------------------
def read_timeclock(path: Path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    issues: list[Issue] = []
    if len(wb.worksheets) > 1:
        issues.append(Issue("WARN", path.name,
                            f"Workbook has {len(wb.worksheets)} sheets; only the first "
                            f"('{wb.worksheets[0].title}') was read."))
    rows = wb.worksheets[0].iter_rows(values_only=True)
    try:
        header = next(rows)
    except StopIteration:
        raise InputError(f"{path.name}: the sheet is empty.")
    cols = find_columns(header, {
        "name": "Client Number and Name", "total": "Total Hours",
        "start": "Start Date", "end": "End Date", "hours": "Shift hours",
    }, path.name)

    entries: list[ClientEntry] = []
    cur: ClientEntry | None = None
    for rownum, row in enumerate(rows, start=2):
        row = list(row) + [None] * (max(cols.values()) + 1 - len(row))
        name = row[cols["name"]]
        if name not in (None, ""):
            fam = re.match(r"\s*(\d+)(?=\s|\D|$)", str(name))
            cur = ClientEntry(rownum, clean(name), int(fam.group(1)) if fam else None,
                              to_float(row[cols["total"]]))
            if cur.family is None:
                issues.append(Issue("ERROR", f"{path.name} row {rownum}",
                                    f"Header '{cur.header}' does not start with a family number; "
                                    "its hours are NOT in the claim sheet."))
            entries.append(cur)
            continue

        hours = to_float(row[cols["hours"]])
        start, end = to_date(row[cols["start"]]), to_date(row[cols["end"]])
        if hours is None and start is None:
            continue                       # blank spacer or the export's grand-total row
        if start is None and end is None:
            continue                       # total row: a number but no date (this row used to
                                           # be added to the last client by the macro)
        if cur is None:
            issues.append(Issue("ERROR", f"{path.name} row {rownum}",
                                "Shift row appears before any client header; skipped."))
            continue
        if hours is None:
            issues.append(Issue("ERROR", f"{path.name} row {rownum}",
                                f"Shift for '{cur.header}' has no hours; skipped."))
            continue
        cur.shifts.append(Shift(rownum, start, end, hours))
    wb.close()
    if not entries:
        raise InputError(f"{path.name}: no client header rows found.")
    return entries, issues


def read_chart(path: Path, sheet: str):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    if sheet not in wb.sheetnames:
        raise InputError(f"{path.name}: no sheet named '{sheet}'. Sheets: {wb.sheetnames}. "
                         "Use --chart-sheet to name it.")
    rows = wb[sheet].iter_rows(values_only=True)
    header_idx, header = None, None
    for i, row in enumerate(rows, start=1):       # header is on row 3 today; search rows 1-10
        if any(norm(v) == "family#" for v in row if v is not None):
            header_idx, header = i, row
            break
        if i >= 10:
            break
    if header is None:
        raise InputError(f"{path.name}/{sheet}: could not find a 'Family #' header row.")
    cols = find_columns(header, {
        "family": "Family #", "client": "Client #", "name": "Name", "ahs": "AHS ID #",
        "hm": "Homemaking hours", "pc": "Personal Care Hours", "rs": "Respite hours",
    }, f"{path.name}/{sheet}")
    need = max(cols.values()) + 1
    chart: dict[int, list[ChartRecord]] = collections.defaultdict(list)
    issues: list[Issue] = []
    seen = set()
    for rownum, row in enumerate(rows, start=header_idx + 1):
        row = list(row) + [None] * (need - len(row))
        fam, cli = to_int(row[cols["family"]]), to_int(row[cols["client"]])
        if fam is None or cli is None:
            continue                       # continuation rows under each client are blank
        if (fam, cli) in seen:
            issues.append(Issue("WARN", f"{sheet} row {rownum}",
                                f"Family {fam} client {cli} appears more than once; first one used."))
            continue
        seen.add((fam, cli))
        chart[fam].append(ChartRecord(
            rownum, fam, cli, clean(row[cols["name"]] or ""), row[cols["ahs"]],
            to_float(row[cols["hm"]]), to_float(row[cols["pc"]]), to_float(row[cols["rs"]])))
    wb.close()
    if not chart:
        raise InputError(f"{path.name}/{sheet}: no family/client rows found.")
    return chart, issues


# --- planning ---------------------------------------------------------------
def pick_month(entries, date_field: str, forced: str | None):
    if forced:
        try:
            y, m = (int(p) for p in forced.split("-"))
            return y, m
        except ValueError:
            raise InputError(f"--month must look like 2026-09, got '{forced}'.")
    counts = collections.Counter()
    for e in entries:
        for s in e.shifts:
            d = s.start if date_field == "start" else s.end
            if d:
                counts[(d.year, d.month)] += 1
    if not counts:
        raise InputError("No dated shifts found, so the month can't be worked out. Use --month.")
    return counts.most_common(1)[0][0]


def plan(entries, chart, year, month, date_field):
    """Return (blocks, per-entry summary rows, input rows, issues)."""
    issues: list[Issue] = []
    by_family: dict[int, list[ClientEntry]] = collections.defaultdict(list)
    for e in entries:
        if e.family is not None:
            by_family[e.family].append(e)

    summary, input_rows = [], []
    daily_by_family: dict[int, dict[int, float]] = {}
    for e in entries:
        placed = outside = 0.0
        for s in e.shifts:
            d = s.start if date_field == "start" else s.end
            if d is None:
                issues.append(Issue("ERROR", f"time clock row {s.row}",
                                    f"{e.header}: shift has no {date_field} date; hours not placed."))
                outside += s.hours
                counted = None
            elif (d.year, d.month) != (year, month):
                outside += s.hours
                counted = None
                issues.append(Issue("WARN", f"time clock row {s.row}",
                                    f"{e.header}: {s.hours:g} h dated {d} is outside {year}-{month:02d}; "
                                    "not in the claim sheet."))
            else:
                placed += s.hours
                counted = d
                if e.family is not None:
                    daily_by_family.setdefault(e.family, {})
                    daily_by_family[e.family][d.day] = daily_by_family[e.family].get(d.day, 0) + s.hours
            input_rows.append((e.family, e.header, s.start, s.end, s.hours, s.row, counted))
        in_report = placed if e.family is not None else 0.0
        if e.family is None:
            outside = sum(s.hours for s in e.shifts)
        summary.append([e.family, e.header, e.total, len(e.shifts), round(in_report, 2),
                        round(outside, 2)])
        if e.total is not None:
            diff = round(sum(s.hours for s in e.shifts) - e.total, 2)
            if abs(diff) > 0.05:
                issues.append(Issue("WARN", f"time clock row {e.row}",
                                    f"{e.header}: shifts add to {sum(s.hours for s in e.shifts):.2f} h "
                                    f"but the export's Total Hours says {e.total:g}."))

    blocks: list[Block] = []
    for fam in sorted(by_family):
        group = by_family[fam]
        if len(group) > 1:
            issues.append(Issue("WARN", f"family {fam}",
                                f"{len(group)} time clock headers share this family number "
                                f"({'; '.join(g.header for g in group)}); their hours were combined."))
        name = group[0].header
        recs = chart.get(fam)
        daily = daily_by_family.get(fam, {})
        if not recs:
            issues.append(Issue("ERROR", f"family {fam}",
                                f"'{name}' is not in the chart. A block was created with no client "
                                "number, AHS ID or monthly hours. Add the family to the chart and re-run."))
            blocks.append(Block(fam, None, name, None, (None, None, None), daily,
                                "NOT IN CHART - add family to chart and re-run"))
            continue
        for i, r in enumerate(recs):
            note = ""
            if len(recs) > 1:
                note = ("hours for the whole family are on this client's block" if i == 0 else
                        "second client in family - hours NOT split, enter manually")
            blocks.append(Block(fam, r.client, name, r.ahs_id,
                                (r.homemaking, r.personal_care, r.respite),
                                daily if i == 0 else {}, note))
    return blocks, summary, input_rows, issues


# --- writer -----------------------------------------------------------------
def paint(ws, row, first_col, last_col, rgb):
    fill = PatternFill("solid", fgColor=rgb)
    for c in range(first_col, last_col + 1):
        ws.cell(row, c).fill = fill


def write_block(ws, top, b: Block, rate, days_in_month):
    r = lambda n: top + n          # noqa: E731  (row helper)
    for off, c1, c2, rgb in BLOCK_FILLS:
        paint(ws, r(off), c1, c2, rgb)
    ws.cell(top, 1, b.family)
    ws.cell(top, 2, b.client)
    ws.cell(top, 3, b.ahs_id)
    ws.cell(top, 1).font = ws.cell(top, 2).font = BOLD
    if b.note:
        c = ws.cell(top, FIRST_DAY_COL, b.note)
        c.font = WARN_FONT
        for k in (1, 2, 3):
            ws.cell(top, k).fill = WARN_FILL
    ws.cell(r(2), 2, "Client Name")
    ws.cell(r(2), 3, b.name).font = BOLD
    ws.cell(r(2), 4, "Days in month")
    ws.cell(r(3), 4, days_in_month)
    ws.cell(r(4), 3, "rate $/hr")
    ws.cell(r(4), 4, rate)

    labels = ["Homemaking hours", "Personal Care Hours", "Respite hours"]
    for k in range(3):
        ws.cell(r(5 + k), 2, labels[k])
        if b.monthly[k] is not None:
            ws.cell(r(5 + k), 4, b.monthly[k])
    ws.cell(r(8), 2, f"=SUM(D{r(5)}:D{r(7)})")
    ws.cell(r(8), 4, "Adj. hrs")
    ws.cell(r(9), 2, "AHS#")
    ws.cell(r(9), 3, f"=C{top}")
    ws.cell(r(10), 4, "Total")
    ws.cell(r(12), 4, "Claim/day")
    for k, lab in enumerate(["Homemaking $$", "Personal Care $$", "Respite $$", "TotaL $$ "]):
        ws.cell(r(14 + k), 4, lab)

    for day, hrs in b.daily.items():                       # personal care row, as the macro did
        c = ws.cell(r(6), FIRST_DAY_COL + day - 1, round(hrs, 2))
        c.number_format = "0.00"
    rate_ref = f"$D{r(4)}"
    for c in range(FIRST_DAY_COL, LAST_DAY_COL + 1):
        L = col(c)
        ws.cell(r(10), c, f"=SUM({L}{r(5)}:{L}{r(8)})")
        ws.cell(r(12), c, f"={L}{r(10)}*{rate_ref}")
        for k in range(3):
            ws.cell(r(14 + k), c, f"={L}{r(5 + k)}*{rate_ref}")
        ws.cell(r(17), c, f"=SUM({L}{r(14)}:{L}{r(16)})")
        for rr in (10, 12, 14, 15, 16, 17):
            ws.cell(r(rr), c).number_format = "0.00"
    first, last = col(FIRST_DAY_COL), col(LAST_DAY_COL)
    for rr in (5, 6, 7, 8, 10, 12):
        ws.cell(r(rr), SUM_COL, f"=SUM({first}{r(rr)}:{last}{r(rr)})").number_format = "0.00"
    ws.cell(r(7), GRAND_COL, f"=SUM({col(SUM_COL)}{r(5)}:{col(SUM_COL)}{r(7)})")
    ws.cell(r(12), GRAND_COL, f"={col(SUM_COL)}{r(10)}*{rate_ref}").number_format = "0.00"


def write_workbook(out_path, blocks, summary, input_rows, issues, year, month, rate, date_field):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Claim {year}-{month:02d}"
    days = calendar.monthrange(year, month)[1]

    ws["A14"], ws["B14"] = "Family Number", "Client Number"
    ws["D14"] = rate
    ws.cell(14, FIRST_DAY_COL, dt.date(year, month, 1)).number_format = "m/d"
    for c in range(FIRST_DAY_COL + 1, LAST_DAY_COL + 1):
        prev = col(c - 1)
        ws.cell(14, c, f'=IFERROR(IF({prev}14=EOMONTH({prev}14,0),"",{prev}14+1),"")').number_format = "m/d"
    for c in range(FIRST_DAY_COL, LAST_DAY_COL + 1):
        ws.cell(13, c, f'=TEXT({col(c)}14,"ddd")').alignment = Alignment(horizontal="center")
    for c in ("A14", "B14", "D14"):
        ws[c].font = BOLD
    paint(ws, 13, FIRST_DAY_COL, LAST_DAY_COL, GREY)
    paint(ws, 14, 4, LAST_DAY_COL, GREEN_L)
    ws.cell(14, FIRST_DAY_COL).fill = PatternFill("solid", fgColor=YELLOW)
    ws["A11"] = f"Hours counted by shift {date_field} date. Generated by excel_converter.py."
    ws["A11"].font = Font(italic=True, color="666666")

    for i, b in enumerate(blocks):
        write_block(ws, FIRST_BLOCK_ROW + i * BLOCK_PITCH, b, rate, days)
    for c, w in zip("ABCD", (13, 14, 30, 14)):
        ws.column_dimensions[c].width = w
    for c in range(FIRST_DAY_COL, GRAND_COL + 1):
        ws.column_dimensions[col(c)].width = 7.5
    ws.freeze_panes = "E15"

    s = wb.create_sheet("Summary")
    s.append(["Family", "Time clock header", "Export 'Total Hours'", "Shift rows",
              "Hours in claim sheet", "Hours NOT in claim sheet"])
    for row in summary:
        s.append(row)
    n = len(summary) + 1
    s.append(["TOTAL", None, f"=SUM(C2:C{n})", f"=SUM(D2:D{n})", f"=SUM(E2:E{n})", f"=SUM(F2:F{n})"])
    for c in s[1]:
        c.font, c.fill = BOLD, HEAD_FILL
    for c in s[n + 1]:
        c.font = BOLD
    for c, w in zip("ABCDEF", (9, 48, 20, 11, 20, 24)):
        s.column_dimensions[c].width = w
    s.freeze_panes = "A2"

    i = wb.create_sheet("Input")
    i.append(["Family", "Time clock header", "Start Date", "End Date", "Shift hours",
              "Time clock row", "Date used"])
    for row in sorted(input_rows, key=lambda x: (x[0] is None, x[0] or 0, x[5])):
        i.append(list(row))
    for c in i[1]:
        c.font, c.fill = BOLD, HEAD_FILL
    for r_ in i.iter_rows(min_row=2, min_col=3, max_col=4):
        for c in r_:
            c.number_format = "yyyy-mm-dd"
    for r_ in i.iter_rows(min_row=2, min_col=7, max_col=7):
        for c in r_:
            c.number_format = "yyyy-mm-dd"
    for c, w in zip("ABCDEFG", (9, 48, 12, 12, 12, 14, 12)):
        i.column_dimensions[c].width = w
    i.freeze_panes = "A2"

    x = wb.create_sheet("Issues")
    x.append(["Level", "Where", "Message"])
    for iss in sorted(issues, key=lambda z: (z.level != "ERROR", z.where)):
        x.append([iss.level, iss.where, iss.message])
        if iss.level == "ERROR":
            x.cell(x.max_row, 1).font = WARN_FONT
    if not issues:
        x.append(["OK", "", "No issues found."])
    for c in x[1]:
        c.font, c.fill = BOLD, HEAD_FILL
    for c, w in zip("ABC", (8, 26, 120)):
        x.column_dimensions[c].width = w
    wb.save(out_path)


# --- command line / GUI -----------------------------------------------------
def run(timeclock: Path, chart_path: Path, chart_sheet: str, out: Path | None,
        month: str | None, rate: float, date_field: str):
    entries, issues = read_timeclock(timeclock)
    chart, chart_issues = read_chart(chart_path, chart_sheet)
    issues += chart_issues
    year, mon = pick_month(entries, date_field, month)
    blocks, summary, input_rows, plan_issues = plan(entries, chart, year, mon, date_field)
    issues += plan_issues
    if out is None:
        OUTPUT_DIR.mkdir(exist_ok=True)
        out = OUTPUT_DIR / f"{dt.datetime.now():%Y-%m-%d_%H%M%S}.xlsx"
    write_workbook(out, blocks, summary, input_rows, issues, year, mon, rate, date_field)
    total_in = sum(r[4] for r in summary)
    total_out = sum(r[5] for r in summary)
    return out, dict(blocks=len(blocks), clients=len(entries), month=f"{year}-{mon:02d}",
                     hours_in=round(total_in, 2), hours_out=round(total_out, 2),
                     errors=sum(i.level == "ERROR" for i in issues),
                     warnings=sum(i.level == "WARN" for i in issues))


def config_chart() -> str | None:
    """Chart path saved in config/excel_converter.json, if it is set and the file exists."""
    try:
        chart = json.loads(CONFIG_FILE.read_text()).get("chart")
    except (OSError, ValueError, AttributeError):
        return None
    return chart if chart and Path(chart).exists() else None


def gui_pick(args):
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox
    except ImportError:
        sys.exit("No file pickers available here. Run: python excel_converter.py TIMECLOCK.xlsx --chart FILE.xlsx")
    root = tk.Tk()
    root.withdraw()
    ft = [("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")]
    tc = filedialog.askopenfilename(title="Time clock export", filetypes=ft)
    if not tc:
        sys.exit(0)
    chart = args.chart or config_chart()
    if not chart:
        chart = filedialog.askopenfilename(title="Workbook containing the Family# Client# Chart sheet",
                                           filetypes=ft)
        if not chart:
            sys.exit(0)
        try:
            CONFIG_FILE.parent.mkdir(exist_ok=True)
            CONFIG_FILE.write_text(json.dumps({"chart": chart}))
        except OSError:
            pass
    args.timeclock, args.chart = tc, chart
    return root, messagebox


def main(argv=None):
    p = argparse.ArgumentParser(description="Build the client claim workbook from a time clock export.")
    p.add_argument("timeclock", nargs="?", help="time clock export (.xlsx)")
    p.add_argument("--chart", help="workbook with the 'Family# Client# Chart' sheet (.xlsx/.xlsm); "
                                   "default: the \"chart\" path in config/excel_converter.json")
    p.add_argument("--chart-sheet", default=DEFAULT_CHART_SHEET)
    p.add_argument("-o", "--out", help="output file (default: output/<date>_<time>.xlsx beside the script)")
    p.add_argument("--month", help="claim month as YYYY-MM (default: the month most shifts fall in)")
    p.add_argument("--rate", type=float, default=DEFAULT_RATE, help=f"$/hr (default {DEFAULT_RATE})")
    p.add_argument("--date-field", choices=("start", "end"), default="start",
                   help="which date a shift counts on; matters for overnight shifts (default: start)")
    args = p.parse_args(argv)

    gui = None
    if not args.timeclock:
        gui = gui_pick(args)
    elif not args.chart:
        args.chart = config_chart()
        if not args.chart:
            p.error(f"--chart is required (or set \"chart\" to an existing file in {CONFIG_FILE})")
    try:
        out, stats = run(Path(args.timeclock), Path(args.chart), args.chart_sheet,
                         Path(args.out) if args.out else None, args.month, args.rate, args.date_field)
    except InputError as e:
        if gui:
            gui[1].showerror("Cannot convert", str(e))
        sys.exit(f"ERROR: {e}")
    msg = (f"Wrote {out}\n{stats['blocks']} claim blocks for {stats['clients']} time clock clients "
           f"({stats['month']}); {stats['hours_in']:g} h in the claim sheet, "
           f"{stats['hours_out']:g} h left out.\n"
           f"{stats['errors']} error(s), {stats['warnings']} warning(s) - see the Issues sheet.")
    print(msg)
    if gui:
        gui[1].showinfo("Done", msg)
    return 1 if stats["errors"] else 0


if __name__ == "__main__":
    sys.exit(main())
