#!/usr/bin/env python3
"""
twig-gpt.py — TeacherWIse [timetable] Generator (modernized)

This script:
1) Generates a TEACHERWISE timetable from a CLASSWISE timetable in the same workbook.
2) Generates per-class printable sheets into a (possibly separate) workbook.
3) Compares two CLASSWISE timetables and highlights differences.

Author (original): Sunil Sangwal <sunil.sangwal@gmail.com>
Modernized refactor: 2025-09-05

License: GPLv3 (same as original)
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import argparse
import logging
import re
import sys
import time
from typing import Dict, List, Tuple, Iterable, Optional

import openpyxl
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
from openpyxl.utils import get_column_letter


# ----------------------------- Logging ---------------------------------------

def setup_logging(verbosity: int) -> None:
    level = logging.WARNING if verbosity == 0 else logging.INFO if verbosity == 1 else logging.DEBUG
    logging.basicConfig(
        level=level,
        format="%(levelname)s: %(message)s",
    )


# ----------------------------- Constants -------------------------------------

DAYS_ALL: List[int] = [1, 2, 3, 4, 5, 6]
PERIOD_COLUMNS = range(2, 10)  # columns B..I → periods 1..8
TEACHERS_SHEET = "TEACHERS"
CLASSWISE_SHEET = "CLASSWISE"
TEACHERWISE_SHEET = "TEACHERWISE"
MASTER_SHEET = "MASTER"

# SUBJECT (1-3,5-6) TCODE
RE_CLASSWISE_LINE = re.compile(
    r"^(?P<subject>[\w \-.]+)\s*\((?P<days>[1-6,\- ]+)\)\s*(?P<teacher>[A-Z]+)$",
    re.IGNORECASE,
)
# CLASS (1-3,5-6) SUBJECT
RE_TEACHERWISE_LINE = re.compile(
    r"^(?P<class_name>[\w]+)\s*\((?P<days>.*)\)\s*(?P<subject>[\w \-.]+)$",
    re.IGNORECASE,
)


# ----------------------------- Data Models -----------------------------------

@dataclass(frozen=True)
class PeriodEntry:
    period: int            # 1..8 (maps to sheet column 2..9)
    klass: str             # "10A" or alias if provided
    days_raw: str          # "1-3, 5-6" (not expanded)
    subject: str


@dataclass
class Context:
    separator: str
    expand_fullnames: bool
    keep_timestamp: bool


# ----------------------------- Utilities -------------------------------------

def now_stamp() -> str:
    return "Last updated on " + time.ctime()


def escape_visible(s: str) -> str:
    return s.replace("\n", r"\n").replace("\t", r"\t")


def expand_days(days: str) -> List[int]:
    """
    Input: "1-2, 4, 6-5"
    Output: [1, 2, 4, 5, 6]
    """
    spans = [g.strip() for g in days.split(",")] if "," in days else [days.strip()]
    out: List[int] = []
    for span in spans:
        if "-" in span:
            a, b = (int(x) for x in span.split("-", 1))
            start, end = (a, b) if a <= b else (b, a)
            out.extend(range(start, end + 1))
        else:
            if span:
                out.append(int(span))
    return out


def compress_days(days: Iterable[int]) -> str:
    """
    Input: [1,2,3,5,6] → "1-3, 5-6"
    """
    ds = sorted(set(days))
    if not ds:
        return ""
    start = ds[0]
    prev = ds[0]
    chunks: List[str] = []
    for d in ds[1:]:
        if d == prev + 1:
            prev = d
            continue
        chunks.append(f"{start}" if start == prev else f"{start}-{prev}")
        start = prev = d
    chunks.append(f"{start}" if start == prev else f"{start}-{prev}")
    return ", ".join(chunks)


def count_days(days: str) -> int:
    return len(set(expand_days(days)))


# ----------------------- TEACHERS sheet helpers ------------------------------

def load_teacher_names(wb: Workbook) -> Dict[str, str]:
    """
    TEACHERS sheet columns:
    A: CODE, B: NAME, (others optional)
    """
    if TEACHERS_SHEET not in wb:
        return {}
    ws = wb[TEACHERS_SHEET]
    names: Dict[str, str] = {}
    row = 2
    while True:
        code = ws.cell(row, 1).value
        if not code:
            break
        if code in names:
            raise ValueError(f"Duplicate teacher code '{code}' in TEACHERS.")
        names[str(code)] = ws.cell(row, 2).value or str(code)
        row += 1
    return names


def load_teacher_details(wb: Workbook) -> Dict[str, Dict[str, str]]:
    """
    Returns: { CODE: { header:value, ... } }, taking headers from row 1.
    """
    if TEACHERS_SHEET not in wb:
        return {}
    ws = wb[TEACHERS_SHEET]
    headers = []
    col = 1
    while True:
        h = ws.cell(1, col).value
        if not h:
            break
        headers.append(str(h).strip().upper())
        col += 1

    details: Dict[str, Dict[str, str]] = {}
    row = 2
    while True:
        code = ws.cell(row, 1).value
        if not code:
            break
        if code in details:
            raise ValueError(f"Duplicate teacher code '{code}' in TEACHERS.")
        details[str(code)] = {}
        for i, h in enumerate(headers, start=1):
            details[str(code)][h] = ws.cell(row, i).value
        row += 1
    return details


# ----------------------------- Clash detection -------------------------------

def class_number(klass: str) -> str:
    """'10A' -> '10' ; '7B' -> '7'"""
    return klass[:-1] if klass and klass[-1].isalpha() else klass


def highlight_clashes(ws: Worksheet, separator: str) -> int:
    """
    Reads TEACHERWISE and prepends '**CLASH** [days]:\\n' to offending cells.
    Clash: same teacher has same 'class-number + subject' overlapping on a day.
    """
    CLASH_MARK = "**CLASH** "
    total = 0

    row = 2
    while True:
        if not ws.cell(row=row, column=1).value:
            break
        for col in PERIOD_COLUMNS:
            cell_val = ws.cell(row=row, column=col).value
            if not cell_val:
                continue

            entries_by_day: Dict[int, List[str]] = {}
            for raw_line in str(cell_val).split(separator):
                line = raw_line.strip()
                if not line:
                    continue
                m = RE_TEACHERWISE_LINE.match(line)
                if not m:
                    logging.warning(
                        "Format issue in TEACHERWISE %s%d: %s",
                        get_column_letter(col), row, line
                    )
                    continue
                klass, days, subject = m.group("class_name", "days", "subject")
                for d in expand_days(days):
                    entries_by_day.setdefault(d, []).append(f"{class_number(klass)}-{subject.strip()}")

            clash_days = [d for d, items in entries_by_day.items() if len(set(items)) > 1]
            if clash_days:
                total += len(clash_days)
                ws.cell(row=row, column=col).value = (
                    f"{CLASH_MARK}{clash_days}:\n" + str(cell_val)
                )
        row += 1
    return total


# -------------------------- Sheet formatting helpers -------------------------

def clear_sheet(ws: Worksheet) -> None:
    row = 2
    while True:
        empty_next = ws.cell(row=row, column=1).value in (None, "")
        for col in range(1, 12):  # include the extra daywise/periods column too
            ws.cell(row=row, column=col).value = ""
        row += 1
        if empty_next:
            break


def format_master_ws(ws: Worksheet) -> None:
    ws.column_dimensions["A"].width = 16
    for col in PERIOD_COLUMNS:
        ws.column_dimensions[get_column_letter(col)].width = 14

    for r in range(1, 4):
        ws.row_dimensions[r].height = 34
    for r in range(4, 10):
        ws.row_dimensions[r].height = 54

    for col in range(1, 10):
        ws[f"{get_column_letter(col)}3"].fill = PatternFill(start_color="c3c3c3", end_color="c3c3c3", fill_type="solid")
    for r in range(4, 10):
        ws[f"A{r}"].fill = PatternFill(start_color="c3c3c3", end_color="c3c3c3", fill_type="solid")

    ws.merge_cells("A1:I1")
    ws.merge_cells("A2:D2")
    ws.merge_cells("E2:I2")

    ws["A1"].font = Font(size=25)
    ws["A2"].font = Font(size=16)
    ws["E2"].font = Font(size=16)

    center_top = Alignment(horizontal="center", vertical="top", wrap_text=True)
    ws["A1"].alignment = center_top
    ws["A2"].alignment = Alignment(horizontal="left", vertical="top")
    ws["E2"].alignment = Alignment(horizontal="right", vertical="top")

    thin = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    for r in range(3, 10):
        for c in range(1, 10):
            cell = ws.cell(r, c)
            cell.border = thin
            cell.alignment = center_top


# ----------------------------- Core generators -------------------------------

def remove_comment(line: str) -> str:
    return line.split("#", 1)[0].strip()


def count_periods_for_teacher(teacher: str, timetable: Dict[str, List[PeriodEntry]]) -> int:
    # per period (1..8) collect unique days
    per_period_days: Dict[int, List[int]] = {}
    for entry in timetable.get(teacher, []):
        per_period_days.setdefault(entry.period, []).extend(expand_days(entry.days_raw))
    return sum(len(set(ds)) for ds in per_period_days.values())


def count_periods_daywise(teacher: str, timetable: Dict[str, List[PeriodEntry]]) -> Dict[int, int]:
    per_day_periods: Dict[int, List[Tuple[int, int]]] = {d: [] for d in DAYS_ALL}
    for entry in timetable.get(teacher, []):
        for day in expand_days(entry.days_raw):
            per_day_periods[day].append((entry.period, day))
    return {d: len(set(per_day_periods[d])) for d in DAYS_ALL}


def generate_teacherwise(wb: Workbook, ctx: Context) -> Tuple[int, int]:
    """
    Reads CLASSWISE and writes TEACHERWISE in same workbook.
    Returns: (warnings_count, clashes_count)
    """
    if CLASSWISE_SHEET not in wb:
        raise FileNotFoundError("CLASSWISE sheet not found.")

    input_ws = wb[CLASSWISE_SHEET]
    teacher_names = load_teacher_names(wb)
    timetable: Dict[str, List[PeriodEntry]] = {}
    warnings = 0

    row = 2
    while True:
        class_cell = input_ws.cell(row, 1).value
        if not class_cell:
            break

        raw = str(class_cell)
        parts = [p.strip() for p in raw.split("@", 1)]
        klass = parts[0]
        alias = parts[1] if len(parts) == 2 else ""

        logging.info("Processing class %s", klass)

        days_check: Dict[int, int] = {}
        periods_assigned: Dict[str, int] = {}

        for col in PERIOD_COLUMNS:
            content = input_ws.cell(row, col).value
            if not content:
                warnings += 1
                logging.warning("Empty cell %s%d in CLASSWISE.", get_column_letter(col), row)
                continue

            for piece in str(content).split(ctx.separator):
                line = remove_comment(piece)
                if not line or line.startswith("##"):
                    if line.startswith("##"):
                        warnings += 1
                        logging.warning("Commented cell %s%d ignored.", get_column_letter(col), row)
                    continue

                m = RE_CLASSWISE_LINE.match(line.upper())
                if not m:
                    warnings += 1
                    logging.warning("Format issue in CLASSWISE %s%d: %s", get_column_letter(col), row, line)
                    continue

                subject, days, teacher = m.group("subject", "days", "teacher")
                subject = subject.strip()
                expanded = expand_days(days)
                for d in expanded:
                    days_check[d] = 1

                periods_assigned[subject] = periods_assigned.get(subject, 0) + count_days(days)
                key_klass = alias if alias else klass
                timetable.setdefault(teacher, []).append(PeriodEntry(period=col - 1, klass=key_klass, days_raw=days, subject=subject))

        # summary per subject into column J
        summary = [f"{subj}: {cnt}" for subj, cnt in sorted(periods_assigned.items())]
        summary.append(f"TOTAL: {sum(periods_assigned.values())}")
        input_ws.cell(row=row, column=10).value = ", ".join(summary)

        pending = sorted(set(DAYS_ALL) - set(days_check.keys()))
        if pending:
            warnings += 1
            logging.warning("Pending days %s in CLASSWISE %s%d.", pending, "row ", row)

        row += 1

    if not ctx.keep_timestamp:
        input_ws.cell(row=row, column=2).value = now_stamp()

    # Prepare TEACHERWISE
    if TEACHERWISE_SHEET in wb:
        out_ws = wb[TEACHERWISE_SHEET]
        clear_sheet(out_ws)
    else:
        out_ws = wb.create_sheet(title=TEACHERWISE_SHEET, index=1)

    header = ["Name", 1, 2, 3, 4, 5, 6, 7, 8, "Periods", "Daywise"]
    for i, h in enumerate(header, start=1):
        out_ws.cell(row=1, column=i).value = h

    # teacher ordering: those present in TEACHERS sheet first, then any extras
    timetable_teachers = set(timetable.keys())
    ordered = [t for t in teacher_names if t in timetable_teachers] + [t for t in sorted(timetable_teachers) if t not in teacher_names]

    row = 2
    totals: Dict[str, int] = {t: count_periods_for_teacher(t, timetable) for t in timetable_teachers}
    for t in ordered:
        display = f"{teacher_names[t]}, {t}" if (ctx.expand_fullnames and t in teacher_names) else t
        out_ws.cell(row=row, column=1).value = display

        # sort entries by days string to group visually
        entries = sorted(timetable.get(t, []), key=lambda e: e.days_raw)
        for e in entries:
            col = e.period + 1  # map back to sheet column
            cell = out_ws.cell(row=row, column=col)
            payload = f"{e.klass} ({e.days_raw}) {e.subject}"
            cell.value = (str(cell.value) + f"{ctx.separator}{payload}") if cell.value else payload

        out_ws.cell(row=row, column=10).value = totals.get(t, 0)
        out_ws.cell(row=row, column=11).value = str(count_periods_daywise(t, timetable))[1:-1]
        row += 1

    if not ctx.keep_timestamp:
        out_ws.cell(row=len(ordered) + 2, column=2).value = now_stamp()

    # clashes
    clashes = highlight_clashes(out_ws, ctx.separator)
    return warnings, clashes


def generate_classwise(input_wb: Workbook, outfile: Path, separator: str) -> int:
    """
    Creates/updates a workbook `outfile` with MASTER and per-class sheets,
    copying timetable from CLASSWISE in `input_wb`.
    """
    if CLASSWISE_SHEET not in input_wb:
        raise FileNotFoundError("CLASSWISE sheet not found in input workbook.")

    input_ws = input_wb[CLASSWISE_SHEET]

    try:
        output_wb = openpyxl.load_workbook(outfile)
    except Exception:
        output_wb = openpyxl.Workbook()

    # Ensure MASTER template
    if MASTER_SHEET not in output_wb:
        master = output_wb.create_sheet(MASTER_SHEET)
        master["A1"] = "GSSS AMARPURA (FAZILKA)"
        master["A4"] = "Mon"
        master["A5"] = "Tue"
        master["A6"] = "Wed"
        master["A7"] = "Thu"
        master["A8"] = "Fri"
        master["A9"] = "Sat"
        for c in PERIOD_COLUMNS:
            master.cell(3, c).value = c - 1
        format_master_ws(master)
    else:
        master = output_wb[MASTER_SHEET]

    # read incharges from TEACHERS
    details = load_teacher_details(input_wb)
    class_incharge: Dict[str, str] = {}
    # try to auto-detect column header names
    # expect 'INCHARGE' in headers (case-insensitive)
    for code, rowvals in details.items():
        klass = (rowvals.get("INCHARGE") or rowvals.get("CLASS") or "").strip() if isinstance(rowvals.get("INCHARGE") or rowvals.get("CLASS"), str) else rowvals.get("INCHARGE") or rowvals.get("CLASS")
        if klass:
            class_incharge[str(klass)] = code

    # (Re)create per-class sheets
    r = 2
    class_names: List[str] = []
    while True:
        klass = input_ws.cell(r, 1).value
        if not klass:
            break
        klass = str(klass)
        class_names.append(klass)
        if klass in output_wb:
            del output_wb[klass]
        ws = output_wb.copy_worksheet(master)
        ws.title = klass
        r += 1

    warnings = 0
    r = 2
    while True:
        klass = input_ws.cell(r, 1).value
        if not klass:
            break
        klass = str(klass)
        ws_out = output_wb[klass]
        ws_out.cell(2, 1).value = f"Class: {klass}"
        # incharge line
        ws_out.cell(2, 5).value = "Class In-charge: "
        code = class_incharge.get(klass)
        if code and code in details:
            gender = (details[code].get("GENDER") or "").lower()
            title = "Ms" if gender == "f" else "Mr"
            name = details[code].get("NAME") or code
            ws_out.cell(2, 5).value = f"Class In-charge: {title} {name}"
        else:
            ws_out.cell(2, 5).value += "_" * 25

        for col in PERIOD_COLUMNS:
            content = input_ws.cell(r, col).value
            if not content:
                warnings += 1
                logging.warning("Empty cell %s%d in CLASSWISE.", get_column_letter(col), r)
                continue

            for raw_line in str(content).split(separator):
                line = raw_line.strip()
                if not line or line.startswith("#"):
                    continue
                m = RE_CLASSWISE_LINE.match(line)
                if not m:
                    warnings += 1
                    logging.warning("Format issue in CLASSWISE %s%d: %s", get_column_letter(col), r, line)
                    continue
                subject, days, teacher = m.group("subject", "days", "teacher")
                for d in expand_days(days):
                    rr = d + 3
                    if ws_out.cell(rr, col).value is None:
                        ws_out.cell(rr, col).value = ""
                    ws_out.cell(rr, col).value += f"{subject.strip()} ({teacher})\n"

        r += 1

    # propagate timestamp from input CLASSWISE B2 to each class sheet at (10,2)
    timestamp = input_ws.cell(2, 2).value
    for ws in output_wb.worksheets:
        if ws.title[:1].isdigit():
            ws.cell(10, 2).value = timestamp

    output_wb.save(outfile)
    return warnings


# ----------------------------- Diff utilities --------------------------------

def _teachers_in_cell(ws: Worksheet, cell_name: str, separator: str) -> List[str]:
    cell_val = ws[cell_name].value or ""
    out: List[str] = []
    for raw in str(cell_val).split(separator):
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        m = RE_CLASSWISE_LINE.match(line)
        if not m:
            raise ValueError(f"Bad format in {cell_name!s}: {line!r}")
        out.append(m.group("teacher"))
    return out


def _affected_teachers(ws_a: Worksheet, ws_b: Worksheet, cell_name: str, separator: str) -> List[str]:
    a = _teachers_in_cell(ws_a, cell_name, separator)
    b = _teachers_in_cell(ws_b, cell_name, separator)
    return sorted(set(a + b))


def show_differences(base_xlsx: Path, current_xlsx: Path, separator: str) -> int:
    wb_a = openpyxl.load_workbook(base_xlsx)
    wb_b = openpyxl.load_workbook(current_xlsx)
    ws_a = wb_a[CLASSWISE_SHEET]
    ws_b = wb_b[CLASSWISE_SHEET]

    diffs: List[str] = []
    affected: List[str] = []

    r = 2
    while True:
        klass = ws_a.cell(r, 1).value
        if not klass:
            break
        for c in range(1, 10):
            name = f"{get_column_letter(c)}{r}"
            if (ws_a.cell(r, c).value or "") != (ws_b.cell(r, c).value or ""):
                diffs.append(name)
                try:
                    affected.extend(_affected_teachers(ws_a, ws_b, name, separator))
                except Exception as exc:
                    logging.debug("Skipping teacher parse in %s: %s", name, exc)
                ws_b.cell(r, c).fill = PatternFill(start_color="c3c3c3", end_color="c3c3c3", fill_type="solid")
        r += 1

    wb_b.save(current_xlsx)
    affected = sorted(set(affected))
    if diffs:
        logging.info("Differences at cells: %s", ", ".join(diffs))
        logging.info("Likely affected teachers: %s.", ", ".join(affected))
    else:
        logging.info("No differences found.")
    return len(diffs)


# --------------------------------- CLI ---------------------------------------

def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="twig.py",
        description="Generates TEACHERWISE (or CLASSWISE) timetables and compares timetables.",
    )
    p.add_argument("-f", "--fullname", action="store_true", help="replace short names with full names")
    p.add_argument("-k", "--keepstamp", action="store_true", help="keep timestamp intact")
    p.add_argument("-s", "--separator", default="\\n", help="line separator inside cells; default is \\n")
    p.add_argument("-V", "--verbosity", action="count", default=1, help="increase logging verbosity (-V, -VV)")
    p.add_argument("-v", "--version", action="store_true", help="display version information and exit")

    sub = p.add_subparsers(dest="command", required=True)

    tw = sub.add_parser("teacherwise", help="Generate TEACHERWISE timetable")
    tw.add_argument("infile", type=Path, help="Workbook containing CLASSWISE timetable (xlsx)")

    cw = sub.add_parser("classwise", help="Generate per-class sheets")
    cw.add_argument("infile", type=Path, help="Workbook containing CLASSWISE timetable (xlsx)")
    cw.add_argument("outfile", type=Path, help="Workbook to write classwise sheets (xlsx)")

    df = sub.add_parser("diff", help="Compare two CLASSWISE timetables")
    df.add_argument("base", type=Path, help="Base CLASSWISE timetable")
    df.add_argument("current", type=Path, help="Current CLASSWISE timetable to annotate")

    return p


def main(argv: Optional[List[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.version:
        print("twig.py: version 250905")
        return 0

    setup_logging(args.verbosity)

    separator = "\n" if args.separator == r"\n" else args.separator
    if separator != "\n":
        logging.info("Using separator '%s'", escape_visible(separator))

    ctx = Context(
        separator=separator,
        expand_fullnames=args.fullname,
        keep_timestamp=args.keepstamp,
    )

    start = time.time()

    try:
        if args.command in {"teacherwise", "classwise"}:
            logging.info("Reading workbook: %s", args.infile)
            wb = openpyxl.load_workbook(args.infile)

        if args.command == "teacherwise":
            warnings, clashes = generate_teacherwise(wb, ctx)
            wb.save(args.infile)
            print(f"Teacherwise timetable saved to '{args.infile}'.")
            print(f"Clashes: {clashes}")
            if warnings:
                print(f"Warnings: {warnings}")

        elif args.command == "classwise":
            warnings = generate_classwise(wb, args.outfile, separator=separator)
            print(f"Classwise timetables saved to '{args.outfile}'.")
            if warnings:
                print(f"Warnings: {warnings}")

        elif args.command == "diff":
            print(f"Comparing '{args.base}' with '{args.current}' ...")
            differences = show_differences(args.base, args.current, separator=separator)
            print(f"Found {differences} differences between {args.base} and {args.current}.")

    except FileNotFoundError as e:
        logging.error(str(e))
        return 2
    except KeyError as e:
        logging.error("Missing sheet or key: %s", e)
        return 2
    except Exception as e:
        logging.exception("Unexpected error: %s", e)
        return 1
    finally:
        elapsed = time.time() - start
        print(f"Finished processing in {elapsed:.3f} seconds.")
        print("Have a nice day!\n")

    return 0


if __name__ == "__main__":
    sys.exit(main())
