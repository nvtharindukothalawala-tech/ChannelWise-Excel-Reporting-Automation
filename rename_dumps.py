#!/usr/bin/env python3
"""
Rename GPOS downloads to the fixed names run_reports.py expects.

It opens each spreadsheet, works out whether it is a sales or a returns
export, reads the date column to see which period the data covers, and
renames the file accordingly.

Typical use — dumps land in Downloads, then:

    python rename_dumps.py --source "C:\\Users\\You\\Downloads"

Check first without touching anything:

    python rename_dumps.py --source "C:\\Users\\You\\Downloads" --dry-run

Periods are worked out relative to today:

    this month            -> sales_current_month.xlsx    / returns_current_month.xlsx
    last month            -> sales_previous_month.xlsx   / returns_previous_month.xlsx
    the month before that -> sales_two_months_ago.xlsx   / returns_two_months_ago.xlsx
    this month last year  -> sales_last_year.xlsx        / returns_last_year.xlsx

Anything else is left alone and reported.
"""

import argparse
import os
import re
import shutil
import sys
from collections import Counter
from datetime import date, datetime, timedelta

try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl is not installed.  Run:  python -m pip install openpyxl")

HERE = os.path.dirname(os.path.abspath(__file__))

# Columns that identify each export type
SALES_MARKERS = {"sold qty", "sales price", "invoice type", "sales amount"}
RETURN_MARKERS = {"return code", "return qty", "return date"}
STOCK_MARKERS = {"opening stock", "available qty", "rol variance", "stock take qty"}

# The two stock exports are identical in every field, so the folder they are
# dropped into is what identifies them.
STOCK_FOLDERS = {"stock_this_month": "stock_this_month.xlsx",
                 "stock_last_month": "stock_last_month.xlsx"}

# The MBR reports cover the month that has just finished, so their dumps are a
# month behind the daily ones. --mbr shifts the reference month back by one and
# writes mbr_ prefixed names, keeping the two sets apart.
SUMMARY_MARKERS = {"bid", "items", "tracking no", "sale amount"}
MBR_NAMES = {
    ("sales", "current_month"): "mbr_sales_month.xlsx",
    ("sales", "previous_month"): "mbr_sales_prev.xlsx",
    ("sales", "two_months_ago"): "mbr_sales_prev2.xlsx",
    ("sales", "last_year"): "mbr_sales_ly.xlsx",
    ("returns", "current_month"): "mbr_returns_month.xlsx",
    ("returns", "previous_month"): "mbr_returns_prev.xlsx",
    ("returns", "two_months_ago"): "mbr_returns_prev2.xlsx",
    ("returns", "last_year"): "mbr_returns_ly.xlsx",
    ("summary", "current_month"): "mbr_summary_month.xlsx",
    ("summary", "previous_month"): "mbr_summary_prev.xlsx",
}
MBR_STOCK_FOLDER = {"stock": "mbr_stock.xlsx"}

# Preferred date column, in order
SALES_DATE_COLS = ["date", "created on"]
RETURN_DATE_COLS = ["return date", "created on"]

EXCEL_EPOCH = date(1899, 12, 30)   # Excel's day 0, accounting for its 1900 leap-year quirk

DATE_PATTERNS = [
    "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d",
    "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y",
    "%m/%d/%Y %H:%M:%S", "%m/%d/%Y %H:%M", "%m/%d/%Y",
    "%Y/%m/%d", "%d-%m-%Y", "%d.%m.%Y",
]


# ----------------------------------------------------------------------
def to_date(value):
    """Best-effort conversion of a cell value to a date. None if not a date."""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        # Excel serial number
        if 20000 < float(value) < 80000:
            return EXCEL_EPOCH + timedelta(days=int(float(value)))
        return None
    text = str(value).strip()
    if not text:
        return None
    for pattern in DATE_PATTERNS:
        try:
            return datetime.strptime(text, pattern).date()
        except ValueError:
            continue
    # last resort: pull a yyyy-mm-dd or dd/mm/yyyy out of the string
    m = re.search(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})", text)
    if m:
        try:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    return None


def month_key(d):
    return (d.year, d.month)


def shift_month(ref, months_back):
    y, m = ref.year, ref.month - months_back
    while m <= 0:
        m += 12
        y -= 1
    return (y, m)


def period_name(ym, today):
    """Map a (year, month) to one of the four period labels, or None."""
    if ym == shift_month(today, 0):
        return "current_month"
    if ym == shift_month(today, 1):
        return "previous_month"
    if ym == shift_month(today, 2):
        return "two_months_ago"
    if ym == (today.year - 1, today.month):
        return "last_year"
    return None


# ----------------------------------------------------------------------
def inspect(path, sample_rows=None):
    """Return (kind, dominant year-month, rows read, date column used)."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = ws.iter_rows(values_only=True)
    try:
        header = [("" if h is None else str(h).strip()) for h in next(rows)]
    except StopIteration:
        wb.close()
        return None, None, 0, None, {}

    lower = [h.lower() for h in header]
    kind = None
    if RETURN_MARKERS & set(lower):
        kind = "returns"
    elif SALES_MARKERS & set(lower):
        kind = "sales"
    if kind == "sales" and SUMMARY_MARKERS <= set(lower):
        kind = "summary"          # the bill-wise summary export, not the detail one
    if kind is None:
        wb.close()
        return None, None, 0, None, {}

    wanted = RETURN_DATE_COLS if kind == "returns" else SALES_DATE_COLS
    idx = None
    used = None
    for name in wanted:
        if name in lower:
            idx = lower.index(name)
            used = header[idx]
            break
    if idx is None:
        wb.close()
        return kind, None, 0, None, {}

    months = Counter()
    seen = 0
    for row in rows:
        if idx >= len(row):
            continue
        d = to_date(row[idx])
        if d:
            months[month_key(d)] += 1
            seen += 1
        if sample_rows and seen >= sample_rows:
            break
    wb.close()

    if not months:
        return kind, None, 0, used, {}
    return kind, months.most_common(1)[0][0], seen, used, dict(months)


# ----------------------------------------------------------------------
def read_headers(path):
    """First row of the first sheet, as plain strings."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    try:
        row = next(ws.iter_rows(values_only=True))
    except StopIteration:
        row = ()
    wb.close()
    return [("" if h is None else str(h).strip()) for h in row]


def collect_stock(source, dest, dry_run, keep_source, folders=None):
    """Route the two stock exports by which subfolder they were dropped into."""
    done, problems = [], []
    for folder, target in (folders or STOCK_FOLDERS).items():
        path = os.path.join(source, folder)
        if not os.path.isdir(path):
            continue
        files = [f for f in sorted(os.listdir(path))
                 if f.lower().endswith((".xlsx", ".xls")) and not f.startswith("~$")]
        if not files:
            problems.append(f"{folder}\\ is empty")
            continue
        if len(files) > 1:
            problems.append(f"{folder}\\ holds {len(files)} files - leave exactly one")
            continue
        src = os.path.join(path, files[0])
        try:
            headers = read_headers(src)
        except Exception as exc:
            problems.append(f"{folder}\\{files[0]} could not be read: {exc}")
            continue
        if not (STOCK_MARKERS & {h.lower() for h in headers}):
            problems.append(f"{folder}\\{files[0]} does not look like a stock export")
            continue
        print(f"  {folder}\\{files[0]}")
        print(f"      stock   | identified by folder")
        print(f"      -> {target}")
        if not dry_run:
            out = os.path.join(dest, target)
            (shutil.copy2 if keep_source else shutil.move)(src, out)
        done.append(target)
    return done, problems


def main():
    ap = argparse.ArgumentParser(
        description="Rename GPOS sales/returns exports to the names run_reports.py expects.")
    ap.add_argument("--source", default=None,
                    help="folder holding the freshly downloaded files (default: the dumps folder)")
    ap.add_argument("--dest", default=None,
                    help="where the renamed files go (default: ./dumps)")
    ap.add_argument("--dry-run", action="store_true", help="show what would happen, change nothing")
    ap.add_argument("--keep-source", action="store_true", help="copy instead of move")
    ap.add_argument("--today", default=None, help="override today's date, as YYYY-MM-DD")
    ap.add_argument("--mbr", action="store_true",
                    help="name the MBR dump set instead: periods shift back one month, "
                         "because MBR covers the month that has just finished")
    ap.add_argument("--explain", action="store_true",
                    help="show the full month-by-month breakdown of every file")
    args = ap.parse_args()

    dest = os.path.abspath(args.dest or os.path.join(HERE, "dumps"))
    source = os.path.abspath(args.source or dest)
    today = datetime.strptime(args.today, "%Y-%m-%d").date() if args.today else date.today()
    # The reports always cover YESTERDAY, so the period names follow yesterday's
    # month, not today's. They only differ on the 1st - when today is 1 October
    # the run is about 30 September, so September is the current month, August
    # the previous one and July the one before that.
    #
    # MBR is different again: it covers the month that has just finished, which
    # is the previous calendar month whether it is run on the 1st or the 5th.
    if args.mbr:
        y, m = shift_month(today, 1)
        today = date(y, m, 1)
        print("\n--mbr: treating %s as the reporting month" % today.strftime("%B %Y"))
    else:
        anchor = today - timedelta(days=1)
        if anchor.month != today.month:
            print("\nrun is about %s, so %s is the current month"
                  % (anchor.strftime("%d %B %Y"), anchor.strftime("%B %Y")))
        today = anchor

    if not os.path.isdir(source):
        sys.exit(f"Source folder not found: {source}")
    os.makedirs(dest, exist_ok=True)

    files = [os.path.join(source, f) for f in sorted(os.listdir(source))
             if f.lower().endswith((".xlsx", ".xls")) and not f.startswith("~$")]
    if not files and not any(os.path.isdir(os.path.join(source, d)) for d in STOCK_FOLDERS):
        sys.exit(f"No spreadsheets found in {source}")

    print(f"\nSource : {source}")
    print(f"Target : {dest}")
    print(f"Today  : {today}   (current month = {today:%B %Y})")
    print("=" * 74)

    planned = {}
    skipped = []
    folders = MBR_STOCK_FOLDER if args.mbr else STOCK_FOLDERS
    stock_done, stock_problems = collect_stock(source, dest, args.dry_run,
                                               args.keep_source, folders)
    skipped.extend((p, "stock folder") for p in stock_problems)

    for path in files:
        name = os.path.basename(path)
        try:
            kind, ym, seen, datecol, spread = inspect(path)
        except Exception as exc:
            skipped.append((name, f"could not read: {exc}"))
            continue

        if kind is None:
            skipped.append((name, "not a sales or returns export"))
            continue
        if ym is None:
            skipped.append((name, "no readable dates found"))
            continue

        period = period_name(ym, today)
        if period is None:
            extra = ""
            if spread and len(spread) > 1:
                extra = " (file spans " + ", ".join(
                    f"{y}-{m:02d}" for (y, m) in sorted(spread)) + ")"
            skipped.append((name,
                f"dominant month is {ym[0]}-{ym[1]:02d}, not a period we track{extra}"))
            continue

        if args.mbr:
            target = MBR_NAMES.get((kind, period))
            if not target:
                skipped.append((name, f"{kind} for {period} is not part of the MBR set"))
                continue
        else:
            if kind == "summary":
                skipped.append((name, "summary exports are only used by the MBR reports "
                                      "- run with --mbr"))
                continue
            target = f"{kind}_{period}.xlsx"
        print(f"  {name}")
        print(f"      {kind:<7} | {ym[0]}-{ym[1]:02d} | {seen} dated rows | column {datecol!r}")
        if args.explain and spread:
            parts = ", ".join(f"{y}-{m:02d}: {n}" for (y, m), n in sorted(spread.items()))
            print(f"      months  | {parts}")
        print(f"      -> {target}")

        if target in planned:
            skipped.append((name, f"clashes with {os.path.basename(planned[target])}, both map to {target}"))
            print(f"      CLASH — already claimed by {os.path.basename(planned[target])}")
            continue
        planned[target] = path

    if skipped:
        print("\nSkipped:")
        for name, why in skipped:
            print(f"  {name}\n      {why}")

    if not planned and not stock_done:
        print("\nNothing to rename.\n")
        return

    print("\n" + "=" * 74)
    if args.dry_run:
        print(f"Dry run — {len(planned)} file(s) would be renamed. Nothing changed.\n")
        return

    moved = 0
    for target, src in sorted(planned.items()):
        out = os.path.join(dest, target)
        try:
            if args.keep_source:
                shutil.copy2(src, out)
            else:
                shutil.move(src, out)
            moved += 1
        except Exception as exc:
            print(f"  FAILED {target}: {exc}")

    verb = "copied" if args.keep_source else "moved"
    print(f"{moved} file(s) {verb} into {dest}\n")

    expected = {f"{k}_{p}.xlsx" for k in ("sales", "returns")
                for p in ("current_month", "previous_month", "last_year")}
    missing = sorted(e for e in expected if not os.path.exists(os.path.join(dest, e)))
    if missing:
        print("Still missing (needed by all 5 reports):")
        for m in missing:
            print(f"  {m}")
        print()
    else:
        print("All six core dumps are present. Safety Sure also needs the two_months_ago pair.\n")


if __name__ == "__main__":
    main()
