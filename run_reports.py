#!/usr/bin/env python3
"""
Channel Report Automation
=========================
Fills the 5 brand channel reports from the daily POS dumps.

Usage:
    python run_reports.py            # normal daily run
    python run_reports.py --audit    # verify layout, change nothing
    python run_reports.py --dry-run  # do everything except save
    python run_reports.py --brand Giggles      # one brand only

The report workbooks are never restructured. Only the raw paste zone is
replaced, and the helper formulas to its left are re-stretched to match
the new row count. Daily Summary / Daily Target / Stp are untouched.
"""

import argparse, csv, json, logging, os, re, shutil, sys, time
from datetime import date, datetime, timedelta
from pathlib import Path
from hashlib import sha256

try:
    import cfpatch
    import chartpatch
    import external_link_patch
except ImportError:
    sys.exit("Required preservation module missing: cfpatch.py, chartpatch.py, or external_link_patch.py")

try:
    import openpyxl
    from openpyxl.formula.translate import Translator
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.formula import ArrayFormula
except ImportError:
    sys.exit("openpyxl is not installed.  Run:  pip install openpyxl")

HERE = os.path.dirname(os.path.abspath(__file__))
log = logging.getLogger("reports")
MATCH_FIELDS = {}          # set from config in main(); used by include_keywords


# ----------------------------------------------------------------------
# setup
# ----------------------------------------------------------------------
def load_config(path):
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


def setup_logging(log_dir):
    os.makedirs(log_dir, exist_ok=True)
    logfile = os.path.join(log_dir, f"run_{datetime.now():%Y-%m-%d}.log")
    fmt = logging.Formatter("%(asctime)s  %(levelname)-7s %(message)s", "%Y-%m-%d %H:%M:%S")
    log.setLevel(logging.INFO)
    log.handlers.clear()
    for h in (logging.FileHandler(logfile, encoding="utf-8"), logging.StreamHandler(sys.stdout)):
        h.setFormatter(fmt)
        log.addHandler(h)
    return logfile


# ----------------------------------------------------------------------
# dump reading
# ----------------------------------------------------------------------
def read_dump(path):
    """Return (headers, rows). Row 1 is the header row in every POS export."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)
    try:
        headers = [("" if h is None else str(h).strip()) for h in next(rows_iter)]
    except StopIteration:
        wb.close()
        return [], []
    rows = [list(r) for r in rows_iter if any(v is not None and str(v).strip() for v in r)]
    wb.close()
    return headers, rows


def required_dumps(targets):
    """Just the dump keys the selected reports actually read."""
    need = set()
    for bc in targets.values():
        for key, spec in bc["sheets"].items():
            need.update(spec.get("stack") or [key])
    return need


def load_all_dumps(cfg, dumps_dir, only=None):
    """Load the dumps that exist. Missing ones are reported, not fatal."""
    dumps = {}
    for key, fname in cfg["dump_files"].items():
        if only is not None and key not in only:
            continue
        path = os.path.join(dumps_dir, fname)
        if not os.path.exists(path):
            log.warning("dump missing, will skip sheets that need it: %s", fname)
            continue
        headers, rows = read_dump(path)
        # Drop accidental export-only VAT fields by header NAME, not position.
        if dump_kind(key) == "sales":
            excluded = {"vat amount", "total (incl. vat)"}
            indexes = [i for i, h in enumerate(headers)
                       if str(h).strip().casefold() not in excluded]
            ignored = [h for i, h in enumerate(headers) if i not in indexes]
            if ignored:
                log.info("%s: ignored accidental sales fields: %s", fname,
                         ", ".join(ignored))
            headers = [headers[i] for i in indexes]
            rows = [[row[i] if i < len(row) else None for i in indexes]
                    for row in rows]
            if len(headers) != 26:
                raise ValueError(
                    f"{fname}: expected 26 standard sales columns after filtering; "
                    f"found {len(headers)}. Check dump headers before writing reports."
                )
        dumps[key] = {"headers": headers, "rows": rows, "file": fname}
        log.info("loaded %-24s %6d rows, %2d cols", fname, len(rows), len(headers))
    return dumps


EXCEL_EPOCH = date(1899, 12, 30)
DATE_PATTERNS = ["%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d",
                 "%d/%m/%Y %H:%M:%S", "%d/%m/%Y", "%m/%d/%Y %H:%M:%S", "%m/%d/%Y"]
DATE_COLS = {"sales": ["date", "created on"], "returns": ["return date", "created on"],
             "stock": ["created date"], "summary": ["date"]}


def dump_kind(key):
    k = key[4:] if key.startswith("mbr_") else key
    if k.startswith("returns"):
        return "returns"
    if k.startswith("stock"):
        return "stock"
    if k.startswith("summary"):
        return "summary"
    return "sales"


def to_date(value):
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        if 20000 < float(value) < 80000:
            return EXCEL_EPOCH + timedelta(days=int(float(value)))
        return None
    text = str(value).strip()
    if re.fullmatch(r"\d{5}(?:\.\d+)?", text):
        serial = float(text)
        if 20000 < serial < 80000:
            return EXCEL_EPOCH + timedelta(days=int(serial))
    for pat in DATE_PATTERNS:
        try:
            return datetime.strptime(text, pat).date()
        except ValueError:
            continue
    m = re.search(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})", text)
    if m:
        try:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    return None


def date_column(headers, kind):
    lower = [h.lower() for h in headers]
    for name in DATE_COLS[kind]:
        if name in lower:
            return lower.index(name)
    return None


def reference_day(dumps):
    """Latest day-of-month present in the current-month sales dump."""
    dump = dumps.get("sales_current_month")
    if not dump:
        return None
    col = date_column(dump["headers"], "sales")
    if col is None:
        return None
    best = 0
    for row in dump["rows"]:
        d = to_date(row[col]) if col < len(row) else None
        if d and d.day > best:
            best = d.day
    return best or None


def keyword_filter(rows, headers, kind, keywords, exclude=None):
    """Keep rows whose match field contains one of `keywords`.

    Used where a report wants a narrower slice than the shared brand bucket -
    e.g. the Be Desi revenue report wants only Be Desi locations, not the
    Lux Desi / Wedding Studio / Raja & Rani sub-brands that the Be Desi
    CHANNEL report includes.
    """
    field = MATCH_FIELDS.get(kind)
    if not field or field not in headers:
        log.error("cannot apply include_keywords: %r not in dump headers", field)
        return []
    col = headers.index(field)
    words = [w.lower() for w in keywords]
    bad = [w.lower() for w in (exclude or [])]
    out = []
    for row in rows:
        v = row[col] if col < len(row) else None
        if v is None:
            continue
        text = str(v).lower()
        if not any(w in text for w in words):
            continue
        if bad and any(w in text for w in bad):
            continue
        out.append(row)
    return out


def truncate(rows, headers, kind, day):
    """Keep only rows dated on or before `day` of their month."""
    col = date_column(headers, kind)
    if col is None or not day:
        return rows
    out = []
    for row in rows:
        d = to_date(row[col]) if col < len(row) else None
        if d is None or d.day <= day:
            out.append(row)
    return out


# ----------------------------------------------------------------------
# brand matching
# ----------------------------------------------------------------------
def build_matcher(cfg):
    kw = {b: [k.lower() for k in words] for b, words in cfg["brand_keywords"].items()}

    def match(value):
        """Return (brand, None) or (None, reason)."""
        if value is None:
            return None, "empty"
        text = str(value).strip().lower()
        if not text:
            return None, "empty"
        hits = [b for b, words in kw.items() if any(w in text for w in words)]
        if len(hits) == 1:
            return hits[0], None
        if not hits:
            return None, "no rule matched"
        return None, "matched several brands: " + ", ".join(sorted(hits))

    return match


def select_by_keywords(dump, kind, keywords, exclude, field_for):
    """Pick rows straight from a dump by matching its location/warehouse field.

    Used where a report needs a narrower set than the shared brand buckets --
    e.g. a BU revenue report that wants Be Desi but not Lux Desi or Raja & Rani.
    """
    field = field_for[kind]
    headers = dump["headers"]
    if field not in headers:
        return []
    col = headers.index(field)
    inc = [k.lower() for k in keywords]
    exc = [k.lower() for k in (exclude or [])]
    out = []
    for row in dump["rows"]:
        v = row[col] if col < len(row) else None
        if v is None:
            continue
        text = str(v).strip().lower()
        if not any(k in text for k in inc):
            continue
        if exc and any(k in text for k in exc):
            continue
        out.append(row)
    return out


def split_by_brand(dumps, cfg, match):
    """Return {dump_key: {brand: [rows]}} plus a list of unmatched rows."""
    field_for = cfg["match_fields"]
    result, unmatched = {}, []

    for key, dump in dumps.items():
        kind = dump_kind(key)
        field = field_for[kind]
        headers = dump["headers"]
        if field not in headers:
            log.error("%s: column %r not found — headers are: %s",
                      dump["file"], field, ", ".join(headers))
            result[key] = {}
            continue

        col = headers.index(field)
        buckets, seen = {}, {}
        buckets["__ALL__"] = list(dump["rows"])
        for row in dump["rows"]:
            value = row[col] if col < len(row) else None
            brand, reason = match(value)
            if brand is None:
                key_v = str(value)
                seen[key_v] = seen.get(key_v, 0) + 1
                unmatched.append({"dump": dump["file"], "field": field,
                                  "value": key_v, "reason": reason})
                continue
            buckets.setdefault(brand, []).append(row)
        result[key] = buckets

        for value, count in sorted(seen.items(), key=lambda x: -x[1]):
            log.warning("%s: %d row(s) unclassified — %s = %r", dump["file"], count, field, value)

    return result, unmatched



# ----------------------------------------------------------------------
# Fashion / Non-Fashion tracker extensions (exact matching, month upsert)
# ----------------------------------------------------------------------
def normalized_location(value):
    """Normalize harmless spacing/case only; NEVER fuzzy-match businesses."""
    return " ".join(str(value or "").split()).casefold()


def select_by_exact_locations(dump, kind, locations, match_fields):
    """Select only explicitly approved Location / Warehouse labels."""
    name = match_fields[kind]
    headers = dump["headers"]
    if name not in headers:
        raise ValueError(f"Missing classification field {name!r} in {dump['file']}")
    allowed = {normalized_location(x) for x in locations}
    if len(allowed) != len(locations):
        raise ValueError("Duplicate approved locations in config")
    pos = headers.index(name)
    return [r for r in dump["rows"]
            if normalized_location(r[pos] if pos < len(r) else None) in allowed]


def real_excel_datetime(value):
    """Convert POS date strings/numeric serials to real Excel date-time values."""
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    if isinstance(value, (int, float)) and 20000 < float(value) < 80000:
        return datetime(1899, 12, 30) + timedelta(days=float(value))
    if re.fullmatch(r"\d{5}(?:\.\d+)?", str(value).strip()):
        serial = float(value)
        if 20000 < serial < 80000:
            return datetime(1899, 12, 30) + timedelta(days=serial)
    for pat in DATE_PATTERNS:
        try:
            return datetime.strptime(str(value).strip(), pat)
        except ValueError:
            pass
    raise ValueError(f"Cannot convert source date-time {value!r}")


def map_source_columns(headers, rows, columns):
    """Match by HEADER NAME, not export column number; None means blank."""
    hmap = {}
    for i, h in enumerate(headers):
        key = str(h).strip().casefold()
        if key in hmap:
            raise ValueError(f"Duplicate source column: {h!r}")
        hmap[key] = i
    missing = [h for h in columns if h is not None and h.strip().casefold() not in hmap]
    if missing:
        raise ValueError(f"Missing required dump headers: {missing}")
    positions = [None if h is None else hmap[h.strip().casefold()] for h in columns]
    mapped = []
    for r in rows:
        new = []
        for name, p in zip(columns, positions):
            val = None if p is None else (r[p] if p < len(r) else None)
            if name in ('Date', 'Return Date', 'Created On') and val is not None:
                val = real_excel_datetime(val)
            new.append(val)
        mapped.append(new)
    return mapped


def dump_month(dump, kind):
    """Only accept complete, single-month dumps for controlled replacement."""
    pos = date_column(dump["headers"], kind)
    if pos is None:
        raise ValueError(f"Date column missing in {dump['file']}")
    months = set()
    for r in dump["rows"]:
        d = to_date(r[pos] if pos < len(r) else None)
        if d is None:
            raise ValueError(f"Invalid source date in {dump['file']}: {r[pos]!r}")
        months.add((d.year, d.month))
    if len(months) != 1:
        raise ValueError(f"Expected exactly one month in {dump['file']}; found {months}")
    return months.pop()


def keep_other_months(ws, spec, replace_month):
    """Preserve historical raw values; replace ONLY the specified month."""
    start = spec['data_start_row']
    width = len(spec['source_columns'])
    fields = spec['source_columns']
    datepos = fields.index(spec['date_target_header'])
    kept = []
    bad_dates = 0
    removed = 0
    for r in range(start, ws.max_row + 1):
        values = [ws.cell(r, spec['paste_start_col'] + col).value
                  for col in range(width)]
        if all(v is None or str(v).strip() == '' for v in values):
            continue
        d = to_date(values[datepos])
        if d is None:
            # Preserve anomaly; don't silently destroy a historical record.
            bad_dates += 1
            kept.append(values)
        elif (d.year, d.month) == replace_month:
            removed += 1
        else:
            kept.append(values)
    if bad_dates:
        log.warning("%s: preserved %d historical rows with invalid dates; REVIEW",
                    ws.title, bad_dates)
    log.info("%s: retained %d history rows; replacing %d rows in %04d-%02d",
             ws.title, len(kept), removed, *replace_month)
    return kept


def check_cluster_coverage(dumps, reports, match_fields):
    """Fail closed: every current record must belong to exactly one cluster."""
    cluster_reports = {n:v for n,v in reports.items() if v.get('group') == 'cluster'}
    if not cluster_reports:
        return
    for dump_key in ('sales_current_month', 'returns_current_month'):
        if dump_key not in dumps:
            continue
        kind = dump_kind(dump_key)
        name = match_fields[kind]
        dump = dumps[dump_key]
        if name not in dump['headers']:
            raise ValueError(f"{name!r} absent from {dump['file']}")
        loc_col = dump['headers'].index(name)
        buckets = {}
        for cluster, bc in cluster_reports.items():
            spec = bc['sheets'].get(dump_key)
            if not spec:
                continue
            for loc in spec.get('exact_locations', []):
                key = normalized_location(loc)
                if key in buckets:
                    raise ValueError(f"Location {loc!r} assigned to {buckets[key]} AND {cluster}")
                buckets[key] = cluster
        unknown = {}
        count = {}
        amount = {}
        amount_name = 'Sales Amount' if kind == 'sales' else 'Sale Amount'
        money_col = (dump['headers'].index(amount_name)
                     if amount_name in dump['headers'] else None)
        for row in dump['rows']:
            loc = row[loc_col]
            cluster = buckets.get(normalized_location(loc))
            if cluster is None:
                key = str(loc)
                unknown[key] = unknown.get(key, 0) + 1
                continue
            count[cluster] = count.get(cluster, 0) + 1
            if money_col is not None:
                try:
                    amount[cluster] = amount.get(cluster, 0.0) + float(row[money_col] or 0)
                except (ValueError, TypeError):
                    raise ValueError(f"Non-numeric {amount_name} for {loc}")
        for cluster in cluster_reports:
            log.info("cluster validation %-12s %-22s rows=%d amount=%.2f",
                     cluster, dump_key, count.get(cluster, 0), amount.get(cluster, 0.0))
        if unknown:
            raise ValueError(f"Unapproved {name} values in {dump['file']}: {unknown}")
        if sum(count.values()) != len(dump['rows']):
            raise ValueError(f"Reconciliation failed for {dump['file']}")


# ----------------------------------------------------------------------
# worksheet writing
# ----------------------------------------------------------------------
CELL_REF = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)")


def shift_ref(ref, delta):
    """Move the row part of every relative reference in `ref` by `delta`."""
    def bump(m):
        if m.group(3) == "$":          # absolute row, leave alone
            return m.group(0)
        return f"{m.group(1)}{m.group(2)}{m.group(3)}{int(m.group(4)) + delta}"
    return CELL_REF.sub(bump, ref)


def capture_template(ws, spec):
    """Grab the helper formulas/styles from the first data row."""
    row = spec["data_start_row"]
    tmpl = []
    for col in range(1, spec["paste_start_col"]):
        cell = ws.cell(row=row, column=col)
        val = cell.value
        if isinstance(val, ArrayFormula):
            kind = "array"
        elif isinstance(val, str) and val.startswith("="):
            kind = "formula"
        else:
            kind = "plain"
        tmpl.append({"col": col, "value": val, "style": cell._style, "kind": kind})
    return tmpl


def clear_from(ws, first_row):
    """Delete every cell at or below first_row, and shrink the used range."""
    dead = [k for k in ws._cells if k[0] >= first_row]
    for k in dead:
        del ws._cells[k]
    for r in [r for r in ws.row_dimensions if r >= first_row]:
        del ws.row_dimensions[r]
    ws._current_row = min(ws._current_row, first_row - 1)


def fill_sheet(ws, spec, headers, rows, template, max_cols=None, skip_cols=None):
    """Write rows into the paste zone and re-stretch the helper formulas."""
    start_row = spec["data_start_row"]
    start_col = spec["paste_start_col"]

    raw_styles = ([ws.cell(start_row, start_col + j)._style
                   for j in range(len(spec.get('source_columns') or []))]
                  if spec.get('source_columns') else [])
    clear_from(ws, start_row)

    drop = {int(c) - 1 for c in (skip_cols or [])}      # 1-based in config
    width = min(len(headers), max_cols) if max_cols else len(headers)

    for i, row in enumerate(rows):
        r = start_row + i
        j = 0
        for src, value in enumerate(row[:width]):
            if src in drop:
                continue
            cell = ws.cell(row=r, column=start_col + j, value=value)
            if j < len(raw_styles) and raw_styles[j] is not None:
                cell._style = raw_styles[j]
            j += 1

        for t in template:
            cell = ws.cell(row=r, column=t["col"])
            if t["value"] is None:
                continue
            letter = get_column_letter(t["col"])
            origin = f"{letter}{start_row}"
            dest = f"{letter}{r}"
            if t["kind"] == "formula":
                cell.value = Translator(t["value"], origin=origin).translate_formula(dest)
            elif t["kind"] == "array":
                # array formulas carry their own anchor range, which has to move too
                af = t["value"]
                cell.value = ArrayFormula(
                    ref=shift_ref(af.ref, r - start_row),
                    text=Translator(af.text, origin=origin).translate_formula(dest),
                )
            else:
                cell.value = t["value"]
            cell._style = t["style"]

    return len(rows)


# ----------------------------------------------------------------------
# fail-closed validation gate for cluster reports
# ----------------------------------------------------------------------
def validate_staged_cluster_report(brand, staged_path, original_path, config_path, month):
    """Return True only if the staged workbook passes the read-only checks."""
    from validate_report_integrity import validate

    try:
        result = validate(
            report_path=Path(staged_path),
            backup_path=Path(original_path),
            config_path=Path(config_path),
            month=month,
            report_name=brand,
        )
        if result != 0:
            log.error("%s: staged integrity validation FAILED", brand)
            return False
        log.info("%s: staged integrity validation PASSED", brand)
        return True
    except Exception:
        log.exception("%s: integrity validation raised an exception", brand)
        return False


# ----------------------------------------------------------------------
# per-brand run
# ----------------------------------------------------------------------
def process_brand(brand, brand_cfg, split, dumps, reports_dir, dry_run, backup_dir=None,
                  ref_day=None, match_fields=None, config_path=None):
    path = os.path.join(reports_dir, brand_cfg["file"])
    if not os.path.exists(path):
        log.error("%-17s report file not found: %s", brand, path)
        return False, {}

    try:
        probe = open(path, "r+b")
        probe.close()
    except (PermissionError, OSError):
        log.error("%-17s file is locked — close it in Excel and re-run", brand)
        return False, {}

    t0 = time.time()
    # snapshot the conditional formatting openpyxl cannot keep
    try:
        cf_snapshot = cfpatch.snapshot(path)
    except Exception as exc:
        if brand_cfg.get("group") == "cluster":
            log.error("%-17s conditional formatting snapshot FAILED: %s", brand, exc)
            return False, {}
        log.warning("%-17s could not snapshot conditional formatting: %s", brand, exc)
        cf_snapshot = {}
    try:
        chart_snapshot = chartpatch.snapshot(path)
    except Exception as exc:
        if brand_cfg.get("group") == "cluster":
            log.error("%-17s chart snapshot FAILED: %s", brand, exc)
            return False, {}
        log.warning("%-17s could not snapshot charts: %s", brand, exc)
        chart_snapshot = {"parts": {}, "overrides": [], "sheets": {}}

    # Preserve external links before openpyxl serializes the workbook.
    # For cluster reports, any snapshot error blocks the update.
    try:
        external_link_snapshot = external_link_patch.snapshot(path)
        log.info("%-17s snapshotted %d external-link component(s)",
                 brand, len(external_link_snapshot))
    except Exception as exc:
        log.error("%-17s external-link snapshot FAILED: %s", brand, exc)
        return False, {}

    wb = openpyxl.load_workbook(path)
    counts = {}

    for dump_key, spec in brand_cfg["sheets"].items():
        if dump_key not in dumps:
            log.warning("%-17s %-22s skipped, dump not available", brand, spec["sheet"])
            continue
        if spec["sheet"] not in wb.sheetnames:
            log.error("%-17s sheet %r not in workbook", brand, spec["sheet"])
            continue

        ws = wb[spec["sheet"]]
        sources = spec.get("stack") or [dump_key]
        missing = [k for k in sources if k not in dumps]
        if missing:
            log.error("%-17s %-22s SKIPPED - missing dump(s): %s",
                      brand, spec["sheet"], ", ".join(missing))
            continue

        kind = dump_kind(sources[0])
        rows = []
        for key in sources:
            pool = split.get(key, {})
            # a sheet may narrow its report's filter - e.g. stock that excludes
            # sub-brands the sales sheets still include
            keywords = (spec.get("keywords") or spec.get("include_keywords")
                        or brand_cfg.get("keywords") or brand_cfg.get("include_keywords"))
            exclude = spec.get("exclude_keywords") or brand_cfg.get("exclude_keywords")
            if "exact_locations" in spec:
                part = select_by_exact_locations(dumps[key], kind,
                                                 spec["exact_locations"], match_fields)
            elif keywords:
                part = keyword_filter(pool.get("__ALL__", []), dumps[key]["headers"],
                                      kind, keywords, exclude)
            elif brand_cfg.get("all_rows"):
                part = pool.get("__ALL__", [])
            elif brand_cfg.get("brands"):
                part = [r for b in brand_cfg["brands"] for r in pool.get(b, [])]
            else:
                part = pool.get(brand, [])
            if spec.get("truncate_to_day") and kind != "stock" \
                    and key not in ("sales_current_month", "returns_current_month"):
                part = truncate(part, dumps[key]["headers"], kind, ref_day)
            rows.extend(part)
        if spec.get('source_columns'):
            rows = map_source_columns(dumps[sources[0]]["headers"],
                                      rows, spec['source_columns'])
            if spec.get('replace_month'):
                month = dump_month(dumps[sources[0]], kind)
                old = keep_other_months(ws, spec, month)
                rows = old + rows
            template = capture_template(ws, spec)
            written = fill_sheet(ws, spec, spec['source_columns'], rows, template)
        else:
            template = capture_template(ws, spec)
            written = fill_sheet(ws, spec, dumps[sources[0]]["headers"], rows, template,
                                 spec.get("max_cols"), spec.get("skip_cols"))
        counts[spec["sheet"]] = written
        if brand_cfg.get("include_keywords"):
            tag = " (" + "+".join(brand_cfg["include_keywords"]) + ")"
        elif brand_cfg.get("all_rows"):
            tag = " (all brands)"
        elif brand_cfg.get("brands"):
            tag = " (" + "+".join(brand_cfg["brands"]) + ")"
        else:
            tag = ""
        log.info("%-17s %-22s %6d rows%s", brand, spec["sheet"], written, tag)

    # The tracker templates contain PivotTables, which must refresh in Excel.
    if brand_cfg.get('group') == 'cluster':
        for sheet in wb:
            for pivot in getattr(sheet, '_pivots', []):
                if pivot.cache is not None:
                    pivot.cache.refreshOnLoad = True
                    pivot.cache.enableRefresh = True

    if dry_run:
        log.info("%-17s dry run — not saved", brand)
        wb.close()
        return True, counts

    is_cluster = brand_cfg.get("group") == "cluster"
    if is_cluster and (not backup_dir or not config_path):
        log.error("%s: cluster update requires backup and configuration path", brand)
        wb.close()
        return False, counts

    # Never allow a stale .tmp workbook to be silently overwritten.
    tmp = path + ".tmp"
    if os.path.exists(tmp):
        log.error("%s: stale staging file exists: %s; inspect it first", brand, tmp)
        wb.close()
        return False, counts

    backup_path = None
    if backup_dir:
        backup_path = os.path.join(backup_dir, brand_cfg["file"])
        os.makedirs(os.path.dirname(backup_path), exist_ok=True)
        if os.path.exists(backup_path):
            log.error("%s: refusing to overwrite existing backup: %s", brand, backup_path)
            wb.close()
            return False, counts
        # Preserve an exact byte-for-byte backup before staging or publishing.
        with open(path, "rb") as original, open(backup_path, "xb") as backup:
            shutil.copyfileobj(original, backup)
        def _digest(filename):
            h = sha256()
            with open(filename, "rb") as stream:
                for block in iter(lambda: stream.read(1024 * 1024), b""):
                    h.update(block)
            return h.hexdigest()
        if _digest(path) != _digest(backup_path):
            wb.close()
            raise RuntimeError("Backup hash mismatch; refusing to publish")
        log.info("%-17s verified backup created: %s", brand, backup_path)

    try:
        # Save staging file, not the original report.
        wb.save(tmp)
    finally:
        wb.close()

    try:
        if cf_snapshot:
            try:
                n = cfpatch.restore(tmp, cf_snapshot)
                log.info("%-17s conditional formatting restored on %d sheet(s)", brand, n)
            except Exception:
                if is_cluster:
                    raise
                log.exception("%-17s conditional formatting restoration FAILED", brand)

        if chart_snapshot["parts"]:
            try:
                n = chartpatch.restore(tmp, chart_snapshot)
                log.info("%-17s charts restored (%d parts)", brand, n)
            except Exception:
                if is_cluster:
                    raise
                log.exception("%-17s chart restoration FAILED", brand)

        # Restore complete original XML and relationship parts before integrity checks.
        # Never publish a report if external-link preservation fails.
        restored_links = external_link_patch.restore(tmp, external_link_snapshot)
        log.info("%-17s external-link components restored: %d", brand, restored_links)

        if is_cluster:
            sales_dump = dumps.get("sales_current_month")
            if not sales_dump:
                raise ValueError("Current sales dump missing; cannot validate cluster report")
            year, month_number = dump_month(sales_dump, "sales")
            month = f"{year:04d}-{month_number:02d}"
            if not validate_staged_cluster_report(
                brand, tmp, path, config_path, month
            ):
                raise RuntimeError("Staged report integrity checks failed; original preserved")

        # Final publication happens ONLY after restoration and validation.
        os.replace(tmp, path)
        log.info("%-17s saved in %.1fs", brand, time.time() - t0)
        return True, counts
    except Exception as exc:
        log.exception("%-17s publication blocked: %s", brand, exc)
        # Only remove OUR stage file. Never remove or replace the original.
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except OSError:
            log.exception("%s: could not remove staged file %s", brand, tmp)
        return False, counts



# ----------------------------------------------------------------------
# audit
# ----------------------------------------------------------------------
def audit(cfg, reports_dir, only=None):
    print("\nLayout audit — compares config.json against the actual workbooks\n")
    problems = 0
    for brand, bc in (only or cfg["brands"]).items():
        path = os.path.join(reports_dir, bc["file"])
        print(f"  {brand}")
        if not os.path.exists(path):
            print(f"      MISSING FILE: {path}")
            problems += 1
            continue
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        for dump_key, spec in bc["sheets"].items():
            name = spec["sheet"]
            if name not in wb.sheetnames:
                print(f"      {name:<22} SHEET NOT FOUND")
                problems += 1
                continue
            ws = wb[name]
            hr = spec["header_row"]
            row = next((r for i, r in enumerate(
                ws.iter_rows(min_row=hr, max_row=hr, max_col=60, values_only=True), 1)), ())
            col = spec["paste_start_col"]
            got = row[col - 1] if col - 1 < len(row) else None
            default = {"returns": "#", "stock": "Supplier"}.get(dump_kind(dump_key), "Bill No")
            expect = spec.get("expect") or default
            ok = str(got).strip().lower() == expect.lower()
            flag = "ok " if ok else "MISMATCH"
            if not ok:
                problems += 1
            print(f"      {name:<22} row{hr} col{get_column_letter(col):<3} "
                  f"expect {expect!r:<10} found {str(got)!r:<14} {flag}")
        wb.close()
        print()
    print("Audit clean.\n" if not problems else f"{problems} problem(s) found.\n")
    return problems == 0


# ----------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="Fill the channel reports from the daily dumps.")
    ap.add_argument("--config", default=os.path.join(HERE, "config.json"))
    ap.add_argument("--audit", action="store_true", help="verify layout only")
    ap.add_argument("--dry-run", action="store_true", help="run but do not save")
    ap.add_argument("--brand", help="process only these reports (comma-separated)")
    ap.add_argument("--group", help="process only reports in this group, e.g. daily")
    ap.add_argument("--list", action="store_true", help="list the reports and their groups, then exit")
    ap.add_argument("--no-backup", action="store_true")
    ap.add_argument("--no-truncate", action="store_true",
                    help="use the comparison dumps whole, without cutting them to the "
                         "current month's latest day - for month-end runs")
    args = ap.parse_args()

    cfg = load_config(args.config)

    if args.list:
        print("\n  report              group")
        print("  " + "-" * 40)
        for name, bc in cfg["brands"].items():
            print(f"  {name:<20}{bc.get('group', 'daily')}")
        print()
        sys.exit(0)

    base = os.path.dirname(os.path.abspath(args.config))
    dumps_dir = os.path.join(base, cfg["paths"]["dumps"])
    reports_dir = os.path.join(base, cfg["paths"]["reports"])
    logs_dir = os.path.join(base, cfg["paths"]["logs"])

    if args.audit:
        if args.brand:
            sel = {b.strip(): cfg["brands"][b.strip()] for b in args.brand.split(",")
                   if b.strip() in cfg["brands"]}
        elif args.group:
            sel = {b: bc for b, bc in cfg["brands"].items()
                   if bc.get("group", "daily") == args.group}
        else:
            sel = None
        sys.exit(0 if audit(cfg, reports_dir, sel) else 1)

    logfile = setup_logging(logs_dir)
    log.info("=" * 62)
    log.info("run started%s", "  (DRY RUN)" if args.dry_run else "")

    if args.brand:
        wanted = [b.strip() for b in args.brand.split(",") if b.strip()]
        unknown = [b for b in wanted if b not in cfg["brands"]]
        if unknown:
            log.error("unknown report name(s): %s", ", ".join(unknown))
            log.error("run with --list to see the valid names")
            sys.exit(1)
        targets = {b: cfg["brands"][b] for b in wanted}
    elif args.group:
        targets = {b: bc for b, bc in cfg["brands"].items()
                   if bc.get("group", "daily") == args.group}
        if not targets:
            log.error("no reports in group %r — run with --list to see the groups", args.group)
            sys.exit(1)
    else:
        targets = cfg["brands"]

    log.info("running %d of %d report(s): %s",
             len(targets), len(cfg["brands"]), ", ".join(targets))

    dumps = load_all_dumps(cfg, dumps_dir, required_dumps(targets))
    if not dumps:
        log.error("no dumps found in %s — nothing to do", dumps_dir)
        sys.exit(1)

    ref_day = None if args.no_truncate else reference_day(dumps)
    if args.no_truncate:
        log.info("--no-truncate: comparison periods used in full, no day cut-off")
    elif ref_day:
        log.info("comparison periods truncated to day %d of their month", ref_day)
    else:
        log.warning("could not read a reference day from the current-month dump; "
                    "comparison periods will NOT be truncated")

    MATCH_FIELDS.update(cfg["match_fields"])
    try:
        check_cluster_coverage(dumps, cfg["brands"], cfg["match_fields"])
    except ValueError as exc:
        log.error("Cluster validation failed; NO reports will be written: %s", exc)
        sys.exit(1)
    split, unmatched = split_by_brand(dumps, cfg, build_matcher(cfg))

    if unmatched:
        path = os.path.join(logs_dir, f"unmatched_{datetime.now():%Y-%m-%d}.csv")
        with open(path, "w", newline="", encoding="utf-8") as fh:
            w = csv.DictWriter(fh, fieldnames=["dump", "field", "value", "reason"])
            w.writeheader()
            w.writerows(unmatched)
        log.warning("%d unclassified row(s) written to %s", len(unmatched), os.path.basename(path))

    if args.no_backup and any(bc.get("group") == "cluster" for bc in targets.values()) and not args.dry_run:
        log.error("--no-backup is not supported for cluster report updates")
        sys.exit(1)
    backup_dir = (None if args.no_backup else
                  os.path.join(base, "backups", f"safe_run_{datetime.now():%Y%m%d_%H%M%S_%f}"))

    ok_count = 0
    for brand, bc in targets.items():
        try:
            ok, _ = process_brand(brand, bc, split, dumps, reports_dir, args.dry_run,
                                  backup_dir, ref_day, cfg["match_fields"],
                                  os.path.abspath(args.config))
            ok_count += int(ok)
        except Exception as exc:
            log.exception("%-17s FAILED: %s", brand, exc)

    log.info("finished — %d of %d report(s) updated", ok_count, len(targets))
    log.info("log: %s", logfile)
    sys.exit(0 if ok_count == len(targets) else 1)


if __name__ == "__main__":
    main()
