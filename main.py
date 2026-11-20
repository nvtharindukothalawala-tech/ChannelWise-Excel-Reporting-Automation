#!/usr/bin/env python3
"""
Infinite Business Ventures - Excel Reporting Automation System
Main CLI Application

Usage:
    python main.py --check
    python main.py --dry-run
    python main.py --report Giggles
    python main.py --all
    python main.py --validate
"""
from __future__ import annotations

import argparse
import sys
import time
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from app.config_loader import load_config
from app.data_loader import load_dumps, required_dump_keys
from app.brand_mapper import (
    build_matcher,
    split_by_brand,
    write_unmatched_log,
    validate_cluster_coverage,
)
from app.date_utils import find_reference_day
from app.logging_utils import setup_logging
from app.report_engine import process_report
from app.integrity_validator import validate_package_integrity

def audit_layout(cfg: dict, reports_dir: Path, targets: dict) -> bool:
    """Verify sheet layout and headers for configured reports."""
    import openpyxl
    print("\n" + "=" * 65)
    print("LAYOUT AUDIT — VERIFYING CONFIG AGAINST WORKBOOKS")
    print("=" * 65)
    problems = 0

    for brand, bc in targets.items():
        path = reports_dir / bc["file"]
        print(f"\n  [{brand}] {bc['file']}")
        if not path.is_file():
            print(f"      MISSING FILE: {path}")
            problems += 1
            continue

        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        for dump_key, spec in bc["sheets"].items():
            name = spec["sheet"]
            if name not in wb.sheetnames:
                print(f"      {name:<24} SHEET NOT FOUND")
                problems += 1
                continue

            ws = wb[name]
            hr = spec["header_row"]
            row = next(
                (r for i, r in enumerate(ws.iter_rows(min_row=hr, max_row=hr, max_col=60, values_only=True), 1)),
                (),
            )
            col = spec["paste_start_col"]
            got = row[col - 1] if col - 1 < len(row) else None

            kind = "returns" if "returns" in dump_key else ("stock" if "stock" in dump_key else "sales")
            default_expect = {"returns": "#", "stock": "Supplier"}.get(kind, "Bill No")
            expect = spec.get("expect") or default_expect
            ok = str(got).strip().lower() == expect.lower()
            flag = "PASS" if ok else "MISMATCH"
            if not ok:
                problems += 1
            print(f"      {name:<24} row {hr:<2} col {col:<2} expect {expect!r:<12} found {str(got)!r:<14} [{flag}]")
        wb.close()

    print("\n" + "=" * 65)
    if problems == 0:
        print("PASS: Layout audit clean. All sheets and header offsets match.")
    else:
        print(f"WARNING: Layout audit found {problems} problem(s).")
    print("=" * 65 + "\n")
    return problems == 0

def main():
    parser = argparse.ArgumentParser(
        description="Infinite Business Ventures - Excel Reporting Automation System"
    )
    default_cfg = HERE / "config" / "reports.json"
    if not default_cfg.is_file():
        default_cfg = HERE / "config_updated.json"

    parser.add_argument("--config", default=str(default_cfg), help="Path to config JSON")
    parser.add_argument("--check", action="store_true", help="Audit file paths, schemas, and templates without processing")
    parser.add_argument("--dry-run", action="store_true", help="Process data and validate updates without saving files")
    parser.add_argument("--report", help="Process specific report(s), comma-separated (e.g. Giggles,BU_Giggles)")
    parser.add_argument("--all", action="store_true", help="Process all 13 reports")
    parser.add_argument("--group", help="Process reports in group (daily, cluster, on_request)")
    parser.add_argument("--validate", action="store_true", help="Run read-only integrity checks on existing report files")
    parser.add_argument("--list", action="store_true", help="List all 13 reports and exit")
    parser.add_argument("--no-backup", action="store_true", help="Skip creating backup copies before saving")
    parser.add_argument("--no-truncate", action="store_true", help="Do not truncate comparison dumps to current reference day")

    args = parser.parse_args()

    cfg = load_config(args.config)
    base_dir = Path(args.config).resolve().parent
    if base_dir.name == "config":
        base_dir = base_dir.parent

    dumps_dir = base_dir / cfg["paths"]["dumps"]
    reports_dir = base_dir / cfg["paths"]["reports"]
    logs_dir = base_dir / cfg["paths"]["logs"]

    if args.list:
        print("\n  Configured Reporting Workbooks:")
        print("  " + "-" * 55)
        for name, bc in cfg["brands"].items():
            print(f"  {name:<24} group={bc.get('group', 'daily'):<12} file={bc['file']}")
        print()
        sys.exit(0)

    # Resolve target reporting jobs
    if args.report:
        wanted = [b.strip() for b in args.report.split(",") if b.strip()]
        unknown = [b for b in wanted if b not in cfg["brands"]]
        if unknown:
            print(f"ERROR: Unknown report name(s): {', '.join(unknown)}")
            print("Use --list to view valid report names.")
            sys.exit(1)
        targets = {b: cfg["brands"][b] for b in wanted}
    elif args.group:
        targets = {b: bc for b, bc in cfg["brands"].items() if bc.get("group") == args.group}
        if not targets:
            print(f"ERROR: No reports found in group {args.group!r}")
            sys.exit(1)
    else:
        # Default or --all: all 13 reports
        targets = cfg["brands"]

    if args.check:
        clean = audit_layout(cfg, reports_dir, targets)
        sys.exit(0 if clean else 1)

    if args.validate:
        print("\n" + "=" * 65)
        print("STANDALONE INTEGRITY VALIDATION")
        print("=" * 65)
        val_errors = 0
        for brand, bc in targets.items():
            rpath = reports_dir / bc["file"]
            errs = validate_package_integrity(rpath)
            if errs:
                print(f"FAIL: {brand} ({bc['file']}): {errs}")
                val_errors += 1
            else:
                print(f"PASS: {brand:<20} ({bc['file']}) - package integrity valid")
        print("=" * 65 + "\n")
        sys.exit(0 if val_errors == 0 else 1)

    # Live or Dry-Run Execution
    logger, logfile = setup_logging(logs_dir)
    logger.info("=" * 65)
    logger.info("REPORTING RUN STARTED%s", "  (DRY RUN)" if args.dry_run else "")
    logger.info("Target reports (%d): %s", len(targets), ", ".join(targets))
    logger.info("=" * 65)

    needed = required_dump_keys(targets)
    dumps, missing_dump_files = load_dumps(dumps_dir, cfg["dump_files"], needed)

    if not dumps:
        logger.error("No dumps found in %s — stopping", dumps_dir)
        sys.exit(1)

    if missing_dump_files:
        logger.warning("Missing dump files detected: %s", ", ".join(missing_dump_files))
        if any("stock" in f for f in missing_dump_files):
            logger.warning("Stock dumps are missing. Dependent stock sheets will be preserved and skipped.")

    # Reference day calculation for comparison period truncation
    ref_day = None if args.no_truncate else find_reference_day(dumps.get("sales_current_month"))
    if args.no_truncate:
        logger.info("--no-truncate: comparison periods used in full without day cut-off")
    elif ref_day:
        logger.info("Comparison periods truncated to day %d of their month", ref_day)
    else:
        logger.warning("Could not determine reference day; comparison periods will not be truncated")

    # Cluster coverage validation (fail-closed if unapproved locations exist)
    try:
        validate_cluster_coverage(dumps, cfg["brands"], cfg["match_fields"])
    except ValueError as exc:
        logger.error("Cluster coverage validation failed; stopping run: %s", exc)
        sys.exit(1)

    matcher = build_matcher(cfg)
    split_data, unmatched = split_by_brand(dumps, cfg, matcher)

    if unmatched:
        write_unmatched_log(unmatched, logs_dir)

    backup_dir = None
    if not args.no_backup and not args.dry_run:
        backup_dir = base_dir / "backups" / f"run_{datetime.now():%Y%m%d_%H%M%S}"

    results = []
    for brand, bc in targets.items():
        res = process_report(
            brand=brand,
            brand_cfg=bc,
            split_data=split_data,
            dumps=dumps,
            reports_dir=reports_dir,
            dry_run=args.dry_run,
            backup_dir=backup_dir,
            ref_day=ref_day,
            match_fields=cfg["match_fields"],
        )
        results.append(res)

    # Print summary execution table
    print("\n" + "=" * 70)
    print("REPORTING EXECUTION SUMMARY")
    print("=" * 70)
    print(f"{'Report':<22} {'Status':<15} {'Details / Rows Written'}")
    print("-" * 70)

    pass_count = 0
    incomplete_count = 0
    fail_count = 0

    for r in results:
        status = r.get("status", "UNKNOWN")
        bname = r.get("brand", "")
        if status == "PASS":
            pass_count += 1
            counts_str = ", ".join(f"{s}:{c}" for s, c in r.get("counts", {}).items())
            print(f"{bname:<22} [PASS]          {counts_str}")
        elif status == "INCOMPLETE":
            incomplete_count += 1
            counts_str = ", ".join(f"{s}:{c}" for s, c in r.get("counts", {}).items())
            print(f"{bname:<22} [INCOMPLETE]    Stock dump missing; sales/returns updated ({counts_str})")
        else:
            fail_count += 1
            reason = r.get("reason", "Error")
            print(f"{bname:<22} [FAILED]        {reason}")

    print("=" * 70)
    logger.info("Summary: %d PASS, %d INCOMPLETE (missing stock), %d FAILED of %d target reports",
                pass_count, incomplete_count, fail_count, len(targets))
    logger.info("Log file: %s", logfile)

    if fail_count > 0:
        sys.exit(1)
    sys.exit(0)

if __name__ == "__main__":
    main()
