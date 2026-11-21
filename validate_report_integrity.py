
"""Read-only integrity validator for cluster reporting workbooks.

Compares monthly source dumps with an updated Excel report.
Checks non-data formulas and workbook components against a backup.

Does not save or modify any Excel workbook.
"""

from __future__ import annotations

import argparse
import json
import warnings

from collections import Counter
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import openpyxl
from openpyxl.utils.datetime import from_excel
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parent

CONFIG_FILE = ROOT / "config_updated.json"

NAMESPACE = (
    "http://schemas.openxmlformats.org/"
    "spreadsheetml/2006/main"
)


def parse_date(value, epoch):
    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            result = from_excel(value, epoch=epoch)
            if isinstance(result, datetime):
                return result.date()
            if isinstance(result, date):
                return result
        except (ValueError, TypeError, OverflowError):
            pass

    if isinstance(value, str):
        value = value.strip()

        for fmt in (
            "%Y-%m-%d %H:%M:%S",
            "%Y-%m-%d",
            "%d/%m/%Y %H:%M:%S",
            "%d/%m/%Y",
        ):
            try:
                return datetime.strptime(value, fmt).date()
            except ValueError:
                continue

        try:
            return datetime.fromisoformat(value).date()
        except ValueError:
            pass

    return None


def money(value):
    if value is None or value == "":
        return Decimal("0")

    try:
        result = Decimal(str(value).replace(",", ""))
        return result if result.is_finite() else None
    except (ValueError, InvalidOperation):
        return None


def read_month_data(path, month, location_field, date_field,
                    amount_field, allowed_locations):
    """Count and sum transactions belonging to the configured locations."""

    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        wb = openpyxl.load_workbook(
            BytesIO(Path(path).read_bytes()),
            read_only=True,
            data_only=True,
            keep_links=False,
        )

    try:
        ws = wb.active
        records = ws.iter_rows(values_only=True)
        headers = [str(x or "").strip() for x in next(records)]

        required = (location_field, date_field, amount_field)
        missing = [field for field in required if field not in headers]

        if missing:
            raise ValueError(
                f"{path.name}: missing columns {missing}"
            )

        li = headers.index(location_field)
        di = headers.index(date_field)
        ai = headers.index(amount_field)

        count = 0
        total = Decimal("0")
        locations = Counter()

        year, month_number = map(int, month.split("-"))
        allowed = set(allowed_locations)

        for row in records:
            location = str(row[li] or "").strip()

            if location not in allowed:
                continue

            transaction_date = parse_date(row[di], wb.epoch)

            if transaction_date is None:
                raise ValueError(
                    f"{path.name}: invalid date for configured location"
                )

            if (
                transaction_date.year,
                transaction_date.month,
            ) != (year, month_number):
                continue

            amount = money(row[ai])

            if amount is None:
                raise ValueError(
                    f"{path.name}: invalid financial amount"
                )

            count += 1
            total += amount
            locations[location] += 1

        return {
            "rows": count,
            "amount": total.quantize(Decimal("0.01")),
            "locations": locations,
        }

    finally:
        wb.close()


def formula_value(cell):
    """Normalize ordinary and array formulas for comparison."""

    if cell.data_type != "f":
        return None

    value = cell.value

    if isinstance(value, str):
        return ("formula", value)

    if hasattr(value, "text"):
        return (
            "array_formula",
            value.text,
            getattr(value, "ref", None),
        )

    return ("formula_object", str(value))


def read_dashboard_formulas(path):
    """Read formulas outside the transaction data worksheets."""

    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        wb = openpyxl.load_workbook(
            BytesIO(Path(path).read_bytes()),
            read_only=True,
            data_only=False,
            keep_links=False,
        )

    formulas = {}

    try:
        for ws in wb.worksheets:
            if ws.title in ("DataSet", "Returns"):
                continue

            for row in ws.iter_rows():
                for cell in row:
                    value = formula_value(cell)

                    if value is not None:
                        formulas[(ws.title, cell.coordinate)] = value

        return formulas

    finally:
        wb.close()


def workbook_components(path):
    """Inspect workbook XML without saving or recalculating it."""

    with ZipFile(path) as archive:
        names = archive.namelist()

        pivots = [
            name for name in names
            if name.startswith("xl/pivotTables/pivotTable")
            and name.endswith(".xml")
        ]

        caches = [
            name for name in names
            if name.startswith("xl/pivotCache/pivotCacheDefinition")
            and name.endswith(".xml")
        ]

        charts = [
            name for name in names
            if name.startswith("xl/charts/chart")
            and name.endswith(".xml")
        ]

        sheet_xmls = [
            name for name in names
            if name.startswith("xl/worksheets/sheet")
            and name.endswith(".xml")
            and "/" not in name[len("xl/worksheets/"):]
        ]

        cf_counts = {}

        for name in sheet_xmls:
            root = ET.fromstring(archive.read(name))
            cf_counts[name] = len(
                root.findall(f"{{{NAMESPACE}}}conditionalFormatting")
            )

        refresh = []

        for name in sorted(caches):
            root = ET.fromstring(archive.read(name))
            refresh.append({
                "cache": name,
                "refreshOnLoad": root.get("refreshOnLoad"),
                "enableRefresh": root.get("enableRefresh"),
            })

        return {
            "pivots": len(pivots),
            "caches": len(caches),
            "charts": len(charts),
            "conditional_formatting": cf_counts,
            "refresh": refresh,
        }


def check_equal(label, actual, expected, errors):
    if actual == expected:
        print(f"PASS: {label}: {actual}")
    else:
        print(
            f"FAIL: {label}: actual={actual}, "
            f"expected={expected}"
        )
        errors.append(label)


def validate(report_path, backup_path, config_path, month,
             report_name):
    errors = []

    year_month = date.fromisoformat(month + "-01")
    month = year_month.strftime("%Y-%m")

    config = json.loads(config_path.read_text(encoding="utf-8"))

    report_config = config["brands"][report_name]

    expected_name = report_config["file"]

    allowed_names = {
        expected_name,
        expected_name + ".tmp",
    }

    if report_path.name not in allowed_names:
        raise ValueError(
            f"Unexpected report filename: {report_path.name}"
        )

    sales_config = report_config["sheets"]["sales_current_month"]
    returns_config = report_config["sheets"]["returns_current_month"]

    dump_folder = config_path.parent / config["paths"]["dumps"]

    dump_names = config.get("dump_files", {})

    # Current project uses these standard source filenames.
    sales_dump = dump_folder / "sales_current_month.xlsx"
    returns_dump = dump_folder / "returns_current_month.xlsx"

    print("\n=== MONTHLY TRANSACTION VALIDATION ===")

    checks = (
        (
            "DataSet",
            sales_dump,
            sales_config,
            "Location",
            "Date",
            "Sales Amount",
        ),
        (
            "Returns",
            returns_dump,
            returns_config,
            "Warehouse",
            "Return Date",
            "Sale Amount",
        ),
    )

    for sheet_name, source_path, sheet_config, location_field, \
            date_field, amount_field in checks:

        expected = read_month_data(
            source_path,
            month,
            location_field,
            date_field,
            amount_field,
            sheet_config["exact_locations"],
        )

        actual = read_month_data_from_report(
            report_path,
            sheet_name,
            month,
            date_field,
            amount_field,
        )

        check_equal(
            f"{sheet_name} row count",
            actual["rows"],
            expected["rows"],
            errors,
        )

        check_equal(
            f"{sheet_name} financial total",
            actual["amount"],
            expected["amount"],
            errors,
        )

        if sheet_name == "DataSet":
            fit_location = (
                "Fit N Finish Shop - Fit N Finish Bin"
            )

            if fit_location in sheet_config["exact_locations"]:
                check_equal(
                    "Fit N Finish rows",
                    actual["locations"][fit_location],
                    expected["locations"][fit_location],
                    errors,
                )

    print("\n=== DASHBOARD FORMULA VALIDATION ===")

    old_formulas = read_dashboard_formulas(backup_path)
    new_formulas = read_dashboard_formulas(report_path)

    changed = [
        key
        for key in set(old_formulas) | set(new_formulas)
        if old_formulas.get(key) != new_formulas.get(key)
    ]

    check_equal(
        "Changed dashboard formulas",
        len(changed),
        0,
        errors,
    )

    if changed:
        print("First changed formula cells:", sorted(changed)[:10])

    print("\n=== EXCEL COMPONENT VALIDATION ===")

    before = workbook_components(backup_path)
    after = workbook_components(report_path)

    # Exact preservation and relationship integrity of original external-link parts.
    # This catches the previously missed externalBook rId1 -> missing .rels entry.
    print("\n=== EXTERNAL-LINK INTEGRITY ===")
    try:
        import external_link_patch
        original_parts = external_link_patch.snapshot(backup_path)
        staged_parts = external_link_patch.snapshot(report_path)
        check_equal("External-link component names", sorted(staged_parts),
                    sorted(original_parts), errors)
        check_equal("External-link component bytes preserved",
                    staged_parts == original_parts, True, errors)
    except Exception as exc:
        errors.append("External-link relationship integrity")
        print(f"FAIL: External-link relationship integrity: {type(exc).__name__}: {exc}")

    for key in ("pivots", "caches", "charts"):
        check_equal(
            key,
            after[key],
            before[key],
            errors,
        )

    check_equal(
        "Conditional formatting block counts",
        after["conditional_formatting"],
        before["conditional_formatting"],
        errors,
    )

    for cache in after["refresh"]:
        if (
            cache["refreshOnLoad"] not in ("1", "true")
            or cache["enableRefresh"] not in ("1", "true")
        ):
            errors.append("Pivot cache refresh configuration")
            print("FAIL: Pivot refresh configuration:", cache)
        else:
            print("PASS: Pivot refresh:", cache["cache"])

    print("\n=== FINAL VALIDATION RESULT ===")

    if errors:
        print("FAIL: REPORT REQUIRES REVIEW")
        print("Failed checks:", ", ".join(sorted(set(errors))))
        return 1

    print("PASS: ALL IMPLEMENTED INTEGRITY CHECKS PASSED")
    print("NOTE: Excel refresh and visual formatting still require review.")
    return 0


def read_month_data_from_report(
    path, sheet_name, month, date_field, amount_field
):
    """Read current-month transactions in the output report."""

    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        wb = openpyxl.load_workbook(
            BytesIO(Path(path).read_bytes()),
            read_only=True,
            data_only=True,
            keep_links=False,
        )

    try:
        ws = wb[sheet_name]
        records = ws.iter_rows(values_only=True)
        headers = [str(x or "").strip() for x in next(records)]

        location_field = (
            "Location" if sheet_name == "DataSet" else "Warehouse"
        )

        li = headers.index(location_field)
        di = headers.index(date_field)
        ai = headers.index(amount_field)

        year, month_number = map(int, month.split("-"))

        count = 0
        total = Decimal("0")
        locations = Counter()

        for row in records:
            transaction_date = parse_date(row[di], wb.epoch)

            if transaction_date is None:
                continue

            if (
                transaction_date.year,
                transaction_date.month,
            ) != (year, month_number):
                continue

            amount = money(row[ai])

            if amount is None:
                raise ValueError(
                    f"{sheet_name}: invalid financial amount"
                )

            count += 1
            total += amount
            locations[str(row[li] or "").strip()] += 1

        return {
            "rows": count,
            "amount": total.quantize(Decimal("0.01")),
            "locations": locations,
        }

    finally:
        wb.close()


def main():
    parser = argparse.ArgumentParser(
        description="Read-only Excel report integrity validator"
    )

    parser.add_argument("--report", required=True, type=Path)
    parser.add_argument("--backup", required=True, type=Path)
    parser.add_argument("--month", required=True)
    parser.add_argument(
        "--brand",
        default="NonFashionTracker",
    )
    parser.add_argument(
        "--config",
        type=Path,
        default=CONFIG_FILE,
    )

    args = parser.parse_args()

    for path in (args.report, args.backup, args.config):
        if not path.is_file():
            parser.error(f"File not found: {path}")

    return validate(
        args.report,
        args.backup,
        args.config,
        args.month,
        args.brand,
    )


if __name__ == "__main__":
    raise SystemExit(main())
