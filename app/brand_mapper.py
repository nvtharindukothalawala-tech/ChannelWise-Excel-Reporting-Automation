"""Brand mapping, location normalization, and cluster classification."""
from __future__ import annotations

import csv
import logging
from collections import Counter
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from .data_cleaner import to_decimal

log = logging.getLogger("reports")

def normalized_location(value: any) -> str:
    """Normalize whitespace and case without altering business names."""
    return " ".join(str(value or "").split()).casefold()

def build_matcher(cfg: dict):
    """Build brand matching function based on keyword lists."""
    kw = {b: [k.lower() for k in words] for b, words in cfg["brand_keywords"].items()}

    def match(value: any) -> tuple[str | None, str | None]:
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

def select_by_exact_locations(dump: dict, kind: str, allowed_locations: list[str], match_fields: dict) -> list[list[any]]:
    """Select only rows matching explicitly configured exact locations."""
    field_name = match_fields[kind]
    headers = dump["headers"]
    if field_name not in headers:
        raise ValueError(f"Classification field {field_name!r} not in {dump['file']}")

    col_idx = headers.index(field_name)
    allowed_set = {normalized_location(loc) for loc in allowed_locations}
    if len(allowed_set) != len(allowed_locations):
        raise ValueError(f"Duplicate approved locations configured in {dump['file']}")

    return [
        row for row in dump["rows"]
        if normalized_location(row[col_idx] if col_idx < len(row) else None) in allowed_set
    ]

def keyword_filter(rows: list[list[any]], headers: list[str], kind: str,
                   match_fields: dict, keywords: list[str], exclude: list[str] | None = None) -> list[list[any]]:
    """Filter rows by keywords in match field."""
    field = match_fields.get(kind)
    if not field or field not in headers:
        log.error("cannot apply keywords: %r not in dump headers", field)
        return []
    col = headers.index(field)
    inc = [k.lower() for k in keywords]
    exc = [k.lower() for k in (exclude or [])]
    out = []
    for row in rows:
        val = row[col] if col < len(row) else None
        if val is None:
            continue
        text = str(val).strip().lower()
        if not any(k in text for k in inc):
            continue
        if exc and any(k in text for k in exc):
            continue
        out.append(row)
    return out

def split_by_brand(dumps: dict, cfg: dict, matcher) -> tuple[dict, list[dict]]:
    """Split dumps into brand buckets, capturing unmatched rows."""
    field_for = cfg["match_fields"]
    result = {}
    unmatched = []

    for key, dump in dumps.items():
        kind = "returns" if "returns" in key else ("stock" if "stock" in key else "sales")
        field = field_for.get(kind, "Location")
        headers = dump["headers"]
        if field not in headers:
            log.error("%s: column %r not found in headers", dump["file"], field)
            result[key] = {}
            continue

        col = headers.index(field)
        buckets = {"__ALL__": list(dump["rows"])}
        seen_unmatched = Counter()

        for row in dump["rows"]:
            val = row[col] if col < len(row) else None
            brand, reason = matcher(val)
            if brand is None:
                key_v = str(val)
                seen_unmatched[key_v] += 1
                unmatched.append({
                    "dump": dump["file"],
                    "field": field,
                    "value": key_v,
                    "reason": reason,
                })
                continue
            buckets.setdefault(brand, []).append(row)

        result[key] = buckets
        for val, count in seen_unmatched.most_common():
            log.warning("%s: %d row(s) unclassified — %s = %r", dump["file"], count, field, val)

    return result, unmatched

def write_unmatched_log(unmatched: list[dict], logs_dir: Path) -> Path | None:
    """Save unclassified rows to CSV log."""
    if not unmatched:
        return None
    logs_dir.mkdir(parents=True, exist_ok=True)
    csv_path = logs_dir / f"unmatched_{datetime.now():%Y-%m-%d}.csv"
    with open(csv_path, "w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=["dump", "field", "value", "reason"])
        writer.writeheader()
        writer.writerows(unmatched)
    log.warning("%d unclassified row(s) written to %s", len(unmatched), csv_path.name)
    return csv_path

def validate_cluster_coverage(dumps: dict, reports_cfg: dict, match_fields: dict) -> dict[str, dict]:
    """
    Validate that 100% of current-month sales and returns rows are mapped
    to exactly one cluster (NonFashion or Fashion) with zero unapproved locations.
    """
    cluster_reports = {n: v for n, v in reports_cfg.items() if v.get("group") == "cluster"}
    if not cluster_reports:
        return {}

    reconciliation_summary = {}

    for dump_key in ("sales_current_month", "returns_current_month"):
        if dump_key not in dumps:
            continue

        kind = "returns" if dump_key.startswith("returns") else "sales"
        field_name = match_fields[kind]
        dump = dumps[dump_key]
        headers = dump["headers"]

        if field_name not in headers:
            raise ValueError(f"Match field {field_name!r} absent from {dump['file']}")

        loc_col = headers.index(field_name)
        location_cluster_map = {}

        for cluster_name, bc in cluster_reports.items():
            spec = bc["sheets"].get(dump_key)
            if not spec:
                continue
            for loc in spec.get("exact_locations", []):
                norm_loc = normalized_location(loc)
                if norm_loc in location_cluster_map:
                    raise ValueError(f"Location {loc!r} assigned to multiple clusters: {location_cluster_map[norm_loc]} and {cluster_name}")
                location_cluster_map[norm_loc] = cluster_name

        amount_name = "Sales Amount" if kind == "sales" else "Sale Amount"
        money_col = headers.index(amount_name) if amount_name in headers else None

        row_counts = Counter()
        amount_totals = {c: Decimal("0.00") for c in cluster_reports}
        unapproved = Counter()

        for row in dump["rows"]:
            loc_raw = row[loc_col] if loc_col < len(row) else None
            norm_loc = normalized_location(loc_raw)
            cluster = location_cluster_map.get(norm_loc)

            if cluster is None:
                unapproved[str(loc_raw)] += 1
                continue

            row_counts[cluster] += 1
            if money_col is not None and money_col < len(row):
                amount_totals[cluster] += to_decimal(row[money_col], amount_name)

        if unapproved:
            raise ValueError(f"Unapproved {field_name} values in {dump['file']}: {dict(unapproved)}")

        total_mapped_rows = sum(row_counts.values())
        if total_mapped_rows != len(dump["rows"]):
            raise ValueError(f"Cluster reconciliation failed for {dump['file']}: {total_mapped_rows} mapped vs {len(dump['rows'])} source rows")

        for cluster in cluster_reports:
            log.info("cluster validation %-18s %-22s rows=%5d amount=%12.2f",
                     cluster, dump_key, row_counts[cluster], amount_totals[cluster])

        reconciliation_summary[dump_key] = {
            "row_counts": dict(row_counts),
            "amount_totals": amount_totals,
            "total_rows": len(dump["rows"]),
        }

    return reconciliation_summary
