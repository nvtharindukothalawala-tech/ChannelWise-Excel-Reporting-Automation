"""Configuration loader and validator."""
from __future__ import annotations

import json
from pathlib import Path

REQUIRED_13_REPORTS = (
    "Giggles",
    "Protecto",
    "SafetySure",
    "BeDapper",
    "BeDesi",
    "BU_DailyRevenue",
    "NonFashionTracker",
    "FashionTracker",
    "BU_Giggles",
    "BU_Protecto",
    "BU_SafetySure",
    "BU_BeDapper",
    "BU_BeDesi",
)

def load_config(config_path: str | Path) -> dict:
    """Load and validate configuration JSON."""
    path = Path(config_path).resolve()
    if not path.is_file():
        raise FileNotFoundError(f"Configuration file not found: {path}")

    with open(path, "r", encoding="utf-8") as fh:
        cfg = json.load(fh)

    # Validate top-level keys
    for req_key in ("paths", "dump_files", "match_fields", "brand_keywords", "brands"):
        if req_key not in cfg:
            raise ValueError(f"Configuration missing required section: {req_key!r}")

    # Validate all 13 reports are present
    missing_reports = [r for r in REQUIRED_13_REPORTS if r not in cfg["brands"]]
    if missing_reports:
        raise ValueError(f"Configuration is missing reporting jobs: {missing_reports}")

    return cfg
