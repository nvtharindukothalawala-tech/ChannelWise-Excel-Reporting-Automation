"""Logging utilities for ChannelReports automation system."""
import logging
import os
import sys
from datetime import datetime
from pathlib import Path

def setup_logging(log_dir: str | Path, log_name_prefix: str = "run") -> tuple[logging.Logger, Path]:
    log_dir = Path(log_dir)
    log_dir.mkdir(parents=True, exist_ok=True)
    logfile = log_dir / f"{log_name_prefix}_{datetime.now():%Y-%m-%d}.log"

    logger = logging.getLogger("reports")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    fmt = logging.Formatter("%(asctime)s  %(levelname)-7s %(message)s", "%Y-%m-%d %H:%M:%S")

    file_handler = logging.FileHandler(logfile, encoding="utf-8")
    file_handler.setFormatter(fmt)
    logger.addHandler(file_handler)

    stream_handler = logging.StreamHandler(sys.stdout)
    stream_handler.setFormatter(fmt)
    logger.addHandler(stream_handler)

    return logger, logfile
