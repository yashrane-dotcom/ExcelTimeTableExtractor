"""
utils.py — Shared utility helpers.
"""

import logging, sys


def configure_logging(level: str = "INFO"):
    logging.basicConfig(
        stream=sys.stdout,
        level=getattr(logging, level.upper(), logging.INFO),
        format="%(asctime)s | %(levelname)-7s | %(name)s | %(message)s",
        datefmt="%H:%M:%S",
    )


def truncate(s: str, n: int = 60) -> str:
    return s[:n] + "…" if len(s) > n else s
