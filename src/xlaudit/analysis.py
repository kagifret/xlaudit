"""Complexity scoring and analysis utilities."""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from xlaudit.models import WorkbookResult

# ── Complexity weights ──────────────────────────────────────────────────────
WEIGHTS = {
    "external_links": 5.0,
    "volatile_functions": 2.0,
    "cross_sheet_refs": 0.1,
    "formula_density": 15.0,
    "sheet_count": 0.5,
    "named_ranges": 0.3,
}

# ── Band thresholds ─────────────────────────────────────────────────────────
BAND_LOW = 10
BAND_HIGH = 25


def compute_complexity(
    *,
    external_links: int,
    volatile_functions: int,
    cross_sheet_refs: int,
    formula_density: float,
    sheet_count: int,
    named_ranges: int,
) -> float:
    """Return a weighted complexity score."""
    score = (
        external_links * WEIGHTS["external_links"]
        + volatile_functions * WEIGHTS["volatile_functions"]
        + cross_sheet_refs * WEIGHTS["cross_sheet_refs"]
        + formula_density * WEIGHTS["formula_density"]
        + sheet_count * WEIGHTS["sheet_count"]
        + named_ranges * WEIGHTS["named_ranges"]
    )
    return round(score, 1)


def classify_band(score: float) -> str:
    """Return LOW / MED / HIGH label for a given complexity score."""
    if score < BAND_LOW:
        return "LOW"
    elif score <= BAND_HIGH:
        return "MED"
    else:
        return "HIGH"


def score_workbook(wb: "WorkbookResult") -> None:
    """Compute and attach complexity_score and complexity_band to *wb* in‑place."""
    wb.complexity_score = compute_complexity(
        external_links=wb.total_external_links,
        volatile_functions=wb.total_volatile,
        cross_sheet_refs=wb.total_cross_sheet_refs,
        formula_density=wb.formula_density,
        sheet_count=wb.sheet_count,
        named_ranges=wb.named_range_count,
    )
    wb.complexity_band = classify_band(wb.complexity_score)
