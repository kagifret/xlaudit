"""High-level scanning API — public interface of xlaudit."""

from __future__ import annotations

import os
from pathlib import Path
from typing import List, Optional

from xlaudit.analysis import score_workbook
from xlaudit.models import ScanReport, SheetResult, WorkbookResult
from xlaudit.parser import (
    count_cross_sheet_refs,
    count_volatile,
    detect_external_links,
    get_named_ranges,
    iter_formulas,
    load_workbook,
)


def scan_workbook(path: str | Path) -> WorkbookResult:
    """Scan a single .xlsx workbook and return a :class:`WorkbookResult`."""
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")
    if path.suffix.lower() != ".xlsx":
        raise ValueError(f"Not an .xlsx file: {path}")

    wb = load_workbook(str(path))

    file_size_kb = round(path.stat().st_size / 1024)
    named_ranges = get_named_ranges(wb)

    sheet_results: List[SheetResult] = []
    total_formulas = 0
    total_volatile = 0
    total_cross_sheet = 0
    all_external: set[str] = set()
    total_cells = 0

    for ws in wb.worksheets:
        formulas = iter_formulas(ws)
        ext_links = detect_external_links(formulas)
        vol = count_volatile(formulas)
        cross = count_cross_sheet_refs(formulas)

        # Count total cells for density calculation
        cell_count = 0
        for row in ws.iter_rows():
            cell_count += len(row)
        total_cells += cell_count

        sr = SheetResult(
            name=ws.title,
            formula_count=len(formulas),
            volatile_count=vol,
            cross_sheet_ref_count=cross,
            external_refs=sorted(ext_links),
        )
        sheet_results.append(sr)
        total_formulas += len(formulas)
        total_volatile += vol
        total_cross_sheet += cross
        all_external.update(ext_links)

    wb.close()

    # formula density = fraction of cells that contain formulas (0.0–1.0)
    formula_density = total_formulas / total_cells if total_cells > 0 else 0.0

    result = WorkbookResult(
        file_path=str(path),
        file_name=path.name,
        file_size_kb=file_size_kb,
        sheet_count=len(sheet_results),
        total_formulas=total_formulas,
        total_external_links=len(all_external),
        total_volatile=total_volatile,
        total_cross_sheet_refs=total_cross_sheet,
        named_range_count=len(named_ranges),
        formula_density=round(formula_density, 4),
        complexity_score=0.0,
        complexity_band="LOW",
        sheets=sheet_results,
    )
    score_workbook(result)
    return result


def scan_directory(
    path: str | Path,
    *,
    recursive: bool = False,
) -> ScanReport:
    """Scan all .xlsx files in a directory and return a :class:`ScanReport`."""
    path = Path(path)
    if not path.is_dir():
        raise NotADirectoryError(f"Not a directory: {path}")

    pattern = "**/*.xlsx" if recursive else "*.xlsx"
    xlsx_files = sorted(path.glob(pattern))

    report = ScanReport(scan_path=str(path))
    for f in xlsx_files:
        try:
            result = scan_workbook(f)
            report.workbooks.append(result)
        except Exception as exc:
            # Skip files that fail to parse, but could add logging later
            import sys
            print(f"⚠ Skipping {f.name}: {exc}", file=sys.stderr)

    return report
