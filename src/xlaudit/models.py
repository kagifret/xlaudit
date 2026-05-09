"""Data models for xlaudit scan results."""

from __future__ import annotations

import dataclasses
from typing import Dict, List, Optional


@dataclasses.dataclass
class SheetResult:
    """Analysis result for a single worksheet."""

    name: str
    formula_count: int = 0
    volatile_count: int = 0
    cross_sheet_ref_count: int = 0
    external_refs: List[str] = dataclasses.field(default_factory=list)
    cross_sheet_targets: Dict[str, int] = dataclasses.field(default_factory=dict)

    def to_dict(self) -> dict:
        return dataclasses.asdict(self)


@dataclasses.dataclass
class WorkbookResult:
    """Analysis result for an entire workbook."""

    file_path: str
    file_name: str
    file_size_kb: int
    sheet_count: int
    total_formulas: int
    total_external_links: int
    total_volatile: int
    total_cross_sheet_refs: int
    named_range_count: int
    formula_density: float
    complexity_score: float
    complexity_band: str
    sheets: List[SheetResult] = dataclasses.field(default_factory=list)

    def to_dict(self) -> dict:
        return dataclasses.asdict(self)


@dataclasses.dataclass
class ScanReport:
    """Container for a full scan run (one or more workbooks)."""

    workbooks: List[WorkbookResult] = dataclasses.field(default_factory=list)
    scan_path: Optional[str] = None

    @property
    def total_files(self) -> int:
        return len(self.workbooks)

    def sorted_by_complexity(self, descending: bool = True) -> List[WorkbookResult]:
        return sorted(
            self.workbooks,
            key=lambda w: w.complexity_score,
            reverse=descending,
        )

    def to_dict(self) -> dict:
        return {
            "scan_path": self.scan_path,
            "total_files": self.total_files,
            "workbooks": [w.to_dict() for w in self.workbooks],
        }
