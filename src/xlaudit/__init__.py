"""xlaudit — Audit Excel workbooks for structure, dependencies, and migration complexity."""

__version__ = "0.1.0"

from xlaudit.scanner import scan_workbook, scan_directory  # noqa: F401

__all__ = ["scan_workbook", "scan_directory", "__version__"]
