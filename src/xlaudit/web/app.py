"""FastAPI web application for xlaudit dashboard."""

from __future__ import annotations

import tempfile
import shutil
from pathlib import Path
from typing import List

from fastapi import FastAPI, File, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from xlaudit import __version__
from xlaudit.scanner import scan_workbook
from xlaudit.models import ScanReport, WorkbookResult

_WEB_DIR = Path(__file__).parent
_STATIC_DIR = _WEB_DIR / "static"

app = FastAPI(title="xlaudit", version=__version__)
app.mount("/static", StaticFiles(directory=str(_STATIC_DIR)), name="static")


def _build_graph(workbooks: List[WorkbookResult]) -> dict:
    """Build a D3-compatible graph from scan results."""
    nodes = []
    links = []
    external_set: set[str] = set()

    for wb in workbooks:
        # Workbook group node
        nodes.append({
            "id": wb.file_name,
            "type": "workbook",
            "band": wb.complexity_band,
            "score": wb.complexity_score,
            "kb": wb.file_size_kb,
        })
        for s in wb.sheets:
            sid = f"{wb.file_name}/{s.name}"
            nodes.append({
                "id": sid,
                "type": "sheet",
                "parent": wb.file_name,
                "formulas": s.formula_count,
                "volatile": s.volatile_count,
            })
            # Edge: sheet → workbook (containment)
            links.append({
                "source": sid,
                "target": wb.file_name,
                "type": "contains",
                "value": 1,
            })
            # Edges: cross-sheet targets
            for target_name, count in s.cross_sheet_targets.items():
                tid = f"{wb.file_name}/{target_name}"
                links.append({
                    "source": sid,
                    "target": tid,
                    "type": "cross_sheet",
                    "value": count,
                })
            # Edges: external links
            for ext in s.external_refs:
                if ext not in external_set:
                    external_set.add(ext)
                    nodes.append({
                        "id": ext,
                        "type": "external",
                        "band": "UNKNOWN",
                        "score": 0,
                    })
                links.append({
                    "source": sid,
                    "target": ext,
                    "type": "external",
                    "value": 3,
                })

    return {"nodes": nodes, "links": links}


@app.get("/", response_class=HTMLResponse)
async def index():
    """Serve the single-page dashboard."""
    html = (_WEB_DIR / "templates" / "index.html").read_text(encoding="utf-8")
    return HTMLResponse(content=html)


@app.post("/api/scan")
async def api_scan(files: List[UploadFile] = File(...)):
    """Accept uploaded .xlsx files, scan them, return results + graph."""
    results: List[WorkbookResult] = []
    errors: List[str] = []
    tmp_dir = Path(tempfile.mkdtemp(prefix="xlaudit_"))

    try:
        for f in files:
            if not f.filename or not f.filename.lower().endswith(".xlsx"):
                errors.append(f"Skipped {f.filename}: not an .xlsx file")
                continue
            tmp_path = tmp_dir / f.filename
            with open(tmp_path, "wb") as out:
                content = await f.read()
                out.write(content)
            try:
                result = scan_workbook(tmp_path)
                results.append(result)
            except Exception as exc:
                errors.append(f"Error scanning {f.filename}: {exc}")
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

    report = ScanReport(workbooks=results, scan_path="upload")
    graph = _build_graph(results)

    return JSONResponse({
        "version": __version__,
        "report": report.to_dict(),
        "graph": graph,
        "errors": errors,
    })


def create_app() -> FastAPI:
    """Factory for external use."""
    return app
