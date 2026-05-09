"""Report rendering — JSON, Markdown, and HTML."""

from __future__ import annotations

import json
from pathlib import Path
from typing import TYPE_CHECKING

from jinja2 import Environment, BaseLoader

if TYPE_CHECKING:
    from xlaudit.models import ScanReport

# ── JSON ─────────────────────────────────────────────────────────────────

def render_json(report: "ScanReport", *, indent: int = 2) -> str:
    return json.dumps(report.to_dict(), indent=indent, ensure_ascii=False)

# ── Markdown ─────────────────────────────────────────────────────────────

_MD = """\
# xlaudit Scan Report

**Path:** `{{ r.scan_path }}`  **Files:** {{ r.total_files }}

| File | KB | Sheets | Formulas | Ext. Links | Volatile | Named Ranges | Complexity |
|------|---:|-------:|---------:|-----------:|---------:|-------------:|-----------:|
{% for w in wbs -%}
| {{ w.file_name }} | {{ w.file_size_kb }} | {{ w.sheet_count }} | {{ w.total_formulas }} | {{ w.total_external_links }} | {{ w.total_volatile }} | {{ w.named_range_count }} | {{ w.complexity_score }} [{{ w.complexity_band }}] |
{% endfor %}
> **Bands:** LOW < 10 · MED 10–25 · HIGH > 25
{% if detail %}
{% for w in wbs %}
### {{ w.file_name }}
| Sheet | Formulas | Volatile | Cross-sheet | External |
|-------|-------:|---------:|------------:|----------|
{% for s in w.sheets -%}
| {{ s.name }} | {{ s.formula_count }} | {{ s.volatile_count }} | {{ s.cross_sheet_ref_count }} | {{ s.external_refs | join(', ') or '—' }} |
{% endfor %}
{% endfor %}
{% endif %}
"""

def render_markdown(report: "ScanReport", *, detail: bool = False) -> str:
    env = Environment(loader=BaseLoader(), autoescape=False)
    return env.from_string(_MD).render(r=report, wbs=report.sorted_by_complexity(), detail=detail)

# ── HTML ─────────────────────────────────────────────────────────────────

_CSS = """
:root{--bg:#0f1117;--sf:#1a1d27;--bd:#2a2d3a;--tx:#e4e4e7;--tm:#9ca3af;--ac:#818cf8;--ag:rgba(129,140,248,.15);--lo:#34d399;--md:#fbbf24;--hi:#f87171;--fn:'Segoe UI',system-ui,sans-serif;--mo:'Cascadia Code','Consolas',monospace}
*{margin:0;padding:0;box-sizing:border-box}
body{background:var(--bg);color:var(--tx);font-family:var(--fn);line-height:1.6;padding:2rem}
.c{max-width:1100px;margin:0 auto}
h1{font-size:1.75rem;font-weight:700;margin-bottom:.25rem;background:linear-gradient(135deg,var(--ac),#c084fc);-webkit-background-clip:text;-webkit-text-fill-color:transparent}
.sub{color:var(--tm);margin-bottom:2rem;font-size:.95rem}
.stats{display:flex;gap:1rem;margin-bottom:2rem;flex-wrap:wrap}
.sc{background:var(--sf);border:1px solid var(--bd);border-radius:12px;padding:1rem 1.5rem;min-width:140px;flex:1}
.sc .v{font-size:1.5rem;font-weight:700;color:var(--ac)}
.sc .l{font-size:.8rem;color:var(--tm);text-transform:uppercase;letter-spacing:.05em}
table{width:100%;border-collapse:collapse;background:var(--sf);border-radius:12px;overflow:hidden;border:1px solid var(--bd);margin-bottom:2rem}
th{text-align:left;padding:.75rem 1rem;font-size:.75rem;text-transform:uppercase;letter-spacing:.06em;color:var(--tm);background:rgba(255,255,255,.03);border-bottom:1px solid var(--bd)}
td{padding:.65rem 1rem;border-bottom:1px solid var(--bd);font-family:var(--mo);font-size:.85rem}
tr:last-child td{border-bottom:none}tr:hover td{background:var(--ag)}
.b{padding:.15rem .55rem;border-radius:6px;font-size:.75rem;font-weight:600}
.b-low{background:rgba(52,211,153,.15);color:var(--lo)}
.b-med{background:rgba(251,191,36,.15);color:var(--md)}
.b-high{background:rgba(248,113,113,.15);color:var(--hi)}
.ft{text-align:center;color:var(--tm);font-size:.8rem;margin-top:2rem;padding-top:1rem;border-top:1px solid var(--bd)}
.lg{color:var(--tm);font-size:.8rem;margin-bottom:1.5rem}
.ds{margin-top:1.5rem}.ds h3{font-size:1rem;font-weight:600;margin-bottom:.5rem;color:var(--ac)}
"""

_HTML = """\
<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>xlaudit Report</title><style>""" + _CSS + """</style></head><body><div class="c">
<h1>📊 xlaudit Report</h1><p class="sub">Scanned <code>{{ r.scan_path }}</code></p>
<div class="stats">
<div class="sc"><div class="v">{{ r.total_files }}</div><div class="l">Workbooks</div></div>
<div class="sc"><div class="v">{{ tf }}</div><div class="l">Total Formulas</div></div>
<div class="sc"><div class="v">{{ te }}</div><div class="l">External Links</div></div>
<div class="sc"><div class="v">{{ tv }}</div><div class="l">Volatile Calls</div></div>
</div>
<table><thead><tr><th>File</th><th>KB</th><th>Sheets</th><th>Formulas</th><th>Ext. Links</th><th>Volatile</th><th>Named Ranges</th><th>Complexity</th></tr></thead><tbody>
{% for w in wbs %}<tr><td>{{ w.file_name }}</td><td>{{ w.file_size_kb }}</td><td>{{ w.sheet_count }}</td><td>{{ w.total_formulas }}</td><td>{{ w.total_external_links }}</td><td>{{ w.total_volatile }}</td><td>{{ w.named_range_count }}</td><td>{{ w.complexity_score }} <span class="b b-{{ w.complexity_band|lower }}">{{ w.complexity_band }}</span></td></tr>{% endfor %}
</tbody></table>
<p class="lg"><strong>Bands:</strong> LOW &lt; 10 · MED 10–25 · HIGH &gt; 25</p>
{% if detail %}{% for w in wbs %}<div class="ds"><h3>{{ w.file_name }}</h3><table><thead><tr><th>Sheet</th><th>Formulas</th><th>Volatile</th><th>Cross-sheet</th><th>External refs</th></tr></thead><tbody>
{% for s in w.sheets %}<tr><td>{{ s.name }}</td><td>{{ s.formula_count }}</td><td>{{ s.volatile_count }}</td><td>{{ s.cross_sheet_ref_count }}</td><td>{{ s.external_refs|join(', ') or '—' }}</td></tr>{% endfor %}
</tbody></table></div>{% endfor %}{% endif %}
<div class="ft">Generated by <strong>xlaudit {{ ver }}</strong></div></div></body></html>
"""

def render_html(report: "ScanReport", *, detail: bool = False) -> str:
    from xlaudit import __version__
    env = Environment(loader=BaseLoader(), autoescape=True)
    wbs = report.sorted_by_complexity()
    return env.from_string(_HTML).render(
        r=report, wbs=wbs, detail=detail, ver=__version__,
        tf=sum(w.total_formulas for w in wbs),
        te=sum(w.total_external_links for w in wbs),
        tv=sum(w.total_volatile for w in wbs),
    )

# ── Save helper ──────────────────────────────────────────────────────────

def save_report(content: str, path: str | Path) -> Path:
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(content, encoding="utf-8")
    return p.resolve()
