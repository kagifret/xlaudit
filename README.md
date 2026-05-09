# 📊 xlaudit

> Audit Excel workbooks for structure, dependencies, and migration complexity.

**xlaudit** scans `.xlsx` files and reports formula density, volatile functions, external links, named ranges, cross-sheet references, and computes a weighted complexity score — helping you assess migration risk before moving spreadsheets to modern platforms.

## Installation

```bash
pip install -e ".[dev]"
```

## Quick Start

### CLI

```bash
# Scan a single file
xlaudit scan budget_2024.xlsx

# Scan a directory (recursive)
xlaudit scan ./reports -r

# With per-sheet detail
xlaudit scan ./reports --detail

# Export to JSON / Markdown / HTML
xlaudit scan ./reports --output json --save report.json
xlaudit scan ./reports --output html --save report.html

# Quick one-liner summary
xlaudit summary ./reports
```


### GUI

```bash
xlaudit serve # launches at http://localhost:8000, auto-opens browser

xlaudit serve --port 9000 #custom port
```



### Terminal Output

```
 File                   KB   Sheets  Formulas  Ext. Links  Volatile  Named Ranges  Complexity
────────────────────────────────────────────────────────────────────────────────────────────────
 budget_2024.xlsx       84      6       312          3          2           8        35.4 [HIGH]
 sales_summary.xlsx     21      2        48          0          0           3         4.1 [LOW]
 kpi_dashboard.xlsx     56      4       201          1          5          12        28.7 [HIGH]

3 workbook(s) scanned.  Complexity: LOW < 10  MED 10–25  HIGH > 25
```

### Library API

```python
from xlaudit import scan_workbook, scan_directory

# Single file
result = scan_workbook("budget_2024.xlsx")
print(result.complexity_score)   # 35.4
print(result.complexity_band)    # "HIGH"
print(result.total_external_links)  # 3

for sheet in result.sheets:
    print(f"  {sheet.name}: {sheet.external_refs}")

# Directory
report = scan_directory("./reports", recursive=True)
for wb in report.sorted_by_complexity():
    print(f"{wb.file_name}: {wb.complexity_score} [{wb.complexity_band}]")
```

## Complexity Scoring

| Metric | Weight |
|--------|-------:|
| External links | × 5.0 |
| Volatile functions (NOW, TODAY, RAND, …) | × 2.0 |
| Cross-sheet references | × 0.1 |
| Formula density (ratio) | × 15.0 |
| Sheet count | × 0.5 |
| Named ranges | × 0.3 |

### Bands

| Band | Score Range |
|------|------------|
| 🟢 LOW | < 10 |
| 🟡 MED | 10 – 25 |
| 🔴 HIGH | > 25 |

## Development

```bash
# Install dev dependencies
pip install -e ".[dev]"

# Run tests
pytest -v

# Run with coverage
pytest --cov=xlaudit --cov-report=term-missing
```

## Project Structure

```
xlaudit/
├── src/xlaudit/
│   ├── __init__.py      # Public API
│   ├── cli.py           # Click CLI
│   ├── models.py        # Dataclass models
│   ├── parser.py        # openpyxl workbook parsing
│   ├── scanner.py       # High-level scan API
│   ├── analysis.py      # Complexity scoring
│   └── reports.py       # JSON / Markdown / HTML rendering
├── tests/
│   ├── test_parser.py   # Link & volatile detection tests
│   └── test_analysis.py # Scoring & band tests
├── pyproject.toml
└── README.md
```

## License

GPL-3.0-or-later
