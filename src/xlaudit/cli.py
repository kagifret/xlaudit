"""CLI entry point for xlaudit."""

from __future__ import annotations

import sys
from pathlib import Path

import click
from rich.console import Console
from rich.table import Table

from xlaudit import __version__
from xlaudit.scanner import scan_workbook, scan_directory
from xlaudit.models import ScanReport
from xlaudit.reports import render_json, render_markdown, render_html, save_report

console = Console()


def _build_table(report: ScanReport, *, detail: bool = False) -> Table:
    """Build a Rich table matching the spec's terminal output."""
    tbl = Table(
        title=None,
        show_header=True,
        header_style="bold cyan",
        border_style="dim",
        row_styles=["", "dim"],
        pad_edge=True,
    )
    tbl.add_column("File", style="white", no_wrap=True)
    tbl.add_column("KB", justify="right")
    tbl.add_column("Sheets", justify="right")
    tbl.add_column("Formulas", justify="right")
    tbl.add_column("Ext. Links", justify="right")
    tbl.add_column("Volatile", justify="right")
    tbl.add_column("Named Ranges", justify="right")
    tbl.add_column("Complexity", justify="right")

    band_colors = {"LOW": "green", "MED": "yellow", "HIGH": "red"}

    for wb in report.sorted_by_complexity():
        color = band_colors.get(wb.complexity_band, "white")
        score_str = f"{wb.complexity_score} [{wb.complexity_band}]"
        tbl.add_row(
            wb.file_name,
            str(wb.file_size_kb),
            str(wb.sheet_count),
            str(wb.total_formulas),
            str(wb.total_external_links),
            str(wb.total_volatile),
            str(wb.named_range_count),
            f"[{color}]{score_str}[/{color}]",
        )

    return tbl


def _scan_path(path: str, recursive: bool) -> ScanReport:
    """Scan a file or directory and return a ScanReport."""
    p = Path(path)
    if p.is_file():
        result = scan_workbook(p)
        report = ScanReport(scan_path=str(p), workbooks=[result])
    elif p.is_dir():
        report = scan_directory(p, recursive=recursive)
    else:
        console.print(f"[red]Error:[/red] Path not found: {path}")
        sys.exit(1)
    return report


@click.group()
@click.version_option(__version__, prog_name="xlaudit")
def cli():
    """xlaudit -- Audit Excel workbooks for complexity and migration risk."""
    pass


@cli.command()
@click.argument("path")
@click.option("--detail", is_flag=True, help="Show per-sheet breakdown.")
@click.option("-r", "--recursive", is_flag=True, help="Scan subdirectories.")
@click.option(
    "--output",
    type=click.Choice(["json", "markdown", "html"], case_sensitive=False),
    default=None,
    help="Export format (prints to stdout if --save is omitted).",
)
@click.option("--save", "save_path", default=None, help="Save report to file.")
def scan(path: str, detail: bool, recursive: bool, output: str | None, save_path: str | None):
    """Scan .xlsx file(s) and display a complexity report."""
    report = _scan_path(path, recursive)

    if report.total_files == 0:
        console.print("[yellow]No .xlsx files found.[/yellow]")
        return

    # Always show terminal table
    console.print()
    console.print(_build_table(report, detail=detail))
    console.print()
    console.print(
        f"[bold]{report.total_files}[/bold] workbook(s) scanned.  "
        f"Complexity: [green]LOW < 10[/green]  [yellow]MED 10–25[/yellow]  [red]HIGH > 25[/red]"
    )

    # Export if requested
    if output:
        renderers = {
            "json": lambda: render_json(report),
            "markdown": lambda: render_markdown(report, detail=detail),
            "html": lambda: render_html(report, detail=detail),
        }
        content = renderers[output]()
        if save_path:
            out = save_report(content, save_path)
            console.print(f"\n[green]✓[/green] Report saved to [bold]{out}[/bold]")
        else:
            console.print()
            console.print(content)


@cli.command()
@click.argument("path")
@click.option("-r", "--recursive", is_flag=True, help="Scan subdirectories.")
def summary(path: str, recursive: bool):
    """One-liner summary sorted by complexity."""
    report = _scan_path(path, recursive)

    if report.total_files == 0:
        console.print("[yellow]No .xlsx files found.[/yellow]")
        return

    for wb in report.sorted_by_complexity():
        band_colors = {"LOW": "green", "MED": "yellow", "HIGH": "red"}
        c = band_colors.get(wb.complexity_band, "white")
        console.print(
            f"  [{c}]{wb.complexity_band:>4}[/{c}]  "
            f"{wb.complexity_score:6.1f}  "
            f"{wb.file_name}"
        )

    console.print(
        f"\n[bold]{report.total_files}[/bold] workbook(s).  "
        f"[green]LOW < 10[/green]  [yellow]MED 10–25[/yellow]  [red]HIGH > 25[/red]"
    )


if __name__ == "__main__":
    cli()
