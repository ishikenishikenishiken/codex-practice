"""Command-line Excel numeric analysis tool.

This script reads an Excel workbook and prints summary statistics for numeric
columns in each sheet. It is designed to operate without modifying any HTML
assets in the repository.
"""

from __future__ import annotations

import argparse
import math
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Optional

import pandas as pd


def _format_float(value: float, precision: int = 4) -> str:
    """Format a floating point number with a fixed precision.

    Non-finite values (NaN or infinities) are rendered as ``"-"`` to keep the
    output compact and readable.
    """

    if value is None or not math.isfinite(value):
        return "-"
    return f"{value:.{precision}f}"


@dataclass
class ColumnSummary:
    """A container for numeric column summary statistics."""

    name: str
    count: int
    mean: Optional[float]
    std: Optional[float]
    min: Optional[float]
    max: Optional[float]
    missing: int

    def render(self) -> str:
        return (
            f"{self.name}\n"
            f"  Count   : {self.count}\n"
            f"  Missing : {self.missing}\n"
            f"  Mean    : {_format_float(self.mean)}\n"
            f"  Std dev : {_format_float(self.std)}\n"
            f"  Min     : {_format_float(self.min)}\n"
            f"  Max     : {_format_float(self.max)}"
        )


@dataclass
class SheetReport:
    """Holds the analysis report for a single sheet."""

    name: str
    numeric_columns: List[ColumnSummary]
    correlation: Optional[pd.DataFrame]

    def has_numeric_columns(self) -> bool:
        return bool(self.numeric_columns)

    def render(self) -> str:
        lines: List[str] = [f"Sheet: {self.name}"]
        if not self.numeric_columns:
            lines.append("  (No numeric columns found)")
            return "\n".join(lines)

        lines.append("Numeric column summary:")
        for summary in self.numeric_columns:
            lines.append("  " + summary.render().replace("\n", "\n  "))

        if self.correlation is not None and not self.correlation.empty:
            lines.append("Correlation matrix:")
            corr = self.correlation.round(4)
            # Align columns for readability
            header = ["    "] + [f"{col:>12}" for col in corr.columns]
            lines.append("".join(header))
            for idx, row in corr.iterrows():
                values = "".join(f"{val:12.4f}" for val in row)
                lines.append(f"    {idx:>12}{values}")
        return "\n".join(lines)


def build_column_summaries(frame: pd.DataFrame) -> List[ColumnSummary]:
    numeric_frame = frame.select_dtypes(include="number")
    summaries: List[ColumnSummary] = []
    for column_name in numeric_frame.columns:
        series = numeric_frame[column_name]
        summaries.append(
            ColumnSummary(
                name=str(column_name),
                count=int(series.count()),
                mean=float(series.mean()) if not series.empty else None,
                std=float(series.std(ddof=0)) if series.count() > 1 else None,
                min=float(series.min()) if not series.empty else None,
                max=float(series.max()) if not series.empty else None,
                missing=int(frame[column_name].isna().sum()),
            )
        )
    return summaries


def build_sheet_report(sheet_name: str, frame: pd.DataFrame) -> SheetReport:
    summaries = build_column_summaries(frame)
    correlation = None
    if summaries:
        numeric_frame = frame.select_dtypes(include="number")
        if numeric_frame.shape[1] >= 2:
            correlation = numeric_frame.corr()
    return SheetReport(name=sheet_name, numeric_columns=summaries, correlation=correlation)


def analyze_workbook(path: Path, sheet: Optional[str] = None) -> List[SheetReport]:
    if sheet is None:
        content = pd.read_excel(path, sheet_name=None)
        # ``read_excel`` returns a DataFrame if there is only a single sheet. Normalize
        if isinstance(content, pd.DataFrame):
            content = {content.columns.name or "Sheet1": content}
    else:
        content = {sheet: pd.read_excel(path, sheet_name=sheet)}

    reports: List[SheetReport] = []
    for sheet_name, frame in content.items():
        reports.append(build_sheet_report(str(sheet_name), frame))
    return reports


def print_reports(reports: Iterable[SheetReport]) -> None:
    first = True
    for report in reports:
        if not first:
            print()
        print(report.render())
        first = False


def parse_arguments(argv: Optional[List[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Analyze numeric columns in an Excel workbook and display summary statistics.",
    )
    parser.add_argument("path", type=Path, help="Path to the Excel workbook (e.g., .xlsx file)")
    parser.add_argument(
        "--sheet",
        type=str,
        default=None,
        help="Name of the sheet to analyze. By default all sheets are processed.",
    )
    return parser.parse_args(argv)


def main(argv: Optional[List[str]] = None) -> int:
    args = parse_arguments(argv)
    if not args.path.exists():
        print(f"Error: {args.path} does not exist.", file=sys.stderr)
        return 1

    try:
        reports = analyze_workbook(args.path, sheet=args.sheet)
    except ValueError as exc:
        print(f"Failed to read Excel file: {exc}", file=sys.stderr)
        return 2

    print_reports(reports)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
