"""Consolidate selected monthly worksheets into one report workbook per input file.

For each input workbook the processor:

1. Reads the selected month worksheets (default August-December).
2. Locates the first tax-type column by its HEADER LABEL (default ``IT``) and
   treats every column from there up to the column before ``GRAND TOTAL`` as a
   tax type. Column positions may differ between months.
3. Copies each month into the output workbook, recomputing GRAND TOTAL per row
   as the sum of its tax-type columns.
4. Matches taxpayers across months by normalised name (trimmed, single-spaced,
   upper-case) and sums every tax type per taxpayer.
5. Writes a summary sheet with per-tax-type totals, one grand total per month
   and a combined grand total for the whole period.
6. Adds a TOTAL row (``=SUM`` formulas) to every sheet and formats zeros as blank.
"""

from __future__ import annotations

import logging
import os
import tempfile
import time
from collections import OrderedDict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable, Sequence

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .config import ReportConfig

log = logging.getLogger(__name__)

HEADER_ROW = 1
FIRST_DATA_ROW = 2

# positive ; negative ; zero ; text  -> the empty zero section displays 0 as blank
ZERO_AS_BLANK_NUMBER_FORMAT = "#,##0.00;-#,##0.00;;@"


class ReportError(Exception):
    """A workbook could not be processed because of its content."""


# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------
@dataclass
class TaxpayerRecord:
    metadata: dict[str, Any] = field(default_factory=dict)
    tax_values: dict[str, float] = field(default_factory=dict)


@dataclass
class MonthData:
    month_name: str
    records: OrderedDict[str, TaxpayerRecord]
    tax_headers: list[str]
    data_row_count: int
    skipped_rows: int


@dataclass
class WorkbookResult:
    input_path: Path
    output_path: Path
    month_rows: dict[str, int]
    combined_taxpayers: int
    combined_total: float
    seconds: float


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def clean_text(value: Any) -> str:
    if value is None:
        return ""
    return " ".join(str(value).strip().split())


def normalise_text(value: Any) -> str:
    return clean_text(value).upper()


def to_number(value: Any) -> float:
    """Convert a cell value to float. Blank/non-numeric -> 0; '(1,000)' -> -1000."""
    if value in (None, ""):
        return 0.0
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip().replace(",", "")
    if not text:
        return 0.0
    if text.startswith("(") and text.endswith(")"):
        text = "-" + text[1:-1]
    try:
        return float(text)
    except ValueError:
        return 0.0


def build_header_map(headers: Sequence[Any], sheet_title: str) -> dict[str, int]:
    """Return {NORMALISED_HEADER: 1-based column}. Duplicates: last one wins (warned)."""
    result: dict[str, int] = {}
    for col, value in enumerate(headers, 1):
        header = normalise_text(value)
        if not header:
            continue
        if header in result:
            log.warning(
                "Sheet '%s': duplicate header '%s' in columns %s and %s; using column %s.",
                sheet_title,
                header,
                get_column_letter(result[header]),
                get_column_letter(col),
                get_column_letter(col),
            )
        result[header] = col
    return result


def pad_row(row: Sequence[Any], length: int) -> list[Any]:
    values = list(row)
    if len(values) < length:
        values.extend([None] * (length - len(values)))
    return values[:length]


def find_sheet_name(workbook, wanted_name: str) -> str:
    """Find a worksheet case-insensitively, ignoring stray spaces."""
    wanted = normalise_text(wanted_name)
    for actual_name in workbook.sheetnames:
        if normalise_text(actual_name) == wanted:
            return actual_name
    raise ReportError(
        f"Missing worksheet '{wanted_name}'. Available: {', '.join(workbook.sheetnames)}"
    )


# ---------------------------------------------------------------------------
# Formatting
# ---------------------------------------------------------------------------
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)
TOTAL_FILL = PatternFill("solid", fgColor="D9EAF7")
DOUBLE_TOP = Side(style="double", color="1F4E78")
THIN_SIDE = Side(style="thin", color="B4C6E7")

PREFERRED_WIDTHS = (34, 42, 58, 16, 52, 24)  # by metadata position


def style_header(ws, last_col: int) -> None:
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 35
    for col in range(1, last_col + 1):
        cell = ws.cell(1, col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=THIN_SIDE)


def style_total_row(ws, row: int, first_numeric_col: int, last_col: int) -> None:
    for col in range(1, last_col + 1):
        cell = ws.cell(row, col)
        cell.font = Font(bold=True)
        cell.fill = TOTAL_FILL
        cell.border = Border(top=DOUBLE_TOP)
        cell.alignment = Alignment(vertical="center")
    for col in range(first_numeric_col, last_col + 1):
        ws.cell(row, col).number_format = ZERO_AS_BLANK_NUMBER_FORMAT


def set_widths(ws, last_col: int, first_numeric_col: int, metadata_cols: Iterable[int]) -> None:
    for col in range(1, min(first_numeric_col, last_col + 1)):
        ws.column_dimensions[get_column_letter(col)].width = 18
    for col, width in zip(metadata_cols, PREFERRED_WIDTHS):
        if col <= last_col:
            ws.column_dimensions[get_column_letter(col)].width = width
    for col in range(first_numeric_col, last_col + 1):
        ws.column_dimensions[get_column_letter(col)].width = 16


def format_numeric_area(
    ws, first_numeric_col: int, last_col: int, last_row: int, id_col: int | None
) -> None:
    for row in ws.iter_rows(
        min_row=FIRST_DATA_ROW, max_row=last_row, min_col=first_numeric_col, max_col=last_col
    ):
        for cell in row:
            cell.number_format = ZERO_AS_BLANK_NUMBER_FORMAT
    if id_col:
        for row in range(FIRST_DATA_ROW, last_row + 1):
            ws.cell(row, id_col).number_format = "@"


def write_total_row(ws, total_row: int, label_col: int, first_col: int, last_col: int, data_rows: int) -> None:
    ws.cell(total_row, label_col, "TOTAL")
    last_data_row = FIRST_DATA_ROW + data_rows - 1
    for col in range(first_col, last_col + 1):
        letter = get_column_letter(col)
        value = f"=SUM({letter}{FIRST_DATA_ROW}:{letter}{last_data_row})" if data_rows else 0
        ws.cell(total_row, col, value)


# ---------------------------------------------------------------------------
# Month processing
# ---------------------------------------------------------------------------
def process_month_sheet(month_name: str, source_ws, output_ws, config: ReportConfig) -> MonthData:
    """Copy one month into ``output_ws`` and collect its taxpayer/tax-type totals."""
    rows = source_ws.iter_rows(values_only=True)
    try:
        raw_headers = next(rows)
    except StopIteration:
        raise ReportError(f"Worksheet '{source_ws.title}' is empty.") from None

    headers = [clean_text(value) for value in raw_headers]
    lookup = build_header_map(headers, source_ws.title)

    required = [*config.metadata_headers, config.first_tax_header, config.grand_total_header]
    missing = [h for h in required if normalise_text(h) not in lookup]
    if missing:
        raise ReportError(f"Worksheet '{source_ws.title}' is missing header(s): {', '.join(missing)}")

    name_col = lookup[normalise_text(config.name_header)]
    first_tax_col = lookup[normalise_text(config.first_tax_header)]
    grand_total_col = lookup[normalise_text(config.grand_total_header)]
    if first_tax_col >= grand_total_col:
        raise ReportError(
            f"Worksheet '{source_ws.title}': header '{config.first_tax_header}' must appear "
            f"before '{config.grand_total_header}'."
        )

    headers = pad_row(headers, grand_total_col)
    output_ws.append(headers)

    tax_headers = [
        clean_text(headers[col - 1]) or f"UNNAMED_TAX_COLUMN_{get_column_letter(col)}"
        for col in range(first_tax_col, grand_total_col)
    ]
    metadata_cols = {h: lookup[normalise_text(h)] for h in config.metadata_headers}
    skip_labels = {normalise_text(x) for x in config.skip_name_labels}

    records: OrderedDict[str, TaxpayerRecord] = OrderedDict()
    data_rows = 0
    skipped = 0

    for raw_row in rows:
        row = pad_row(raw_row, grand_total_col)
        display_name = clean_text(row[name_col - 1])
        key = normalise_text(display_name)

        if not key or key in skip_labels:
            if key or any(v not in (None, "") for v in row):
                skipped += 1
            continue

        tax_numbers = [to_number(row[col - 1]) for col in range(first_tax_col, grand_total_col)]
        row[grand_total_col - 1] = sum(tax_numbers)  # recompute, never trust source total
        output_ws.append(row)
        data_rows += 1

        record = records.get(key)
        if record is None:
            record = TaxpayerRecord(
                metadata={h: row[c - 1] for h, c in metadata_cols.items()},
                tax_values={h: 0.0 for h in tax_headers},
            )
            record.metadata[config.name_header] = display_name
            records[key] = record

        for tax_header, value in zip(tax_headers, tax_numbers):
            record.tax_values[tax_header] += value
        for header, col in metadata_cols.items():  # latest non-blank metadata wins
            value = row[col - 1]
            if value not in (None, ""):
                record.metadata[header] = value

    total_row = FIRST_DATA_ROW + data_rows
    write_total_row(output_ws, total_row, name_col, first_tax_col, grand_total_col, data_rows)

    style_header(output_ws, grand_total_col)
    style_total_row(output_ws, total_row, first_tax_col, grand_total_col)
    set_widths(output_ws, grand_total_col, first_tax_col, metadata_cols.values())
    format_numeric_area(
        output_ws, first_tax_col, grand_total_col, total_row,
        metadata_cols.get(config.id_header),
    )
    if data_rows:
        output_ws.auto_filter.ref = f"A1:{get_column_letter(grand_total_col)}{total_row - 1}"

    if skipped:
        log.debug("  %s: skipped %d blank/total label row(s).", month_name, skipped)

    return MonthData(month_name, records, tax_headers, data_rows, skipped)


# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
def ordered_tax_headers(months: Sequence[MonthData]) -> list[str]:
    """All tax headers across the period, in first-seen order."""
    output: list[str] = []
    seen: set[str] = set()
    for month in months:
        for header in month.tax_headers:
            key = normalise_text(header)
            if key not in seen:
                seen.add(key)
                output.append(header)
    return output


def tax_value(record: TaxpayerRecord | None, wanted_header: str) -> float:
    if record is None:
        return 0.0
    wanted = normalise_text(wanted_header)
    for header, value in record.tax_values.items():
        if normalise_text(header) == wanted:
            return value
    return 0.0


def record_total(record: TaxpayerRecord | None) -> float:
    return sum(record.tax_values.values()) if record else 0.0


def preferred_metadata(key: str, months: Sequence[MonthData], header: str) -> Any:
    """Latest non-blank value, checking the last selected month first."""
    for month in reversed(months):
        record = month.records.get(key)
        if record is None:
            continue
        value = record.metadata.get(header)
        if value not in (None, ""):
            return value
    return ""


def create_summary_sheet(wb: Workbook, months: Sequence[MonthData], config: ReportConfig) -> tuple[int, float]:
    ws = wb.create_sheet(config.summary_sheet_name)

    tax_headers = ordered_tax_headers(months)
    month_total_headers = [f"{m.month_name.upper()} GRAND TOTAL" for m in months]
    summary_headers = [*config.metadata_headers, *tax_headers, *month_total_headers,
                       config.combined_total_header]
    ws.append(summary_headers)

    all_keys: set[str] = set()
    for month in months:
        all_keys.update(month.records.keys())
    sorted_keys = sorted(
        all_keys, key=lambda k: normalise_text(preferred_metadata(k, months, config.name_header))
    )

    grand_total = 0.0
    for key in sorted_keys:
        row: list[Any] = [preferred_metadata(key, months, h) for h in config.metadata_headers]
        for tax_header in tax_headers:
            row.append(sum(tax_value(m.records.get(key), tax_header) for m in months))
        monthly_totals = [record_total(m.records.get(key)) for m in months]
        combined = sum(monthly_totals)
        row.extend(monthly_totals)
        row.append(combined)
        grand_total += combined
        ws.append(row)

    count = len(sorted_keys)
    first_numeric_col = len(config.metadata_headers) + 1
    last_col = len(summary_headers)
    name_col = config.metadata_headers.index(config.name_header) + 1
    total_row = FIRST_DATA_ROW + count
    write_total_row(ws, total_row, name_col, first_numeric_col, last_col, count)

    style_header(ws, last_col)
    style_total_row(ws, total_row, first_numeric_col, last_col)
    set_widths(ws, last_col, first_numeric_col, range(1, len(config.metadata_headers) + 1))
    id_col = (config.metadata_headers.index(config.id_header) + 1
              if config.id_header in config.metadata_headers else None)
    format_numeric_area(ws, first_numeric_col, last_col, total_row, id_col)
    if count:
        ws.auto_filter.ref = f"A1:{get_column_letter(last_col)}{total_row - 1}"

    return count, grand_total


# ---------------------------------------------------------------------------
# File processing
# ---------------------------------------------------------------------------
def find_input_files(config: ReportConfig) -> list[Path]:
    return [
        p for p in sorted(config.input_dir.glob(config.input_pattern))
        if not p.name.startswith("~$") and not p.name.endswith(config.output_suffix)
    ]


def output_path_for(input_path: Path, config: ReportConfig) -> Path:
    return config.output_dir / f"{input_path.stem}{config.output_suffix}"


def _save_atomically(wb: Workbook, target: Path) -> None:
    """Save to a temp file first so a failure never leaves a half-written report."""
    fd, tmp_name = tempfile.mkstemp(suffix=".xlsx", dir=target.parent)
    os.close(fd)
    try:
        wb.save(tmp_name)
        os.replace(tmp_name, target)
    except BaseException:
        Path(tmp_name).unlink(missing_ok=True)
        raise


def process_workbook(input_path: Path, config: ReportConfig) -> WorkbookResult:
    start = time.perf_counter()
    output_path = output_path_for(input_path, config)
    if output_path.exists() and not config.overwrite:
        raise FileExistsError(f"Output already exists (use --overwrite): {output_path}")

    source_wb = load_workbook(input_path, read_only=True, data_only=True)
    try:
        sheet_names = {m: find_sheet_name(source_wb, m) for m in config.months}

        output_wb = Workbook()
        output_wb.remove(output_wb.active)

        months: list[MonthData] = []
        for month in config.months:
            months.append(
                process_month_sheet(month, source_wb[sheet_names[month]],
                                    output_wb.create_sheet(month), config)
            )

        count, total = create_summary_sheet(output_wb, months, config)

        # Formulas are written without cached values; ask Excel to calculate on open.
        output_wb.calculation.fullCalcOnLoad = True
        output_wb.calculation.calcMode = "auto"

        config.output_dir.mkdir(parents=True, exist_ok=True)
        _save_atomically(output_wb, output_path)
    finally:
        source_wb.close()

    return WorkbookResult(
        input_path=input_path,
        output_path=output_path,
        month_rows={m.month_name: m.data_row_count for m in months},
        combined_taxpayers=count,
        combined_total=total,
        seconds=time.perf_counter() - start,
    )
