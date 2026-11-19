"""Configuration for the monthly report consolidation."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

DEFAULT_MONTHS = ("August", "September", "October", "November", "December")

DEFAULT_METADATA_HEADERS = (
    "DISTRICT",
    "TAXPAYER_NAME",
    "ADDRESS",
    "TAXPAYER_ID",
    "INDUSTRY_CODE",
    "FORM_TYPE",
)


@dataclass(frozen=True)
class ReportConfig:
    input_dir: Path
    output_dir: Path
    input_pattern: str = "*_MONTHLY_COLLECTIONS_2025.xlsx"
    months: tuple[str, ...] = DEFAULT_MONTHS
    metadata_headers: tuple[str, ...] = DEFAULT_METADATA_HEADERS
    name_header: str = "TAXPAYER_NAME"
    id_header: str = "TAXPAYER_ID"
    # Header LABEL (not a column letter) of the first tax-type column.
    first_tax_header: str = "IT"
    grand_total_header: str = "GRAND TOTAL"
    # Rows whose taxpayer name equals one of these labels are treated as
    # existing total rows in the source and are not counted as taxpayers.
    skip_name_labels: frozenset[str] = field(
        default_factory=lambda: frozenset({"TOTAL", "GRAND TOTAL"})
    )
    overwrite: bool = False

    @property
    def period_label(self) -> str:
        """e.g. 'AUG_TO_DEC' for the default months."""
        return f"{self.months[0][:3]}_TO_{self.months[-1][:3]}".upper()

    @property
    def summary_sheet_name(self) -> str:
        return f"{self.period_label}_SUMMARY"

    @property
    def combined_total_header(self) -> str:
        return f"{self.months[0]} TO {self.months[-1]} GRAND TOTAL".upper()

    @property
    def output_suffix(self) -> str:
        return f"_{self.period_label}_PROCESSED.xlsx"

    def validate(self) -> None:
        if not self.months:
            raise ValueError("At least one month must be selected.")
        if self.name_header not in self.metadata_headers:
            raise ValueError(f"name_header '{self.name_header}' must be a metadata header.")
        if not self.input_dir.is_dir():
            raise FileNotFoundError(f"Input folder not found: {self.input_dir}")
