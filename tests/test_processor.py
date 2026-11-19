"""Tests for the monthly report consolidation (all data is synthetic)."""

from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from monthly_report.cli import main  # noqa: E402
from monthly_report.config import ReportConfig  # noqa: E402
from monthly_report.processor import (  # noqa: E402
    ZERO_AS_BLANK_NUMBER_FORMAT,
    ReportError,
    normalise_text,
    process_workbook,
    to_number,
)

META = ["DISTRICT", "TAXPAYER_NAME", "ADDRESS", "TAXPAYER_ID", "INDUSTRY_CODE", "FORM_TYPE"]
MONTHS = ["August", "September", "October", "November", "December"]


def meta(name: str, tp_id: str = "111-222-333-000", address: str = "1 Example St.") -> list:
    return ["DISTRICT TEST", name, address, tp_id, "RETAIL TRADE", "FORM-A"]


def make_workbook(path: Path, sheets: dict[str, list[list]]) -> Path:
    wb = Workbook()
    wb.remove(wb.active)
    for title, rows in sheets.items():
        ws = wb.create_sheet(title)
        for row in rows:
            ws.append(row)
    wb.save(path)
    return path


def standard_month(rows: list[list], tax=("IT", "VAT")) -> list[list]:
    return [META + list(tax) + ["GRAND TOTAL"], *rows]


@pytest.fixture
def cfg(tmp_path: Path) -> ReportConfig:
    (tmp_path / "in").mkdir()
    return ReportConfig(input_dir=tmp_path / "in", output_dir=tmp_path / "out")


def sheet_rows(ws) -> list[tuple]:
    return list(ws.iter_rows(values_only=True))


# ---------------------------------------------------------------------------
def test_to_number_handles_text_blanks_and_parentheses():
    assert to_number(None) == 0.0
    assert to_number("") == 0.0
    assert to_number("1,234.50") == 1234.5
    assert to_number("(250.00)") == -250.0
    assert to_number("n/a") == 0.0
    assert to_number(7) == 7.0


def test_consolidates_by_normalised_name_and_sums_months(cfg):
    sheets = {m: standard_month([meta("Acme Trading", ), ]) for m in MONTHS}
    sheets["August"] = standard_month([meta("Acme Trading") + [100, 50, 999]])  # stale total
    sheets["September"] = standard_month([meta("  acme   TRADING ") + ["1,000.00", None, 0]])
    sheets["October"] = standard_month([meta("Beta Foods") + [10, 20, 30]])
    sheets["November"] = standard_month([])
    sheets["December"] = standard_month([meta("Acme Trading") + [1, 2, 3],
                                         meta("Acme Trading") + [4, 5, 9]])  # duplicate in month
    src = make_workbook(cfg.input_dir / "X_MONTHLY_COLLECTIONS_2025.xlsx", sheets)

    result = process_workbook(src, cfg)

    assert result.combined_taxpayers == 2
    assert result.combined_total == pytest.approx(150 + 1000 + 30 + 3 + 9)
    assert result.month_rows == {"August": 1, "September": 1, "October": 1,
                                 "November": 0, "December": 2}

    wb = load_workbook(result.output_path)
    assert wb.sheetnames == MONTHS + ["AUG_TO_DEC_SUMMARY"]

    # Row GRAND TOTAL is recomputed from tax columns, not copied from the source.
    aug = sheet_rows(wb["August"])
    assert aug[1][-1] == 150

    summary = sheet_rows(wb["AUG_TO_DEC_SUMMARY"])
    header = summary[0]
    assert header[:6] == tuple(META)
    assert header[6:] == ("IT", "VAT", "AUGUST GRAND TOTAL", "SEPTEMBER GRAND TOTAL",
                          "OCTOBER GRAND TOTAL", "NOVEMBER GRAND TOTAL",
                          "DECEMBER GRAND TOTAL", "AUGUST TO DECEMBER GRAND TOTAL")
    acme = summary[1]
    assert acme[1] == "Acme Trading"  # latest non-blank display name (December)
    assert acme[6:] == (1105, 57, 150, 1000, 0, 0, 12, 1162)
    beta = summary[2]
    assert beta[1] == "Beta Foods"
    assert beta[-1] == 30
    total = summary[3]
    assert total[1] == "TOTAL"
    assert total[6] == "=SUM(G2:G3)"
    assert total[-1] == "=SUM(N2:N3)"


def test_tax_columns_located_by_header_not_position(cfg):
    sheets = {m: standard_month([meta("Acme") + [1, 2, 3]]) for m in MONTHS}
    # September has an extra column before the tax types and an extra tax type.
    sheets["September"] = [META + ["REMARKS", "IT", "VAT", "CGT", "GRAND TOTAL", "IGNORED"],
                           meta("Acme") + ["note", 10, 20, 5, 0, 12345]]
    src = make_workbook(cfg.input_dir / "X_MONTHLY_COLLECTIONS_2025.xlsx", sheets)

    result = process_workbook(src, cfg)
    wb = load_workbook(result.output_path)

    sep = sheet_rows(wb["September"])
    assert sep[0][-1] == "GRAND TOTAL"  # columns after GRAND TOTAL are dropped
    assert sep[1][-1] == 35
    summary = sheet_rows(wb["AUG_TO_DEC_SUMMARY"])
    assert summary[0][6:9] == ("IT", "VAT", "CGT")
    assert summary[1][6:9] == (14, 28, 5)
    assert summary[1][-1] == 47


def test_total_and_grand_total_label_rows_are_not_counted(cfg):
    rows = [meta("Acme") + [100, 0, 100],
            [None] * 9,
            [None, "TOTAL", None, None, None, None, None, None, 100],
            [None, "GRAND TOTAL", None, None, None, None, 100, 0, 100]]
    sheets = {m: standard_month(rows) for m in MONTHS}
    src = make_workbook(cfg.input_dir / "X_MONTHLY_COLLECTIONS_2025.xlsx", sheets)

    result = process_workbook(src, cfg)

    assert result.combined_taxpayers == 1
    assert result.combined_total == 500


def test_missing_month_or_header_raises_clear_error(cfg):
    sheets = {m: standard_month([]) for m in MONTHS[:-1]}
    src = make_workbook(cfg.input_dir / "A_MONTHLY_COLLECTIONS_2025.xlsx", sheets)
    with pytest.raises(ReportError, match="Missing worksheet 'December'"):
        process_workbook(src, cfg)

    sheets = {m: [META + ["VAT", "GRAND TOTAL"]] for m in MONTHS}
    src = make_workbook(cfg.input_dir / "B_MONTHLY_COLLECTIONS_2025.xlsx", sheets)
    with pytest.raises(ReportError, match="missing header"):
        process_workbook(src, cfg)

    sheets = {m: [META + ["GRAND TOTAL", "IT"]] for m in MONTHS}
    src = make_workbook(cfg.input_dir / "C_MONTHLY_COLLECTIONS_2025.xlsx", sheets)
    with pytest.raises(ReportError, match="must appear before"):
        process_workbook(src, cfg)


def test_existing_output_is_not_overwritten_without_flag(cfg):
    sheets = {m: standard_month([meta("Acme") + [1, 1, 2]]) for m in MONTHS}
    src = make_workbook(cfg.input_dir / "X_MONTHLY_COLLECTIONS_2025.xlsx", sheets)
    first = process_workbook(src, cfg)
    with pytest.raises(FileExistsError):
        process_workbook(src, cfg)
    again = process_workbook(src, ReportConfig(cfg.input_dir, cfg.output_dir, overwrite=True))
    assert again.output_path == first.output_path


def test_formatting_zero_as_blank_and_id_as_text(cfg):
    sheets = {m: standard_month([meta("Acme") + [0, 5, 5]]) for m in MONTHS}
    src = make_workbook(cfg.input_dir / "X_MONTHLY_COLLECTIONS_2025.xlsx", sheets)
    ws = load_workbook(process_workbook(src, cfg).output_path)["August"]
    assert ws["G2"].number_format == ZERO_AS_BLANK_NUMBER_FORMAT
    assert ws["D2"].number_format == "@"
    assert ws["B3"].font.bold and ws["B3"].value == "TOTAL"
    assert ws.freeze_panes == "A2"
    assert ws.auto_filter.ref == "A1:I2"


def test_cli_end_to_end_with_generated_sample_data(tmp_path):
    sys.path.insert(0, str(ROOT / "scripts"))
    import generate_sample_data as gen

    paths = gen.generate(tmp_path / "in", taxpayers_per_district=40, districts=["T1", "T2"], seed=7)
    rc = main(["--input-dir", str(tmp_path / "in"), "--output-dir", str(tmp_path / "out")])
    assert rc == 0

    # Independent reference: recompute the period total straight from the inputs.
    for path in paths:
        src = load_workbook(path, read_only=True)
        expected: dict[str, float] = {}
        for month in MONTHS:
            rows = list(src[month].iter_rows(values_only=True))
            hdr = [normalise_text(h) for h in rows[0]]
            first, last, name = hdr.index("IT"), hdr.index("GRAND TOTAL"), hdr.index("TAXPAYER_NAME")
            for r in rows[1:]:
                key = normalise_text(r[name])
                if key and key not in ("TOTAL", "GRAND TOTAL"):
                    expected[key] = expected.get(key, 0.0) + sum(to_number(v) for v in r[first:last])
        out = load_workbook(tmp_path / "out" / f"{path.stem}_AUG_TO_DEC_PROCESSED.xlsx")
        summary = sheet_rows(out["AUG_TO_DEC_SUMMARY"])
        got = {normalise_text(r[1]): r[-1] for r in summary[1:-1]}
        assert got.keys() == expected.keys()
        for key, value in expected.items():
            assert got[key] == pytest.approx(value)


def test_cli_returns_error_codes(tmp_path):
    assert main(["--input-dir", str(tmp_path / "missing"), "--output-dir", str(tmp_path)]) == 1
    (tmp_path / "in").mkdir()
    assert main(["--input-dir", str(tmp_path / "in"), "--output-dir", str(tmp_path / "o")]) == 1
    make_workbook(tmp_path / "in" / "BAD_MONTHLY_COLLECTIONS_2025.xlsx", {"August": [["x"]]})
    assert main(["--input-dir", str(tmp_path / "in"), "--output-dir", str(tmp_path / "o")]) == 2


@pytest.mark.skipif(shutil.which("soffice") is None, reason="LibreOffice not installed")
def test_total_formulas_evaluate_to_python_totals(cfg, tmp_path):
    """Recalculate the output in LibreOffice and compare formula results."""
    sheets = {m: standard_month([meta("Acme") + [100.5, 50, 0], meta("Beta") + [1, "2", 0]])
              for m in MONTHS}
    result = process_workbook(make_workbook(cfg.input_dir / "X_MONTHLY_COLLECTIONS_2025.xlsx",
                                            sheets), cfg)
    conv = tmp_path / "recalc"
    subprocess.run(["soffice", "--headless", "--convert-to", "xlsx", "--outdir", str(conv),
                    str(result.output_path)], check=True, capture_output=True, timeout=120)
    ws = load_workbook(conv / result.output_path.name, data_only=True)["AUG_TO_DEC_SUMMARY"]
    last = list(ws.iter_rows(values_only=True))[-1]
    assert last[1] == "TOTAL"
    assert last[-1] == pytest.approx(result.combined_total) == pytest.approx(5 * 153.5)
