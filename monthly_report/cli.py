"""Command-line entry point: python -m monthly_report --input-dir ... --output-dir ..."""

from __future__ import annotations

import argparse
import logging
import time
from pathlib import Path

from .config import DEFAULT_MONTHS, ReportConfig
from .processor import ReportError, find_input_files, process_workbook

log = logging.getLogger("monthly_report")


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        prog="python -m monthly_report",
        description="Consolidate selected monthly worksheets into one summary report per workbook.",
    )
    p.add_argument("--input-dir", type=Path, required=True, help="Folder with source workbooks.")
    p.add_argument("--output-dir", type=Path, required=True, help="Folder for generated reports.")
    p.add_argument("--pattern", default=ReportConfig.input_pattern,
                   help="Glob for input files (default: %(default)s).")
    p.add_argument("--months", nargs="+", default=list(DEFAULT_MONTHS),
                   help="Worksheet names to include, in order (default: August..December).")
    p.add_argument("--first-tax-header", default=ReportConfig.first_tax_header,
                   help="Header label of the first tax-type column (default: %(default)s).")
    p.add_argument("--overwrite", action="store_true", help="Replace existing output files.")
    p.add_argument("--log-file", type=Path, help="Also write the log to this file.")
    p.add_argument("-v", "--verbose", action="store_true", help="Debug logging.")
    return p.parse_args(argv)


def setup_logging(verbose: bool, log_file: Path | None) -> None:
    handlers: list[logging.Handler] = [logging.StreamHandler()]
    if log_file:
        log_file.parent.mkdir(parents=True, exist_ok=True)
        handlers.append(logging.FileHandler(log_file, encoding="utf-8"))
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s %(levelname)-7s %(message)s",
        datefmt="%H:%M:%S",
        handlers=handlers,
        force=True,
    )


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    setup_logging(args.verbose, args.log_file)

    config = ReportConfig(
        input_dir=args.input_dir,
        output_dir=args.output_dir,
        input_pattern=args.pattern,
        months=tuple(args.months),
        first_tax_header=args.first_tax_header,
        overwrite=args.overwrite,
    )
    try:
        config.validate()
    except (ValueError, FileNotFoundError) as exc:
        log.error("%s", exc)
        return 1

    files = find_input_files(config)
    if not files:
        log.error("No workbooks matching '%s' in %s", config.input_pattern, config.input_dir)
        return 1

    log.info("Found %d workbook(s). Months: %s", len(files), ", ".join(config.months))
    log.info("Tax types start at header '%s' and end before '%s'.",
             config.first_tax_header, config.grand_total_header)

    start = time.perf_counter()
    ok = failed = rows_total = 0
    for path in files:
        log.info("Processing %s", path.name)
        try:
            result = process_workbook(path, config)
        except PermissionError:
            failed += 1
            log.error("  Cannot access %s (is it open in Excel?)", path.name)
            continue
        except (ReportError, FileExistsError, OSError) as exc:
            failed += 1
            log.error("  FAILED: %s", exc)
            continue
        except Exception:  # unexpected: keep going, but show the traceback
            failed += 1
            log.exception("  FAILED with an unexpected error")
            continue

        ok += 1
        rows = sum(result.month_rows.values())
        rows_total += rows
        for month, n in result.month_rows.items():
            log.info("  %-10s rows: %s", month, f"{n:,}")
        log.info("  Combined taxpayers: %s | Period total: %s | %.2fs",
                 f"{result.combined_taxpayers:,}", f"{result.combined_total:,.2f}", result.seconds)
        log.info("  Saved: %s", result.output_path)

    elapsed = time.perf_counter() - start
    log.info("Finished: %d succeeded, %d failed, %s monthly rows in %.2fs.",
             ok, failed, f"{rows_total:,}", elapsed)
    return 0 if failed == 0 else 2
