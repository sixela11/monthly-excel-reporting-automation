"""Measure processing time on synthetic workbooks of increasing size.

Usage:
    python scripts/benchmark.py            # prints a Markdown table
    python scripts/benchmark.py --save docs/BENCHMARK.md
"""

from __future__ import annotations

import argparse
import platform
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "scripts"))

from generate_sample_data import generate  # noqa: E402
from monthly_report.config import ReportConfig  # noqa: E402
from monthly_report.processor import find_input_files, process_workbook  # noqa: E402

SCENARIOS = [(3, 250), (3, 2_000), (10, 1_000)]  # (workbooks, taxpayers per workbook)


def run(workbooks: int, taxpayers: int) -> tuple[int, float]:
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        generate(tmp / "in", taxpayers, [f"D{i:02d}" for i in range(workbooks)], seed=1)
        cfg = ReportConfig(input_dir=tmp / "in", output_dir=tmp / "out")
        start = time.perf_counter()
        rows = sum(sum(process_workbook(f, cfg).month_rows.values()) for f in find_input_files(cfg))
        return rows, time.perf_counter() - start


def main() -> None:
    p = argparse.ArgumentParser()
    p.add_argument("--save", type=Path)
    args = p.parse_args()

    lines = [
        "# Benchmark (synthetic data)",
        "",
        f"Environment: Python {platform.python_version()} on {platform.system()} "
        f"({platform.machine()}). Times exclude data generation. Each workbook has 12 month "
        "sheets; 5 are processed.",
        "",
        "| Workbooks | Taxpayers / workbook | Monthly rows processed | Total time (s) | Rows / s |",
        "|---:|---:|---:|---:|---:|",
    ]
    for workbooks, taxpayers in SCENARIOS:
        rows, seconds = run(workbooks, taxpayers)
        lines.append(f"| {workbooks} | {taxpayers:,} | {rows:,} | {seconds:.2f} | {rows / seconds:,.0f} |")
        print(lines[-1], flush=True)
    text = "\n".join(lines) + "\n"
    if args.save:
        args.save.write_text(text, encoding="utf-8")
        print(f"Saved {args.save}")


if __name__ == "__main__":
    main()
