# Monthly Excel Reporting Automation

A Python command-line tool that turns a folder of monthly collection workbooks into
consolidated period reports: one output workbook per input file, with cleaned monthly
sheets, a per-taxpayer summary, and verified totals.

> **About this repository.** I built this from a script I use each month in my own work.
> This version has been rewritten for a portfolio: all data is synthetic, and the column
> names, file names and paths are neutral and configurable. It contains no real records
> and no organisation-specific code.

## The manual process it replaces

Every month, each district workbook (12 month sheets) has to be turned into an
August–December report. By hand, that means doing this for every workbook:

1. Copy the five month sheets into a new workbook.
2. Find the tax-type columns in each sheet. Their positions change between months
   because columns get added (for example a `REMARKS` column, or a new tax type).
3. Recalculate each row's GRAND TOTAL, since source totals can be stale.
4. Match the same taxpayer across five sheets even when the name is spaced or capitalised
   differently, and add up every tax type.
5. Build a summary with one grand total per month and one for the whole period.
6. Add TOTAL rows, number formats, filters and frozen headers.

Doing this by hand is slow and easy to get wrong: a missed row, a shifted column or a
mismatched name changes the totals. The script does all six steps the same way every
time. I still start it by hand each month; it does not run on a schedule.

## What it does

| Step | Behaviour |
|---|---|
| Discover | Processes every workbook in `--input-dir` that matches `--pattern`, skipping Excel lock files and previous outputs. |
| Validate | Checks that every selected month sheet exists (case- and space-insensitive) and that all required headers are present, and that the first tax header comes before `GRAND TOTAL`. If a workbook fails, the error is logged and the run continues with the next one. |
| Locate columns | Finds tax types **by header label** (default: from `IT` to the column before `GRAND TOTAL`), not by fixed column letters. |
| Clean | Reads numbers stored as text (`"1,234.50"`, `"(250.00)"`), treats blanks as 0, and skips blank rows and existing `TOTAL` / `GRAND TOTAL` label rows. |
| Recompute | Recalculates each row's GRAND TOTAL from its tax-type columns. |
| Match | Groups records by normalised taxpayer name (trimmed, single-spaced, upper-case). The latest non-blank metadata wins. |
| Summarise | Writes a `AUG_TO_DEC_SUMMARY` sheet with every tax type seen in the period (first-seen order), five monthly grand totals, and the combined period total. |
| Format | Adds `=SUM()` TOTAL rows, shows zeros as blanks, stores IDs as text, and adds frozen headers, autofilter and column widths. |
| Save safely | Writes to a temporary file and then renames it, so a failed run never leaves a half-written report. It will not overwrite an existing report unless you pass `--overwrite`. |
| Report | Logs row counts, taxpayers, period totals and timings to the console, and optionally to a file. Exit code `0` means success, `1` means bad setup, `2` means at least one workbook failed. |

## Quick start

```bash
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt

python scripts/generate_sample_data.py --output-dir sample_data/input
python -m monthly_report --input-dir sample_data/input --output-dir output
```

Options:

```text
--pattern GLOB            input file pattern (default *_MONTHLY_COLLECTIONS_2025.xlsx)
--months M [M ...]        sheets to include, in order (default August..December)
--first-tax-header LABEL  header of the first tax-type column (default IT)
--overwrite               replace existing reports
--log-file PATH           also write the log to a file
-v / --verbose            debug logging
```

## Repository layout

```text
monthly_report/
  config.py        settings (headers, months, labels) in one dataclass
  processor.py     reading, matching, summarising, formatting, saving
  cli.py           command-line interface, logging, timing, exit codes
scripts/
  generate_sample_data.py   reproducible fictional workbooks (seeded)
  benchmark.py              timing on synthetic data of increasing size
sample_data/input/          3 synthetic district workbooks (12 months each)
examples/output/            reports generated from them, plus run.log
tests/test_processor.py     automated tests
docs/BENCHMARK.md           measured run times
docs/DEMO_SCRIPT.md         2–3 minute walkthrough
```

## Testing

```bash
pip install -r requirements-dev.txt
python -m pytest -v
```

The tests cover:
- number parsing,
- matching across months by normalised name (including duplicates within one month),
- stale source totals,
- columns that shift between months, and tax types that appear part-way through the period,
- skipping total label rows,
- clear errors for missing sheets or headers,
- refusing to overwrite existing reports,
- formatting,
- an end-to-end CLI run checked against an independent recalculation of the inputs,
- exit codes,
- a LibreOffice recalculation check that the `=SUM()` TOTAL rows give the same figures as Python. This test is skipped if LibreOffice is not installed.

The rewrite was also checked against my original script on the same synthetic data: every
cell value and number format in the generated reports matched (59,970 cells, 0 differences).
One difference is deliberate: `GRAND TOTAL` label rows are now skipped as well as `TOTAL` rows.

## Measured results (synthetic data)

Measured with `scripts/benchmark.py`; full table in [`docs/BENCHMARK.md`](docs/BENCHMARK.md):

| Workbooks | Monthly rows processed | Time |
|---:|---:|---:|
| 3 | 3,264 | 1.2 s |
| 3 | 26,351 | 9.3 s |
| 10 | 43,944 | 15.0 s |

Run times depend on the machine. These figures were measured in a Linux container (4 vCPU)
on synthetic data only. They are processing times, not measurements of time saved in real
work. Re-run the benchmark on your own hardware.

## Known limitations

- **Name-based matching.** Taxpayers are matched by normalised name, the same as the original
  workflow. Matching ignores capitalisation and extra spaces only; it is not fuzzy matching. Spelling variants (`Acme Trading Co.` vs `Acme Trading Company`) are treated as
  different taxpayers, and two different taxpayers with the same name are merged. Matching on
  an ID column would be more reliable if the source data has one for every row.
- Headers must be in row 1 and every selected month must exist in every workbook.
- Columns to the right of `GRAND TOTAL` are not carried into the output.
- If the same header appears twice in a sheet, the last one is used, and a warning is logged.
- The TOTAL rows are Excel formulas. Excel calculates them when the file is opened;
  tools that read the file without calculating it will see the formulas, not the numbers.
- The whole workbook is held in memory. Very large files (hundreds of thousands of rows)
  would need a streaming write mode.
