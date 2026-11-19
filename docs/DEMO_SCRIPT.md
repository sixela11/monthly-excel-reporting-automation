# Demo video script (about 2 minutes)

**Goal:** show the manual problem, the automation running, the generated reports, and measured
results. Use synthetic data only.

## Before recording

```bash
cd monthly-excel-reporting-automation
python -m venv .venv
.venv\Scripts\activate            # macOS/Linux: source .venv/bin/activate
pip install -r requirements-dev.txt
```

- Delete any old `output/` folder.
- Open `sample_data/input/DISTRICT_NORTH_MONTHLY_COLLECTIONS_2025.xlsx` in Excel and click the **August** tab (the file opens on January).
- Open a terminal in the project folder, with a large font.

Speaking time is about 270 words.

---

## 1. The problem (0:00–0:30) — input workbook

**Show:**
- The 12 month tabs.
- **August, B2 and M2:** the name `  Golden  Builders  Enterprises ` has extra spaces. Its GRAND TOTAL in M2 is **439,340.48**, but its tax columns G2:L2 add up to **399,400.44**.
- **September:** `IT` has moved to column H because of an extra `REMARKS` column in G.
- **November:** a new tax type, `CGT`, appears in column M.

> "Every month I turn workbooks like this into an August-to-December report. By hand, that
> means copying five sheets, finding tax columns that move between months, fixing totals that
> don't add up, and matching each taxpayer across the months, even when a name has extra
> spaces or different capitalisation."

## 2. Run it (0:30–0:55) — terminal

```bash
python -m monthly_report --input-dir sample_data/input --output-dir output --log-file output/run.log
```

**Point at:**
- the row counts per month;
- `Combined taxpayers: 250`;
- the final `Finished: 3 succeeded, 0 failed` line.

> "One command processes every workbook in the folder. It checks the sheets and headers,
> finds columns by header name, recalculates each row's total, and logs counts and timing."

## 3. The report (0:55–1:30) — Excel

**Open:** `output/DISTRICT_NORTH_MONTHLY_COLLECTIONS_2025_AUG_TO_DEC_PROCESSED.xlsx`

**Show:**
- **August, M2:** GRAND TOTAL is now **399,400.44**.
- **AUG_TO_DEC_SUMMARY, row 92:** Golden Builders Enterprises is a single row: August **399,400.44** (N92), September **111,605.35** (O92), November **185,898.77** (Q92), period total **696,904.56** (S92). October and December are blank because zeros are hidden.
- The `CGT` column (M).
- The TOTAL row (row 252): click S252 to show `=SUM(S2:S251)`.

> "Each report contains the five cleaned months and a summary. Names are matched regardless
> of capitalisation and extra spaces; spelling differences are not corrected. The stale total
> has been recalculated, and the TOTAL row uses live Excel formulas, so anyone can check it."

## 4. Verification and results (1:30–2:00) — terminal

```bash
python -m pytest -q
```

Then show `docs/BENCHMARK.md`.

> "Ten automated tests check the calculations, including against an independent
> recalculation. On fictional data in a test environment, ten workbooks with about 44,000 rows
> processed in about 15 seconds. That's processing time, not a measurement of time saved at
> work. I start it manually each month. The code is on GitHub."

---

### Notes
- Only quote numbers that appear on screen. Re-run `python scripts/benchmark.py` on your
  recording machine and use those times.
- Do not show real workbooks, real folders or the original internal script.
