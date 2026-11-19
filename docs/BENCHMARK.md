# Benchmark (synthetic data)

Environment: Python 3.13.18 on Linux (x86_64). Times exclude data generation. Each workbook has 12 month sheets; 5 are processed.
These are processing times on fictional data, not measurements of time saved in real work.

| Workbooks | Taxpayers / workbook | Monthly rows processed | Total time (s) | Rows / s |
|---:|---:|---:|---:|---:|
| 3 | 250 | 3,264 | 1.19 | 2,735 |
| 3 | 2,000 | 26,351 | 9.30 | 2,832 |
| 10 | 1,000 | 43,944 | 14.96 | 2,938 |
