"""Generate fully fictional monthly collection workbooks for demos and benchmarks.

All names, addresses, IDs and amounts are random. The data deliberately contains
the irregularities the automation has to handle:

* column positions that shift between months (an extra REMARKS column),
* a tax type that only appears in later months (CGT from November),
* names with inconsistent spacing/case across months,
* numbers stored as text ("1,234.50", "(250.00)"), blanks,
* a stale GRAND TOTAL value on some rows,
* TOTAL / GRAND TOTAL label rows at the bottom of each sheet.

Usage:
    python scripts/generate_sample_data.py --output-dir sample_data/input
"""

from __future__ import annotations

import argparse
import random
from pathlib import Path

from openpyxl import Workbook

MONTHS = ["January", "February", "March", "April", "May", "June", "July",
          "August", "September", "October", "November", "December"]
TAX_TYPES = ["IT", "VAT", "WT", "PT", "DST", "ET"]
LATE_TAX_TYPE = "CGT"  # appears from November onwards
DISTRICTS = ["NORTH", "SOUTH", "EAST"]

ADJECTIVES = ["Bluefin", "Golden", "Silver", "Evergreen", "Summit", "Harbor", "Crescent",
              "Maple", "Coral", "Granite", "Sunrise", "Pioneer", "Riverside", "Orchid", "Atlas"]
NOUNS = ["Trading", "Foods", "Logistics", "Hardware", "Textiles", "Pharmacy", "Builders",
         "Motors", "Bakery", "Electronics", "Farms", "Printing", "Supplies", "Ventures"]
SUFFIXES = ["Co.", "Inc.", "Corp.", "Enterprises", "Store", "Ltd."]
STREETS = ["Example St.", "Sample Ave.", "Demo Road", "Test Blvd.", "Placeholder Lane"]
INDUSTRIES = ["RETAIL TRADE", "WHOLESALE TRADE", "MANUFACTURING", "FOOD SERVICE",
              "TRANSPORT", "CONSTRUCTION", "PROFESSIONAL SERVICES"]
FORMS = ["FORM-A", "FORM-B", "FORM-C"]


def make_taxpayers(rng: random.Random, district: str, count: int) -> list[dict]:
    seen: set[str] = set()
    taxpayers = []
    while len(taxpayers) < count:
        name = f"{rng.choice(ADJECTIVES)} {rng.choice(NOUNS)} {rng.choice(SUFFIXES)}"
        if len(taxpayers) >= len(ADJECTIVES) * 3:  # make names unique at scale
            name = f"{name[:-1] if name.endswith('.') else name} {len(taxpayers):04d}"
        if name.upper() in seen:
            continue
        seen.add(name.upper())
        taxpayers.append({
            "district": f"DISTRICT {district}",
            "name": name,
            "address": f"{rng.randint(1, 999)} {rng.choice(STREETS)}, Sample City",
            "id": f"{rng.randint(100, 999)}-{rng.randint(100, 999)}-{rng.randint(100, 999)}-000",
            "industry": rng.choice(INDUSTRIES),
            "form": rng.choice(FORMS),
            "scale": rng.choice([1_000, 5_000, 20_000, 100_000]),
        })
    return taxpayers


def messy_name(rng: random.Random, name: str) -> str:
    roll = rng.random()
    if roll < 0.08:
        return "  " + name.replace(" ", "  ") + " "
    if roll < 0.14:
        return name.upper()
    return name


def amount_cell(rng: random.Random, value: float):
    roll = rng.random()
    if roll < 0.05:
        return f"{value:,.2f}"  # number stored as text
    if roll < 0.07:
        return f"({abs(value):,.2f})" if value else None
    return value


def build_month(ws, rng: random.Random, month_index: int, taxpayers: list[dict]) -> None:
    month = MONTHS[month_index]
    tax_types = TAX_TYPES + ([LATE_TAX_TYPE] if month in ("November", "December") else [])
    with_remarks = month in ("September", "December")  # shifts tax columns right

    headers = ["DISTRICT", "TAXPAYER_NAME", "ADDRESS", "TAXPAYER_ID", "INDUSTRY_CODE",
               "FORM_TYPE"] + (["REMARKS"] if with_remarks else []) + tax_types + ["GRAND TOTAL"]
    ws.append(headers)

    sheet_total = 0.0
    for tp in taxpayers:
        if rng.random() < 0.12:  # not every taxpayer pays every month
            continue
        values = []
        for _ in tax_types:
            if rng.random() < 0.45:
                values.append(0.0 if rng.random() < 0.3 else None)
            else:
                values.append(round(rng.uniform(0.05, 1.0) * tp["scale"], 2))
        numeric = sum(v or 0.0 for v in values)
        grand_total = numeric if rng.random() > 0.05 else round(numeric * 1.1, 2)  # stale total
        sheet_total += numeric
        row = [tp["district"], messy_name(rng, tp["name"]), tp["address"], tp["id"],
               tp["industry"], tp["form"]]
        if with_remarks:
            row.append(rng.choice(["", "", "late filing", "amended"]))
        row += [amount_cell(rng, v) if v is not None else None for v in values]
        row.append(grand_total)
        ws.append(row)

    ws.append([None] * len(headers))
    for label in ("TOTAL", "GRAND TOTAL"):
        row = [None] * len(headers)
        row[1] = label
        row[-1] = round(sheet_total, 2)
        ws.append(row)


def generate(output_dir: Path, taxpayers_per_district: int, districts: list[str], seed: int) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    rng = random.Random(seed)
    paths = []
    for district in districts:
        taxpayers = make_taxpayers(rng, district, taxpayers_per_district)
        wb = Workbook()
        wb.remove(wb.active)
        for i, month in enumerate(MONTHS):
            build_month(wb.create_sheet(month), rng, i, taxpayers)
        path = output_dir / f"DISTRICT_{district}_MONTHLY_COLLECTIONS_2025.xlsx"
        wb.save(path)
        paths.append(path)
    return paths


def main() -> None:
    p = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    p.add_argument("--output-dir", type=Path, default=Path("sample_data/input"))
    p.add_argument("--taxpayers", type=int, default=250, help="Taxpayers per district.")
    p.add_argument("--districts", nargs="+", default=DISTRICTS)
    p.add_argument("--seed", type=int, default=2025)
    args = p.parse_args()
    for path in generate(args.output_dir, args.taxpayers, args.districts, args.seed):
        print(f"Created {path}")


if __name__ == "__main__":
    main()
