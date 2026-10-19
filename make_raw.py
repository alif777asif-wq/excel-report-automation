"""
Generate twelve messy monthly sales exports (plus an optional next month) with KNOWN problems,
and the ground truth they were made from, so the cleaning can be checked rather than trusted.

Everything is synthetic: fictional reps and products, no real company data.
"""
import csv
import json
import os
import random
from datetime import date, datetime
from decimal import Decimal

import openpyxl

import clean

REGION_LIST = ["North", "South", "East", "West", "Central"]
REGION_WEIGHTS = [0.27, 0.23, 0.2, 0.18, 0.12]
REPS = ["Avery Lindqvist", "Noor Haddad", "Mateo Ruiz", "Priya Raman", "Jonas Weber", "Imani Okafor",
        "Elena Petrova", "Kenji Mori", "Sofia Duarte", "Tariq Aziz"]
PRODUCTS = [("Standing Desk", "Furniture", 32900), ("Desk Lamp", "Lighting", 3990), ("Monitor Arm", "Accessories", 8900),
            ("Ergonomic Chair", "Furniture", 24900), ("Wireless Keyboard", "Peripherals", 5900),
            ("Webcam HD", "Peripherals", 7450), ("Laptop Stand", "Accessories", 4500), ("Task Light", "Lighting", 5400),
            ("Filing Cabinet", "Furniture", 13900), ("USB-C Hub", "Peripherals", 6400),
            ("Cable Tray", "Accessories", 1950), ("Bookshelf", "Furniture", 11900)]
HEADERS = [
    ["Order ID", "Order Date", "Region", "Sales Rep", "Product", "Category", "Units", "Unit Price", "Discount", "Amount"],
    ["order_id", "date", "region", "rep", "product", "category", "qty", "price", "disc_pct", "net_amount"],
    ["OrderID", "OrderDate", "Area", "Rep", "Item", "Cat", "Units", "UnitPrice", "Disc", "Total"],
]
# one entry per monthly export: file name, format, delimiter, header style, date style, money style, region style, footer
FILE_PLAN = [
    ("sales_2025-01.csv", "csv", ",", 0, "iso", "num", "proper", False),
    ("Feb_2025_branch_export.csv", "csv", ",", 1, "dmy", "num", "lower", True),
    ("branch_export_mar.xlsx", "xlsx", None, 0, "native", "num", "proper", False),
    ("APR-sales-report.csv", "csv", ",", 2, "mdy_text", "usd", "upper", True),
    ("sales_may_2025.csv", "csv", ";", 1, "dmy", "num", "abbr", False),
    ("June sales (final).xlsx", "xlsx", None, 2, "mdy_text", "usd", "padded", False),
    ("export_2025_07.csv", "csv", ",", 0, "dmy", "num", "proper", True),
    ("Aug-2025-orders.csv", "csv", ",", 1, "iso", "usd", "lower", False),
    ("september_branch.xlsx", "xlsx", None, 0, "native", "num", "abbr", True),
    ("OCT_export.csv", "csv", ",", 2, "dmy", "usd", "proper", False),
    ("sales_nov_2025.csv", "csv", ",", 0, "mdy_text", "num", "upper", True),
    ("dec2025_orders.xlsx", "xlsx", None, 1, "native", "usd", "padded", False),
]
NEXT_MONTH = ("sales_2026-01.csv", "csv", ",", 0, "dmy", "usd", "lower", True)


def fmt_region(region, style):
    return {"proper": region, "lower": region.lower(), "upper": region.upper(), "abbr": region[0],
            "padded": f" {region} "}[style]


def fmt_date(d, style):
    return {"iso": d.isoformat(), "dmy": d.strftime("%d/%m/%Y"), "mdy_text": d.strftime("%b %d, %Y"),
            "native": datetime(d.year, d.month, d.day)}[style]


def fmt_money(cents, style):
    return f"${cents / 100:,.2f}" if style == "usd" else round(cents / 100, 2)


def make_month(rng, year, month, plan, order_base):
    fname, fmt, delim, hstyle, dstyle, mstyle, rstyle, footer = plan
    n = rng.randint(820, 1080)
    days = [28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][month - 2] if month > 1 else 31
    days = 28 if month == 2 else {1: 31, 3: 31, 4: 30, 5: 31, 6: 30, 7: 31, 8: 31, 9: 30, 10: 31, 11: 30, 12: 31}[month]
    truth, raw = [], []
    expected = {k: 0 for k in ("blank_rows", "footer_rows", "missing_id_rejected", "duplicates_removed",
                               "amounts_recomputed", "text_numbers_converted", "dates_normalized",
                               "region_labels_standardized", "cells_trimmed")}
    for i in range(n):
        name, cat, base_c = rng.choice(PRODUCTS)
        region = rng.choices(REGION_LIST, REGION_WEIGHTS)[0]
        units = rng.randint(1, 12)
        price_c = int(round(base_c * rng.uniform(0.96, 1.04)))
        disc = rng.choice([0, 0, 0.05, 0.1, 0.15])
        d = date(year, month, rng.randint(1, days))
        oid = f"SO-{year}{month:02d}-{order_base + i + 1:04d}"
        amount_c = clean.calc_amount_c(units, price_c, disc)
        rep = rng.choice(REPS)
        row = {"order_id": oid, "order_date": d, "region": region, "rep": rep, "product": name, "category": cat,
               "units": units, "unit_price_c": price_c, "discount": disc, "amount_c": amount_c,
               "source_file": fname}
        cells = [oid, fmt_date(d, dstyle), fmt_region(region, rstyle), rep, name, cat, units,
                 fmt_money(price_c, mstyle), f"{int(disc * 100)}%" if mstyle == "usd" else disc,
                 fmt_money(amount_c, mstyle)]
        padded = 0
        if rng.random() < 0.03:  # stray whitespace around a text cell
            col = rng.choice([3, 4, 5])
            cells[col] = f" {cells[col]}  "
            padded += 1
        r = rng.random()
        if r < 0.0025:  # order id missing: cannot be repaired, must be rejected (no other fixes are counted)
            cells[0] = None
            expected["missing_id_rejected"] += 1
            raw.append(cells)
            continue
        # fixes are counted only for rows that are kept, which is what the cleaner counts too
        expected["dates_normalized"] += int(dstyle in ("dmy", "mdy_text"))
        expected["region_labels_standardized"] += int(fmt_region(region, rstyle) != region)
        expected["cells_trimmed"] += padded + int(rstyle == "padded")
        expected["text_numbers_converted"] += 3 if mstyle == "usd" else 0
        if r < 0.0105:  # amount cell empty: recomputed from units x price x (1 - discount)
            cells[9] = None
            expected["amounts_recomputed"] += 1
            if mstyle == "usd":
                expected["text_numbers_converted"] -= 1
        truth.append(row)
        raw.append(cells)
    for _ in range(int(len(raw) * 0.015)):  # exact duplicate rows
        candidates = [c for c in raw if c[0] is not None]
        src = rng.choice(candidates)
        if any(isinstance(c, str) and c != c.strip() for c in src[3:6]):
            continue  # keep the expected whitespace count simple: never duplicate a padded row
        if src[9] is None:
            continue
        # duplicates carry the same formatted text, so they add the same text/date/region fix counts
        raw.insert(rng.randint(1, len(raw)), list(src))
        expected["duplicates_removed"] += 1
    for _ in range(rng.randint(3, 9)):
        raw.insert(rng.randint(0, len(raw)), [None] * 10)
        expected["blank_rows"] += 1
    if footer:
        total = sum(t["amount_c"] for t in truth)
        raw.append(["Total"] + [None] * 8 + [fmt_money(total, mstyle)])
        expected["footer_rows"] = 1
    return truth, raw, expected


def write_file(folder, plan, raw):
    fname, fmt, delim, hstyle = plan[0], plan[1], plan[2], plan[3]
    path = os.path.join(folder, fname)
    if fmt == "xlsx":
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(HEADERS[hstyle])
        for r in raw:
            ws.append(r)
        wb.save(path)
    else:
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f, delimiter=delim)
            w.writerow(HEADERS[hstyle])
            for r in raw:
                w.writerow(["" if c is None else c for c in r])
    return path


def generate(folder, seed=11, include_next=False):
    """Write the exports into `folder`. Returns (truth_records, per_file_expected_counts)."""
    os.makedirs(folder, exist_ok=True)
    rng = random.Random(seed)
    truth, expected = [], {}
    plans = [(2025, i + 1, p) for i, p in enumerate(FILE_PLAN)] + ([(2026, 1, NEXT_MONTH)] if include_next else [])
    for year, month, plan in plans:
        t, raw, exp = make_month(rng, year, month, plan, 0)
        write_file(folder, plan, raw)
        truth += t
        expected[plan[0]] = exp
    return truth, expected


def save_truth(truth, expected, path):
    serial = [{**t, "order_date": t["order_date"].isoformat()} for t in truth]
    with open(path, "w", encoding="utf-8") as f:
        json.dump({"orders": serial, "expected_fixes": expected}, f)


if __name__ == "__main__":
    t, e = generate("raw_exports")
    print(len(t), "true orders in", len(e), "files ->", "./raw_exports")
