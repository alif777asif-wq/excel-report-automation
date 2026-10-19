"""
Cleaning layer for messy monthly sales exports (CSV and XLSX from different branches).

Reads every export exactly as delivered (all cells as raw values), maps the different header
spellings to one schema, repairs what can be repaired, rejects what cannot, and counts every fix
so the cleaning can be audited. Money is handled in integer cents to avoid rounding drift.
"""
import csv
import os
import re
from datetime import date, datetime
from decimal import ROUND_HALF_UP, Decimal

import openpyxl

SCHEMA = ["order_id", "order_date", "region", "rep", "product", "category", "units", "unit_price", "discount", "amount"]
HEADER_ALIASES = {
    "orderid": "order_id", "orderdate": "order_date", "date": "order_date", "region": "region", "area": "region",
    "salesrep": "rep", "rep": "rep", "product": "product", "item": "product", "category": "category",
    "cat": "category", "units": "units", "qty": "units", "unitprice": "unit_price", "price": "unit_price",
    "discount": "discount", "disc": "discount", "discpct": "discount", "amount": "amount", "total": "amount",
    "netamount": "amount",
}
REGIONS = {"north": "North", "n": "North", "south": "South", "s": "South", "east": "East", "e": "East",
           "west": "West", "w": "West", "central": "Central", "c": "Central"}
FOOTER_LABELS = {"total", "grand total"}
TEXT_COLS = ("order_id", "region", "rep", "product", "category")


def calc_amount_c(units, price_c, discount):
    """Net amount in cents: units x unit price x (1 - discount), rounded half up."""
    value = Decimal(units) * Decimal(price_c) * (Decimal(1) - Decimal(str(discount)))
    return int(value.quantize(Decimal(1), rounding=ROUND_HALF_UP))


def parse_money_c(value):
    """'$1,299.00' / 1299.0 / '1299' -> cents. Returns (cents, was_text_formatted)."""
    if isinstance(value, (int, float)):
        return int((Decimal(str(value)) * 100).quantize(Decimal(1), rounding=ROUND_HALF_UP)), False
    s = str(value).strip()
    formatted = "$" in s or "," in s
    s = s.replace("$", "").replace(",", "")
    return int((Decimal(s) * 100).quantize(Decimal(1), rounding=ROUND_HALF_UP)), formatted


def parse_discount(value):
    """'10%' / 0.1 / '0.1' -> Decimal fraction. Returns (fraction, was_text_formatted)."""
    if isinstance(value, (int, float)):
        return Decimal(str(value)), False
    s = str(value).strip()
    if s.endswith("%"):
        return Decimal(s[:-1]) / 100, True
    return Decimal(s), False


def read_raw(path):
    """Return (header, rows) with every cell as the raw value (str, number, datetime or None)."""
    if path.lower().endswith(".xlsx"):
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        rows = [list(r) for r in wb.worksheets[0].iter_rows(values_only=True)]
        wb.close()
    else:
        with open(path, newline="", encoding="utf-8-sig") as f:
            first = f.readline()
            f.seek(0)
            delim = ";" if first.count(";") > first.count(",") else ","
            rows = [[c if c != "" else None for c in r] for r in csv.reader(f, delimiter=delim)]
    return rows[0], rows[1:]


def map_header(header):
    mapping = {}
    for idx, h in enumerate(header):
        key = re.sub(r"[^a-z0-9]", "", str(h).lower())
        if key in HEADER_ALIASES:
            mapping[HEADER_ALIASES[key]] = idx
    missing = [c for c in SCHEMA if c not in mapping]
    if missing:
        raise ValueError(f"unrecognised header, missing {missing}")
    return mapping


def detect_date_format(values):
    """Pick one date format for the whole file; returns (format, ambiguous)."""
    strings = [v for v in values if isinstance(v, str)]
    if not strings:
        return None, False
    s = strings[0].strip()
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s):
        return "%Y-%m-%d", False
    if re.fullmatch(r"[A-Za-z]{3} \d{1,2}, \d{4}", s):
        return "%b %d, %Y", False
    if re.fullmatch(r"\d{2}/\d{2}/\d{4}", s):
        firsts = [int(v.strip()[:2]) for v in strings]
        seconds = [int(v.strip()[3:5]) for v in strings]
        if max(firsts) > 12:
            return "%d/%m/%Y", False
        if max(seconds) > 12:
            return "%m/%d/%Y", False
        return "%d/%m/%Y", True  # genuinely ambiguous: day first, and it is flagged
    raise ValueError(f"unrecognised date format: {s!r}")


def is_blank(row):
    return all(c is None or str(c).strip() == "" for c in row)


def clean_file(path):
    """Clean one export. Returns (records, log). records are dicts in the canonical schema."""
    header, rows = read_raw(path)
    m = map_header(header)
    date_fmt, ambiguous = detect_date_format([r[m["order_date"]] for r in rows if not is_blank(r)])
    log = {"file": os.path.basename(path), "rows_read": len(rows), "blank_rows": 0, "footer_rows": 0,
           "missing_id_rejected": 0, "duplicates_removed": 0, "amounts_recomputed": 0, "text_numbers_converted": 0,
           "dates_normalized": 0, "region_labels_standardized": 0, "cells_trimmed": 0, "rows_kept": 0,
           "ambiguous_dates": int(ambiguous), "amount_mismatches": 0}
    seen, records = set(), []
    for r in rows:
        if is_blank(r):
            log["blank_rows"] += 1
            continue
        get = lambda col: r[m[col]] if m[col] < len(r) else None  # noqa: E731
        first = str(r[0]).strip().lower() if r and r[0] is not None else ""
        if first in FOOTER_LABELS:
            log["footer_rows"] += 1
            continue
        oid = get("order_id")
        if oid is None or str(oid).strip() == "":
            log["missing_id_rejected"] += 1
            continue
        oid = str(oid).strip()
        if oid in seen:
            log["duplicates_removed"] += 1
            continue
        seen.add(oid)
        log["cells_trimmed"] += sum(1 for c in TEXT_COLS if isinstance(get(c), str) and get(c) != get(c).strip())
        raw_date = get("order_date")
        if isinstance(raw_date, datetime):
            d = raw_date.date()
        elif isinstance(raw_date, date):
            d = raw_date
        else:
            d = datetime.strptime(str(raw_date).strip(), date_fmt).date()
            log["dates_normalized"] += int(date_fmt != "%Y-%m-%d")
        region_raw = str(get("region"))
        region = REGIONS[region_raw.strip().lower()]
        log["region_labels_standardized"] += int(region_raw != region)
        units = int(float(get("units")))
        price_c, p_text = parse_money_c(get("unit_price"))
        disc, d_text = parse_discount(get("discount"))
        amount_raw = get("amount")
        calc_c = calc_amount_c(units, price_c, disc)
        if amount_raw is None or str(amount_raw).strip() == "":
            amount_c, a_text = calc_c, False
            log["amounts_recomputed"] += 1
        else:
            amount_c, a_text = parse_money_c(amount_raw)
            log["amount_mismatches"] += int(amount_c != calc_c)
        log["text_numbers_converted"] += int(p_text) + int(d_text) + int(a_text)
        records.append({"order_id": oid, "order_date": d, "region": region, "rep": str(get("rep")).strip(),
                        "product": str(get("product")).strip(), "category": str(get("category")).strip(),
                        "units": units, "unit_price_c": price_c, "discount": float(disc), "amount_c": amount_c,
                        "source_file": os.path.basename(path)})
    log["rows_kept"] = len(records)
    return records, log


def clean_folder(folder):
    """Clean every export in a folder (sorted by name). Returns (records, logs)."""
    paths = sorted(os.path.join(folder, f) for f in os.listdir(folder) if f.lower().endswith((".csv", ".xlsx")))
    all_records, logs = [], []
    for p in paths:
        rec, log = clean_file(p)
        all_records += rec
        logs.append(log)
    seen, unique = {}, []
    for r in all_records:  # the same order id arriving in two different files is also a duplicate
        if r["order_id"] in seen:
            logs[[l_["file"] for l_ in logs].index(r["source_file"])]["duplicates_removed"] += 1
            logs[[l_["file"] for l_ in logs].index(r["source_file"])]["rows_kept"] -= 1
            continue
        seen[r["order_id"]] = True
        unique.append(r)
    unique.sort(key=lambda r: (r["order_date"], r["order_id"]))
    return unique, logs
