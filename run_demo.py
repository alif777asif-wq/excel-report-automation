#!/usr/bin/env python3
"""
End-to-end run: messy monthly exports -> cleaned data -> Excel report -> independent verification,
then a refresh run when next month's export arrives. Writes everything to ./output and a copy of
the small, real artefacts to ./sample_output (used by boards/build_boards.py).

Run:  python run_demo.py        (needs LibreOffice for the verification step)
"""
import json
import os
import re
import shutil
import subprocess
import sys
import time
from datetime import datetime

import openpyxl
import pandas as pd
import xlsxwriter

import build_report
import clean
import make_raw
import verify

ROOT = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(ROOT, "output")
SAMPLE = os.path.join(ROOT, "sample_output")
FIX_KEYS = ["blank_rows", "footer_rows", "missing_id_rejected", "duplicates_removed", "amounts_recomputed",
            "text_numbers_converted", "dates_normalized", "region_labels_standardized", "cells_trimmed"]


def log(level, msg):
    line = f"{datetime.now():%H:%M:%S} [{level}] {msg}"
    print(line, flush=True)
    with open(os.path.join(OUT, "run.log"), "a", encoding="utf-8") as f:
        f.write(line + "\n")


def frame(records):
    df = pd.DataFrame(records)
    df["unit_price"] = df["unit_price_c"] / 100
    df["amount"] = df["amount_c"] / 100
    return df


def pipeline(raw_dir, xlsx_path):
    """clean -> build. Returns (records, logs, df, expected_values, seconds_clean, seconds_build)."""
    t0 = time.time()
    records, logs = clean.clean_folder(raw_dir)
    t1 = time.time()
    df = frame(records)
    expected = build_report.build(df, logs, xlsx_path)
    return records, logs, df, expected, t1 - t0, time.time() - t1


def tool_versions():
    lo = subprocess.run(["soffice", "--version"], capture_output=True, text=True).stdout.strip()
    return {"python": sys.version.split()[0], "pandas": pd.__version__, "xlsxwriter": xlsxwriter.__version__,
            "openpyxl": openpyxl.__version__, "libreoffice": re.sub(r"\s+\(.*", "", lo)}


def raw_examples(raw_dir):
    """Verbatim raw lines that show the problems, taken from the real export files."""
    out = {"headers": {}, "blank_row": None, "footer": None, "duplicate": None, "panels": []}
    for fname in sorted(os.listdir(raw_dir)):
        if not fname.endswith(".csv"):
            continue
        with open(os.path.join(raw_dir, fname), encoding="utf-8") as f:
            lines = f.read().splitlines()
        out["headers"][fname] = lines[0]
        for no, ln in enumerate(lines[1:], start=2):
            if out["blank_row"] is None and ln.replace(",", "").replace(";", "").strip() == "":
                out["blank_row"] = {"file": fname, "line_no": no, "text": ln}
            if out["footer"] is None and ln.lower().startswith("total"):
                out["footer"] = {"file": fname, "line_no": no, "text": ln}
        seen = {}
        for no, ln in enumerate(lines[1:], start=2):
            if ln.strip(",; ") and ln in seen and out["duplicate"] is None:
                out["duplicate"] = {"file": fname, "line_nos": [seen[ln], no], "text": ln}
            seen.setdefault(ln, no)
    for fname in ("sales_2025-01.csv", "Feb_2025_branch_export.csv", "APR-sales-report.csv"):
        with open(os.path.join(raw_dir, fname), encoding="utf-8") as f:
            lines = f.read().splitlines()
        out["panels"].append({"file": fname, "lines": lines[:6]})
    return out


def main():
    shutil.rmtree(OUT, ignore_errors=True)
    os.makedirs(OUT)
    rep = {"environment": "synthetic monthly sales exports generated locally (fictional reps and products), no real data",
           "generated_at": datetime.now().isoformat(timespec="seconds"), "tools": tool_versions()}

    # ---------------------------------------------------------------- 1. the raw exports
    raw_dir = os.path.join(OUT, "raw_exports")
    truth, expected_fixes = make_raw.generate(raw_dir)
    files = sorted(os.listdir(raw_dir))
    types = {t: sum(f.lower().endswith(t) for f in files) for t in ("csv", "xlsx")}
    log("INFO", f"Raw exports: {len(files)} files ({types['csv']} CSV, {types['xlsx']} XLSX), "
                f"3 header spellings, 4 date styles, comma and semicolon delimiters")

    # ---------------------------------------------------------------- 2. clean + build
    xlsx_path = os.path.join(OUT, "Sales_Report.xlsx")
    records, logs, df, exp, t_clean, t_build = pipeline(raw_dir, xlsx_path)
    tot = {k: sum(l_[k] for l_ in logs) for k in FIX_KEYS + ["rows_read", "rows_kept", "ambiguous_dates", "amount_mismatches"]}
    tot["fixes_total"] = sum(tot[k] for k in FIX_KEYS)
    log("INFO", f"Cleaning: {tot['rows_read']:,} raw rows read, {tot['rows_kept']:,} orders kept in {t_clean:.1f}s")
    log("INFO", f"Fixed: {tot['blank_rows']} blank rows, {tot['footer_rows']} footer rows, "
                f"{tot['duplicates_removed']} duplicates, {tot['amounts_recomputed']} missing amounts recomputed, "
                f"{tot['missing_id_rejected']} rows rejected (no order ID)")
    log("INFO", f"Normalised: {tot['dates_normalized']:,} dates, {tot['region_labels_standardized']:,} region labels, "
                f"{tot['text_numbers_converted']:,} text numbers, {tot['cells_trimmed']:,} padded cells")
    matched = sum(all(l_[k] == expected_fixes[l_["file"]][k] for k in FIX_KEYS) for l_ in logs)
    truth_check = verify.compare_with_truth(records, truth)
    log("INFO", f"Cleaning log equals the injected problems in {matched}/{len(logs)} files")
    log("INFO", f"Ground truth: {truth_check['cleaned_orders']:,} of {truth_check['true_orders']:,} orders recovered, "
                f"{truth_check['orders_with_wrong_field']} with a wrong field, revenue matches: {truth_check['revenue_matches']}")
    log("INFO", f"Report built in {t_build:.1f}s: {os.path.getsize(xlsx_path) / 1e6:.2f} MB, 5 sheets")

    # ---------------------------------------------------------------- 3. independent verification
    rc, t_recalc = verify.recalculate(xlsx_path, os.path.join(OUT, "recalculated"))
    cmp_ = verify.compare_formulas(xlsx_path, rc)
    tamper = os.path.join(OUT, "recalculated", "tampered_for_selftest.xlsx")
    verify.tamper_copy(xlsx_path, tamper)
    rc2, _ = verify.recalculate(tamper, os.path.join(OUT, "recalculated", "selftest"))
    selftest = verify.compare_formulas(tamper, rc2)
    log("INFO", f"LibreOffice recalculation ({t_recalc:.1f}s): {cmp_['formulas']} formulas, "
                f"{cmp_['mismatches']} mismatches, {cmp_['errors']} errors, max deviation {cmp_['max_abs_deviation']:.1e}")
    log("INFO", f"Verifier self-test: a planted wrong value was caught ({selftest['mismatches']} mismatch found)")

    # ---------------------------------------------------------------- 4. refresh: next month arrives
    refresh_raw = os.path.join(OUT, "refresh_exports")
    truth2, exp2 = make_raw.generate(refresh_raw, include_next=True)
    refresh_xlsx = os.path.join(OUT, "refresh", "Sales_Report.xlsx")
    os.makedirs(os.path.dirname(refresh_xlsx))
    t0 = time.time()
    rec2, logs2, df2, e2, tc2, tb2 = pipeline(refresh_raw, refresh_xlsx)
    t_refresh = time.time() - t0
    rc3, _ = verify.recalculate(refresh_xlsx, os.path.join(OUT, "refresh", "recalculated"))
    cmp2 = verify.compare_formulas(refresh_xlsx, rc3)
    truth2_check = verify.compare_with_truth(rec2, truth2)
    log("INFO", f"Refresh: next month's export added ({len(logs2)} files, {len(e2['months'])} months), "
                f"report rebuilt in {t_refresh:.1f}s, {len(rec2) - len(records):,} new orders")
    log("INFO", f"Refresh check: {cmp2['formulas']} formulas, {cmp2['mismatches']} mismatches, {cmp2['errors']} errors, "
                f"{truth2_check['orders_with_wrong_field']} orders with a wrong field")

    # ---------------------------------------------------------------- 5. report + sample artefacts
    rep.update({
        "files": {"count": len(files), **types, "header_styles": 3, "date_styles": 4},
        "cleaning": {**tot, "seconds": round(t_clean, 2), "log_matches_injected_files": matched,
                     "files_checked": len(logs), "ground_truth": truth_check},
        "workbook": {"orders": exp["orders"], "revenue": exp["total_revenue"], "units": exp["units"],
                     "sheets": 5, "size_bytes": os.path.getsize(xlsx_path), "build_seconds": round(t_build, 2),
                     "by_region": {g: {"orders": v[0], "units": v[1], "revenue": v[2]} for g, v in exp["by_region"].items()},
                     "monthly_total": [exp["month_total"][ym] for ym in exp["months"]],
                     "months": [f"{y}-{m:02d}" for y, m in exp["months"]], "last_row": exp["last_row"]},
        "verification": {"engine": rep["tools"]["libreoffice"], "formulas": cmp_["formulas"],
                         "by_sheet": cmp_["by_sheet"], "mismatches": cmp_["mismatches"], "errors": cmp_["errors"],
                         "max_abs_deviation": cmp_["max_abs_deviation"], "recalc_seconds": round(t_recalc, 1),
                         "selftest_caught": selftest["mismatches"] >= 1},
        "pipeline_seconds": round(t_clean + t_build, 2),
        "refresh": {"new_file": "sales_2026-01.csv", "files": len(logs2), "months_after": len(e2["months"]),
                    "orders_added": len(rec2) - len(records), "formulas": cmp2["formulas"],
                    "mismatches": cmp2["mismatches"], "errors": cmp2["errors"],
                    "pipeline_seconds": round(t_refresh, 2), "orders_with_wrong_field": truth2_check["orders_with_wrong_field"]},
    })
    shutil.rmtree(SAMPLE, ignore_errors=True)
    os.makedirs(SAMPLE)
    with open(os.path.join(OUT, "run_report.json"), "w", encoding="utf-8") as f:
        json.dump(rep, f, indent=2, default=str)
    for src in ("run_report.json", "run.log", "Sales_Report.xlsx"):
        shutil.copy(os.path.join(OUT, src), os.path.join(SAMPLE, src))
    pd.DataFrame(logs).to_csv(os.path.join(SAMPLE, "cleaning_log.csv"), index=False)
    with open(os.path.join(SAMPLE, "raw_examples.json"), "w", encoding="utf-8") as f:
        json.dump(raw_examples(raw_dir), f, indent=1)
    shutil.copytree(raw_dir, os.path.join(SAMPLE, "raw_exports"))
    log("INFO", "Done")
    return 0


if __name__ == "__main__":
    sys.exit(main())
