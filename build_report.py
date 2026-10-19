"""
Build the sales report workbook from the cleaned orders.

Everything on the report sheets is a live formula over the 'Clean Data' sheet, so the numbers
recalculate in Excel. The expected result of every formula is stored in the file as its cached
value, which means previewers that never recalculate still show real numbers. verify.py then
recalculates the workbook independently in LibreOffice and compares.
"""
from datetime import datetime

import xlsxwriter

REGIONS = ["North", "South", "East", "West", "Central"]
GREEN, GREEN_DARK, PAPER, YELLOW = "#0B3D2E", "#082B20", "#EEF2EC", "#F5C93B"
LOG_COLUMNS = [
    ("rows_read", "Rows read"), ("blank_rows", "Blank rows"), ("footer_rows", "Footer rows"),
    ("missing_id_rejected", "Missing ID (rejected)"), ("duplicates_removed", "Duplicates removed"),
    ("amounts_recomputed", "Amounts recomputed"), ("text_numbers_converted", "Text numbers converted"),
    ("dates_normalized", "Dates normalized"), ("region_labels_standardized", "Region labels standardized"),
    ("cells_trimmed", "Cells trimmed"), ("rows_kept", "Rows kept"),
]


def monthly_formula(last, row, col):
    """One grid cell: a region's revenue in one calendar month."""
    amt = f"'Clean Data'!$J$2:$J${last}"
    day = f"'Clean Data'!$B$2:$B${last}"
    reg = f"'Clean Data'!$C$2:$C${last}"
    return (f'=SUMIFS({amt},{day},">="&$A{row},'
            f'{day},"<"&EDATE($A{row},1),{reg},{col}$4)')


def region_formulas(last, row):
    """Orders, units and revenue for one region row of the By Region sheet."""
    reg = f"'Clean Data'!$C$2:$C${last}"
    return (f"=COUNTIFS({reg},$A{row})",
            f"=SUMIFS('Clean Data'!$G$2:$G${last},{reg},$A{row})",
            f"=SUMIFS('Clean Data'!$J$2:$J${last},{reg},$A{row})")


def month_starts(df):
    first, last = df["order_date"].min(), df["order_date"].max()
    months, y, m = [], first.year, first.month
    while (y, m) <= (last.year, last.month):
        months.append((y, m))
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    return months


def build(df, logs, path):
    """df: cleaned orders with columns order_id, order_date, region, rep, product, category, units,
    unit_price, discount, amount (dollars), source_file. Returns a dict of the expected cell values."""
    n = len(df)
    last = n + 1
    wb = xlsxwriter.Workbook(path, {"default_date_format": "yyyy-mm-dd"})
    wb.set_properties({"title": "Sales report", "subject": "Automated Excel report (synthetic demo data)",
                       "comments": "Built by build_report.py; every figure on the report sheets is a live formula."})
    f = lambda **k: wb.add_format({"font_name": "Calibri", "font_size": 11, **k})  # noqa: E731
    title, sub = f(font_size=20, bold=True, font_color=GREEN), f(font_color="#5B6B63", italic=True)
    head = f(bold=True, font_color="#FFFFFF", bg_color=GREEN, border=0, valign="vcenter", text_wrap=True)
    head_r = f(bold=True, font_color="#FFFFFF", bg_color=GREEN, align="right", valign="vcenter", text_wrap=True)
    txt, intg = f(), f(num_format="#,##0")
    usd, usd2, pct, mon = f(num_format="$#,##0.00"), f(num_format="$#,##0"), f(num_format="0.0%"), f(num_format="mmm yyyy", align="left")
    tot_t, tot_i = f(bold=True, top=1), f(bold=True, top=1, num_format="#,##0")
    tot_u, tot_p = f(bold=True, top=1, num_format="$#,##0.00"), f(bold=True, top=1, num_format="0.0%")
    kpi_l = f(font_color="#5B6B63", font_size=10, valign="bottom")
    kpi_usd = f(font_size=22, bold=True, font_color=GREEN, num_format="$#,##0", align="left")
    kpi_int = f(font_size=22, bold=True, font_color=GREEN, num_format="#,##0", align="left")
    kpi_usd2 = f(font_size=22, bold=True, font_color=GREEN, num_format="$#,##0.00", align="left")
    kpi_txt = f(font_size=22, bold=True, font_color=GREEN, align="left")
    kpi_mon = f(font_size=22, bold=True, font_color=GREEN, num_format="mmm yyyy", align="left")
    kpi_pct = f(font_size=22, bold=True, font_color=GREEN, num_format="+0.0%;-0.0%", align="left")

    ws_sum = wb.add_worksheet("Summary")
    ws_reg = wb.add_worksheet("By Region")
    ws_mon = wb.add_worksheet("Monthly")
    ws_dat = wb.add_worksheet("Clean Data")
    ws_log = wb.add_worksheet("Cleaning Log")

    # ---------------------------------------------------------------- Clean Data (a real Excel table)
    cols = ["Order ID", "Date", "Region", "Rep", "Product", "Category", "Units", "Unit Price", "Discount", "Amount",
            "Source File"]
    ws_dat.add_table(0, 0, n, len(cols) - 1, {"name": "SalesData", "style": "Table Style Medium 7",
                                              "columns": [{"header": c} for c in cols]})
    date_f, money_f, pct_f, int_f = f(num_format="yyyy-mm-dd"), f(num_format="$#,##0.00"), f(num_format="0%"), f()
    for i, r in enumerate(df.itertuples(index=False), start=1):
        ws_dat.write_string(i, 0, r.order_id)
        ws_dat.write_datetime(i, 1, datetime(r.order_date.year, r.order_date.month, r.order_date.day), date_f)
        for c, v in enumerate((r.region, r.rep, r.product, r.category), start=2):
            ws_dat.write_string(i, c, v)
        ws_dat.write_number(i, 6, r.units, int_f)
        ws_dat.write_number(i, 7, r.unit_price, money_f)
        ws_dat.write_number(i, 8, r.discount, pct_f)
        ws_dat.write_number(i, 9, r.amount, money_f)
        ws_dat.write_string(i, 10, r.source_file)
    for c, w in enumerate([16, 12, 10, 18, 20, 13, 7, 11, 9, 12, 28]):
        ws_dat.set_column(c, c, w)
    ws_dat.freeze_panes(1, 0)

    # ---------------------------------------------------------------- expected values (for cached results)
    by_reg = {g: (len(x), int(x["units"].sum()), round(float(x["amount"].sum()), 2)) for g, x in df.groupby("region")}
    total_orders, total_units = n, int(df["units"].sum())
    total_rev = round(float(df["amount"].sum()), 2)
    months = month_starts(df)
    df = df.assign(_m=[(d.year, d.month) for d in df["order_date"]])
    grid = {(ym, g): round(float(x["amount"].sum()), 2) for (ym, g), x in df.groupby(["_m", "region"])}
    month_tot = {ym: round(sum(grid.get((ym, g), 0.0) for g in REGIONS), 2) for ym in months}

    # ---------------------------------------------------------------- By Region
    ws_reg.write("A1", "Revenue by region", title)
    ws_reg.write("A2", "Every figure is a formula over the Clean Data sheet.", sub)
    for c, h in enumerate(["Region", "Orders", "Units", "Revenue", "Average order", "Share of revenue"]):
        ws_reg.write(3, c, h, head if c == 0 else head_r)
    for i, g in enumerate(REGIONS):
        row = i + 5
        o_f, u_f, r_f = region_formulas(last, row)
        orders, units, rev = by_reg[g]
        ws_reg.write_string(i + 4, 0, g, txt)
        ws_reg.write_formula(i + 4, 1, o_f, intg, orders)
        ws_reg.write_formula(i + 4, 2, u_f, intg, units)
        ws_reg.write_formula(i + 4, 3, r_f, usd2, rev)
        ws_reg.write_formula(i + 4, 4, f"=IF(B{row}=0,0,D{row}/B{row})", usd, rev / orders)
        ws_reg.write_formula(i + 4, 5, f"=D{row}/$D$10", pct, rev / total_rev)
    ws_reg.write_string(9, 0, "Total", tot_t)
    ws_reg.write_formula(9, 1, "=SUM(B5:B9)", tot_i, total_orders)
    ws_reg.write_formula(9, 2, "=SUM(C5:C9)", tot_i, total_units)
    ws_reg.write_formula(9, 3, "=SUM(D5:D9)", tot_u, total_rev)
    ws_reg.write_formula(9, 4, "=D10/B10", tot_u, total_rev / total_orders)
    ws_reg.write_formula(9, 5, "=SUM(F5:F9)", tot_p, 1.0)
    ws_reg.conditional_format("D5:D9", {"type": "data_bar", "bar_color": YELLOW, "bar_solid": True})
    ws_reg.set_column("A:A", 14)
    ws_reg.set_column("B:C", 11)
    ws_reg.set_column("D:F", 17)
    ws_reg.set_row(3, 30)

    # ---------------------------------------------------------------- Monthly
    ws_mon.write("A1", "Revenue by month and region", title)
    ws_mon.write("A2", "Month starts are real dates; each cell is a SUMIFS over the Clean Data sheet.", sub)
    ws_mon.write(3, 0, "Month", head)
    for c, g in enumerate(REGIONS, start=1):
        ws_mon.write(3, c, g, head_r)
    ws_mon.write(3, 6, "Total", head_r)
    ws_mon.write(3, 7, "Change vs previous month", head_r)
    first_r, last_r = 5, 4 + len(months)
    for i, ym in enumerate(months):
        row = 5 + i
        ws_mon.write_datetime(row - 1, 0, datetime(ym[0], ym[1], 1), mon)
        for c, g in enumerate(REGIONS, start=1):
            col = "ABCDEFGH"[c]
            ws_mon.write_formula(row - 1, c, monthly_formula(last, row, col), usd2, grid.get((ym, g), 0.0))
        ws_mon.write_formula(row - 1, 6, f"=SUM(B{row}:F{row})", f(num_format="$#,##0", bold=True), month_tot[ym])
        if i == 0:
            ws_mon.write_string(row - 1, 7, "n/a", f(align="right", font_color="#8A968F"))
        else:
            prev = month_tot[months[i - 1]]
            ws_mon.write_formula(row - 1, 7, f'=IF(G{row - 1}=0,"",G{row}/G{row - 1}-1)', f(num_format="+0.0%;-0.0%"),
                                 month_tot[ym] / prev - 1 if prev else "")
    tr = last_r + 1
    ws_mon.write_string(tr - 1, 0, "Total", tot_t)
    for c, g in enumerate(REGIONS, start=1):
        col = "ABCDEFGH"[c]
        ws_mon.write_formula(tr - 1, c, f"=SUM({col}{first_r}:{col}{last_r})", f(bold=True, top=1, num_format="$#,##0"),
                             round(sum(grid.get((ym, g), 0.0) for ym in months), 2))
    ws_mon.write_formula(tr - 1, 6, f"=SUM(G{first_r}:G{last_r})", f(bold=True, top=1, num_format="$#,##0"),
                         round(sum(month_tot.values()), 2))
    ws_mon.conditional_format(f"B{first_r}:F{last_r}", {"type": "3_color_scale", "min_color": "#FFFFFF",
                                                       "mid_color": "#DDEBD3", "max_color": "#8FBF8A"})
    ws_mon.set_column("A:A", 12)
    ws_mon.set_column("B:G", 13)
    ws_mon.set_column("H:H", 17)
    ws_mon.set_row(3, 30)
    ws_mon.freeze_panes(4, 1)

    # ---------------------------------------------------------------- Summary
    first_d, last_d = df["order_date"].min(), df["order_date"].max()
    ws_sum.hide_gridlines(2)
    ws_sum.write("B1", "Sales report", title)
    ws_sum.write("B2", f"{first_d:%d %b %Y} to {last_d:%d %b %Y}, {n:,} orders consolidated from {len(logs)} export files",
                 sub)
    best_region = max(by_reg, key=lambda g: by_reg[g][2])
    best_month = max(months, key=lambda ym: month_tot[ym])
    last_mom = month_tot[months[-1]] / month_tot[months[-2]] - 1
    kpis = [
        ("Total revenue", "='By Region'!D10", kpi_usd, total_rev),
        ("Orders", "='By Region'!B10", kpi_int, total_orders),
        ("Units sold", "='By Region'!C10", kpi_int, total_units),
        ("Average order value", "='By Region'!D10/'By Region'!B10", kpi_usd2, total_rev / total_orders),
        ("Top region by revenue", "=INDEX('By Region'!A5:A9,MATCH(MAX('By Region'!D5:D9),'By Region'!D5:D9,0))",
         kpi_txt, best_region),
        ("Best month by revenue", f"=INDEX(Monthly!A{first_r}:A{last_r},MATCH(MAX(Monthly!G{first_r}:G{last_r}),Monthly!G{first_r}:G{last_r},0))",
         kpi_mon, datetime(best_month[0], best_month[1], 1)),
        ("Latest month vs previous", f"=Monthly!H{last_r}", kpi_pct, last_mom),
    ]
    for i, (label, formula, fmt, value) in enumerate(kpis):
        r = 3 + i * 2
        ws_sum.write(r, 1, label, kpi_l)
        if hasattr(value, "year"):
            serial = (value - datetime(1899, 12, 30)).days
            ws_sum.write_formula(r + 1, 1, formula, fmt, serial)
        else:
            ws_sum.write_formula(r + 1, 1, formula, fmt, value)
        ws_sum.set_row(r + 1, 32)
    ws_sum.set_column("A:A", 3)
    ws_sum.set_column("B:B", 26)
    ws_sum.set_column("C:C", 3)
    c1 = wb.add_chart({"type": "column"})
    c1.add_series({"name": "Revenue", "categories": "='By Region'!$A$5:$A$9", "values": "='By Region'!$D$5:$D$9",
                   "fill": {"color": GREEN}, "gap": 60})
    c1.set_title({"name": "Revenue by region", "name_font": {"size": 12, "bold": True, "color": GREEN}})
    c1.set_legend({"none": True})
    c1.set_y_axis({"num_format": "$#,##0", "min": 0, "major_gridlines": {"visible": True, "line": {"color": "#D9DFD8"}}})
    c1.set_size({"width": 520, "height": 300})
    ws_sum.insert_chart("D4", c1)
    c2 = wb.add_chart({"type": "column"})
    c2.add_series({"name": "Revenue", "categories": f"=Monthly!$A${first_r}:$A${last_r}",
                   "values": f"=Monthly!$G${first_r}:$G${last_r}", "fill": {"color": GREEN}, "gap": 40})
    c2.set_title({"name": "Revenue by month", "name_font": {"size": 12, "bold": True, "color": GREEN}})
    c2.set_legend({"none": True})
    c2.set_x_axis({"num_format": "mmm"})
    c2.set_y_axis({"num_format": "$#,##0", "min": 0, "major_gridlines": {"visible": True, "line": {"color": "#D9DFD8"}}})
    c2.set_size({"width": 520, "height": 300})
    ws_sum.insert_chart("D20", c2)

    # ---------------------------------------------------------------- Cleaning Log
    ws_log.write("A1", "What the cleaning step fixed", title)
    ws_log.write("A2", "One row per export file. Rows read = blank + footer + rejected + duplicates + kept.", sub)
    ws_log.write(3, 0, "Export file", head)
    for c, (_, h) in enumerate(LOG_COLUMNS, start=1):
        ws_log.write(3, c, h, head_r)
    ws_log.write(3, len(LOG_COLUMNS) + 1, "Reconciles", head_r)
    for i, lg in enumerate(logs):
        row = 5 + i
        ws_log.write_string(i + 4, 0, lg["file"], txt)
        for c, (k, _) in enumerate(LOG_COLUMNS, start=1):
            ws_log.write_number(i + 4, c, lg[k], intg)
        ok = lg["rows_read"] == lg["blank_rows"] + lg["footer_rows"] + lg["missing_id_rejected"] + lg["duplicates_removed"] + lg["rows_kept"]
        ws_log.write_formula(i + 4, len(LOG_COLUMNS) + 1, f"=B{row}=SUM(C{row}:F{row},L{row})", f(align="right"), ok)
    lr = 5 + len(logs)
    ws_log.write_string(lr - 1, 0, "Total", tot_t)
    for c, (k, _) in enumerate(LOG_COLUMNS, start=1):
        col = "ABCDEFGHIJKLMN"[c]
        ws_log.write_formula(lr - 1, c, f"=SUM({col}5:{col}{lr - 1})", tot_i, sum(l_[k] for l_ in logs))
    ws_log.set_column("A:A", 30)
    ws_log.set_column("B:M", 13)
    ws_log.set_row(3, 44)
    ws_log.freeze_panes(4, 1)

    for ws in (ws_sum, ws_reg, ws_mon, ws_log):
        ws.set_landscape()
        ws.fit_to_pages(1, 0)
    ws_sum.activate()
    wb.close()
    return {"total_revenue": total_rev, "orders": total_orders, "units": total_units, "by_region": by_reg,
            "months": months, "month_total": month_tot, "last_row": last}
