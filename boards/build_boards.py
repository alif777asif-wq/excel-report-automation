#!/usr/bin/env python3
"""
Build the six Upwork portfolio boards (self-contained HTML + CSS, 1280 x 720 artboards).

Every number, label, raw data line and code line on the boards is read from files produced by a real
run: ../sample_output/run_report.json, raw_examples.json and the real source ../build_report.py.
Nothing is typed in by hand, so nothing can drift from the data.

Run:  python build_boards.py        then   python export.py
"""
import base64
import html
import json
import os
import re
import sys
import textwrap

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SAMPLE = os.path.join(ROOT, "sample_output")
sys.path.insert(0, ROOT)
import build_report  # noqa: E402  (the real module: its output is quoted on board 5)
GITHUB = "https://github.com/alif777asif-wq/excel-report-automation"

REP = json.load(open(os.path.join(SAMPLE, "run_report.json"), encoding="utf-8"))
RAW = json.load(open(os.path.join(SAMPLE, "raw_examples.json"), encoding="utf-8"))
SRC = open(os.path.join(ROOT, "build_report.py"), encoding="utf-8").read().splitlines()

CL, WB, VF, RF = REP["cleaning"], REP["workbook"], REP["verification"], REP["refresh"]
REGIONS = ["North", "South", "East", "West", "Central"]
TOT_REV = WB["revenue"]


# ------------------------------------------------------------------ formatting helpers
def n(v):
    return f"{v:,}"


def usd(v, cents=True):
    return f"${v:,.2f}" if cents else f"${v:,.0f}"


def esc(s):
    return html.escape(s, quote=False)


def contrast(a, b):
    def lum(h):
        h = h.lstrip("#")
        c = [int(h[i:i + 2], 16) / 255 for i in (0, 2, 4)]
        c = [x / 12.92 if x <= 0.03928 else ((x + 0.055) / 1.055) ** 2.4 for x in c]
        return 0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2]
    la, lb = lum(a), lum(b)
    hi, lo = max(la, lb), min(la, lb)
    return (hi + 0.05) / (lo + 0.05)


PALETTE = {"green": "#0B3D2E", "green_2": "#14503E", "green_deep": "#082B20", "paper": "#E8EEE7", "panel": "#F3F6F1",
           "ink": "#0B241B", "ink_2": "#46594F", "mute": "#9DB8AB", "yellow": "#F5C93B"}


# ------------------------------------------------------------------ fonts (embedded, so no network is needed)
def font_face(family, weight, style, fname):
    data = base64.b64encode(open(os.path.join(HERE, "fonts", fname), "rb").read()).decode()
    return (f"@font-face{{font-family:'{family}';font-weight:{weight};font-style:{style};font-display:block;"
            f"src:url(data:font/woff2;base64,{data}) format('woff2')}}")


FONTS = "\n".join([
    font_face("Source Serif 4", 600, "normal", "source-serif-4-latin-600-normal.woff2"),
    font_face("Source Serif 4", 700, "normal", "source-serif-4-latin-700-normal.woff2"),
    font_face("Source Serif 4", 600, "italic", "source-serif-4-latin-600-italic.woff2"),
    font_face("Hanken Grotesk", 400, "normal", "hanken-grotesk-latin-400-normal.woff2"),
    font_face("Hanken Grotesk", 500, "normal", "hanken-grotesk-latin-500-normal.woff2"),
    font_face("Hanken Grotesk", 600, "normal", "hanken-grotesk-latin-600-normal.woff2"),
    font_face("JetBrains Mono", 400, "normal", "jetbrains-mono-latin-400-normal.woff2"),
    font_face("JetBrains Mono", 500, "normal", "jetbrains-mono-latin-500-normal.woff2"),
    font_face("JetBrains Mono", 700, "normal", "jetbrains-mono-latin-700-normal.woff2"),
])

BASE_CSS = """
:root{
  --green:#0B3D2E; --green-2:#14503E; --green-deep:#082B20; --paper:#E8EEE7; --panel:#F3F6F1;
  --ink:#0B241B; --ink-2:#46594F; --mute:#9DB8AB; --yellow:#F5C93B;
  --line-d:rgba(157,184,171,.22); --line-l:#C3CEC2;
  --serif:'Source Serif 4','Source Serif Pro',Charter,'Iowan Old Style','Palatino Linotype',Georgia,serif;
  --sans:'Hanken Grotesk','Helvetica Neue',Helvetica,Arial,sans-serif;
  --mono:'JetBrains Mono','SF Mono',Menlo,Consolas,'Liberation Mono',monospace;
}
*{box-sizing:border-box;margin:0;padding:0}
html,body{width:1280px;height:720px;overflow:hidden;background:var(--paper)}
body{-webkit-font-smoothing:antialiased;text-rendering:geometricPrecision;font-kerning:normal;font-variant-numeric:tabular-nums}
.art{position:relative;width:1280px;height:720px;overflow:hidden;background:var(--paper);color:var(--ink)}
.abs{position:absolute}
.dark{color:#EEF3EC;background:linear-gradient(160deg,#14503E 0%,#0B3D2E 46%,#082B20 100%)}
.grid-l{position:absolute;inset:0;background-image:linear-gradient(to right,rgba(11,61,46,.07) 1px,transparent 1px),linear-gradient(to bottom,rgba(11,61,46,.07) 1px,transparent 1px);background-size:80px 80px}
.grid-d{position:absolute;inset:0;background-image:linear-gradient(to right,rgba(157,184,171,.10) 1px,transparent 1px),linear-gradient(to bottom,rgba(157,184,171,.10) 1px,transparent 1px);background-size:80px 80px}
.eyebrow{font:500 12px/16px var(--mono);letter-spacing:.18em;text-transform:uppercase;white-space:nowrap}
.serif{font-family:var(--serif);font-weight:600;letter-spacing:-.03em;font-optical-sizing:auto}
.foot{position:absolute;top:632px;font:500 10.5px/16px var(--mono);letter-spacing:.08em;text-transform:uppercase;white-space:nowrap}
.foot.l{left:160px}.foot.r{right:160px}
.on-paper .foot{color:var(--ink-2)}.on-dark .foot{color:var(--mute)}
"""


def page(title, comment, css, body, dark=False):
    cls = "art dark on-dark" if dark else "art on-paper"
    return f"""<!--
{comment}
-->
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=1280">
<title>{esc(title)}</title>
<style>
{FONTS}
{BASE_CSS}
{css}
</style>
</head>
<body>
<div class="{cls}">
{body}
</div>
</body>
</html>
"""


def comment_block(no, role, provenance, system):
    return (f"  UPWORK PORTFOLIO IMAGE {no:02d} / 06: {role}\n"
            f"  GitHub link to the real project: {GITHUB}\n"
            f"  (the repository created by the push command in portfolio-entry-sheet.md)\n\n"
            f"  Artboard 1280 x 720 CSS px (16:9). Export 3840 x 2160 px at device scale factor 3.\n"
            f"  System: {system}\n"
            f"  Provenance: {provenance}\n"
            f"  All figures are read from sample_output/run_report.json by build_boards.py. Synthetic demo data.")


FOOT_L = "Synthetic demo data, generated for this project"


def foot(no):
    return f'<div class="foot l">{FOOT_L}</div><div class="foot r">{no:02d} / 06</div>'


boards = []

# ================================================================== 01 COVER: asymmetric 5 | 7 split
region_rows = "".join(
    f'<div class="row"><i class="rn">{5 + i}</i><b class="c a">{g}</b><b class="c b">{n(WB["by_region"][g]["orders"])}</b>'
    f'<b class="c c2">{n(WB["by_region"][g]["units"])}</b><b class="c d{" sel" if i == 0 else ""}">{usd(WB["by_region"][g]["revenue"], False)}</b></div>'
    for i, g in enumerate(REGIONS))
tot_orders, tot_units = WB["orders"], WB["units"]
last = WB["last_row"]
f_line1 = f"=SUMIFS('Clean Data'!$J$2:$J${last},"
f_line2 = f"'Clean Data'!$C$2:$C${last},$A5)"
css = """
.panel{position:absolute;left:560px;top:0;width:720px;height:720px;background:linear-gradient(160deg,#14503E 0%,#0B3D2E 46%,#082B20 100%)}
.pgrid{position:absolute;left:560px;top:0;width:720px;height:720px;background-image:linear-gradient(to right,rgba(157,184,171,.10) 1px,transparent 1px),linear-gradient(to bottom,rgba(157,184,171,.10) 1px,transparent 1px);background-size:80px 80px}
.eyebrow{left:160px;top:96px;color:var(--ink-2)}
.h{left:160px;top:148px;font-size:84px;line-height:88px;white-space:nowrap}
.h span{display:block}
.sub{left:160px;top:456px;width:370px;font:400 18px/27px var(--sans);color:var(--ink-2)}
.card{left:640px;top:150px;width:480px;height:398px;background:#0A3427;border:1px solid #2F6B56;color:#EEF3EC;box-shadow:0 38px 56px -34px rgba(0,0,0,.55)}
.fbar{display:flex;height:56px;border-bottom:1px solid #2F6B56}
.nb{width:56px;border-right:1px solid #2F6B56;font:500 12px/56px var(--mono);text-align:center;color:var(--mute)}
.fx{width:36px;border-right:1px solid #2F6B56;font:italic 600 15px/56px var(--serif);text-align:center;color:var(--yellow)}
.ft{padding:10px 12px;font:400 11px/18px var(--mono);color:#EEF3EC;white-space:nowrap}
.cols,.row{display:flex}
.cols{height:28px;background:#0E4A3A;border-bottom:1px solid #2F6B56}
.cols i{font:500 11px/28px var(--mono);text-align:center;color:var(--mute);border-left:1px solid #2F6B56;font-style:normal}
.cols i.g{width:36px;border-left:0}.cols i.a{width:112px}.cols i.b{width:88px}.cols i.c2{width:88px}.cols i.d{width:156px;color:var(--ink);background:var(--yellow);border-left-color:var(--yellow)}
.row{height:40px;border-bottom:1px solid rgba(47,107,86,.7)}
.row .rn{width:36px;font:500 11px/40px var(--mono);text-align:center;color:var(--mute);background:#0E4A3A;font-style:normal;border-right:1px solid #2F6B56}
.row .c{font:400 13px/40px var(--mono);padding:0 12px;border-right:1px solid rgba(47,107,86,.7);font-weight:400}
.row .a{width:112px}.row .b{width:88px;text-align:right}.row .c2{width:88px;text-align:right}.row .d{width:156px;text-align:right;border-right:0}
.row.head .c{font-weight:700;color:var(--mute);font-size:11px;letter-spacing:.06em;text-transform:uppercase}
.row .d.sel{background:var(--yellow);color:var(--ink);font-weight:700;box-shadow:inset 0 0 0 2px var(--ink)}
.row.total{border-top:1.5px solid #EEF3EC;border-bottom:0}
.row.total .c{font-weight:700}
.tabs{position:absolute;left:0;right:0;bottom:0;height:34px;display:flex;background:#0E4A3A;border-top:1px solid #2F6B56}
.tabs span{font:500 11px/33px var(--mono);padding:0 14px;color:var(--mute);border-right:1px solid #2F6B56}
.tabs span.on{background:#0A3427;color:#EEF3EC;box-shadow:inset 0 -3px 0 var(--yellow)}
.cap{left:640px;top:122px;font:500 10.5px/16px var(--mono);letter-spacing:.14em;text-transform:uppercase;color:var(--mute)}
.on-paper .foot.r{color:var(--mute)}
"""
body = f"""
<div class="grid-l"></div><div class="panel"></div><div class="pgrid"></div>
<div class="abs eyebrow">Excel report automation</div>
<h1 class="abs h serif" data-essential><span>{REP['files']['count']} exports.</span><span>One Excel</span><span>report.</span></h1>
<p class="abs sub">Messy monthly sales files, cleaned and reported by one script. Re-run it when next month's file arrives.</p>
<div class="abs cap">Report sheet: By Region</div>
<div class="abs card" data-essential>
  <div class="fbar"><div class="nb">D5</div><div class="fx">fx</div><div class="ft">{esc(f_line1)}<br>{esc(f_line2)}</div></div>
  <div class="cols"><i class="g"></i><i class="a">A</i><i class="b">B</i><i class="c2">C</i><i class="d">D</i></div>
  <div class="row head"><i class="rn">4</i><b class="c a">Region</b><b class="c b">Orders</b><b class="c c2">Units</b><b class="c d">Revenue</b></div>
  {region_rows}
  <div class="row total"><i class="rn">10</i><b class="c a">Total</b><b class="c b">{n(tot_orders)}</b><b class="c c2">{n(tot_units)}</b><b class="c d">{usd(TOT_REV)}</b></div>
  <div class="tabs"><span>Summary</span><span class="on">By Region</span><span>Monthly</span><span>Clean Data</span><span>Cleaning Log</span></div>
</div>
<div class="foot l" style="color:var(--ink-2)">{FOOT_L}</div><div class="foot r" style="color:var(--mute)">01 / 06</div>
"""
boards.append(("01-excel-report-automation-python-cover", "Excel Report Automation: cover",
               comment_block(1, "COVER (asymmetric 5 | 7 split)",
                             "files.count, cleaning.rows_kept via workbook.by_region, workbook.revenue, the SUMIFS formula text of By Region!D5",
                             "80 px module (16 x 9 grid); left text column x 160-540, sheet card x 640-1120 on a grid line; all essentials inside x 160-1120."),
               css, body, False))

# ================================================================== 02 THE MESSY EXPORTS: specimen strips
def split_raw(line, delim):
    out, cur, q = [], "", False
    for ch in line:
        if ch == '"':
            q = not q
        if ch == delim and not q:
            out.append(cur)
            cur = ""
        else:
            cur += ch
    return out + [cur]


def strip_html(panel):
    fname = panel["file"]
    delim = ";" if panel["lines"][0].count(";") > panel["lines"][0].count(",") else ","
    rows = []
    for i, ln in enumerate(panel["lines"][:2]):
        toks = split_raw(ln, delim)
        if i == 1:
            toks = [f"<mark>{esc(t)}</mark>" if k in (1, 2, 9) and t else esc(t) for k, t in enumerate(toks)]
        else:
            toks = [esc(t) for t in toks]
        text = esc(delim).join(toks)
        rows.append(f'<div class="ln hd"><span>{text}</span><em class="fn">{esc(fname)}</em></div>' if i == 0
                    else f'<div class="ln">{text}</div>')
    return f'<div class="strip">{"".join(rows)}</div>'


strips = "".join(strip_html(p) for p in RAW["panels"])
d = RAW


def defect(label, info, text):
    return (f'<div class="df"><span class="dl">{label}</span><span class="dt">{esc(text)}</span>'
            f'<span class="dm">{esc(info)}</span></div>')


same_file = len({d["blank_row"]["file"], d["footer"]["file"], d["duplicate"]["file"]}) == 1
pre = lambda x: "" if same_file else x["file"] + ", "  # noqa: E731
defects = (f'<div class="dh">Defects found in {esc(d["blank_row"]["file"])}</div>' if same_file else "") + "".join([
    defect("Blank row", f'{pre(d["blank_row"])}line {d["blank_row"]["line_no"]}', d["blank_row"]["text"]),
    defect("Footer row", f'{pre(d["footer"])}line {d["footer"]["line_no"]}', d["footer"]["text"]),
    defect("Duplicate", f'{pre(d["duplicate"])}lines {d["duplicate"]["line_nos"][0]} and {d["duplicate"]["line_nos"][1]}',
           d["duplicate"]["text"]),
])
css = """
.grid-l{background-image:linear-gradient(to bottom,rgba(11,61,46,.07) 1px,transparent 1px);background-size:80px 80px}
.eyebrow{left:160px;top:88px;color:var(--ink-2)}
.h{left:160px;top:120px;font-size:62px;line-height:66px;white-space:nowrap}
.h span{display:block}
.strips{left:160px;top:278px;width:960px}
.strip{background:var(--panel);border:1px solid var(--line-l);padding:8px 16px;margin-bottom:10px;box-shadow:0 1px 0 #fff inset}
.ln{font:400 12px/21px var(--mono);white-space:nowrap;color:var(--ink)}
.ln.hd{color:var(--ink-2);font-weight:700;display:flex;justify-content:space-between}
.fn{font-style:normal;font-weight:600;font-size:10.5px;letter-spacing:.06em;color:var(--ink-2)}
mark{background:rgba(245,201,59,.78);color:var(--ink);padding:1px 2px;margin:0 -2px}
.dfs{left:160px;top:498px;width:960px}
.dh{height:26px;border-top:1.5px solid var(--ink);border-bottom:1px solid var(--line-l);font:500 10.5px/26px var(--mono);letter-spacing:.14em;text-transform:uppercase;color:var(--ink-2)}
.df{display:flex;align-items:center;height:26px;border-bottom:1px solid var(--line-l);font:400 11px/26px var(--mono);white-space:nowrap}
.dl{width:92px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;font-size:10.5px}
.dt{flex:1;overflow:hidden;color:var(--ink)}
.dm{margin-left:16px;color:var(--ink-2);font-size:10.5px}
"""
body = f"""
<div class="grid-l"></div>
<div class="abs eyebrow">{REP['files']['count']} files, {REP['files']['header_styles']} header spellings, {REP['files']['date_styles']} date styles</div>
<h1 class="abs h serif" data-essential><span>Every export,</span><span>formatted differently.</span></h1>
<div class="abs strips" data-essential>{strips}</div>
<div class="abs dfs">{defects}</div>
{foot(2)}
"""
boards.append(("02-messy-excel-exports-cleaned-csv-xlsx", "Messy Excel Exports, Cleaned",
               comment_block(2, "THE TASK (specimen strips)",
                             "files.count, files.header_styles, files.date_styles; raw lines quoted verbatim from raw_examples.json (three export files plus a blank, a footer and a duplicate row)",
                             "12-column grid, three 960 px strips on a 12 px rhythm; highlighted cells mark date, region and amount."),
               css, body, False))

# ================================================================== 03 FIX LEDGER: golden ratio, dark
ledger_items = [("Text numbers converted", CL["text_numbers_converted"]), ("Region labels standardized", CL["region_labels_standardized"]),
                ("Dates normalized", CL["dates_normalized"]), ("Padded cells trimmed", CL["cells_trimmed"]),
                ("Duplicate rows removed", CL["duplicates_removed"]), ("Missing amounts recomputed", CL["amounts_recomputed"]),
                ("Blank rows removed", CL["blank_rows"]), ("Rows rejected, no order ID", CL["missing_id_rejected"]),
                ("Footer rows removed", CL["footer_rows"])]
ledger_items.sort(key=lambda x: -x[1])
vmax = max(v for _, v in ledger_items)
TRACK = 250
rows = "".join(
    f'<div class="lr"><span class="ll">{lab}</span><span class="lt"><i style="width:{max(1.0, v / vmax * TRACK):.2f}px"></i></span>'
    f'<span class="lv">{n(v)}</span></div>' for lab, v in ledger_items)
assert sum(v for _, v in ledger_items) == CL["fixes_total"]
css = """
.eyebrow{left:160px;top:96px;color:var(--mute)}
.h{left:160px;top:124px;font-size:50px;line-height:54px;white-space:nowrap}
.h span{display:block}
.hero{left:160px;top:272px;font:700 112px/112px var(--serif);letter-spacing:-.04em;color:var(--yellow);white-space:nowrap}
.hlab{left:160px;top:398px;width:340px;font:500 18px/26px var(--sans);color:#EEF3EC}
.hcap{left:160px;top:456px;width:340px;font:400 14px/22px var(--sans);color:var(--mute)}
.led{left:560px;top:112px;width:560px}
.lh{display:flex;height:28px;border-bottom:1.5px solid #EEF3EC;font:500 10.5px/28px var(--mono);letter-spacing:.14em;text-transform:uppercase;color:var(--mute)}
.lh .ll{width:230px}.lh .lt{width:262px}.lh .lv{width:68px;text-align:right}
.lr{display:flex;align-items:center;height:42px;border-bottom:1px solid var(--line-d)}
.ll{width:230px;font:500 15px/20px var(--sans);color:#EEF3EC;white-space:nowrap}
.lt{width:262px;height:42px;display:flex;align-items:center}
.lt i{display:block;height:14px;background:var(--mute)}
.lv{width:68px;text-align:right;font:500 15px/20px var(--mono);color:#EEF3EC}
.ltot{display:flex;align-items:center;height:46px;border-top:1.5px solid #EEF3EC;margin-top:-1px;font:700 15px/20px var(--mono);color:#EEF3EC}
.ltot .ll{font:600 13px/20px var(--mono);letter-spacing:.14em;text-transform:uppercase;width:492px}
"""
body = f"""
<div class="grid-d"></div>
<div class="abs eyebrow">Cleaning log, {REP['files']['count']} export files</div>
<h1 class="abs h serif" data-essential><span>Every fix,</span><span>logged.</span></h1>
<div class="abs hero" data-essential>{n(CL['fixes_total'])}</div>
<div class="abs hlab">fixes recorded, one count per type per file</div>
<div class="abs hcap">{n(CL['rows_read'])} raw rows read. {n(CL['rows_kept'])} orders kept.<br>Bars start at zero and share one scale.</div>
<div class="abs led" data-essential>
  <div class="lh"><span class="ll">Fix type</span><span class="lt"></span><span class="lv">Count</span></div>
  {rows}
  <div class="ltot"><span class="ll">Total</span><span class="lv" style="width:68px">{n(CL['fixes_total'])}</span></div>
</div>
{foot(3)}
"""
boards.append(("03-excel-data-cleaning-log-fixes", "Excel Data Cleaning: every fix logged",
               comment_block(3, "THE DATA (fix ledger, golden-ratio split)",
                             "cleaning.* counts (nine fix types, fixes_total, rows_read, rows_kept); bar widths = count / largest count, zero-based",
                             "Hero column 38.2% | ledger 61.8% of the 960 px content width; 42 px ledger rows."),
               css, body, True))

# ================================================================== 04 DASHBOARD: 3 x 2 modules
months = WB["months"]
mt = WB["monthly_total"]
mmax = 800000.0
PLOT_H, PLOT_W = 214, 480
best_i = max(range(len(mt)), key=lambda i: mt[i])
bw = PLOT_W / len(mt)
grid_lines = "".join(
    f'<div class="gl" style="bottom:{v / mmax * PLOT_H:.1f}px"><span>{("$0" if v == 0 else "$" + str(int(v / 1000)) + "k")}</span></div>'
    for v in (0, 200000, 400000, 600000, 800000))
mlet = "JFMAMJJASOND"
bars = "".join(
    f'<div class="bar{" best" if i == best_i else ""}" style="left:{i * bw + bw * 0.16:.1f}px;width:{bw * 0.68:.1f}px;height:{v / mmax * PLOT_H:.1f}px"></div>'
    f'<div class="ml" style="left:{i * bw:.1f}px;width:{bw:.1f}px">{mlet[i]}</div>' for i, v in enumerate(mt))
bx = best_i * bw + bw / 2
best_lab = f"{['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][best_i]} {usd(mt[best_i])}"
rmax = max(WB["by_region"][g]["revenue"] for g in REGIONS)
rrows = "".join(
    f'<div class="rr"><span class="rg">{g}</span><span class="rb"><i style="width:{WB["by_region"][g]["revenue"] / rmax * 100:.1f}%"></i></span>'
    f'<span class="rv">{usd(WB["by_region"][g]["revenue"], False)}</span><span class="rp">{WB["by_region"][g]["revenue"] / TOT_REV * 100:.1f}%</span></div>'
    for g in REGIONS)
css = """
.grid-l{background-image:none}
.eyebrow{left:160px;top:92px;color:var(--ink-2)}
.h{left:160px;top:120px;font-size:50px;line-height:54px;white-space:nowrap}
.h span{display:block}.h em{font-style:italic}
.kpis{left:700px;top:116px;width:420px;display:flex}
.kpi{flex:1;padding:0 0 0 16px;border-left:1px solid var(--line-l)}
.kpi:first-child{flex:1.75}
.kl{font:500 10px/16px var(--mono);letter-spacing:.14em;text-transform:uppercase;color:var(--ink-2)}
.kv{font:600 24px/38px var(--serif);letter-spacing:-.02em;margin-top:4px;white-space:nowrap}
.pan{position:absolute;background:var(--panel);border:1px solid var(--line-l);box-shadow:0 24px 40px -30px rgba(11,61,46,.4)}
.pt{position:absolute;left:20px;top:14px;font:500 10.5px/16px var(--mono);letter-spacing:.14em;text-transform:uppercase;color:var(--ink-2)}
.chart{left:160px;top:262px;width:560px;height:350px}
.plot{position:absolute;left:56px;top:52px;width:480px;height:214px}
.gl{position:absolute;left:0;right:0;border-top:1px solid var(--line-l);height:0}
.gl span{position:absolute;left:-46px;top:-8px;font:400 10px/16px var(--mono);color:var(--ink-2);width:40px;text-align:right}
.bar{position:absolute;bottom:0;background:var(--green)}
.bar.best{background:var(--yellow);box-shadow:inset 0 0 0 1.5px var(--green-deep)}
.ml{position:absolute;bottom:-24px;text-align:center;font:500 11px/16px var(--mono);color:var(--ink-2)}
.blab{position:absolute;font:700 11px/16px var(--mono);color:var(--ink);white-space:nowrap}
.cnote{position:absolute;left:20px;bottom:14px;font:400 10.5px/16px var(--mono);color:var(--ink-2)}
.reg{left:744px;top:262px;width:376px;height:350px}
.rl{position:absolute;left:20px;right:20px;top:46px}
.rr{display:flex;align-items:center;height:48px;border-bottom:1px solid var(--line-l)}
.rg{width:64px;font:600 14px/20px var(--sans)}
.rb{width:96px;height:12px;margin-right:10px}
.rb i{display:block;height:12px;background:var(--green)}
.rv{width:92px;text-align:right;font:500 12px/20px var(--mono)}
.rp{width:50px;margin-left:10px;text-align:right;font:400 12px/20px var(--mono);color:var(--ink-2)}
.rt{display:flex;align-items:center;height:44px;border-top:1.5px solid var(--ink);margin-top:-1px;font:700 12px/20px var(--mono)}
.rt .rg{font:700 14px/20px var(--sans)}
.rt .rb{width:82px}
"""
body = f"""
<div class="grid-l"></div>
<div class="abs eyebrow">Monthly and regional sales, live formulas</div>
<h1 class="abs h serif" data-essential><span>Excel dashboard,</span><span><em>built by script.</em></span></h1>
<div class="abs kpis">
  <div class="kpi"><div class="kl">Total revenue</div><div class="kv">{usd(TOT_REV)}</div></div>
  <div class="kpi"><div class="kl">Orders</div><div class="kv">{n(WB['orders'])}</div></div>
  <div class="kpi"><div class="kl">Average order</div><div class="kv">{usd(TOT_REV / WB['orders'])}</div></div>
</div>
<div class="pan chart" data-essential><div class="pt">Revenue by month, {months[0][:4]}</div>
  <div class="plot">{grid_lines}{bars}<div class="blab" style="left:{bx - 4:.1f}px;bottom:{mt[best_i] / mmax * PLOT_H + 6:.1f}px;transform:translateX(-50%)">{best_lab}</div></div>
  <div class="cnote">Zero-based axis. Highlighted: the highest month.</div></div>
<div class="pan reg" data-essential><div class="pt">Revenue by region</div>
  <div class="rl">{rrows}<div class="rt"><span class="rg">Total</span><span class="rb"></span><span class="rv" style="width:108px">{usd(TOT_REV)}</span><span class="rp">100.0%</span></div></div></div>
{foot(4)}
"""
boards.append(("04-excel-dashboard-automation-monthly-regional-sales", "Excel Dashboard Automation: monthly and regional sales",
               comment_block(4, "THE DELIVERABLE (3 x 2 modular dashboard)",
                             "workbook.monthly_total (12 values), workbook.by_region, workbook.revenue, workbook.orders; average order = revenue / orders as in the Summary sheet",
                             "Two rows by three modules; chart panel 7 modules, region panel 4.7; bar heights = value / 800,000 on a zero-based axis."),
               css, body, False))

# ================================================================== 05 THE CODE: framed and layered, dark
start = next(i for i, ln in enumerate(SRC) if ln.startswith("def monthly_formula"))
end = start
while end + 1 < len(SRC) and SRC[end + 1].strip():
    end += 1
code_lines = [(start + 1 + k, SRC[start + k]) for k in range(end - start + 1)]
TOKEN = re.compile(r'(?P<doc>""".*?""")|(?P<str>[fF]?"[^"]*"|[fF]?\'[^\']*\')|(?P<kw>\b(?:def|return)\b)|(?P<fn>\b[a-z_]+(?=\())|(?P<num>\b\d+\b)')


def highlight(line):
    out, pos = "", 0
    for m in TOKEN.finditer(line):
        out += esc(line[pos:m.start()])
        kind = m.lastgroup
        out += f'<span class="t-{kind}">{esc(m.group(0))}</span>'
        pos = m.end()
    return out + esc(line[pos:])


code_html = "".join(f'<div class="cl"><span class="no">{no}</span><span class="cc">{highlight(ln)}</span></div>' for no, ln in code_lines)
T = REP["tools"]
lo_ver = re.search(r"LibreOffice\s+([\d.]+)", T["libreoffice"]).group(1)
stack = [f"Python {T['python']}", f"pandas {T['pandas']}", f"XlsxWriter {T['xlsxwriter']}", f"openpyxl {T['openpyxl']}",
         f"LibreOffice {lo_ver}"]
stack_html = "".join(f'<div class="sk"><b>{esc(s.split()[0])}</b><span>{esc(s.split()[1])}</span></div>' for s in stack)
n_cells = len(months) * len(REGIONS)
def wrap_args(text, width=64):
    """Wrap a formula at argument boundaries (after commas) so no sheet name or range is split."""
    lines, cur = [], ""
    for chunk in re.split(r"(?<=,)", text):
        if cur and len(cur) + len(chunk) > width:
            lines.append(cur)
            cur = ""
        cur += chunk
    return "\n".join(lines + [cur])


res_text = wrap_args(build_report.monthly_formula(WB["last_row"], 5, "B"))
css = """
.eyebrow{left:160px;top:96px;color:var(--mute)}
.h{left:160px;top:124px;font-size:44px;line-height:48px;white-space:nowrap}
.h span{display:block}
.stack{left:160px;top:300px;width:340px}
.sk{display:flex;justify-content:space-between;height:34px;border-bottom:1px solid var(--line-d);font:400 13px/34px var(--mono)}
.sk:first-child{border-top:1px solid var(--line-d)}
.sk b{font-weight:500;color:#EEF3EC}.sk span{color:var(--mute)}
.stl{left:160px;top:274px;font:500 10.5px/16px var(--mono);letter-spacing:.14em;text-transform:uppercase;color:var(--mute)}
.g1,.g2{position:absolute;border:1px solid rgba(157,184,171,.30);background:rgba(8,43,32,.55)}
.g1{left:592px;top:134px;width:540px;height:362px}
.g2{left:586px;top:128px;width:540px;height:362px;border-color:rgba(157,184,171,.42)}
.win{left:580px;top:116px;width:540px;height:362px;background:#061F17;border:1px solid #3B7A63;box-shadow:0 40px 60px -34px rgba(0,0,0,.65)}
.wb{height:40px;background:#0E4A3A;border-bottom:1px solid #2F6B56;position:relative}
.wb i{position:absolute;top:15px;width:10px;height:10px;border-radius:50%;background:#3B7A63}
.wb em{position:absolute;left:0;right:0;top:0;text-align:center;font:500 12px/40px var(--mono);color:var(--mute);font-style:normal}
.code{padding:22px 0 0}
.cl{display:flex;height:26px;font:400 12px/26px var(--mono);white-space:pre;font-variant-ligatures:none;font-feature-settings:'liga' 0,'calt' 0}
.no{width:46px;text-align:right;padding-right:14px;color:#5E8576;flex:none}
.cc{color:#DCE6DD}
.t-kw{color:var(--yellow);font-weight:700}.t-fn{color:#FFFFFF;font-weight:700}.t-str{color:#A9D4BE}.t-doc{color:#7FA392}.t-num{color:var(--yellow)}
.res{border-top:1px solid #2F6B56;margin:14px 0 0;padding:12px 20px 0 46px;font:400 10.5px/17px var(--mono);color:#A9D4BE;white-space:pre;font-variant-ligatures:none}
.res b{display:block;font:500 10px/16px var(--mono);letter-spacing:.14em;text-transform:uppercase;color:var(--mute);margin-bottom:4px;margin-left:-26px}
.note{left:580px;top:500px;width:540px;font:500 15px/22px var(--sans);color:#EEF3EC}
.note b{color:var(--yellow);font-weight:600}
.note small{display:block;margin-top:4px;font:400 12px/18px var(--mono);color:var(--mute)}
"""
body = f"""
<div class="grid-d"></div>
<div class="abs eyebrow">Python, pandas, XlsxWriter</div>
<h1 class="abs h serif" data-essential><span>Excel automation</span><span>in plain Python.</span></h1>
<div class="abs stl">Tools used</div>
<div class="abs stack">{stack_html}</div>
<div class="abs g1"></div><div class="abs g2"></div>
<div class="abs win" data-essential><div class="wb"><i style="left:16px"></i><i style="left:32px"></i><i style="left:48px"></i><em>build_report.py</em></div>
  <div class="code">{code_html}</div>
  <div class="res"><b>Returns for Monthly!B5</b>{esc(res_text)}</div></div>
<div class="abs note"><b>{n_cells} of the {n(VF['formulas'])} formulas</b> in the report are written by this one function.<small>{len(months)} months x {len(REGIONS)} regions, one SUMIFS per cell</small></div>
{foot(5)}
"""
boards.append(("05-excel-automation-python-xlsxwriter-code", "Excel Automation in Python: the code",
               comment_block(5, "THE TOOLKIT (framed, layered code window)",
                             f"tools.* versions; verification.formulas; months x regions = {n_cells}; code lines {code_lines[0][0]}-{code_lines[-1][0]} quoted verbatim from build_report.py",
                             "Two ghost frames offset by 12 px behind the window; window x 540-1096, left column x 160-460."),
               css, body, True))

# ================================================================== 06 VERIFIED: monolith on yellow
checks = [("Formula mismatches", str(VF["mismatches"])), ("Excel error values", str(VF["errors"])),
          ("Orders vs ground truth", f"{n(CL['ground_truth']['cleaned_orders'])} / {n(CL['ground_truth']['true_orders'])}"),
          ("Self-test: planted error caught", "Yes" if VF["selftest_caught"] else "No"),
          ("13th month: formulas verified", f"{RF['formulas']} / {RF['formulas']}")]
check_html = "".join(f'<div class="ck"><span class="cl2">{a}</span><span class="cv">{b}</span></div>' for a, b in checks)
assert VF["mismatches"] == 0 and VF["errors"] == 0 and RF["mismatches"] == 0 and RF["errors"] == 0
css = """
.grid-l{background-image:linear-gradient(to right,rgba(11,61,46,.06) 1px,transparent 1px),linear-gradient(to bottom,rgba(11,61,46,.06) 1px,transparent 1px)}
.eyebrow{left:160px;top:96px;color:var(--ink-2)}
.mono{left:160px;top:128px;width:640px;height:208px;background:var(--yellow)}
.big{left:176px;top:136px;font:700 168px/192px var(--serif);letter-spacing:-.05em;color:var(--green-deep);white-space:nowrap}
.h{left:160px;top:364px;font-size:40px;line-height:44px;white-space:nowrap}
.sub{left:160px;top:420px;width:560px;font:400 18px/27px var(--sans);color:var(--ink-2)}
.cks{left:844px;top:128px;width:276px;border-top:1.5px solid var(--ink)}
.ck{height:62px;border-bottom:1px solid var(--line-l)}
.cl2{display:block;font:500 10.5px/16px var(--mono);letter-spacing:.12em;text-transform:uppercase;color:var(--ink-2);padding-top:9px}
.cv{display:block;font:600 26px/32px var(--serif);letter-spacing:-.02em}
.band{left:160px;top:536px;width:960px;border-top:1.5px solid var(--ink);padding-top:12px;display:flex;justify-content:space-between;font:500 14px/22px var(--sans)}
.band b{font-weight:600}.band span.m{font:500 13px/22px var(--mono);color:var(--ink-2)}
"""
body = f"""
<div class="grid-l"></div>
<div class="abs eyebrow">Independent recalculation</div>
<div class="abs mono" data-essential></div>
<div class="abs big" data-essential>{VF['formulas']}/{VF['formulas']}</div>
<h1 class="abs h serif" data-essential>Every formula, checked.</h1>
<p class="abs sub">Each formula in the workbook was recalculated in {esc(lo_ver and 'LibreOffice ' + lo_ver)} and compared with the value stored in the file.</p>
<div class="abs cks" data-essential>{check_html}</div>
<div class="abs band"><span>When a 13th monthly export was added, the report rebuilt in <b>{RF['pipeline_seconds']} s</b>.</span><span class="m">Monthly sheet: {len(months)} to {RF['months_after']} rows</span></div>
{foot(6)}
"""
boards.append(("06-excel-formulas-verified-libreoffice", "Excel Formulas Verified",
               comment_block(6, "THE PROOF (typographic monolith)",
                             "verification.formulas, mismatches, errors, selftest_caught, tools.libreoffice; cleaning.ground_truth; refresh.formulas, refresh.pipeline_seconds, refresh.months_after",
                             "Yellow monolith 640 x 208 at x 160 (a grid line); check ledger in a 276 px column on the right."),
               css, body, False))

# ------------------------------------------------------------------ write
if __name__ == "__main__":
    for name, title, comment, css_, body_, dark in boards:
        path = os.path.join(HERE, name + ".html")
        open(path, "w", encoding="utf-8").write(page(title, comment, css_, body_, dark))
        print("wrote", os.path.basename(path), f"{os.path.getsize(path) / 1024:.0f} KB")
    pairs = [("ink on paper", "ink", "paper"), ("ink-2 on paper", "ink_2", "paper"), ("ink-2 on panel", "ink_2", "panel"),
             ("paper on green", "paper", "green"), ("mute on green", "mute", "green"), ("mute on green-deep", "mute", "green_deep"),
             ("yellow on green", "yellow", "green"), ("ink on yellow", "ink", "yellow"), ("green-deep on yellow", "green_deep", "yellow")]
    lines = [f"{a:24s} {contrast(PALETTE[b], PALETTE[c]):5.1f} : 1" for a, b, c in pairs]
    open(os.path.join(HERE, "contrast.txt"), "w").write("\n".join(lines) + "\n")
    print("\n".join(lines))
