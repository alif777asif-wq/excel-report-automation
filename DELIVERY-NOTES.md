# Excel report automation: six Upwork portfolio images, delivery notes

## What I had, and what I did about it

The three placeholders in your brief ({enter your subject}, {enter your real project data}, {enter your GitHub link}) were still blank in the text I received. That is why the first run produced empty typographic slides: there was nothing real to put in them. The brief also forbids inventing anything, so I did not fill the gap with made-up figures.

Instead I built a real, working Excel automation and measured it. Every number, table, raw data line and code line on the six images comes from that run, and each image says so in a quiet caption: "Synthetic demo data, generated for this project". This is **not client work**, and the images do not claim it is. If you have a real project and permission to show it, send the figures and I will rebuild the boards around them.

Two things this project does **not** contain, so no image claims them:
- **VBA and macros.** I cannot run Excel here, so I cannot verify a VBA run. The automation is Python that writes an Excel workbook. Your Fiverr gig sells macros and VBA; this item shows Python-driven Excel automation, which matches your Upwork headline.
- **Pivot tables.** The report uses SUMIFS summary tables and charts, which I could verify. They are called summaries, never pivot tables.

## The project behind the images (all measured)

| Item | Result |
|---|---|
| Inputs | 12 monthly sales exports (8 CSV, 4 XLSX), 3 header spellings, 4 date styles |
| Raw rows read | 11,716 |
| Orders kept | 11,442, all matching the ground truth the exports were generated from (0 wrong fields) |
| Fixes logged | 31,110 across nine types; the cleaning log equals the injected problems in 12 of 12 files |
| Report | 5 sheets, 0.64 MB, total revenue 7,496,987.86 |
| Pipeline time | 1.72 s (clean 0.72 s, build 1.01 s) |
| Verification | 149 formulas recalculated in LibreOffice 24.2.7.2: 0 mismatches, 0 errors, largest difference 4.5e-13 |
| Verifier self-test | a deliberately wrong cached value was caught: True |
| Refresh | a 13th export added, report rebuilt in 2.05 s, 157 formulas verified, 0 mismatches |

The verification step first forces LibreOffice to recalculate the workbook, because it otherwise trusts the stored values of Excel-looking files. The self-test proves the check can fail.

## Design system

- Artboard 1280 x 720 CSS px (16:9) on an 80 px module (16 x 9 grid). Six different proportional structures: 5 | 7 asymmetric split, stacked specimen strips, golden-ratio 38.2 | 61.8, a two-by-three module dashboard, a framed and layered window, and a single typographic monolith.
- Type: Source Serif 4 for headlines, Hanken Grotesk for supporting text, JetBrains Mono for data and code. All three are embedded (SIL Open Font License), so the files render identically offline.
- Palette: deep green, sage paper and one yellow accent, the same family as your Fiverr banners. **Reasoning:** green signals accuracy and stability to someone handing over their reporting, sage paper keeps it calm and document-like, and a single yellow marks the one place the eye should land on each board.

| Pair | Contrast |
|---|---|
| ink on paper | 13.9 : 1 |
| ink-2 on paper | 6.4 : 1 |
| ink-2 on panel | 6.9 : 1 |
| paper on green | 10.4 : 1 |
| mute on green | 5.7 : 1 |
| mute on green-deep | 7.2 : 1 |
| yellow on green | 7.7 : 1 |
| ink on yellow | 10.4 : 1 |
| green-deep on yellow | 9.7 : 1 |

## The six images

**01: 12 exports. One Excel report.** · `01-excel-report-automation-python-cover`
- Rationale: An asymmetric 5 | 7 split sets the promise in plain serif type beside the real By Region sheet, with one selected cell showing the formula behind the number.
- Provenance: files.count (12); By Region values (orders, units, revenue per region and totals); the SUMIFS text of By Region!D5 as written by build_report.py.

**02: Every export, formatted differently.** · `02-messy-excel-exports-cleaned-csv-xlsx`
- Rationale: Three verbatim raw export lines, each with its own header spelling and date style, make the task visible, and the highlighted cells are exactly what the script has to fix.
- Provenance: files.count, header_styles (3), date_styles (4); raw lines quoted from raw_examples.json (three export files plus one blank, one footer and one duplicate row).

**03: Every fix, logged.** · `03-excel-data-cleaning-log-fixes`
- Rationale: A golden-ratio split pairs one dominant real figure with a zero-based ledger of every fix type, so the scale of the cleaning is the whole picture.
- Provenance: cleaning.* counts: nine fix types, fixes_total (31,110), rows_read (11,716), rows_kept (11,442).

**04: Excel dashboard, built by script.** · `04-excel-dashboard-automation-monthly-regional-sales`
- Rationale: A two-by-three module dashboard is drawn only from the workbook's own totals, with a zero-based monthly chart and a regional table that match the delivered file.
- Provenance: workbook.monthly_total (12 values), by_region, revenue (7,496,987.86), orders; average order = revenue / orders as in the Summary sheet.

**05: Excel automation in plain Python.** · `05-excel-automation-python-xlsxwriter-code`
- Rationale: A layered code window quotes the real function that writes 60 of the 149 formulas, with its actual output beneath it.
- Provenance: tools.* versions; verification.formulas (149); 12 months x 5 regions = 60; build_report.py lines 24-30 quoted verbatim; the function's real return value for Monthly!B5.

**06: Every formula, checked.** · `06-excel-formulas-verified-libreoffice`
- Rationale: A yellow monolith turns the verification result into one number that survives the thumbnail test, backed by a hairline ledger of the checks.
- Provenance: verification.formulas / mismatches / errors / selftest_caught; tools.libreoffice (24.2.7.2); cleaning.ground_truth (11,442 of 11,442); refresh.formulas (157), pipeline_seconds (2.05), months 12 to 13.

## Thumbnail test (about 300 px wide)

I rendered all six at 300 px wide side by side. The message reads on every board: "12 exports. One Excel report.", "Every export, formatted differently.", the 31,110 hero figure, "Excel dashboard, built by script." with the chart, "Excel automation in plain Python.", and the 149/149 monolith.

## Audit against the brief

| Requirement | Result |
|---|---|
| Exactly 16:9, authored 1280 x 720, exported 3840 x 2160 | Yes. All 12 files (6 PNG, 6 JPG) measured 3840 x 2160. |
| 300 DPI, sRGB | Yes. DPI tag and embedded sRGB profile verified on every file. |
| JPG highest quality, 4:4:4 | Quality 100, subsampling 4:4:4 verified. |
| Under 5 MB | All six JPG masters are under it (2.01, 1.58, 2.78, 1.02, 2.80 and 1.03 MB), so no reduced copy was needed. |
| Safe zones | Measured in the browser for every board: all text sits inside the 70 px margin and every headline and focal element sits inside the central 960 x 720 region. |
| HTML and CSS only, no JavaScript, images or URLs | Checked in all six files: no script, image, iframe, SVG, external URL or event handler. |
| No Upwork logos, ratings, badges, contact details, watermarks, website screenshots or borrowed imagery | None present. The word "Excel" appears as text only. |
| No invented figures | Every figure is generated from the run report by `boards/build_boards.py`. Nothing is typed by hand. |
| GitHub link carried verbatim in an HTML comment | Yes, in all six files, and in the entry sheet. |
| Zero-based charts | The monthly chart and every bar use a zero baseline. |

## Omissions (nothing was filled in)

- **Real client data:** none was supplied, so there is no real project figure on any board. The data is a measured demo, labelled as such.
- **GitHub link:** you did not supply one. The link in the comments and the entry sheet is the repository the push command creates, and it only exists after you push. The URL-free marker "Source code linked in project" is left out of the artwork for the same reason.
- **VBA, macros, pivot tables:** not claimed anywhere (see above).
- **Upwork's image limit:** I could not confirm the 5 MB figure from Upwork's pages, so I used the one in your brief.
- **The caption "Synthetic demo data":** it is a provenance label, not a watermark. I recommend keeping it, because the description says the same thing.

## Exporting at exactly 3840 x 2160

The finished files are in `boards/exports/`. To re-export:

1. Once: `pip install playwright pillow`, then `playwright install chromium`.
2. In the `boards` folder run `python3 export.py` (no internet needed, the fonts are embedded). Add a file name to export one board.
3. It writes a lossless PNG master and a JPG master at quality 100 with 4:4:4 chroma, both 3840 x 2160, 300 DPI, sRGB profile embedded, and prints the safe-zone audit.
4. If a JPG master ever exceeds 5 MB, it also writes `<name>-upload.jpg`: the same 3840 x 2160 re-encoded at the highest quality that fits under 4.9 MB. Upload that one.

By hand in Chrome DevTools: Ctrl+Shift+M, a 1280 x 720 viewport at device pixel ratio 3, then "Capture screenshot". That gives a 3840 x 2160 PNG but not the 300 DPI tag or the sRGB profile, so `export.py` is the better route.

To rebuild the HTML from the data: `python3 build_boards.py` in the `boards` folder.
