#!/usr/bin/env python3
"""
Export the six Upwork portfolio boards at exactly 3840 x 2160 px (4K UHD, 16:9).

The boards embed their fonts, so no internet connection is needed.

For every board it writes, into ./exports:
  <name>.png          lossless master, sRGB profile embedded, 300 DPI
  <name>.jpg          JPG master, quality 100, 4:4:4 chroma (no subsampling), sRGB profile, 300 DPI
  <name>-upload.jpg   ONLY when the JPG master is over the 5 MB limit: same 3840 x 2160,
                      re-encoded at the highest quality that fits under the limit

It also audits the safe zones in the browser and prints the result:
  * every piece of text sits inside the 70 px safe margin of the 1280 x 720 artboard
  * every element marked data-essential (headline, focal element) sits inside the central 960 x 720 region

Setup (once):   pip install playwright pillow   then   playwright install chromium
Run:            python3 export.py                  # all boards in this folder
                python3 export.py 03-*.html        # just some boards
"""
import argparse
import io
import sys
from pathlib import Path

from PIL import Image, ImageCms
from playwright.sync_api import sync_playwright

ARTBOARD_W, ARTBOARD_H = 1280, 720
SCALE = 3                                   # 1280 x 3 = 3840, 720 x 3 = 2160
OUT_SIZE = (ARTBOARD_W * SCALE, ARTBOARD_H * SCALE)
DPI = 300
UPLOAD_LIMIT_BYTES = 4_900_000              # safely under 5 MB whether MB means 1,000,000 or 1,048,576 bytes
SAFE = (70, 70, 1210, 650)                  # text must stay inside this box (artboard px)
CENTRAL = (160, 70, 1120, 650)              # headline and focal elements must stay inside this box
SRGB_PROFILE = ImageCms.ImageCmsProfile(ImageCms.createProfile("sRGB")).tobytes()

AUDIT_JS = """() => {
  const out = {text: [], essential: []};
  const w = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
  let node;
  while ((node = w.nextNode())) {
    const t = node.textContent.trim();
    if (!t) continue;
    const r = document.createRange(); r.selectNodeContents(node);
    for (const b of r.getClientRects()) if (b.width > 0 && b.height > 0) out.text.push([b.left, b.top, b.right, b.bottom, t.slice(0, 28)]);
  }
  document.querySelectorAll('[data-essential]').forEach(el => {
    const b = el.getBoundingClientRect();
    out.essential.push([b.left, b.top, b.right, b.bottom, (el.className || el.tagName).toString().slice(0, 20)]);
  });
  return out;
}"""


def audit(page, name):
    data = page.evaluate(AUDIT_JS)
    problems = []
    for l, t, r, b, label in data["text"]:
        if l < SAFE[0] - 0.5 or t < SAFE[1] - 0.5 or r > SAFE[2] + 0.5 or b > SAFE[3] + 0.5:
            problems.append(f"text outside safe margin: '{label}' ({l:.0f},{t:.0f})-({r:.0f},{b:.0f})")
    for l, t, r, b, label in data["essential"]:
        if l < CENTRAL[0] - 0.5 or t < CENTRAL[1] - 0.5 or r > CENTRAL[2] + 0.5 or b > CENTRAL[3] + 0.5:
            problems.append(f"focal element outside central 960 x 720: {label} ({l:.0f},{t:.0f})-({r:.0f},{b:.0f})")
    return len(data["text"]), len(data["essential"]), problems


def capture(html_files, raw_dir, do_audit):
    raw_dir.mkdir(parents=True, exist_ok=True)
    raws, report = [], {}
    with sync_playwright() as p:
        browser = p.chromium.launch(args=["--force-color-profile=srgb"])
        ctx = browser.new_context(viewport={"width": ARTBOARD_W, "height": ARTBOARD_H}, device_scale_factor=SCALE)
        for html in html_files:
            page = ctx.new_page()
            page.goto(html.resolve().as_uri(), wait_until="load")
            page.evaluate("document.fonts.ready")
            page.wait_for_timeout(300)
            states = page.evaluate("[...document.fonts].map(f => [f.family.replace(/['\\\"]/g, ''), f.status])")
            broken = sorted({fam for fam, st in states if st == "error"})
            if broken or not any(st == "loaded" for _, st in states):
                sys.exit(f"{html.name}: embedded fonts failed to load ({', '.join(broken) or 'none loaded'}).")
            if do_audit:
                report[html.name] = audit(page, html.name)
            raw = raw_dir / (html.stem + ".raw.png")
            page.screenshot(path=str(raw), clip={"x": 0, "y": 0, "width": ARTBOARD_W, "height": ARTBOARD_H})
            page.close()
            raws.append(raw)
        browser.close()
    return raws, report


def jpeg_bytes(im, quality):
    buf = io.BytesIO()
    im.save(buf, "JPEG", quality=quality, subsampling=0, optimize=True, dpi=(DPI, DPI), icc_profile=SRGB_PROFILE)
    return buf.getvalue()


def finalize(raw_png, out_dir, name):
    out_dir.mkdir(parents=True, exist_ok=True)
    im = Image.open(raw_png).convert("RGB")
    if im.size != OUT_SIZE:
        sys.exit(f"{name}: expected {OUT_SIZE}, got {im.size}")
    png_path = out_dir / f"{name}.png"
    im.save(png_path, "PNG", dpi=(DPI, DPI), icc_profile=SRGB_PROFILE, optimize=True, compress_level=9)
    master = jpeg_bytes(im, 100)
    jpg_path = out_dir / f"{name}.jpg"
    jpg_path.write_bytes(master)
    lines = [f"{png_path.name}: {png_path.stat().st_size / 1e6:.2f} MB",
             f"{jpg_path.name}: {len(master) / 1e6:.2f} MB (q100, 4:4:4)"]
    if len(master) > UPLOAD_LIMIT_BYTES:
        lo, hi, best = 1, 100, None
        while lo <= hi:                       # highest quality that still fits
            mid = (lo + hi) // 2
            data = jpeg_bytes(im, mid)
            if len(data) <= UPLOAD_LIMIT_BYTES:
                best, lo = (mid, data), mid + 1
            else:
                hi = mid - 1
        quality, data = best
        up = out_dir / f"{name}-upload.jpg"
        up.write_bytes(data)
        lines.append(f"{up.name}: {len(data) / 1e6:.2f} MB (q{quality}, 4:4:4) <- upload this one")
    else:
        lines.append("master JPG is under the limit: upload it as is")
    return lines


def main():
    here = Path(__file__).resolve().parent
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("files", nargs="*")
    ap.add_argument("--out", default="exports")
    ap.add_argument("--no-audit", action="store_true")
    args = ap.parse_args()
    html_files = [Path(f) for f in args.files] or sorted(here.glob("0*.html"))
    out_dir = Path(args.out)
    raws, report = capture(html_files, out_dir / ".raw", not args.no_audit)
    failed = False
    for html, raw in zip(html_files, raws):
        print(html.name)
        if html.name in report:
            n_text, n_ess, problems = report[html.name]
            print(f"    safe-zone audit: {n_text} text boxes, {n_ess} focal elements, "
                  f"{'all inside' if not problems else str(len(problems)) + ' problem(s)'}")
            for pr in problems:
                print("      !", pr)
            failed |= bool(problems)
        for line in finalize(raw, out_dir, html.stem):
            print("   ", line)
        raw.unlink()
    (out_dir / ".raw").rmdir()
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
