"""Render chart SVGs to crisp PNGs for PowerPoint (2x scale, transparent background by default).

usage: python chart_png.py <chart.svg> [out.png] [--scale 2]
       (or let charts.py do it: python charts.py spec.json chart.png)

PowerPoint can't keep SVG text crisp across every viewer (older Office, Google Slides, Keynote
import), so decks get a high-resolution PNG. Size the picture in the PPTX at w/140 by h/140 inches
(the px canvas was planned at ~140 px per inch), and keep the SVG next to it for later edits.
Needs: pip install playwright (Chromium is found via PLAYWRIGHT_BROWSERS_PATH or `playwright install chromium`).
"""
import glob
import os
import re
import sys


def launch(p):
    """Launch Chromium; fall back to CHROMIUM_PATH or any browser under PLAYWRIGHT_BROWSERS_PATH."""
    try:
        return p.chromium.launch()
    except Exception:
        root = os.environ.get("PLAYWRIGHT_BROWSERS_PATH", "")
        cands = [os.environ.get("CHROMIUM_PATH", "")] + [os.path.join(root, "chromium")] + glob.glob(os.path.join(root, "chromium*/chrome-linux*/chrome"))
        for c in cands:
            if c and os.path.isfile(c):
                return p.chromium.launch(executable_path=c)
        raise


def svgs_to_pngs(jobs, scale=2):
    """jobs: [(svg_text, w, h, out_png)]. One browser for all charts."""
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        b = launch(p)
        for svg, w, h, out in jobs:
            pg = b.new_page(viewport={"width": int(w), "height": int(h)}, device_scale_factor=scale)
            pg.set_content(f'<html><body style="margin:0;background:transparent">{svg}</body></html>')
            pg.locator("svg").first.screenshot(path=out, omit_background=True)
            pg.close()
        b.close()


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    scale = float(sys.argv[sys.argv.index("--scale") + 1]) if "--scale" in sys.argv else 2
    if "--scale" in sys.argv: args.remove(sys.argv[sys.argv.index("--scale") + 1])
    src = args[0]; out = args[1] if len(args) > 1 else re.sub(r"\.svg$", "", src) + ".png"
    svg = open(src, encoding="utf8").read()
    m = re.search(r'<svg[^>]*\bwidth="([\d.]+)"[^>]*\bheight="([\d.]+)"', svg)
    w, h = (float(m.group(1)), float(m.group(2))) if m else (1600, 900)
    svgs_to_pngs([(svg, w, h, out)], scale)
    print("wrote", out)
