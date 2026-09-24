#!/usr/bin/env python3
"""Rasterize the line icons in assets/icons/svg into brand-coloured PNGs for PowerPoint.

Run again after adding an SVG (and a catalog.json entry) to assets/icons/:
    python3 build_icons.py

Writes assets/icons/<colour>/<name>.png at 256×256, transparent.
Needs Google Chrome (or set CHROME_PATH).
"""
import os
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

from PIL import Image

ROOT = Path(__file__).resolve().parent.parent
ICONS = ROOT / "assets" / "icons"
COLOURS = {"deep-blue": "#0D145F", "mist": "#F3F4FA", "onyx": "#131416", "periwinkle": "#9C90E6"}
CELL, COLS, STROKE = 256, 10, 1.75


def chrome():
    for c in [os.environ.get("CHROME_PATH"), "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
              "/Applications/Chromium.app/Contents/MacOS/Chromium", shutil.which("google-chrome"), shutil.which("chromium")]:
        if c and os.path.exists(c):
            return c
    sys.exit("Chrome not found; set CHROME_PATH.")


def shoot(html, w, h, out):
    with tempfile.TemporaryDirectory() as d:
        page = Path(d) / "sheet.html"
        page.write_text(html)
        proc = subprocess.Popen([chrome(), "--headless=new", "--disable-gpu", "--hide-scrollbars",
                                 "--default-background-color=00000000", f"--window-size={w},{h}",
                                 f"--user-data-dir={d}/profile", f"--screenshot={out}", page.as_uri()],
                                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        last, deadline = -1, time.time() + 60
        while time.time() < deadline:
            if out.exists() and out.stat().st_size == last and last > 0:
                break
            last = out.stat().st_size if out.exists() else -1
            time.sleep(0.4)
        proc.kill()


def main():
    names = sorted(p.stem for p in (ICONS / "svg").glob("*.svg"))
    rows = (len(names) + COLS - 1) // COLS
    for colour, hexv in COLOURS.items():
        cells = []
        for n in names:
            svg = (ICONS / "svg" / f"{n}.svg").read_text()
            svg = svg.replace('stroke="currentColor"', f'stroke="{hexv}"').replace('stroke-width="2"', f'stroke-width="{STROKE}"')
            svg = svg.replace('width="24"', f'width="{CELL}"').replace('height="24"', f'height="{CELL}"')
            cells.append(f'<div style="width:{CELL}px;height:{CELL}px">{svg}</div>')
        html = (f'<html><body style="margin:0;background:transparent;display:grid;'
                f'grid-template-columns:repeat({COLS},{CELL}px)">{"".join(cells)}</body></html>')
        with tempfile.TemporaryDirectory() as d:
            sheet = Path(d) / "sheet.png"
            shoot(html, COLS * CELL, rows * CELL, sheet)
            im = Image.open(sheet)
            outdir = ICONS / colour
            outdir.mkdir(exist_ok=True)
            for i, n in enumerate(names):
                r, c = divmod(i, COLS)
                im.crop((c * CELL, r * CELL, (c + 1) * CELL, (r + 1) * CELL)).save(outdir / f"{n}.png")
        print(f"✓ {len(names)} icons in {colour}")


if __name__ == "__main__":
    main()
