#!/usr/bin/env python3
"""Render Frontier Commons branded graphics (PNG) from HTML templates.

Usage:
    python3 render.py spec.json [-o OUTDIR] [--scale 2]
    python3 render.py --list

A spec is a JSON object (or a list of them):
    {
      "template": "quote",          # file in ../templates/<name>.html
      "size": "square",             # square | portrait | story | landscape | og | linkedin-banner | WxH
      "name": "quote-sept",         # output file name (optional)
      "logo": "wordmark",           # wordmark | icon | none   (optional, default wordmark)
      "theme": "deep-blue",         # deep-blue | mist | papaya | periwinkle | onyx (optional)
      "fields": { "quote": "...", "author": "..." }
    }

Text fields support light markup:
    **words**  -> highlighted (colour depends on template)
    *words*    -> italic
    newline    -> line break
Image fields (any key ending in "image") are paths relative to the spec file.
"""
import argparse
import html
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TEMPLATES = ROOT / "templates"
ASSETS = ROOT / "assets"

SIZES = {
    "square": (1080, 1080),          # Instagram / LinkedIn / Facebook post
    "portrait": (1080, 1350),        # Instagram portrait post
    "story": (1080, 1920),           # IG / FB story, reels cover
    "landscape": (1920, 1080),       # slides, YouTube thumbnail, screens
    "og": (1200, 630),               # website / link preview, X post
    "linkedin-banner": (1584, 396),  # LinkedIn page cover
    "email-header": (1200, 400),     # newsletter header
}

THEMES = ("deep-blue", "mist", "papaya", "periwinkle", "onyx")
DARK_THEMES = ("deep-blue", "onyx")

CHROME_CANDIDATES = [
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/Applications/Chromium.app/Contents/MacOS/Chromium",
    "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
    "/Applications/Brave Browser.app/Contents/MacOS/Brave Browser",
    "google-chrome", "google-chrome-stable", "chromium", "chromium-browser", "chrome",
]


def find_chrome():
    env = os.environ.get("CHROME_PATH")
    if env:
        return env
    for c in CHROME_CANDIDATES:
        if os.path.isabs(c) and os.path.exists(c):
            return c
        found = shutil.which(c)
        if found:
            return found
    return None


def markup(text):
    """Escape text, then apply **highlight**, *italic* and newlines."""
    s = html.escape(str(text))
    s = re.sub(r"\*\*(.+?)\*\*", r'<span class="hl">\1</span>', s, flags=re.S)
    s = re.sub(r"(?<!\*)\*(?!\*)(.+?)\*", r"<em>\1</em>", s, flags=re.S)
    return s.replace("\n", "<br>")


def fill(template, fields):
    """Tiny mustache: {{#key}}...{{/key}} sections, {{key}} values, {{{key}}} raw."""
    def section(m):
        key, body = m.group(1), m.group(2)
        val = fields.get(key)
        if isinstance(val, list):
            out = []
            for i, item in enumerate(val, 1):
                ctx = dict(fields)
                ctx.update(item if isinstance(item, dict) else {"item": item})
                ctx["index"] = str(i)
                out.append(fill(body, ctx))
            return "".join(out)
        return fill(body, fields) if val else ""

    def inverted(m):
        return "" if fields.get(m.group(1)) else fill(m.group(2), fields)

    template = re.sub(r"\{\{#(\w+)\}\}(.*?)\{\{/\1\}\}", section, template, flags=re.S)
    template = re.sub(r"\{\{\^(\w+)\}\}(.*?)\{\{/\1\}\}", inverted, template, flags=re.S)
    template = re.sub(r"\{\{\{(\w+)\}\}\}", lambda m: str(fields.get(m.group(1), "")), template)
    return re.sub(r"\{\{(\w+)\}\}", lambda m: markup(fields.get(m.group(1), "")), template)


def logo_html(kind, on_dark):
    colour = "mist" if on_dark else "deep-blue"
    icon = (ASSETS / f"logo-icon-{colour}.svg").as_uri()
    if kind == "none":
        return ""
    if kind == "icon":
        return f'<img class="logo-icon" src="{icon}" alt="">'
    return (f'<div class="wordmark"><img src="{icon}" alt="">'
            f'<span>FRONTIER COMMONS</span></div>')


def build_html(spec, spec_dir, w, h):
    name = spec["template"]
    path = TEMPLATES / f"{name}.html"
    if not path.exists():
        sys.exit(f"Unknown template '{name}'. Run with --list to see options.")
    tpl = path.read_text()

    fields = dict(spec.get("fields", {}))
    for k, v in list(fields.items()):
        if k.endswith("image") and v and not re.match(r"^(https?|file|data):", str(v)):
            p = (spec_dir / v).resolve()
            if not p.exists():
                sys.exit(f"Image not found: {p}")
            fields[k] = p.as_uri()
        if k.endswith("image") and v:
            fields[k] = html.escape(fields[k], quote=True)  # raw-inserted via {{{key}}}

    m = re.search(r"default-theme:\s*([\w-]+)", tpl)
    theme = spec.get("theme") or (m.group(1) if m else "mist")
    if theme not in THEMES:
        sys.exit(f"Unknown theme '{theme}'. Use one of {', '.join(THEMES)}.")
    fields["logo"] = logo_html(spec.get("logo", "wordmark"), theme in DARK_THEMES)
    fields["theme"] = theme
    fields["base_css"] = (TEMPLATES / "base.css").as_uri()
    r = w / h
    fields["aspect"] = "ultrawide" if r > 2.5 else "wide" if r > 1.2 else "tall" if r < 0.83 else "square"
    return fill(tpl, fields)


def screenshot(chrome, html_file, out_png, w, h, scale, transparent=False):
    cmd = [
        chrome, "--headless=new", "--disable-gpu", "--hide-scrollbars",
        "--no-first-run", "--no-default-browser-check",
        f"--window-size={w},{h}", f"--force-device-scale-factor={scale}",
        "--virtual-time-budget=2500", "--allow-file-access-from-files",
        f"--screenshot={out_png}", html_file.as_uri(),
    ]
    if transparent:
        cmd.insert(1, "--default-background-color=00000000")
    # Headless Chrome on macOS often writes the screenshot and then fails to exit,
    # so watch for the file and stop Chrome as soon as it is complete.
    for _ in range(3):
        with tempfile.TemporaryDirectory() as profile:
            proc = subprocess.Popen(cmd + [f"--user-data-dir={profile}"],
                                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            deadline = time.time() + 30
            last_size = -1
            while time.time() < deadline:
                if proc.poll() is not None and not Path(out_png).exists():
                    break  # Chrome exited without a screenshot; retry
                if Path(out_png).exists():
                    size = Path(out_png).stat().st_size
                    if size > 0 and size == last_size:
                        break
                    last_size = size
                time.sleep(0.25)
            if proc.poll() is None:
                proc.terminate()
                try:
                    proc.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    proc.kill()
        if Path(out_png).exists() and Path(out_png).stat().st_size > 0:
            return
    sys.exit(f"Chrome failed to render {out_png}")


def parse_size(s):
    if s in SIZES:
        return SIZES[s]
    m = re.match(r"^(\d+)x(\d+)$", str(s))
    if not m:
        sys.exit(f"Unknown size '{s}'. Use one of {', '.join(SIZES)} or WxH.")
    return int(m.group(1)), int(m.group(2))


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("spec", nargs="?")
    ap.add_argument("-o", "--outdir", default=None, help="output folder (default: next to spec)")
    ap.add_argument("--scale", type=float, default=1, help="pixel density; 2 = retina/print-sharp")
    ap.add_argument("--keep-html", action="store_true", help="also save the filled HTML for debugging")
    ap.add_argument("--list", action="store_true", help="list templates and sizes")
    a = ap.parse_args()

    if a.list or not a.spec:
        print("Templates:", ", ".join(sorted(p.stem for p in TEMPLATES.glob("*.html"))))
        print("Sizes:", ", ".join(f"{k} ({w}x{h})" for k, (w, h) in SIZES.items()))
        return

    chrome = find_chrome()
    if not chrome:
        sys.exit("No Chrome/Chromium found. Install Google Chrome or set CHROME_PATH.")

    spec_path = Path(a.spec).resolve()
    specs = json.loads(spec_path.read_text())
    specs = specs if isinstance(specs, list) else [specs]
    outdir = Path(a.outdir).resolve() if a.outdir else spec_path.parent
    outdir.mkdir(parents=True, exist_ok=True)

    for i, spec in enumerate(specs, 1):
        w, h = parse_size(spec.get("size", "square"))
        name = spec.get("name") or f"{spec_path.stem}-{i}-{spec['template']}"
        page = build_html(spec, spec_path.parent, w, h)
        with tempfile.NamedTemporaryFile("w", suffix=".html", delete=False) as f:
            f.write(page)
            tmp = Path(f.name)
        out_png = outdir / f"{name}.png"
        if out_png.exists():
            out_png.unlink()
        screenshot(chrome, tmp, out_png, w, h, a.scale, spec.get("transparent", False))
        if a.keep_html:
            shutil.copy(tmp, outdir / f"{name}.html")
        tmp.unlink()
        print(f"✓ {out_png}  ({int(w * a.scale)}x{int(h * a.scale)})")


if __name__ == "__main__":
    main()
