#!/usr/bin/env python3
"""Build a Frontier Commons branded PowerPoint deck from a JSON outline.

Usage:
    python3 build_deck.py deck.json [-o out.pptx] [--plan]

deck.json:
    {
      "title": "Deck title (file metadata)",
      "footer": "Frontier Commons",          # optional footer text on content slides
      "slides": [ { "layout": "title", ... }, { "layout": "auto", "title": "...", "points": [...] }, ... ]
    }

"auto" slides pick their layout from the number and kind of points (see SKILL.md); every list
layout is split into balanced continuation slides when it holds too much. --plan prints the
choice made for each slide; a copy check lists text that is too long.

Layouts and their fields (* = required):
    title       title*, subtitle, date
    agenda      title, items* [str]
    section     title*, kicker
    content     title*, tag, body, bullets [str]
    two_column  title*, tag, left* {heading, body, bullets}, right* {heading, body, bullets}
    cards       title*, tag, cards* [{title*, subtitle, text, icon}]      (2-4)
    icon_grid   title*, tag, items* [{title*, text, icon}], theme          (3-6)
    stats       title*, tag, stats* [{value*, label*, detail, icon}]      (2-4)
    process     title*, tag, steps* [{title*, text, date, icon}]          (3-6)
    chart       title*, chart* {type, categories, series|values, format}, takeaway, source
    quote       quote*, author, role
    image       title*, image* | graphic* (a render.py graphic spec), tag, body, bullets, caption, side
    statement   text*, tag                                                (one big sentence)
    closing     title, subtitle, contact
  teaching & discussion (see references/teaching.md):
    question    question*, format, time, group, prompts, tag, theme
    questions   questions* [str] (2-4), title, format, time, lettered
    think_pair_share  question*, think/pair/share, think_time/pair_time/share_time
    scenario    story*, question, options (2-4)
    activity    title*, steps* (<=5), time, group, materials, output
    recap       points* (<=4), next_step, title
    break       title, time, text

Deck "style": "teaching" (default; the Plug In editorial look), "workshop" (warm, playful) or "formal" (briefings).
Deck "font_scale" resizes all text (default 0.9 in the teaching style, 1.0 otherwise).
Every slide accepts "notes" (speaker notes) and "icons": false.
Text supports **highlight** (bold + brand accent colour) and *italic*.
Image paths are relative to the JSON file.
"""
import argparse
import json
import re
import sys
from pathlib import Path

from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

ROOT = Path(__file__).resolve().parent.parent
ASSETS = ROOT / "assets"

# ---- Brand tokens (Branding Guide, Jan 2026) ---------------------------------
ONYX = RGBColor(0x13, 0x14, 0x16)
DEEP_BLUE = RGBColor(0x0D, 0x14, 0x5F)
MIST = RGBColor(0xF3, 0xF4, 0xFA)
PERIWINKLE = RGBColor(0x9C, 0x90, 0xE6)
PAPAYA = RGBColor(0xFF, 0xEF, 0xD5)
AMBER = RGBColor(0xFA, 0xBC, 0x20)
BLUE_50 = RGBColor(0x1F, 0x2F, 0xE0)     # tonal step — text highlights on light backgrounds
PERI_82 = RGBColor(0xBC, 0xB4, 0xEE)     # tonal step — text highlights on dark backgrounds
ONYX_40 = RGBColor(0x5F, 0x64, 0x6D)     # tonal step — secondary text
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

DISPLAY = "Inter"      # headers, headings
BODY = "DM Sans"       # section headers, body, captions

W, H = Inches(13.333), Inches(7.5)
M = Inches(0.8)        # outer margin
CONTENT_W = W - 2 * M


# ---- Low-level helpers -------------------------------------------------------
def bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    if STYLE == "teaching" and color == MIST:  # white page with the grey dot-grid band on the left
        fill.fore_color.rgb = WHITE
        slide.shapes.add_picture(str(ASSETS / "textures" / "dotband-grey.png"), -Inches(0.05), 0, Inches(1.58), H)
        return
    fill.fore_color.rgb = color


def rect(slide, x, y, w, h, color, rounded=False, radius=0.08):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    shp.line.fill.background()
    shp.shadow.inherit = False
    if rounded:
        shp.adjustments[0] = radius
    return shp


def oval(slide, x, y, d, color):
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, d, d)
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def tracking(run, thousandths):
    """Letter spacing in Figma-style units (1/1000 em), e.g. -20."""
    size = run.font.size.pt if run.font.size else 18
    run.font._element.set("spc", str(int(round(size * thousandths / 10))))


def add_runs(par, text, font, size, color, bold=False, italic=False, hl=None, track=0):
    """Add text to a paragraph, honouring **highlight** and *italic* markup."""
    tokens = re.split(r"(\*\*.+?\*\*|(?<!\*)\*(?!\*).+?\*)", str(text))
    for tok in tokens:
        if not tok:
            continue
        r = par.add_run()
        is_hl = tok.startswith("**") and tok.endswith("**")
        is_it = not is_hl and tok.startswith("*") and tok.endswith("*") and len(tok) > 1
        r.text = tok[2:-2] if is_hl else tok[1:-1] if is_it else tok
        f = r.font
        f.name = font
        f.size = Pt(round(size * FONT_SCALE, 1))
        f.bold = (bold if STYLE == "teaching" else True) if is_hl else bold  # teaching highlights by colour only
        f.italic = True if is_it else italic
        f.color.rgb = (hl or color) if is_hl else color
        if track:
            tracking(r, track)


def textbox(slide, x, y, w, h, text="", font=BODY, size=16, color=ONYX, bold=False, italic=False,
            align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, line=1.15, hl=None, track=0, autofit=False):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    lines = text if isinstance(text, list) else str(text).split("\n")
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line
        add_runs(p, ln, font, size, color, bold, italic, hl, track)
    if autofit:
        bodyPr = tf._txBody.find(qn("a:bodyPr"))
        for child in list(bodyPr):
            bodyPr.remove(child)
        bodyPr.append(bodyPr.makeelement(qn("a:normAutofit"), {}))
    return tb


def bullets(slide, x, y, w, h, items, size=18, color=ONYX, hl=BLUE_50, marker=DEEP_BLUE, gap=10):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.line_spacing = 1.3
        p.space_after = Pt(gap)
        pPr = p._p.get_or_add_pPr()
        indent = int(Pt(size) * 1.1)
        pPr.set("marL", str(indent))
        pPr.set("indent", str(-indent))
        bu_clr = pPr.makeelement(qn("a:buClr"), {})
        srgb = bu_clr.makeelement(qn("a:srgbClr"), {"val": str(marker)})
        bu_clr.append(srgb)
        pPr.append(bu_clr)
        pPr.append(pPr.makeelement(qn("a:buFont"), {"typeface": "Arial"}))
        pPr.append(pPr.makeelement(qn("a:buChar"), {"char": "•"}))
        add_runs(p, item, BODY, size, color, hl=hl, track=-10)
    return tb


def picture(slide, path, x, y, w, h, radius_in=0.25):
    """Insert an image cropped to fill the box, with brand rounded corners."""
    im = Image.open(path).convert("RGBA")
    box_ratio = w / h
    iw, ih = im.size
    if iw / ih > box_ratio:
        nw = int(ih * box_ratio)
        im = im.crop(((iw - nw) // 2, 0, (iw - nw) // 2 + nw, ih))
    else:
        nh = int(iw / box_ratio)
        im = im.crop((0, (ih - nh) // 2, iw, (ih - nh) // 2 + nh))
    px_per_in = im.size[0] / (w / 914400)
    r = int(radius_in * px_per_in)
    mask = Image.new("L", im.size, 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, im.size[0] - 1, im.size[1] - 1], r, fill=255)
    im.putalpha(mask)
    out = Path(TMP_DIR) / f"img_{abs(hash((str(path), w, h)))}.png"
    im.save(out)
    return slide.shapes.add_picture(str(out), x, y, w, h)


def render_graphic(spec, w, h):
    """Render a frontier-commons-graphics template to PNG sized for a slide frame."""
    import subprocess
    renderer = ROOT / "scripts" / "render.py"
    spec = dict(spec)
    spec.setdefault("size", f"{int(w / 914400 * 150)}x{int(h / 914400 * 150)}")
    spec.setdefault("logo", "none")
    spec["name"] = f"graphic_{abs(hash(json.dumps(spec, sort_keys=True)))}"
    spec_file = Path(TMP_DIR) / f"{spec['name']}.json"
    spec_file.write_text(json.dumps(spec))
    r = subprocess.run([sys.executable, str(renderer), str(spec_file), "-o", TMP_DIR, "--scale", "2"],
                       capture_output=True, text=True)
    out = Path(TMP_DIR) / f"{spec['name']}.png"
    if not out.exists():
        sys.exit(f"Graphic render failed:\n{r.stdout}{r.stderr}")
    return out


def logo(slide, kind, x, y, height):
    """kind: icon-deep-blue | icon-mist | wordmark-horizontal-mist | wordmark-stacked-deep-blue ..."""
    name = f"logo-{kind}.png" if kind.startswith("icon") else f"{kind}.png"
    return slide.shapes.add_picture(str(ASSETS / "logos" / name), x, y, height=height)


ICON_DIR = ASSETS / "icons"
PERI_91 = RGBColor(0xDC, 0xD8, 0xF6)     # tonal step — icon badge fill on light backgrounds
ICON_CATALOG = json.loads((ICON_DIR / "catalog.json").read_text())
WARNINGS = []


SUFFIXES = ("ations", "ation", "ities", "ively", "ivity", "ive", "ings", "ing", "ions", "ion", "ed", "es", "s", "e")


def stem(word):
    for suf in SUFFIXES:
        if word.endswith(suf) and len(word) - len(suf) >= 4:
            return word[: -len(suf)]
    return word


def _words(text):
    return re.sub(r"[^a-z0-9 -]", " ", (text or "").lower()).split()


def pick_icon(title, *texts):
    """Choose the catalog icon whose keywords best match; the title counts three times as much as the text."""
    fields = [(title, 3)] + [(t, 1) for t in texts]
    best, best_score = None, 0
    for name, words in ICON_CATALOG.items():
        score = 0
        for text, weight in fields:
            if not text:
                continue
            flat = " " + " ".join(_words(text)) + " "
            stem_list = [stem(w) for w in _words(text)]
            seen = set()
            for kw in words:
                key = kw if (" " in kw or "-" in kw) else stem(kw)
                if key in seen:
                    continue  # "partner" and "partners" count once
                seen.add(key)
                if " " in kw or "-" in kw:
                    pos = flat.find(f" {kw}")
                    hit = pos >= 0
                    pos = flat[:max(pos, 0)].count(" ")
                else:
                    hit = stem(kw) in stem_list
                    pos = stem_list.index(stem(kw)) if hit else 0
                if hit:
                    # earlier words in the title win ties ("Career support" -> career)
                    score += weight * (1 + kw.count(" ")) + (0.5 / (pos + 1) if weight > 1 else 0)
        if score > best_score:
            best, best_score = name, score
    return best


def consistent_icons(names):
    """Icons for every item or none: a row that mixes icons and numbers looks broken."""
    return names if all(names) else [None] * len(names)


def resolve_icon(item, s, *texts):
    """An item's icon: explicit name, false/none to disable, else auto-picked when the slide allows icons."""
    if isinstance(item, dict) and "icon" in item:
        v = item["icon"]
        if not v or v == "none":
            return None
        if v != "auto":
            if v in ICON_CATALOG:
                return v
            WARNINGS.append(f"unknown icon '{v}' (see references/icons.md); picked one automatically")
    if s.get("icons", True) is False:
        return None
    return pick_icon(*texts)


def icon(slide, name, x, y, size, colour="deep-blue"):
    return slide.shapes.add_picture(str(ICON_DIR / colour / f"{name}.png"), x, y, size, size)


def icon_badge(slide, name, x, y, d, dark=False):
    """Icon in a soft Periwinkle disc (light slides) or a plain Periwinkle icon (dark slides)."""
    if dark:
        return icon(slide, name, x, y, d, "periwinkle")
    oval(slide, x, y, d, PERI_91)
    pad = int(d * 0.24)
    return icon(slide, name, x + pad, y + pad, d - 2 * pad, "deep-blue")


def tag(slide, x, y, text):
    """Amber pill label (accent — use sparingly)."""
    size = 12
    w = Emu(int(Pt(size) * 0.62 * len(text) + Inches(0.4)))
    h = Inches(0.36)
    shp = rect(slide, x, y, w, h, AMBER, rounded=True, radius=0.5)
    tf = shp.text_frame
    tf.margin_left = tf.margin_right = Inches(0.12)
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    add_runs(p, text, BODY, size, ONYX, bold=True)
    return h


def footer(slide, n, text, dark=False):
    color = MIST if dark else ONYX_40
    if STYLE == "teaching":
        return
    if STYLE == "workshop":
        text = None
    logo(slide, "icon-mist" if dark else "icon-deep-blue", W - M - Inches(0.32), H - Inches(0.62), Inches(0.32))
    textbox(slide, M, H - Inches(0.58), Inches(6), Inches(0.3), f"{n}   {text}" if text else str(n),
            font=BODY, size=10, color=color)


def est_lines(text, width, size, bold=False):
    """Rough line count for wrapped text: DM Sans/Inter average ~0.52em per character (0.56 bold)."""
    per_line = max(1, int(width / Pt(size) / (0.56 if bold else 0.52)))
    lines, cur = 1, 0
    for word in str(text).split():
        if cur and cur + 1 + len(word) > per_line:
            lines, cur = lines + 1, len(word)
        else:
            cur += (1 if cur else 0) + len(word)
    return lines


def header(slide, s, top=Inches(0.7), color=DEEP_BLUE, size=36):
    """Optional amber tag + slide title. Returns y where content can start."""
    if STYLE == "teaching" and color != MIST:
        return e_header(slide, s)
    y = top
    if s.get("tag"):
        tag(slide, M, y, s["tag"])
        y += Inches(0.55)
    textbox(slide, M, y, CONTENT_W, Inches(1.2), s.get("title", ""), font=DISPLAY, size=size,
            color=color, bold=True, track=-10, line=1.05, hl=PERI_82 if color == MIST else BLUE_50,
            anchor=MSO_ANCHOR.TOP)
    lines = est_lines(s.get("title", ""), CONTENT_W, size, bold=True)
    return y + Pt(size) * 1.15 * lines + Inches(0.45)


# ---- Layouts -----------------------------------------------------------------
def l_title(prs, s, n, deck):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, DEEP_BLUE)
    textbox(sl, M, Inches(1.3), Inches(10), Inches(2.6), s["title"], font=DISPLAY, size=66, color=MIST,
            bold=True, line=1.0, anchor=MSO_ANCHOR.BOTTOM, hl=PERI_82)
    if s.get("subtitle"):
        textbox(sl, M, Inches(4.15), Inches(10), Inches(0.8), s["subtitle"].upper(), font=DISPLAY, size=24,
                color=MIST, bold=True, track=-10)
    rect(sl, M, Inches(6.45), CONTENT_W, Pt(1), MIST)
    if s.get("date"):
        textbox(sl, M, Inches(6.65), Inches(5), Inches(0.4), s["date"], font=BODY, size=14, color=MIST)
    logo(sl, "wordmark-horizontal-mist", W - M - Inches(2.9), Inches(6.68), Inches(0.37))
    return sl


def l_section(prs, s, n, deck):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, PERIWINKLE)
    line_y = int(H * 0.786)
    rect(sl, 0, line_y, W, Pt(1.25), MIST)
    size = 110 if len(s["title"]) <= 14 else 80
    # Box bottom sits one descender below the line so the baseline rests on it (brand signature).
    desc = int(Pt(size) * 0.25)
    textbox(sl, M, line_y + desc - Inches(3), CONTENT_W, Inches(3), s["title"], font=DISPLAY, size=size,
            color=MIST, bold=True, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.BOTTOM, line=0.95, track=-20)
    if s.get("kicker"):
        textbox(sl, M, Inches(0.7), Inches(8), Inches(0.5), s["kicker"], font=BODY, size=16, color=MIST)
    logo(sl, "icon-deep-blue", M, H - Inches(0.85), Inches(0.4))
    return sl


def l_content(prs, s, n, deck):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, MIST)
    y = header(sl, s)
    rect(sl, M, y - Inches(0.25), Inches(0.9), Pt(3), DEEP_BLUE)
    avail = H - y - Inches(1.0)
    if s.get("body"):
        tb = textbox(sl, M, y, Inches(9.5), Inches(1.4), s["body"], font=BODY, size=20, color=ONYX,
                     line=1.4, track=-10, hl=BLUE_50)
        y += Inches(0.5) + Pt(20) * 1.4 * (len(s["body"]) // 85 + 1)
        avail = H - y - Inches(1.0)
    items = s.get("bullets") or []
    if len(items) > 5:  # 6-8 short points: two balanced columns
        half = (len(items) + 1) // 2
        colw = (CONTENT_W - Inches(0.6)) / 2
        bullets(sl, M, y, Emu(int(colw)), avail, items[:half], size=18)
        bullets(sl, M + Emu(int(colw + Inches(0.6))), y, Emu(int(colw)), avail, items[half:], size=18)
    elif items:
        bullets(sl, M, y, Inches(10.5), avail, items, size=20)
    footer(sl, n, deck.get("footer"))
    return sl


def _column(sl, x, y, w, h, col):
    if col.get("heading"):
        textbox(sl, x, y, w, Inches(0.5), col["heading"], font=BODY, size=22, color=DEEP_BLUE, bold=True, track=-10)
        y += Inches(0.6)
    if col.get("body"):
        textbox(sl, x, y, w, Inches(1.5), col["body"], font=BODY, size=17, color=ONYX, line=1.4, track=-10, hl=BLUE_50)
        y += Pt(17) * 1.4 * (len(col["body"]) // 55 + 1) + Inches(0.25)
    if col.get("bullets"):
        bullets(sl, x, y, w, h, col["bullets"], size=17)


def l_two_column(prs, s, n, deck):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, MIST)
    y = header(sl, s)
    gap = Inches(0.6)
    cw = (CONTENT_W - gap) / 2
    h = H - y - Inches(1)
    rect(sl, M, y, Emu(int(cw)), h, WHITE, rounded=True, radius=0.06)
    rect(sl, M + Emu(int(cw + gap)), y, Emu(int(cw)), h, WHITE, rounded=True, radius=0.06)
    pad = Inches(0.4)
    _column(sl, M + pad, y + pad, Emu(int(cw - 2 * pad)), h - 2 * pad, s["left"])
    _column(sl, M + Emu(int(cw + gap)) + pad, y + pad, Emu(int(cw - 2 * pad)), h - 2 * pad, s["right"])
    footer(sl, n, deck.get("footer"))
    return sl


def l_cards(prs, s, n, deck):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, WHITE)
    y = header(sl, s)
    cards = s["cards"][:4]
    gap = Inches(0.35)
    cw = (CONTENT_W - gap * (len(cards) - 1)) / len(cards)
    longest = max(len(c.get("text", "")) for c in cards)
    chars_per_line = max(12, int(cw / Inches(1)) * 9)
    icons = consistent_icons([resolve_icon(c, s, c.get("title"), c.get("subtitle"), c.get("text")) for c in cards])
    icon_h = Inches(0.95) if any(icons) else 0
    ch = min(H - y - Inches(1.0), Inches(2.6) + icon_h + Pt(15) * 1.4 * (longest // chars_per_line + 2))
    y = y + (H - y - Inches(1.0) - ch) / 2  # centre the row vertically
    y = Emu(int(y))
    for i, c in enumerate(cards):
        x = M + Emu(int((cw + gap) * i))
        rect(sl, x, y, Emu(int(cw)), ch, PAPAYA, rounded=True, radius=0.1)
        pad = Inches(0.3)
        iw = Emu(int(cw - 2 * pad))
        textbox(sl, x + pad, y + Inches(0.45), iw, Inches(0.5), c["title"], font=DISPLAY, size=22, color=ONYX,
                bold=True, align=PP_ALIGN.CENTER)
        uw = Inches(1.9)
        rect(sl, x + Emu(int((cw - uw) / 2)), y + Inches(1.0), uw, Pt(2.5), AMBER)
        ty = y + Inches(1.2)
        if c.get("subtitle"):
            textbox(sl, x + pad, ty, iw, Inches(0.8), c["subtitle"], font=DISPLAY, size=18, color=ONYX,
                    align=PP_ALIGN.CENTER)
            ty += Inches(0.85)
        if icons[i]:
            isz = Inches(0.75)
            icon(sl, icons[i], x + Emu(int((cw - isz) / 2)), ty + Inches(0.05), isz)
        ty += icon_h
        if c.get("text"):
            textbox(sl, x + pad, ty + Inches(0.15), iw, ch - (ty - y) - Inches(0.4), c["text"], font=BODY, size=15,
                    color=ONYX, align=PP_ALIGN.CENTER, line=1.4, hl=DEEP_BLUE)
    footer(sl, n, deck.get("footer"))
    return sl


def l_stats(prs, s, n, deck):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, MIST)
    y = header(sl, s)
    stats = s["stats"][:4]
    gap = Inches(0.5)
    cw = (CONTENT_W - gap * (len(stats) - 1)) / len(stats)
    has_icons = any(st.get("icon") in ICON_CATALOG for st in stats)
    top = y + (Inches(0.9) if has_icons else Inches(0.4))
    for i, st in enumerate(stats):
        x = M + Emu(int((cw + gap) * i))
        iw = Emu(int(cw))
        if st.get("icon") in ICON_CATALOG:
            icon_badge(sl, st["icon"], x, top - Inches(0.75), Inches(0.7))
        textbox(sl, x, top, iw, Inches(1.4), st["value"], font=DISPLAY, size=72 if len(stats) < 4 else 60,
                color=DEEP_BLUE, bold=True, anchor=MSO_ANCHOR.BOTTOM, line=0.9, track=-30)
        rect(sl, x, top + Inches(1.6), Inches(0.9), Pt(4), AMBER)
        textbox(sl, x, top + Inches(1.85), iw, Inches(0.9), st["label"], font=DISPLAY, size=20, color=ONYX,
                line=1.2, track=-10)
        if st.get("detail"):
            textbox(sl, x, top + Inches(2.75), iw, Inches(1.2), st["detail"], font=BODY, size=14, color=ONYX_40,
                    line=1.4)
    footer(sl, n, deck.get("footer"))
    return sl


def l_process(prs, s, n, deck):
    """Sequential steps / journey / timeline. Steps may carry a `date` (shown above the dot) and an icon."""
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, MIST if STYLE == "teaching" else PAPAYA)
    y = header(sl, s)
    steps = s["steps"][:6]
    k = len(steps)
    start = s.get("start", 1)  # step numbering continues across split slides
    colw = CONTENT_W / k
    d = Inches(0.9)
    has_dates = any(st.get("date") for st in steps)
    cy = y + (Inches(1.05) if has_dates else Inches(0.7))
    rect(sl, M + Emu(int(colw / 2)), cy + d / 2 - Pt(1.5), Emu(int(colw * (k - 1))), Pt(3), DEEP_BLUE)
    tw = Emu(int(colw - Inches(0.25)))
    step_icons = consistent_icons([resolve_icon(st, s, st.get("title"), st.get("text")) for st in steps])
    for i, st in enumerate(steps):
        cx = M + Emu(int(colw * i + colw / 2))
        ic = step_icons[i]
        c = oval(sl, cx - d / 2, cy, d, PERIWINKLE)
        num = str(start + i)
        if ic:
            pad = int(d * 0.24)
            icon(sl, ic, cx - d / 2 + pad, cy + pad, d - 2 * pad, "onyx")
        else:
            tf = c.text_frame
            tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            add_runs(p, num, DISPLAY, 22, ONYX, bold=True)
        if st.get("date"):
            textbox(sl, cx - tw / 2, cy - Inches(0.5), tw, Inches(0.35), st["date"].upper(), font=BODY, size=12,
                    color=BLUE_50, bold=True, align=PP_ALIGN.CENTER, track=60)
        label = f"{num}. {st['title']}" if ic and s.get("numbered", True) else st["title"]
        textbox(sl, cx - tw / 2, cy + d + Inches(0.3), tw, Inches(0.7), label, font=BODY, size=17,
                color=DEEP_BLUE, bold=True, align=PP_ALIGN.CENTER, line=1.15)
        if st.get("text"):
            textbox(sl, cx - tw / 2, cy + d + Inches(1.0), tw, Inches(2), st["text"], font=BODY, size=14,
                    color=ONYX, align=PP_ALIGN.CENTER, line=1.4)
    footer(sl, n, deck.get("footer"))
    return sl


def l_icon_grid(prs, s, n, deck):
    """3–6 features / benefits / programs, each with an icon, a short title and one line of text."""
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    dark = s.get("theme") == "dark"
    bg(sl, DEEP_BLUE if dark else MIST)
    y = header(sl, s, color=MIST if dark else DEEP_BLUE)
    items = s["items"][:6]
    k = len(items)
    cols = {1: 1, 2: 2, 3: 3, 4: 4 if max(len(it.get("text", "")) for it in items) < 90 else 2, 5: 3, 6: 3}[k]
    rows = (k + cols - 1) // cols
    gap_x, gap_y = Inches(0.5), Inches(0.35)
    cw = (CONTENT_W - gap_x * (cols - 1)) / cols
    avail = H - y - Inches(0.95)
    rh = min(Inches(2.6) if rows == 1 else Inches(1.7), (avail - gap_y * (rows - 1)) / rows)
    d = Inches(0.85) if rows == 1 else Inches(0.7)
    centred = rows == 1
    top = y + (avail - (rh * rows + gap_y * (rows - 1))) / 2
    for i, it in enumerate(items):
        r, c = divmod(i, cols)
        in_row = min(cols, k - r * cols)
        offset = (cols - in_row) * (cw + gap_x) / 2  # centre a short last row
        x = M + Emu(int(offset + (cw + gap_x) * c))
        yy = Emu(int(top + (rh + gap_y) * r))
        ic = resolve_icon(it, s, it.get("title"), it.get("text")) or "sparkles"
        iw = Emu(int(cw))
        if centred:
            icon_badge(sl, ic, x + Emu(int((cw - d) / 2)), yy, d, dark)
            ty = yy + d + Inches(0.25)
            align = PP_ALIGN.CENTER
            tx, tw = x, iw
        else:
            icon_badge(sl, ic, x, yy, d, dark)
            ty = yy + Inches(0.02)
            align = PP_ALIGN.LEFT
            tx, tw = x + d + Inches(0.25), Emu(int(cw - d - Inches(0.25)))
        textbox(sl, tx, ty, tw, Inches(0.5), it["title"], font=BODY, size=18, color=MIST if dark else DEEP_BLUE,
                bold=True, align=align, line=1.15, track=-10)
        title_lines = est_lines(it["title"], tw, 18, bold=True)
        text_y = ty + Pt(18) * 1.15 * title_lines + Inches(0.12)
        if it.get("text"):
            textbox(sl, tx, text_y, tw, rh - (text_y - yy), it["text"], font=BODY, size=14,
                    color=MIST if dark else ONYX, align=align, line=1.4)
    footer(sl, n, deck.get("footer"), dark=dark)
    return sl


CHART_COLOURS = [DEEP_BLUE, PERIWINKLE, AMBER, RGBColor(0x62, 0x6D, 0xEA), RGBColor(0xCC, 0xC6, 0xF2),
                 RGBColor(0x5F, 0x64, 0x6D)]


def l_chart(prs, s, n, deck):
    """A native, editable PowerPoint chart in brand colours, with an optional takeaway panel."""
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION

    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, MIST)
    y = header(sl, s)
    ch = s["chart"]
    kind = ch.get("type", "column")
    types = {"column": XL_CHART_TYPE.COLUMN_CLUSTERED, "bar": XL_CHART_TYPE.BAR_CLUSTERED,
             "stacked": XL_CHART_TYPE.COLUMN_STACKED, "line": XL_CHART_TYPE.LINE_MARKERS,
             "pie": XL_CHART_TYPE.PIE, "doughnut": XL_CHART_TYPE.DOUGHNUT}
    if kind not in types:
        sys.exit(f"Slide {n}: chart type must be one of {', '.join(types)}")
    data = CategoryChartData()
    data.categories = ch["categories"]
    series = ch.get("series") or [{"name": ch.get("name", ""), "values": ch["values"]}]
    for se in series:
        data.add_series(se.get("name", ""), se["values"])
    has_side = bool(s.get("takeaway"))
    cw = Inches(7.6) if has_side else CONTENT_W
    chh = H - y - Inches(1.2)
    gf = sl.shapes.add_chart(types[kind], M, y, cw, chh, data)
    chart = gf.chart
    chart.font.name = BODY
    chart.font.size = Pt(13)
    chart.font.color.rgb = ONYX
    pie = kind in ("pie", "doughnut")
    if pie:
        pts = chart.plots[0].series[0].points
        for i in range(len(ch["categories"])):
            pts[i].format.fill.solid()
            pts[i].format.fill.fore_color.rgb = CHART_COLOURS[i % len(CHART_COLOURS)]
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.RIGHT
        chart.legend.include_in_layout = False
    else:
        for i, se in enumerate(chart.plots[0].series):
            col = CHART_COLOURS[i % len(CHART_COLOURS)]
            if kind == "line":
                se.format.line.color.rgb = col
                se.format.line.width = Pt(3)
                se.smooth = False
                se.marker.format.fill.solid()
                se.marker.format.fill.fore_color.rgb = col
                se.marker.format.line.color.rgb = col
            else:
                se.format.fill.solid()
                se.format.fill.fore_color.rgb = col
        if kind != "line":
            chart.plots[0].gap_width = 60
        chart.has_legend = len(series) > 1
        if chart.has_legend:
            chart.legend.position = XL_LEGEND_POSITION.TOP
            chart.legend.include_in_layout = False
        va = chart.value_axis
        va.has_major_gridlines = True
        va.major_gridlines.format.line.color.rgb = RGBColor(0xD6, 0xD8, 0xDB)
        va.format.line.fill.background()
        va.tick_labels.font.color.rgb = ONYX_40
        chart.category_axis.format.line.color.rgb = RGBColor(0xAD, 0xB1, 0xB8)
        chart.category_axis.tick_labels.font.color.rgb = ONYX
        if kind == "bar":
            chart.category_axis.reverse_order = True  # first category at the top
        if ch.get("format"):
            va.tick_labels.number_format = ch["format"]
            va.tick_labels.number_format_is_linked = False
    plot = chart.plots[0]
    plot.has_data_labels = ch.get("labels", True)
    if plot.has_data_labels:
        dl = plot.data_labels
        dl.font.size = Pt(12)
        dl.font.bold = True
        dl.font.color.rgb = MIST if pie else ONYX
        if ch.get("format"):
            dl.number_format = ch["format"]
            dl.number_format_is_linked = False
        if not pie and kind != "line":
            dl.position = XL_LABEL_POSITION.OUTSIDE_END
    if has_side:
        px = M + cw + Inches(0.5)
        pw = CONTENT_W - cw - Inches(0.5)
        rect(sl, px, y, pw, chh, PAPAYA, rounded=True, radius=0.08)
        rect(sl, px + Inches(0.4), y + Inches(0.5), Inches(0.8), Pt(4), AMBER)
        textbox(sl, px + Inches(0.4), y + Inches(0.75), pw - Inches(0.8), chh - Inches(1.2), s["takeaway"],
                font=DISPLAY, size=22, color=DEEP_BLUE, line=1.25, track=-20, hl=BLUE_50)
    if s.get("source"):
        textbox(sl, M, y + chh + Inches(0.08), cw, Inches(0.3), f"Source: {s['source']}", font=BODY, size=11,
                color=ONYX_40)
    footer(sl, n, deck.get("footer"))
    return sl


def l_agenda(prs, s, n, deck):
    s = dict(s, items=[i.get("title") if isinstance(i, dict) else i for i in s["items"]])
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, MIST)
    s = dict(s, title=s.get("title", "Agenda"))
    y = header(sl, s)
    items = s["items"][:8]
    two = len(items) > 4
    per_col = (len(items) + 1) // 2 if two else len(items)
    cw = (CONTENT_W - Inches(0.6)) / 2 if two else CONTENT_W
    rowh = min(Inches(1.05), (H - y - Inches(1)) / per_col)
    d = Inches(0.6)
    for i, item in enumerate(items):
        col, row = divmod(i, per_col)
        x = M + Emu(int((cw + Inches(0.6)) * col))
        yy = y + Emu(int(rowh * row))
        c = oval(sl, x, yy, d, PERIWINKLE)
        tf = c.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        add_runs(p, str(i + 1), DISPLAY, 18, ONYX, bold=True)
        textbox(sl, x + d + Inches(0.3), yy, Emu(int(cw - d - Inches(0.3))), d, item, font=DISPLAY, size=22,
                color=ONYX, anchor=MSO_ANCHOR.MIDDLE, track=-20)
    footer(sl, n, deck.get("footer"))
    return sl


def l_quote(prs, s, n, deck):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, PAPAYA)
    textbox(sl, M, Inches(0.7), Inches(2), Inches(1.6), "“", font=DISPLAY, size=160, color=AMBER, bold=True,
            line=1.0)
    size = 34 if len(s["quote"]) < 160 else 26
    textbox(sl, M, Inches(2.1), Inches(10.5), Inches(3.2), s["quote"], font=BODY, size=size, color=DEEP_BLUE,
            italic=True, line=1.3, track=-20, anchor=MSO_ANCHOR.MIDDLE, hl=BLUE_50)
    rect(sl, M, Inches(5.55), Inches(1.2), Pt(4), AMBER)
    if s.get("author"):
        textbox(sl, M, Inches(5.8), Inches(8), Inches(0.4), s["author"], font=BODY, size=18, color=DEEP_BLUE, bold=True)
    if s.get("role"):
        textbox(sl, M, Inches(6.2), Inches(8), Inches(0.4), s["role"], font=BODY, size=14, color=ONYX_40)
    footer(sl, n, deck.get("footer"))
    return sl


def l_image(prs, s, n, deck):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, MIST)
    img_w = Inches(5.6)
    img_left = s.get("side", "left") == "left"
    ix = M if img_left else W - M - img_w
    tx = M + img_w + Inches(0.7) if img_left else M
    tw = CONTENT_W - img_w - Inches(0.7)
    if s.get("graphic"):
        img = render_graphic(s["graphic"], img_w, H - Inches(1.6))
    else:
        img = DECK_DIR / s["image"]
    picture(sl, img, ix, Inches(0.7), img_w, H - Inches(1.6))
    y = Inches(1.3)
    if s.get("tag"):
        tag(sl, tx, y, s["tag"])
        y += Inches(0.6)
    textbox(sl, tx, y, tw, Inches(1.8), s["title"], font=DISPLAY, size=34, color=DEEP_BLUE, bold=True, line=1.05,
            track=-10, hl=BLUE_50)
    y += Pt(34) * 1.1 * (len(s["title"]) // 26 + 1) + Inches(0.35)
    if s.get("body"):
        textbox(sl, tx, y, tw, Inches(2), s["body"], font=BODY, size=17, color=ONYX, line=1.4, track=-10, hl=BLUE_50)
        y += Pt(17) * 1.4 * (len(s["body"]) // 48 + 1) + Inches(0.3)
    if s.get("bullets"):
        bullets(sl, tx, y, tw, H - y - Inches(1), s["bullets"], size=17)
    if s.get("caption"):
        textbox(sl, ix, H - Inches(0.8), img_w, Inches(0.3), s["caption"], font=BODY, size=11, color=ONYX_40)
    footer(sl, n, deck.get("footer"))
    return sl


def l_statement(prs, s, n, deck):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, DEEP_BLUE)
    y = Inches(2.0)
    if s.get("tag"):
        tag(sl, M, Inches(1.4), s["tag"])
    textbox(sl, M, y, Inches(11), Inches(3.8), s["text"], font=DISPLAY, size=48, color=MIST, bold=True, line=1.1,
            track=-20, hl=PERI_82, anchor=MSO_ANCHOR.TOP)
    footer(sl, n, deck.get("footer"), dark=True)
    return sl


def l_closing(prs, s, n, deck):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, DEEP_BLUE)
    lh = Inches(1.9)
    pic = logo(sl, "wordmark-stacked-mist", 0, Inches(1.2), lh)
    pic.left = int((W - pic.width) / 2)
    textbox(sl, M, Inches(3.65), CONTENT_W, Inches(1), s.get("title", "Thank you"), font=DISPLAY, size=48,
            color=MIST, bold=True, align=PP_ALIGN.CENTER)
    if s.get("subtitle"):
        textbox(sl, M, Inches(4.6), CONTENT_W, Inches(0.6), s["subtitle"], font=DISPLAY, size=22, color=PERI_82,
                align=PP_ALIGN.CENTER, track=-20)
    if s.get("contact"):
        textbox(sl, M, Inches(5.7), CONTENT_W, Inches(0.8), s["contact"], font=BODY, size=16, color=MIST,
                align=PP_ALIGN.CENTER, line=1.4)
    return sl


# ---- Teaching & discussion style ---------------------------------------------
# The default look: warmer grounds (Papaya, Periwinkle), rounded shapes, big friendly type,
# icons on every list, and facilitation slides (questions, activities, scenarios, recaps).
STYLE = "teaching"
WHITE_ = WHITE


def chip(slide, x, y, text, icon_name=None, fill=None, ink=ONYX, size=14):
    """A rounded label with an optional icon, e.g. [clock] 5 min. Returns its width."""
    fill = fill or WHITE
    h = Inches(0.46)
    pad = Inches(0.18)
    isz = Inches(0.26) if icon_name else 0
    w = Emu(int(Pt(size) * 0.56 * len(text) + pad * 2 + (isz + Inches(0.1) if icon_name else 0)))
    rect(slide, x, y, w, h, fill, rounded=True, radius=0.5)
    tx = x + pad
    if icon_name:
        colour = "mist" if ink == MIST else "onyx" if ink == ONYX else "deep-blue"
        icon(slide, icon_name, tx, y + (h - isz) / 2, isz, colour)
        tx += isz + Inches(0.1)
    textbox(slide, tx, y, w - (tx - x) - pad + Inches(0.1), h, text, font=BODY, size=size, color=ink, bold=True,
            anchor=MSO_ANCHOR.MIDDLE)
    return w


def chips(slide, x, y, items, **kw):
    """Row of chips: items = [(text, icon)]."""
    for text, ic in items:
        if text:
            x += chip(slide, x, y, text, ic, **kw) + Inches(0.15)
    return x


def meta_chips(s):
    """Format / time / group-size chips shared by discussion and activity slides."""
    return [(s.get("format"), "users"), (s.get("time"), "clock"), (s.get("group"), "users-round")]


def t_title(prs, s, n, deck):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, PAPAYA)
    # friendly brand shapes on the right: one Periwinkle disc, a Deep Blue rounded tile with the topic icon, an Amber dot
    oval(sl, Inches(8.3), Inches(0.6), Inches(5.6), PERIWINKLE)
    tile = Inches(2.5)
    rect(sl, Inches(7.6), Inches(3.9), tile, tile, DEEP_BLUE, rounded=True, radius=0.18)
    ic = s.get("icon") or pick_icon(s.get("title", ""), s.get("subtitle", "")) or "sparkles"
    icon(sl, ic, Inches(7.6) + Inches(0.55), Inches(3.9) + Inches(0.55), tile - Inches(1.1), "mist")
    oval(sl, Inches(10.4), Inches(5.55), Inches(0.9), AMBER)
    y = Inches(1.5)
    if s.get("session"):
        tag(sl, M, y, s["session"])
    tsize = 54 if len(s["title"]) <= 24 else 46 if len(s["title"]) <= 44 else 40
    textbox(sl, M, y + Inches(0.6), Inches(6.6), Inches(2.9), s["title"], font=DISPLAY, size=tsize, color=DEEP_BLUE,
            bold=True, line=1.02, track=-20, hl=BLUE_50, anchor=MSO_ANCHOR.TOP)
    lines = est_lines(s["title"].replace("\n", " "), Inches(6.6), tsize, bold=True) + s["title"].count("\n")
    sy = y + Inches(0.6) + Pt(tsize) * 1.25 * lines + Inches(0.3)
    if s.get("subtitle"):
        textbox(sl, M, sy, Inches(6.4), Inches(1.2), s["subtitle"], font=DISPLAY, size=22, color=ONYX, line=1.25,
                track=-20)
    info = "   ·   ".join(x for x in (s.get("presenter"), s.get("date")) if x)
    if info:
        textbox(sl, M, H - Inches(1.45), Inches(6.5), Inches(0.4), info, font=BODY, size=14, color=ONYX)
    logo(sl, "wordmark-horizontal-deep-blue", M, H - Inches(0.85), Inches(0.3))
    return sl


def t_section(prs, s, n, deck):
    sl = l_section(prs, s, n, deck)
    return sl


def t_content(prs, s, n, deck):
    """Friendly content: each point in its own rounded row with an icon (or dots for 6+ short points)."""
    items = s.get("bullets") or []
    if len(items) > 5 or s.get("icons") is False:
        return l_content(prs, s, n, deck)
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, MIST)
    y = header(sl, s, size=38)
    if s.get("body"):
        textbox(sl, M, y - Inches(0.15), Inches(10.5), Inches(1.2), s["body"], font=BODY, size=20, color=ONYX,
                line=1.35, track=-10, hl=BLUE_50)
        y += Pt(20) * 1.35 * est_lines(s["body"], Inches(10.5), 20) + Inches(0.2)
    if items:
        names = [pick_icon(b) for b in items]
        names = [nm or "circle-check" for nm in names]
        avail = H - y - Inches(0.9)
        gap = Inches(0.16)
        rh = min(Inches(0.95), (avail - gap * (len(items) - 1)) / len(items))
        for i, (b, nm) in enumerate(zip(items, names)):
            yy = Emu(int(y + (rh + gap) * i))
            rect(sl, M, yy, CONTENT_W, rh, WHITE, rounded=True, radius=0.25)
            d = min(Inches(0.62), rh - Inches(0.2))
            icon_badge(sl, nm, M + Inches(0.22), yy + (rh - d) / 2, d)
            textbox(sl, M + Inches(0.22) + d + Inches(0.3), yy, CONTENT_W - d - Inches(0.9), rh, b, font=BODY,
                    size=20, color=ONYX, anchor=MSO_ANCHOR.MIDDLE, line=1.2, track=-10, hl=BLUE_50)
    footer(sl, n, deck.get("footer"))
    return sl


def t_agenda(prs, s, n, deck):
    """Agenda as a friendly run-sheet: numbered rows with optional minutes per item."""
    items = [i if isinstance(i, dict) else {"title": i} for i in s["items"]]
    if len(items) > 6:
        return l_agenda(prs, dict(s, items=[i["title"] for i in items]), n, deck)
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, MIST)
    y = header(sl, dict(s, title=s.get("title", "Our time together")), size=38)
    gap = Inches(0.16)
    rh = min(Inches(0.85), (H - y - Inches(0.9) - gap * (len(items) - 1)) / len(items))
    for i, it in enumerate(items):
        yy = Emu(int(y + (rh + gap) * i))
        rect(sl, M, yy, CONTENT_W, rh, WHITE, rounded=True, radius=0.3)
        d = rh - Inches(0.22)
        c = oval(sl, M + Inches(0.14), yy + Inches(0.11), d, PERIWINKLE)
        tf = c.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        add_runs(p, str(i + 1), DISPLAY, 18, ONYX, bold=True)
        textbox(sl, M + d + Inches(0.45), yy, Inches(8.4), rh, it["title"], font=DISPLAY, size=22, color=ONYX,
                anchor=MSO_ANCHOR.MIDDLE, track=-20)
        if it.get("time"):
            tw = Emu(int(Pt(14) * 0.56 * len(it["time"]) + Inches(0.8)))
            chip(sl, M + CONTENT_W - tw - Inches(0.2), yy + (rh - Inches(0.46)) / 2, it["time"], "clock",
                 fill=PAPAYA)
    footer(sl, n, deck.get("footer"))
    return sl


def l_question(prs, s, n, deck):
    """One big open question for the room, in a speech bubble, with format/time chips and optional prompts."""
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    theme = s.get("theme", "periwinkle")
    bg(sl, PAPAYA if theme == "papaya" else PERIWINKLE)
    tag(sl, M, Inches(0.7), s.get("tag", "Let's discuss"))
    prompts = s.get("prompts") or []
    bw = CONTENT_W
    bh = Inches(3.1) if prompts else Inches(3.9)
    by = Inches(1.35)
    bubble = rect(sl, M, by, bw, bh, WHITE, rounded=True, radius=0.14)
    tail = sl.shapes.add_shape(MSO_SHAPE.RIGHT_TRIANGLE, M + Inches(1.0), by + bh - Inches(0.02), Inches(0.6),
                               Inches(0.5))
    tail.fill.solid()
    tail.fill.fore_color.rgb = WHITE
    tail.line.fill.background()
    tail.shadow.inherit = False
    tail.rotation = 0
    tail.flipV = True
    q = s["question"]
    size = 44 if len(q) <= 70 else 36 if len(q) <= 110 else 30
    icon(sl, "message-circle", M + Inches(0.45), by + Inches(0.45), Inches(0.75), "deep-blue")
    textbox(sl, M + Inches(1.5), by + Inches(0.3), bw - Inches(2.0), bh - Inches(0.6), q, font=DISPLAY, size=size,
            color=DEEP_BLUE, bold=True, line=1.1, track=-20, anchor=MSO_ANCHOR.MIDDLE, hl=BLUE_50)
    y = by + bh + Inches(0.75)
    chips(sl, M, y, meta_chips(s))
    if prompts:
        px = M + Inches(4.6) if any(x for x, _ in meta_chips(s)) else M
        textbox(sl, px, y - Inches(0.05), W - px - M, Inches(1.4),
                "\n".join("→  " + p for p in prompts[:3]), font=BODY, size=16, color=ONYX, line=1.35)
    footer(sl, n, deck.get("footer"))
    return sl


def l_questions(prs, s, n, deck):
    """2-4 discussion questions as cards: groups pick one, or work through them in order."""
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, PAPAYA)
    s = dict(s, tag=s.get("tag", "Let's discuss"), title=s.get("title", "Pick one question to discuss"))
    y = header(sl, s, size=36)
    qs = s["questions"][:4]
    cols = 2 if len(qs) == 4 else len(qs)
    rows = (len(qs) + cols - 1) // cols
    gap = Inches(0.3)
    cw = (CONTENT_W - gap * (cols - 1)) / cols
    avail = H - y - Inches(1.5)
    ch = (avail - gap * (rows - 1)) / rows
    for i, q in enumerate(qs):
        r, c = divmod(i, cols)
        x = M + Emu(int((cw + gap) * c))
        yy = Emu(int(y + (ch + gap) * r))
        rect(sl, x, yy, Emu(int(cw)), Emu(int(ch)), WHITE, rounded=True, radius=0.1)
        d = Inches(0.6)
        c_ = oval(sl, x + Inches(0.3), yy + Inches(0.3), d, PERIWINKLE)
        tf = c_.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        add_runs(p, "ABCD"[i] if s.get("lettered") else str(i + 1), DISPLAY, 18, ONYX, bold=True)
        textbox(sl, x + Inches(0.3), yy + Inches(1.05), Emu(int(cw - Inches(0.6))), Emu(int(ch - Inches(1.3))), q,
                font=DISPLAY, size=22 if rows == 1 else 20, color=DEEP_BLUE, bold=True, line=1.2, track=-20)
    chips(sl, M, H - Inches(1.25), meta_chips(s))
    footer(sl, n, deck.get("footer"))
    return sl


def l_think_pair_share(prs, s, n, deck):
    """The question on top; Think / Pair / Share columns with minutes and one instruction each."""
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, MIST)
    s = dict(s, tag=s.get("tag", "Think · Pair · Share"))
    y = header(sl, dict(s, title=s["question"]), size=34)
    steps = [("Think", "brain", s.get("think", "On your own, jot down your first thoughts."), s.get("think_time", "2 min")),
             ("Pair", "users", s.get("pair", "Share with the person next to you. What's similar? Different?"), s.get("pair_time", "4 min")),
             ("Share", "megaphone", s.get("share", "A few pairs share one idea with the whole room."), s.get("share_time", "5 min"))]
    gap = Inches(0.35)
    cw = (CONTENT_W - gap * 2) / 3
    ch = H - y - Inches(1.0)
    fills = [WHITE, PAPAYA, PERI_91]
    for i, (name, ic, text, t) in enumerate(steps):
        x = M + Emu(int((cw + gap) * i))
        inner = Emu(int(cw - Inches(0.7)))
        rect(sl, x, y, Emu(int(cw)), ch, fills[i], rounded=True, radius=0.1)
        d = Inches(0.7)
        if fills[i] != PERI_91:
            icon_badge(sl, ic, x + Inches(0.35), y + Inches(0.35), d)
        else:
            icon(sl, ic, x + Inches(0.35), y + Inches(0.35), d)
        if t:  # time chip top-right, level with the icon, so it never collides with the text
            tw = Emu(int(Pt(14) * 0.56 * len(t) + Inches(0.8)))
            chip(sl, x + Emu(int(cw)) - tw - Inches(0.3), y + Inches(0.47), t, "clock",
                 fill=WHITE if fills[i] != WHITE else MIST)
        textbox(sl, x + Inches(0.35), y + Inches(1.2), inner, Inches(0.6), f"{i + 1}. {name}",
                font=DISPLAY, size=26, color=DEEP_BLUE, bold=True)
        textbox(sl, x + Inches(0.35), y + Inches(1.8), inner, ch - Inches(2.0), text, font=BODY, size=17, color=ONYX,
                line=1.35)
    footer(sl, n, deck.get("footer"))
    return sl


def l_activity(prs, s, n, deck):
    """Hands-on exercise: numbered instructions on the left, logistics panel on the right."""
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, MIST)
    s = dict(s, tag=s.get("tag", "Try it"))
    y = header(sl, s, size=36)
    steps = s.get("steps", [])[:5]
    lw = Inches(7.4)
    gap = Inches(0.18)
    rh = min(Inches(0.95), (H - y - Inches(0.9) - gap * (len(steps) - 1)) / max(1, len(steps)))
    for i, st in enumerate(steps):
        yy = Emu(int(y + (rh + gap) * i))
        rect(sl, M, yy, lw, rh, WHITE, rounded=True, radius=0.25)
        d = min(Inches(0.6), rh - Inches(0.2))
        c = oval(sl, M + Inches(0.2), yy + (rh - d) / 2, d, AMBER if i == 0 else PERIWINKLE)
        tf = c.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        add_runs(p, str(i + 1), DISPLAY, 18, ONYX, bold=True)
        textbox(sl, M + d + Inches(0.45), yy, lw - d - Inches(0.7), rh, st, font=BODY, size=18, color=ONYX,
                anchor=MSO_ANCHOR.MIDDLE, line=1.2, hl=BLUE_50)
    px = M + lw + Inches(0.4)
    pw = CONTENT_W - lw - Inches(0.4)
    ph = H - y - Inches(0.9)
    rect(sl, px, y, pw, ph, PAPAYA, rounded=True, radius=0.1)
    rows = [("Time", s.get("time"), "clock"), ("Groups", s.get("group"), "users"),
            ("You'll need", s.get("materials"), "package"), ("Bring back", s.get("output"), "share-2")]
    ry = y + Inches(0.4)
    for label, val, ic in rows:
        if not val:
            continue
        icon(sl, ic, px + Inches(0.35), ry + Inches(0.02), Inches(0.42), "deep-blue")
        textbox(sl, px + Inches(0.95), ry, pw - Inches(1.2), Inches(0.3), label.upper(), font=BODY, size=11,
                color=BLUE_50, bold=True, track=60)
        textbox(sl, px + Inches(0.95), ry + Inches(0.28), pw - Inches(1.2), Inches(0.8), val, font=BODY, size=16,
                color=ONYX, line=1.25)
        ry += Inches(0.55) + Pt(16) * 1.25 * est_lines(val, pw - Inches(1.2), 16) + Inches(0.15)
    footer(sl, n, deck.get("footer"))
    return sl


def l_scenario(prs, s, n, deck):
    """A short real-life story, then 'What would you do?' with lettered options for a show of hands."""
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, MIST)
    lw = Inches(6.2)
    rect(sl, M, Inches(0.7), lw, H - Inches(1.5), PAPAYA, rounded=True, radius=0.08)
    tag(sl, M + Inches(0.45), Inches(1.1), s.get("tag", "Scenario"))
    icon(sl, "book-heart", M + lw - Inches(1.15), Inches(1.0), Inches(0.7), "deep-blue")
    textbox(sl, M + Inches(0.45), Inches(1.8), lw - Inches(0.9), H - Inches(3.0), s["story"], font=BODY, size=19,
            color=ONYX, italic=True, line=1.45, track=-10, hl=DEEP_BLUE)
    rx = M + lw + Inches(0.6)
    rw = CONTENT_W - lw - Inches(0.6)
    q = s.get("question", "What would you do?")
    textbox(sl, rx, Inches(0.9), rw, Inches(1.4), q, font=DISPLAY, size=30, color=DEEP_BLUE, bold=True, line=1.1,
            track=-20, hl=BLUE_50)
    qy = Inches(0.9) + Pt(30) * 1.15 * est_lines(q, rw, 30, bold=True) + Inches(0.35)
    opts = s.get("options", [])[:4]
    gap = Inches(0.16)
    oh = min(Inches(0.9), (H - qy - Inches(1.0) - gap * (len(opts) - 1)) / max(1, len(opts)))
    for i, o in enumerate(opts):
        yy = Emu(int(qy + (oh + gap) * i))
        rect(sl, rx, yy, rw, oh, WHITE, rounded=True, radius=0.3)
        d = oh - Inches(0.24)
        c = oval(sl, rx + Inches(0.12), yy + Inches(0.12), d, PERIWINKLE)
        tf = c.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        add_runs(p, "ABCD"[i], DISPLAY, 16, ONYX, bold=True)
        textbox(sl, rx + d + Inches(0.3), yy, rw - d - Inches(0.5), oh, o, font=BODY, size=16, color=ONYX,
                anchor=MSO_ANCHOR.MIDDLE, line=1.2)
    footer(sl, n, deck.get("footer"))
    return sl


def l_recap(prs, s, n, deck):
    """What we learned (checked rows) plus one personal next step."""
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, MIST)
    s = dict(s, title=s.get("title", "What we learned today"))
    y = header(sl, s, size=38)
    pts = s.get("points", [])[:4]
    has_next = bool(s.get("next_step"))
    lw = Inches(7.3) if has_next else CONTENT_W
    gap = Inches(0.16)
    rh = min(Inches(0.9), (H - y - Inches(0.9) - gap * (len(pts) - 1)) / max(1, len(pts)))
    for i, p_ in enumerate(pts):
        yy = Emu(int(y + (rh + gap) * i))
        rect(sl, M, yy, lw, rh, WHITE, rounded=True, radius=0.25)
        d = min(Inches(0.56), rh - Inches(0.2))
        icon_badge(sl, "circle-check", M + Inches(0.2), yy + (rh - d) / 2, d)
        textbox(sl, M + d + Inches(0.45), yy, lw - d - Inches(0.7), rh, p_, font=BODY, size=18, color=ONYX,
                anchor=MSO_ANCHOR.MIDDLE, line=1.2, hl=BLUE_50)
    if has_next:
        px = M + lw + Inches(0.4)
        pw = CONTENT_W - lw - Inches(0.4)
        ph = H - y - Inches(0.9)
        rect(sl, px, y, pw, ph, DEEP_BLUE, rounded=True, radius=0.1)
        icon(sl, "footprints", px + Inches(0.4), y + Inches(0.4), Inches(0.7), "periwinkle")
        textbox(sl, px + Inches(0.4), y + Inches(1.3), pw - Inches(0.8), Inches(0.5), s.get("next_label", "Your next step"),
                font=BODY, size=14, color=PERI_82, bold=True, track=40)
        textbox(sl, px + Inches(0.4), y + Inches(1.75), pw - Inches(0.8), ph - Inches(2.1), s["next_step"],
                font=DISPLAY, size=22, color=MIST, bold=True, line=1.2, track=-20, hl=AMBER)
    footer(sl, n, deck.get("footer"))
    return sl


def l_break(prs, s, n, deck):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, DEEP_BLUE)
    d = Inches(1.6)
    icon(sl, s.get("icon", "coffee"), (W - d) / 2, Inches(1.5), d, "periwinkle")
    textbox(sl, M, Inches(3.4), CONTENT_W, Inches(1.2), s.get("title", "Let's take a break"), font=DISPLAY, size=54,
            color=MIST, bold=True, align=PP_ALIGN.CENTER, track=-20)
    if s.get("time"):
        tw = Emu(int(Pt(16) * 0.56 * len(s["time"]) + Inches(0.9)))
        chip(sl, (W - tw) / 2, Inches(4.75), s["time"], "clock", fill=AMBER, size=16)
    if s.get("text"):
        textbox(sl, M, Inches(5.5), CONTENT_W, Inches(0.8), s["text"], font=BODY, size=18, color=PERI_82,
                align=PP_ALIGN.CENTER)
    return sl


def t_closing(prs, s, n, deck):
    """Warm close: thanks, an open invitation for questions, and how to stay in touch."""
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, PERIWINKLE)
    oval(sl, Inches(-1.2), Inches(4.6), Inches(4.2), PAPAYA)
    oval(sl, Inches(11.2), Inches(-0.9), Inches(2.8), DEEP_BLUE)
    oval(sl, Inches(10.6), Inches(1.5), Inches(0.7), AMBER)
    textbox(sl, M, Inches(1.6), CONTENT_W, Inches(1.2), s.get("title", "Thank you!"), font=DISPLAY, size=60,
            color=DEEP_BLUE, bold=True, align=PP_ALIGN.CENTER, track=-20)
    textbox(sl, M, Inches(2.9), CONTENT_W, Inches(0.8), s.get("subtitle", "What questions are still on your mind?"),
            font=DISPLAY, size=26, color=ONYX, align=PP_ALIGN.CENTER, track=-20)
    if s.get("contact"):
        textbox(sl, M, Inches(4.2), CONTENT_W, Inches(0.9), s["contact"], font=BODY, size=16, color=ONYX,
                align=PP_ALIGN.CENTER, line=1.4)
    pic = logo(sl, "wordmark-horizontal-deep-blue", 0, H - Inches(1.05), Inches(0.36))
    pic.left = int((W - pic.width) / 2)
    return sl


TEACHING_OVERRIDES = {"title": t_title, "content": t_content, "agenda": t_agenda, "closing": t_closing}


# ---- Teaching style (default): calm, editorial, white ------------------------
# Modelled on Frontier Commons' own Fall Fellows "Plug In" sessions: white slides with a grey
# dot-grid band on the left, bold Deep Blue titles, one highlighted word in Blue 35%, spark
# bullets, uppercase letter-spaced labels over short rules, hairline dividers, no footers.
BLUE_35 = RGBColor(0x15, 0x21, 0x9D)
RULE = RGBColor(0xD6, 0xD8, 0xDB)
TEX = ASSETS / "textures"
EM = Inches(0.83)            # editorial side margin (120px at 1920)
ETOP = Inches(0.67)          # title top (96px)


def set_alpha(shape, opacity):
    """Fill opacity 0-1 on an autoshape (PowerPoint alpha)."""
    srgb = shape.fill._xPr.find(qn("a:solidFill")).find(qn("a:srgbClr"))
    a = srgb.makeelement(qn("a:alpha"), {"val": str(int(opacity * 100000))})
    srgb.append(a)


def hairline(sl, x, y, w, color=RULE):
    return rect(sl, x, y, w, Pt(1), color)


def label(sl, x, y, w, text, size=20, color=BLUE_35, align=PP_ALIGN.CENTER, h=None):
    """Uppercase, letter-spaced DM Sans label (the guide's section-header treatment)."""
    return textbox(sl, x, y, w, h or Inches(0.8), str(text).upper(), font=BODY, size=size, color=color, bold=True,
                   align=align, track=160, line=1.2)


def auto_highlight(text):
    """Highlight the last line (or last word) of a statement when the writer didn't mark one."""
    if "**" in text:
        return text
    if "\n" in text:
        head, last = text.rsplit("\n", 1)
        return f"{head}\n**{last}**"
    words = text.split(" ")
    return " ".join(words[:-1] + [f"**{words[-1]}**"]) if len(words) > 1 else text


def e_header(sl, s, size=26):
    y = ETOP
    if s.get("tag"):
        label(sl, EM, y, Inches(8), s["tag"], size=13, align=PP_ALIGN.LEFT, h=Inches(0.35))
        y += Inches(0.4)
    textbox(sl, EM, y, W - 2 * EM, Inches(1.3), s.get("title", ""), font=DISPLAY, size=size, color=DEEP_BLUE,
            bold=True, track=-35, line=1.0, hl=BLUE_35)
    y += Pt(size) * 1.05 * est_lines(s.get("title", ""), W - 2 * EM, size, bold=True) + Inches(0.12)
    if s.get("lead"):
        textbox(sl, EM, y, W - 2 * EM, Inches(0.5), s["lead"], font=BODY, size=15, color=ONYX_40)
        y += Inches(0.45)
    return y + Inches(0.25)


def spark_list(sl, x, y, w, h, items, size=22, even=True):
    """Spark-bulleted points, spread evenly down the area like the Plug In slides."""
    heights = [Pt(size) * 1.3 * est_lines(t, w - Inches(0.55), size) for t in items]
    free = h - sum(heights)
    gap = free / (len(items) + 1) if even else Inches(0.3)
    gap = max(Inches(0.18), min(gap, Inches(0.9)))
    yy = y + (gap if even else 0)
    for t, hh in zip(items, heights):
        icon(sl, "spark", x, yy + Pt(size) * 0.18, Inches(0.3), "blue")
        textbox(sl, x + Inches(0.52), yy, w - Inches(0.55), hh + Inches(0.1), t, font=BODY, size=size, color=ONYX,
                line=1.3, hl=BLUE_35)
        yy += hh + gap


def e_content(prs, s, n, deck):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, MIST)
    y = e_header(sl, s)
    if s.get("body"):
        textbox(sl, EM + Inches(0.9), y, Inches(10.4), Inches(1.4), s["body"], font=BODY, size=20, color=ONYX,
                line=1.4, hl=BLUE_35)
        y += Pt(20) * 1.4 * est_lines(s["body"], Inches(10.4), 20) + Inches(0.2)
    items = s.get("bullets") or []
    area = H - y - Inches(0.55)
    if len(items) > 5:
        half = (len(items) + 1) // 2
        cw = (W - 2 * EM - Inches(0.9) - Inches(0.6)) / 2
        spark_list(sl, EM + Inches(0.9), y, Emu(int(cw)), area, items[:half], size=19)
        spark_list(sl, EM + Inches(0.9) + Emu(int(cw + Inches(0.6))), y, Emu(int(cw)), area, items[half:], size=19)
    elif items:
        spark_list(sl, EM + Inches(0.9), y, Inches(10.9), area, items)
    return sl


def e_definition(prs, s, n, deck):
    """Title plus one big centred statement ("What [topic] is")."""
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, MIST)
    y = e_header(sl, s)
    text = s.get("text") or s.get("body", "")
    size = 36 if len(text) <= 110 else 30
    textbox(sl, Inches(1.8), y, W - Inches(3.6), H - y - Inches(0.9), text, font=DISPLAY, size=size,
            color=DEEP_BLUE, line=1.14, track=-30, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, hl=BLUE_35)
    return sl


def e_statement(prs, s, n, deck):
    """A centred statement slide: section openers, pauses, big ideas."""
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, MIST)
    text = s.get("text") or s.get("title", "")
    marked = auto_highlight(text) if s.get("highlight", True) else text
    size = s.get("size") or (42 if len(text) <= 40 else 34 if len(text) <= 90 else 30)
    top = Inches(1.2)
    if s.get("kicker") or s.get("tag"):
        label(sl, Inches(1.5), Inches(2.0), W - Inches(3), s.get("kicker") or s.get("tag"), size=14, h=Inches(0.4))
    textbox(sl, Inches(1.5), top, W - Inches(3), H - Inches(2.4), marked, font=DISPLAY, size=size, color=DEEP_BLUE,
            line=1.05, track=-35, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, hl=BLUE_35)
    if s.get("time"):
        label(sl, Inches(1.5), H - Inches(1.5), W - Inches(3), s["time"], size=14, h=Inches(0.4))
    return sl


def e_section(prs, s, n, deck):
    return e_statement(prs, dict(s, text=s["title"]), n, deck)


def l_pillars(prs, s, n, deck):
    """2-4 typographic columns: LABEL, short rule, a sentence, optional reference ("Where it shows up")."""
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, WHITE)
    y = e_header(sl, s)
    items = s["items"][:4]
    k = len(items)
    spots = consistent_icons([resolve_icon(it, s, it.get("label") or it.get("title"), it.get("text")) for it in items]) \
        if s.get("icons", True) is not False else [None] * k
    spot_h = Inches(1.3) if any(spots) else 0
    pad = Inches(0.3)
    gap = Inches(0.55)
    cw = (W - 2 * EM - 2 * pad - gap * (k - 1)) / k
    size = 17 if k == 4 else 18
    lab_lines = max(est_lines(str(it.get("label") or it.get("title")).upper(), cw, 20, bold=True) for it in items)
    lab_h = Pt(20) * 1.25 * lab_lines * 1.15
    blocks = [spot_h + lab_h + Inches(0.35) + Pt(size) * 1.35 * est_lines(it.get("text", ""), cw, size)
              + (Inches(0.45) if it.get("ref") else 0) for it in items]
    area = H - y - Inches(0.8)
    top = y + max(Inches(0.1), (area - max(blocks)) / 2)
    for i, it in enumerate(items):
        x = EM + pad + Emu(int((cw + gap) * i))
        iw = Emu(int(cw))
        if spots[i]:
            ss = Inches(1.15)
            spot(sl, spots[i], x + Emu(int((cw - ss) / 2)), top, ss, disc=[AMBER, PERIWINKLE][i % 2])
        lt = top + spot_h
        label(sl, x, lt, iw, it.get("label") or it.get("title"), h=lab_h)
        rw = Inches(1.3)
        rect(sl, x + Emu(int((cw - rw) / 2)), lt + lab_h + Inches(0.1), rw, Pt(1), PERIWINKLE)
        text = it.get("text", "")
        quote = it.get("quote") or text.startswith(("“", '"'))
        ty = lt + lab_h + Inches(0.3)
        textbox(sl, x, ty, iw, Inches(2.6), text, font=BODY, size=size, color=ONYX, italic=bool(quote), line=1.35,
                align=PP_ALIGN.CENTER, hl=BLUE_35)
        if it.get("ref"):
            ry = ty + Pt(size) * 1.35 * est_lines(text, cw, size) + Inches(0.15)
            textbox(sl, x, ry, iw, Inches(0.4), it["ref"], font=BODY, size=15, color=DEEP_BLUE, bold=True,
                    align=PP_ALIGN.CENTER)
    if s.get("note"):
        textbox(sl, EM, H - Inches(0.75), W - 2 * EM, Inches(0.35), s["note"], font=BODY, size=12, color=ONYX_40,
                italic=True, align=PP_ALIGN.CENTER)
    return sl


def numbered_rows(sl, x, y, w, items, size=23, num_color=BLUE_35, right=None):
    """Hairline-separated numbered rows (reflection questions, agenda, steps). `right` = optional right labels."""
    for i, t in enumerate(items):
        hairline(sl, x, y, w)
        yy = y + Inches(0.24)
        textbox(sl, x, yy + Pt(size) * 0.1, Inches(0.5), Inches(0.5), str(i + 1), font=DISPLAY, size=int(size * 0.74),
                color=num_color, bold=True, track=-30)
        tw = w - Inches(0.6) - (Inches(1.6) if right else 0)
        textbox(sl, x + Inches(0.6), yy, tw, Inches(1.2), t, font=DISPLAY, size=size, color=DEEP_BLUE, line=1.2,
                track=-25, hl=BLUE_35)
        if right and right[i]:
            label(sl, x + w - Inches(1.6), yy + Pt(size) * 0.2, Inches(1.6), right[i], size=12, align=PP_ALIGN.RIGHT,
                  h=Inches(0.35))
        y = yy + Pt(size) * 1.2 * est_lines(t, tw, size) + Inches(0.34)
    return y


def e_questions(prs, s, n, deck):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, MIST)
    y = e_header(sl, dict(s, title=s.get("title", "Reflection questions")))
    y += Inches(0.2)
    y = numbered_rows(sl, EM, y, Inches(9.9), s["questions"])
    meta = " · ".join(x for x in (s.get("format"), s.get("time")) if x)
    if meta:
        label(sl, EM, y + Inches(0.2), Inches(6), meta, size=13, align=PP_ALIGN.LEFT, h=Inches(0.35))
    return sl


def e_question(prs, s, n, deck):
    """One open question for the room, centred, with format/time and follow-up prompts."""
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, MIST)
    meta = " · ".join(x for x in (s.get("tag", "Discuss"), s.get("format"), s.get("time")) if x)
    label(sl, Inches(1.5), Inches(1.25), W - Inches(3), meta, size=14, h=Inches(0.4))
    q = s["question"]
    size = 40 if len(q) <= 80 else 34
    prompts = s.get("prompts") or []
    qh = Inches(3.2) if prompts else Inches(4.3)
    textbox(sl, Inches(1.6), Inches(1.75), W - Inches(3.2), qh, q, font=DISPLAY, size=size, color=DEEP_BLUE,
            line=1.14, track=-30, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, hl=BLUE_35)
    if prompts:
        yy = Inches(1.75) + qh + Inches(0.25)
        hairline(sl, Inches(3.0), yy, W - Inches(6.0))
        textbox(sl, Inches(2.5), yy + Inches(0.3), W - Inches(5.0), Inches(1.4), "\n".join(prompts[:3]), font=BODY,
                size=16, color=ONYX, italic=True, line=1.5, align=PP_ALIGN.CENTER)
    return sl


def e_quote(prs, s, n, deck):
    """A big centred quote, with an optional reflection question under a hairline."""
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, MIST)
    q = s["quote"].strip()
    if not q.startswith(("“", '"')):
        q = f"“{q}”"
    size = 36 if len(q) <= 120 else 30
    follow = s.get("question") or s.get("prompt")
    textbox(sl, Inches(1.6), Inches(1.0), W - Inches(3.2), Inches(3.6) if follow else Inches(4.8), q, font=DISPLAY,
            size=size, color=DEEP_BLUE, line=1.14, track=-30, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
            hl=BLUE_35)
    y = Inches(4.75)
    if s.get("author"):
        who = s["author"] + (f", {s['role']}" if s.get("role") else "")
        textbox(sl, Inches(2), y - Inches(0.1), W - Inches(4), Inches(0.4), who, font=BODY, size=15, color=ONYX_40,
                align=PP_ALIGN.CENTER)
        y += Inches(0.45)
    if follow:
        hairline(sl, Inches(3.0), y, W - Inches(6.0))
        textbox(sl, Inches(2.5), y + Inches(0.3), W - Inches(5.0), Inches(1.2), follow, font=BODY, size=15,
                color=ONYX, italic=True, line=1.4, align=PP_ALIGN.CENTER)
    return sl


def e_recap(prs, s, n, deck):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, MIST)
    y = e_header(sl, dict(s, title=s.get("title", "Recap")))
    pts = s.get("points", [])[:5]
    bottom = H - Inches(0.55) - (Inches(1.4) if s.get("next_step") else 0)
    spark_list(sl, EM + Inches(0.9), y, Inches(10.9), bottom - y, pts)
    if s.get("next_step"):
        yy = bottom + Inches(0.1)
        hairline(sl, EM + Inches(0.9), yy, Inches(10.9))
        label(sl, EM + Inches(0.9), yy + Inches(0.25), Inches(4), s.get("next_label", "Your next step"), size=13,
              align=PP_ALIGN.LEFT, h=Inches(0.35))
        textbox(sl, EM + Inches(0.9), yy + Inches(0.6), Inches(10.9), Inches(0.7), s["next_step"], font=DISPLAY,
                size=22, color=DEEP_BLUE, line=1.2, track=-25, hl=BLUE_35)
    return sl


def e_break(prs, s, n, deck):
    return e_statement(prs, dict(s, text=s.get("title", "Let's take a\npause")), n, deck)


def l_tool(prs, s, n, deck):
    """A method or template (Empathy Map, Personas): labelled rows on the left, an example image on the right."""
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, MIST)
    y = e_header(sl, s)
    rows = s.get("rows") or [{"label": k.replace("_", " "), "text": s[k]} for k in ("purpose", "goal", "keep_in_mind")
                             if s.get(k)]
    lw = Inches(3.75)
    yy = y + Inches(0.05)
    for i, r in enumerate(rows[:4]):
        if i:
            hairline(sl, EM, yy - Inches(0.12), lw)
        label(sl, EM, yy, lw, r["label"], size=12, align=PP_ALIGN.LEFT, h=Inches(0.3))
        textbox(sl, EM, yy + Inches(0.32), lw, Inches(1.4), r["text"], font=BODY, size=14, color=ONYX, line=1.35)
        yy += Inches(0.32) + Pt(14) * 1.35 * est_lines(r["text"], lw, 14) + Inches(0.35)
    fx = EM + lw + Inches(0.45)
    fw = W - EM - fx
    fh = H - y - Inches(0.55)
    frame = rect(sl, fx, y, fw, fh, WHITE)
    frame.line.color.rgb = RULE
    frame.line.width = Pt(0.75)
    if s.get("image"):
        im = Image.open(DECK_DIR / s["image"])
        iw, ih = im.size
        scale = min((fw - Inches(0.3)) / iw, (fh - Inches(0.3)) / ih)
        pw, ph = int(iw * scale), int(ih * scale)
        sl.shapes.add_picture(str(DECK_DIR / s["image"]), fx + (fw - pw) // 2, y + (fh - ph) // 2, pw, ph)
    else:
        textbox(sl, fx, y, fw, fh, s.get("image_caption", "[Add an example image]"), font=BODY, size=14,
                color=ONYX_40, italic=True, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    return sl


def e_agenda(prs, s, n, deck):
    items = [i if isinstance(i, dict) else {"title": i} for i in s["items"]]
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, MIST)
    y = e_header(sl, dict(s, title=s.get("title", "Today")))
    numbered_rows(sl, EM, y + Inches(0.15), Inches(10.8), [i["title"] for i in items], size=23,
                  right=[i.get("time") for i in items])
    return sl


def e_dark(sl, image=None):
    """Deep Blue ground with the title treatment: optional photo under an overlay, a skewed band and dots."""
    bg(sl, DEEP_BLUE)
    if image:
        picture(sl, DECK_DIR / image, Inches(2.58), Inches(0.62), Inches(9.6), Inches(5.7), radius_in=0)
        ov = rect(sl, 0, 0, W, H, DEEP_BLUE)
        set_alpha(ov, 0.82)
    band = sl.shapes.add_shape(MSO_SHAPE.PARALLELOGRAM, Inches(-1.4), Inches(-0.2), Inches(4.4), H + Inches(0.6))
    band.fill.solid()
    band.fill.fore_color.rgb = BLUE_35
    band.line.fill.background()
    band.shadow.inherit = False
    band.adjustments[0] = 0.28
    set_alpha(band, 0.55)
    sl.shapes.add_picture(str(TEX / "dots-periwinkle.png"), Inches(-0.28), H - Inches(5.28), Inches(2.5), Inches(5.28))


def e_title(prs, s, n, deck):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    e_dark(sl, s.get("image"))
    x = Inches(0.76)
    label(sl, x, Inches(1.35), Inches(8), s.get("kicker", "Frontier Commons"), size=17, color=WHITE,
          align=PP_ALIGN.LEFT, h=Inches(0.45))
    title = s["title"].upper()
    lines = title.count("\n") + 1
    size = 100 if lines <= 2 and max(len(l) for l in title.split("\n")) <= 12 else 72
    ty = Inches(1.95)
    textbox(sl, x, ty, Inches(11.5), Pt(size) * 0.95 * lines + Inches(0.3), title, font=DISPLAY, size=size,
            color=WHITE, bold=True, line=0.88, track=-45)
    sy = ty + Pt(size) * 0.9 * lines + Inches(0.15)
    if s.get("series"):
        textbox(sl, x, sy, Inches(11.5), Inches(1.4), s["series"].upper(), font=DISPLAY, size=int(size * 0.8),
                color=AMBER, bold=True, italic=True, line=0.95, track=-30)
    elif s.get("subtitle"):
        textbox(sl, x, sy + Inches(0.3), Inches(10), Inches(1.2), s["subtitle"], font=DISPLAY, size=28, color=PERI_82,
                line=1.2)
    info = s.get("info") or "  ·  ".join(x_ for x_ in (s.get("session"), s.get("date")) if x_)
    if info:
        textbox(sl, x, H - Inches(0.95), Inches(8), Inches(0.4), info, font=BODY, size=14, color=WHITE)
    if s.get("org"):
        textbox(sl, W - Inches(6.8), H - Inches(0.95), Inches(6.1), Inches(0.4), s["org"], font=BODY, size=14,
                color=WHITE, align=PP_ALIGN.RIGHT)
    else:
        logo(sl, "wordmark-horizontal-mist", W - Inches(3.7), H - Inches(0.98), Inches(0.32))
    return sl


def e_closing(prs, s, n, deck):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    e_dark(sl)
    x = Inches(0.76)
    textbox(sl, x, Inches(2.0), Inches(11.5), Inches(1.6), s.get("title", "Thank you").upper(), font=DISPLAY,
            size=88, color=WHITE, bold=True, line=0.9, track=-45)
    if s.get("subtitle"):
        textbox(sl, x, Inches(3.55), Inches(11), Inches(1.0), s["subtitle"], font=DISPLAY, size=30, color=AMBER,
                bold=True, italic=True, track=-20)
    if s.get("contact"):
        textbox(sl, x, Inches(4.7), Inches(9), Inches(1.0), s["contact"], font=BODY, size=16, color=WHITE, line=1.4)
    logo(sl, "wordmark-horizontal-mist", W - Inches(3.7), H - Inches(0.98), Inches(0.32))
    return sl


def e_tps(prs, s, n, deck):
    steps = [{"label": f"Think · {s.get('think_time', '2 min')}", "text": s.get("think", "On your own, jot down your first thoughts.")},
             {"label": f"Pair · {s.get('pair_time', '4 min')}", "text": s.get("pair", "Share with the person next to you. What's similar? What's different?")},
             {"label": f"Share · {s.get('share_time', '5 min')}", "text": s.get("share", "A few pairs share one idea with the whole room.")}]
    return l_pillars(prs, dict(s, title=s["question"], items=steps, tag=s.get("tag", "Think · Pair · Share")), n, deck)


def e_activity(prs, s, n, deck):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, MIST)
    y = e_header(sl, dict(s, tag=s.get("tag", "Try it")))
    numbered_rows(sl, EM, y + Inches(0.1), Inches(7.4), s.get("steps", [])[:5], size=20)
    px = EM + Inches(8.0)
    pw = W - EM - px
    yy = y + Inches(0.1)
    for lab, val in (("Time", s.get("time")), ("Groups", s.get("group")), ("You'll need", s.get("materials")),
                     ("Bring back", s.get("output"))):
        if not val:
            continue
        hairline(sl, px, yy, pw)
        label(sl, px, yy + Inches(0.2), pw, lab, size=12, align=PP_ALIGN.LEFT, h=Inches(0.3))
        textbox(sl, px, yy + Inches(0.5), pw, Inches(0.9), val, font=BODY, size=16, color=ONYX, line=1.3)
        yy += Inches(0.5) + Pt(16) * 1.3 * est_lines(val, pw, 16) + Inches(0.3)
    return sl


def e_scenario(prs, s, n, deck):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, MIST)
    label(sl, EM, ETOP, Inches(5), s.get("tag", "Scenario"), size=13, align=PP_ALIGN.LEFT, h=Inches(0.35))
    lw = Inches(6.1)
    textbox(sl, EM, ETOP + Inches(0.5), lw, H - ETOP - Inches(1.2), s["story"], font=DISPLAY, size=26,
            color=DEEP_BLUE, line=1.25, track=-25, anchor=MSO_ANCHOR.MIDDLE, hl=BLUE_35)
    rx = EM + lw + Inches(0.7)
    rw = W - EM - rx
    q = s.get("question", "What would you do?")
    textbox(sl, rx, ETOP + Inches(0.5), rw, Inches(1.0), q, font=DISPLAY, size=26, color=DEEP_BLUE, bold=True,
            track=-30, line=1.1)
    y = ETOP + Inches(0.6) + Pt(26) * 1.15 * est_lines(q, rw, 26, bold=True) + Inches(0.2)
    for i, o in enumerate(s.get("options", [])[:4]):
        hairline(sl, rx, y, rw)
        textbox(sl, rx, y + Inches(0.22), Inches(0.5), Inches(0.5), "ABCD"[i], font=DISPLAY, size=17, color=BLUE_35,
                bold=True)
        textbox(sl, rx + Inches(0.5), y + Inches(0.18), rw - Inches(0.5), Inches(1), o, font=BODY, size=17,
                color=ONYX, line=1.3)
        y += Inches(0.18) + Pt(17) * 1.3 * est_lines(o, rw - Inches(0.5), 17) + Inches(0.3)
    return sl


EDITORIAL_OVERRIDES = {"title": e_title, "content": e_content, "section": e_section, "statement": e_statement,
                       "questions": e_questions, "question": e_question, "quote": e_quote, "recap": e_recap,
                       "break": e_break, "agenda": e_agenda, "closing": e_closing, "think_pair_share": e_tps,
                       "activity": e_activity, "scenario": e_scenario}


# ---- Teaching style, illustrated (chosen direction) --------------------------
# Spot illustrations built only from brand shapes: a Periwinkle disc, a Deep Blue rounded tile holding a
# Mist line icon, an Amber dot and a navy dot texture; smaller "spots" are a coloured disc behind a
# Deep Blue line icon (the guide's line-art-with-one-fill look). Bullet slides use a Papaya side panel.
def illo(sl, name, x, y, s, disc=None):
    disc = disc or PERIWINKLE
    x, y, s = int(x), int(y), int(s)
    oval(sl, x, y, int(s * 0.82), disc)
    sl.shapes.add_picture(str(TEX / "dots-navy.png"), x + int(s * 0.66), y + int(s * 0.02), int(s * 0.3), int(s * 0.3))
    rect(sl, x + int(s * 0.30), y + int(s * 0.34), int(s * 0.58), int(s * 0.58), DEEP_BLUE, rounded=True, radius=0.2)
    icon(sl, name, x + int(s * 0.38), y + int(s * 0.42), int(s * 0.42), "mist")
    oval(sl, x + int(s * 0.12), y + int(s * 0.70), int(s * 0.16), AMBER)


def spot(sl, name, x, y, s, disc=AMBER):
    """Coloured disc with a Deep Blue line icon overlapping it."""
    x, y, s = int(x), int(y), int(s)
    oval(sl, x, y + int(s * 0.1), int(s * 0.75), disc)
    icon(sl, name, x + int(s * 0.2), y + int(s * 0.05), int(s * 0.75), "deep-blue")


def topic_icon(s, *texts):
    return s.get("icon") if s.get("icon") in ICON_CATALOG else (pick_icon(*texts) or "sparkles")


def i_section(prs, s, n, deck, text=None, icon_name=None, graphic=None):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, PAPAYA)
    text = text or s["title"]
    marked = auto_highlight(text)
    lines = text.count("\n") + 1
    longest = max(len(l) for l in text.split("\n"))
    size = 62 if longest <= 14 else 50 if longest <= 20 else 40
    block = Pt(size) * 0.98 * lines
    top = (H - block) / 2
    if s.get("kicker"):
        label(sl, Inches(0.9), top - Inches(0.6), Inches(6), s["kicker"], size=14, align=PP_ALIGN.LEFT, h=Inches(0.4))
    textbox(sl, Inches(0.9), top, Inches(6.9), block + Inches(0.4), marked, font=DISPLAY, size=size, color=DEEP_BLUE,
            bold=True, line=0.95, track=-45, hl=BLUE_35)
    if s.get("time"):
        label(sl, Inches(0.9), top + block + Inches(0.35), Inches(6), s["time"], size=14, align=PP_ALIGN.LEFT,
              h=Inches(0.4))
    if graphic:
        graphic(sl)
    else:
        illo(sl, icon_name or topic_icon(s, text.replace("\n", " ")), Inches(8.3), Inches(1.75), Inches(3.8))
    return sl


def ripples(sl, cx=Inches(10.2), cy=Inches(3.75), r=Inches(1.9)):
    """Calm 'breathe' graphic for pauses: concentric Periwinkle rings around an Amber centre."""
    for scale, colour in ((1.0, RGBColor(0xEC, 0xEA, 0xFA)), (0.76, PERI_91), (0.54, RGBColor(0xBC, 0xB4, 0xEE)),
                          (0.33, PERIWINKLE)):
        rr = int(r * scale)
        oval(sl, cx - rr, cy - rr, 2 * rr, colour)
    rr = int(r * 0.14)
    oval(sl, cx - rr, cy - rr, 2 * rr, AMBER)
    sl.shapes.add_picture(str(TEX / "dots-navy.png"), cx + int(r * 0.62), cy - int(r * 1.12), int(r * 0.55), int(r * 0.55))


def i_break(prs, s, n, deck):
    return i_section(prs, s, n, deck, text=s.get("title", "Let's take a\npause"),
                     graphic=None if s.get("icon") else ripples, icon_name=s.get("icon"))


PANEL_ICON = Inches(1.4)


def panel(sl, s, deck, title, icon_name, max_d=None):
    """Papaya panel on the left with a kicker, the title and a big Amber-disc illustration."""
    bg(sl, WHITE)
    pw = Inches(4.86)
    p = rect(sl, -Inches(0.8), 0, pw + Inches(0.8), H, PAPAYA, rounded=True, radius=0.1)
    kicker = s.get("kicker") or deck.get("kicker")
    if kicker:  # small label at the top, as originally
        label(sl, Inches(0.76), Inches(0.76), pw - Inches(1.2), kicker, size=12, align=PP_ALIGN.LEFT, h=Inches(0.3))
    y = Inches(1.68)  # title sits lower and larger than the label
    size = 37 if len(title) <= 40 else 30 if len(title) <= 60 else 25
    tw = pw - Inches(1.3)
    textbox(sl, Inches(0.76), y, tw, Inches(3.2), title, font=DISPLAY, size=size, color=DEEP_BLUE,
            bold=True, line=1.0, track=-35, hl=BLUE_35)
    title_bottom = y + Pt(size) * 1.08 * est_lines(title, tw, size, bold=True)
    room = H - Inches(0.9) - title_bottom - Inches(0.5)          # space left for the illustration
    d = int(max_d or PANEL_ICON)  # one size on every bullet slide, so the deck feels consistent
    if room >= d * 1.12:
        oval(sl, Inches(0.9), H - d - Inches(0.9), d, AMBER)
        icon(sl, icon_name, Inches(0.9) + int(d * 0.17), H - int(d * 1.12) - Inches(0.9), d, "deep-blue")
    return pw


def i_content(prs, s, n, deck, items=None, title=None, icon_name=None, max_d=None):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    items = items if items is not None else (s.get("bullets") or [])
    title = title or s.get("title", "")
    pw = panel(sl, s, deck, title, icon_name or topic_icon(s, title, " ".join(items)), max_d=max_d)
    x = pw + Inches(0.7)
    w = W - x - Inches(0.7)
    y = Inches(0.8)
    if s.get("lead"):
        textbox(sl, x, y, w, Inches(0.5), s["lead"], font=BODY, size=16, color=ONYX_40)
        y += Inches(0.5)
    if s.get("body"):
        textbox(sl, x, y, w, Inches(1.4), s["body"], font=BODY, size=19, color=ONYX, line=1.4, hl=BLUE_35)
        y += Pt(19) * 1.4 * est_lines(s["body"], w, 19) + Inches(0.2)
    bottom = H - Inches(0.8) - (Inches(1.3) if s.get("next_step") else 0)
    size = 21 if len(items) <= 4 else 18
    if items:
        spark_list(sl, x, y, w, bottom - y, items, size=size)
    if s.get("next_step"):
        yy = bottom + Inches(0.05)
        hairline(sl, x, yy, w)
        label(sl, x, yy + Inches(0.2), Inches(4), s.get("next_label", "Your next step"), size=12,
              align=PP_ALIGN.LEFT, h=Inches(0.3))
        textbox(sl, x, yy + Inches(0.52), w, Inches(0.7), s["next_step"], font=DISPLAY, size=21, color=DEEP_BLUE,
                line=1.2, track=-25, hl=BLUE_35)
    return sl


def i_recap(prs, s, n, deck):
    return i_content(prs, s, n, deck, items=s.get("points", [])[:5], title=s.get("title", "Recap"),
                     icon_name=s.get("icon", "list-checks"))


def bubbles(sl, x, y):
    """Two overlapping speech bubbles with text bars: the discussion illustration."""
    b1 = rect(sl, x, y, Inches(2.64), Inches(1.94), WHITE, rounded=True, radius=0.3)
    t1 = sl.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, x + Inches(0.3), y + Inches(1.78), Inches(0.5), Inches(0.42))
    t1.rotation = 180
    for t in (t1,):
        t.fill.solid(); t.fill.fore_color.rgb = WHITE; t.line.fill.background(); t.shadow.inherit = False
    rect(sl, x + Inches(0.42), y + Inches(0.55), Inches(0.85), Inches(0.11), PERIWINKLE, rounded=True, radius=0.5)
    rect(sl, x + Inches(0.42), y + Inches(0.83), Inches(1.55), Inches(0.11), PERIWINKLE, rounded=True, radius=0.5)
    bx, by = x + Inches(0.62), y + Inches(1.52)
    rect(sl, bx, by, Inches(2.5), Inches(1.74), PAPAYA, rounded=True, radius=0.3)
    t2 = sl.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, bx + Inches(1.8), by + Inches(1.58), Inches(0.5), Inches(0.42))
    t2.rotation = 180
    t2.fill.solid(); t2.fill.fore_color.rgb = PAPAYA; t2.line.fill.background(); t2.shadow.inherit = False
    rect(sl, bx + Inches(0.42), by + Inches(0.6), Inches(1.4), Inches(0.11), AMBER, rounded=True, radius=0.5)
    rect(sl, bx + Inches(0.42), by + Inches(0.88), Inches(0.95), Inches(0.11), AMBER, rounded=True, radius=0.5)


def i_questions(prs, s, n, deck):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, WHITE)
    lw = Inches(4.45)
    rect(sl, 0, 0, lw, H, PERI_91)
    bubbles(sl, Inches(0.75), Inches(1.9))
    x = lw + Inches(0.8)
    w = W - x - Inches(0.8)
    textbox(sl, x, Inches(0.76), w, Inches(1.1), s.get("title", "Reflection questions"), font=DISPLAY, size=27,
            color=DEEP_BLUE, bold=True, track=-35, line=1.0)
    ty = Inches(0.76) + Pt(32) * 1.05 * est_lines(s.get("title", "Reflection questions"), w, 32, bold=True)
    meta = " · ".join(v for v in (s.get("format"), s.get("time")) if v)
    if meta:
        textbox(sl, x, ty + Inches(0.08), w, Inches(0.4), meta, font=BODY, size=14, color=ONYX_40)
        ty += Inches(0.4)
    y = ty + Inches(0.45)
    size = 22 if len(s["questions"]) <= 3 else 19
    for i, q in enumerate(s["questions"][:4]):
        hairline(sl, x, y, w)
        yy = y + Inches(0.24)
        c = oval(sl, x, yy + Inches(0.02), Inches(0.44), PERIWINKLE)
        tf = c.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        add_runs(p, str(i + 1), DISPLAY, 15, ONYX, bold=True)
        tw = w - Inches(0.7)
        textbox(sl, x + Inches(0.7), yy, tw, Inches(1.2), q, font=DISPLAY, size=size, color=DEEP_BLUE, line=1.22,
                track=-20, hl=BLUE_35)
        y = yy + Pt(size) * 1.22 * est_lines(q, tw, size) + Inches(0.32)
    return sl


def i_definition(prs, s, n, deck):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, MIST)
    title = s.get("title", "")
    text = s.get("text") or s.get("body", "")
    textbox(sl, EM, ETOP, Inches(7.6), Inches(1.0), title, font=DISPLAY, size=26, color=DEEP_BLUE, bold=True,
            track=-35)
    size = 36 if len(text) <= 110 else 30
    textbox(sl, EM, Inches(1.7), Inches(7.3), H - Inches(2.6), text, font=DISPLAY, size=size, color=DEEP_BLUE,
            line=1.15, track=-30, anchor=MSO_ANCHOR.MIDDLE, hl=BLUE_35)
    illo(sl, topic_icon(s, title, text), Inches(9.0), Inches(2.1), Inches(3.3), disc=PERI_91)
    return sl


def i_quote(prs, s, n, deck):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, WHITE)
    spot(sl, s.get("icon", "message-circle"), (W - Inches(1.3)) / 2, Inches(0.55), Inches(1.3), disc=AMBER)
    q = s["quote"].strip()
    if not q.startswith(("“", '"')):
        q = f"“{q}”"
    size = 34 if len(q) <= 120 else 28
    follow = s.get("question") or s.get("prompt")
    textbox(sl, Inches(1.6), Inches(1.9), W - Inches(3.2), Inches(2.9), q, font=DISPLAY, size=size,
            color=DEEP_BLUE, line=1.14, track=-30, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, hl=BLUE_35)
    y = Inches(5.0)
    if s.get("author"):
        who = s["author"] + (f", {s['role']}" if s.get("role") else "")
        textbox(sl, Inches(2), y - Inches(0.1), W - Inches(4), Inches(0.4), who, font=BODY, size=15, color=ONYX_40,
                align=PP_ALIGN.CENTER)
        y += Inches(0.45)
    if follow:
        rect(sl, Inches(1.8), y, W - Inches(3.6), Inches(1.35), PERI_91, rounded=True, radius=0.3)
        textbox(sl, Inches(2.3), y, W - Inches(4.6), Inches(1.35), follow, font=BODY, size=17, color=DEEP_BLUE,
                italic=True, line=1.4, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    return sl


EDITORIAL_OVERRIDES.update({"section": i_section, "break": i_break, "content": i_content, "recap": i_recap,
                            "questions": i_questions, "definition": i_definition, "quote": i_quote})


def title_icon(s):
    return s.get("icon") if s.get("icon") in ICON_CATALOG else (
        pick_icon(s.get("info", ""), s.get("title", ""), s.get("series", ""), s.get("subtitle", "")) or "sparkles")


def pill(sl, x, y, text, fill, ink, size=15):
    h = Inches(0.5)
    w = Emu(int(Pt(size) * 0.58 * len(text) + Inches(0.6)))
    shp = rect(sl, x, y, w, h, fill, rounded=True, radius=0.5)
    textbox(sl, x, y, w, h, text, font=BODY, size=size, color=ink, bold=True, align=PP_ALIGN.CENTER,
            anchor=MSO_ANCHOR.MIDDLE)
    return w


def title_text(sl, s, x, y, w, ink, series_ink, kicker_ink, upper=False, size=60):
    label(sl, x, y, w, s.get("kicker", "Frontier Commons"), size=13, color=kicker_ink, align=PP_ALIGN.LEFT,
          h=Inches(0.35))
    y += Inches(0.55)
    title = s["title"].upper() if upper else s["title"]
    lines = title.count("\n") + 1 if "\n" in title else est_lines(title, w, size, bold=True)
    textbox(sl, x, y, w, Pt(size) * 1.0 * lines + Inches(0.3), title, font=DISPLAY, size=size, color=ink, bold=True,
            line=0.95, track=-40)
    y += Pt(size) * 0.97 * lines + Inches(0.1)
    if s.get("series"):
        ss = int(size * 0.78)
        textbox(sl, x, y, w, Pt(ss) * 1.2, s["series"].upper() if upper else s["series"], font=DISPLAY, size=ss,
                color=series_ink, bold=True, italic=True, line=1.0, track=-30)
        y += Pt(ss) * 1.1 + Inches(0.35)
    elif s.get("subtitle"):
        textbox(sl, x, y + Inches(0.1), w, Inches(1.0), s["subtitle"], font=DISPLAY, size=24, color=series_ink,
                line=1.2)
        y += Inches(1.0)
    return y


def t_title_illustrated(prs, s, n, deck):
    """Papaya title with a large brand-shape composition around the week's topic icon."""
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, PAPAYA)
    oval(sl, Inches(7.7), Inches(0.75), Inches(5.0), PERIWINKLE)
    sl.shapes.add_picture(str(TEX / "dots-navy.png"), Inches(11.2), Inches(0.45), Inches(1.5), Inches(1.5))
    tile = Inches(2.5)
    rect(sl, Inches(7.2), Inches(3.55), tile, tile, DEEP_BLUE, rounded=True, radius=0.2)
    icon(sl, title_icon(s), Inches(7.2) + Inches(0.55), Inches(3.55) + Inches(0.55), tile - Inches(1.1), "mist")
    oval(sl, Inches(10.55), Inches(5.3), Inches(0.95), AMBER)
    oval(sl, Inches(6.75), Inches(2.7), Inches(0.42), PERI_91)
    y = title_text(sl, s, Inches(0.9), Inches(1.45), Inches(6.2), DEEP_BLUE, BLUE_35, BLUE_35, size=58)
    if s.get("info"):
        parts = [p.strip() for p in s["info"].split("·")]
        x = Inches(0.9)
        for p in parts[:3]:
            x += pill(sl, x, y, p, WHITE, DEEP_BLUE, size=14) + Inches(0.12)
    if s.get("org"):
        textbox(sl, Inches(0.9), H - Inches(0.95), Inches(6), Inches(0.4), s["org"], font=BODY, size=13,
                color=ONYX_40)
    else:
        logo(sl, "wordmark-horizontal-deep-blue", Inches(0.9), H - Inches(0.95), Inches(0.3))
    return sl


def t_title_bold(prs, s, n, deck):
    """Deep Blue title with big bleeding circles, the topic icon and the series in Amber."""
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg(sl, DEEP_BLUE)
    if s.get("icon"):  # no topic graphic by default; set "icon" on the title slide to add one
        icon(sl, title_icon(s), Inches(9.3), Inches(2.2), Inches(3.0), "periwinkle")
    sl.shapes.add_picture(str(TEX / "dots-periwinkle.png"), Inches(-0.2), H - Inches(3.6), Inches(1.7), Inches(3.6))
    y = title_text(sl, s, Inches(0.9), Inches(1.35), Inches(6.4), WHITE, AMBER, PERI_82, upper=True, size=64)
    if s.get("info"):
        textbox(sl, Inches(0.9), y, Inches(6.2), Inches(0.5), s["info"], font=BODY, size=15, color=PERI_82)
    if s.get("org"):
        textbox(sl, Inches(0.9), H - Inches(0.95), Inches(6), Inches(0.4), s["org"], font=BODY, size=13, color=MIST)
    else:
        logo(sl, "wordmark-horizontal-mist", Inches(0.9), H - Inches(0.95), Inches(0.3))
    return sl


def e_title_any(prs, s, n, deck):
    v = s.get("variant") or deck.get("title_variant", "bold")
    return {"illustrated": t_title_illustrated, "bold": t_title_bold}.get(v, e_title)(prs, s, n, deck)


EDITORIAL_OVERRIDES["title"] = e_title_any


LAYOUTS = {
    "title": l_title, "section": l_section, "content": l_content, "two_column": l_two_column,
    "cards": l_cards, "stats": l_stats, "process": l_process, "agenda": l_agenda, "quote": l_quote,
    "image": l_image, "statement": l_statement, "closing": l_closing,
    "icon_grid": l_icon_grid, "chart": l_chart,
    "question": l_question, "questions": l_questions, "think_pair_share": l_think_pair_share,
    "activity": l_activity, "scenario": l_scenario, "recap": l_recap, "break": l_break,
    "pillars": l_pillars, "tool": l_tool, "definition": e_definition,
}

# ---- Planner: pick layouts from content, split overflow, lint copy -----------
# How many items fit one slide, per layout (more are split across balanced continuation slides).
CAPACITY = {"content": 8, "cards": 4, "stats": 4, "process": 6, "agenda": 8, "icon_grid": 6, "questions": 4,
            "activity": 5, "recap": 5, "pillars": 4}
LIST_KEY = {"content": "bullets", "cards": "cards", "stats": "stats", "process": "steps", "agenda": "items",
            "icon_grid": "items", "questions": "questions", "activity": "steps", "recap": "points",
            "pillars": "items"}
CLOSED_Q = re.compile(r"^(is|are|was|were|do|does|did|can|could|will|would|should|have|has|had|am)\b", re.I)
LONG_BULLET = 90   # a bullet longer than this counts as "long": at most 4 long bullets per slide
LIMITS = {"title": 60, "bullet": 110, "body": 320, "card": 140, "step": 90, "stat_label": 60, "quote": 280,
          "statement": 110}
NUMBER_RE = re.compile(r"^\s*(\d+ in \d+|[~≈<>+]?\s*[$€£]?\d[\d,.]*\s*(%|[kKmMbB]\b|x\b|\+)?\+?)\s+(.+)$")


def chunks(items, cap):
    """Split into the fewest slides of at most `cap`, as evenly as possible (7 -> 4+3, not 6+1)."""
    k = max(1, -(-len(items) // cap))
    base, extra = divmod(len(items), k)
    out, i = [], 0
    for j in range(k):
        size = base + (1 if j < extra else 0)
        out.append(items[i:i + size])
        i += size
    return out


def as_point(p):
    return p if isinstance(p, dict) else {"text": str(p)}


def auto_layout(s):
    """Choose a layout for {"layout": "auto", "points": [...], "kind"?: ...}. Returns (layout, slide)."""
    pts = [as_point(p) for p in s.get("points", [])]
    kind = s.get("kind", "")
    base = {k: v for k, v in s.items() if k not in ("points", "kind", "layout")}
    n = len(pts)
    if n == 0:
        return ("statement", dict(base, text=s.get("body") or s.get("title", ""))) if not s.get("title") else \
               ("content", base)
    # numbers: explicit values, or points that start with a figure ("7M international students")
    if kind in ("numbers", "stats") or all("value" in p or NUMBER_RE.match(p.get("text", "")) for p in pts):
        stats = []
        for p in pts:
            if "value" in p:
                stats.append({"value": p["value"], "label": p.get("label") or p.get("title") or p.get("text", ""),
                              "detail": p.get("detail") or (p.get("text") if p.get("label") or p.get("title") else None),
                              "icon": p.get("icon")})
            else:
                m = NUMBER_RE.match(p["text"])
                stats.append({"value": m.group(1).strip(), "label": m.group(3).strip(), "icon": p.get("icon")})
        if kind == "chart":  # only chart numbers that share one unit and compare like with like
            return "chart_or_stats", dict(base, stats=stats)
        return "stats", dict(base, stats=stats)
    if kind in ("steps", "process", "journey", "timeline") or any("date" in p for p in pts):
        return "process", dict(base, steps=[{"title": p.get("title") or p.get("text"), "text": p.get("text") if p.get("title") else None,
                                             "date": p.get("date"), "icon": p.get("icon", "auto")} for p in pts])
    if kind == "agenda":
        return "agenda", dict(base, items=[{"title": p.get("title") or p["text"], "time": p.get("time")} for p in pts])
    texts_ = [p.get("text") or p.get("title") or "" for p in pts]
    if kind == "think_pair_share" or kind == "tps":
        return "think_pair_share", dict(base, question=texts_[0])
    if kind == "activity":
        return "activity", dict(base, steps=texts_)
    if kind == "recap":
        return "recap", dict(base, points=texts_)
    if kind == "scenario":
        return "scenario", dict(base, story=base.get("story") or texts_[0], options=texts_[1:] if not base.get("story") else texts_)
    if kind in ("question", "questions", "discussion") or all(t.strip().endswith("?") for t in texts_):
        if n == 1:
            return "question", dict(base, question=texts_[0])
        return "questions", dict(base, questions=texts_)
    if kind == "quote" or (n == 1 and pts[0].get("author")):
        p = pts[0]
        return "quote", dict(base, quote=p.get("text") or p.get("quote"), author=p.get("author"), role=p.get("role"))
    if kind == "definition" or (n == 1 and STYLE == "teaching" and base.get("title")):
        return "definition", dict(base, text=pts[0].get("text") or pts[0].get("title"))
    if n == 1:
        return "statement", dict(base, text=pts[0].get("title") or pts[0]["text"])
    titled = all(p.get("title") for p in pts)
    if kind == "compare" or (n == 2 and titled):
        side = lambda p: {"heading": p.get("title"), "body": p.get("text") if not p.get("bullets") else None,
                          "bullets": p.get("bullets")}
        if n == 2:
            return "two_column", dict(base, left=side(pts[0]), right=side(pts[1]))
    if titled and (kind == "pillars" or (STYLE == "teaching" and n <= 4 and kind != "features"
                                         and not any(p.get("subtitle") for p in pts))):
        return "pillars", dict(base, items=[{"label": p["title"], "text": p.get("text", ""), "ref": p.get("ref"),
                                             "quote": p.get("quote"), "icon": p.get("icon", "auto")} for p in pts])
    if titled and (kind == "principles" or any(p.get("subtitle") for p in pts)) and n <= 8:
        return "cards", dict(base, cards=[{"title": p["title"], "subtitle": p.get("subtitle"), "text": p.get("text"),
                                           "icon": p.get("icon", "auto")} for p in pts])
    if titled:
        return "icon_grid", dict(base, items=[{"title": p["title"], "text": p.get("text"), "icon": p.get("icon", "auto")}
                                              for p in pts])
    # plain sentences
    texts = [p["text"] for p in pts]
    if 3 <= n <= 6 and kind == "features" and all(len(t) <= 40 for t in texts):
        return "icon_grid", dict(base, items=[{"title": t, "icon": p.get("icon", "auto")} for t, p in zip(texts, pts)])
    return "content", dict(base, bullets=texts)


def plan(deck):
    """Expand auto slides, split overflowing ones, and collect copy warnings."""
    out = []
    for idx, s in enumerate(deck["slides"], 1):
        layout = s.get("layout") or ("auto" if "points" in s else "content")
        why = "given"
        if layout == "auto":
            layout, s = auto_layout(s)
            why = f"auto, {len(s.get(LIST_KEY.get(layout, ''), []) or []) or 'n/a'} points"
            if layout == "chart_or_stats":
                st = s.pop("stats")
                vals = [re.sub(r"[^\d.]", "", str(x["value"])) for x in st]
                if all(v and v.count(".") <= 1 for v in vals):
                    layout = "chart"
                    s["chart"] = {"type": "bar", "categories": [x["label"] for x in st],
                                  "values": [float(v) for v in vals]}
                    if all(str(x["value"]).strip().endswith("%") for x in st):
                        s["chart"]["format"] = '0"%"'
                    why = f"auto, {len(st)} numbers -> chart"
                else:
                    layout, s["stats"] = "stats", st
        if layout not in LAYOUTS:
            sys.exit(f"Slide {idx}: unknown layout '{layout}'. Options: auto, {', '.join(LAYOUTS)}")
        key = LIST_KEY.get(layout)
        items = s.get(key) if key else None
        cap = CAPACITY.get(layout, 99)
        if layout == "content" and items:
            long_ = sum(len(str(b)) > LONG_BULLET for b in items)
            cap = 4 if long_ >= 2 else (8 if all(len(str(b)) <= 60 for b in items) else 5)
        if layout == "cards" and items and len(items) > 4:
            layout, key, cap = "icon_grid", "items", 6  # 5+ principles read better as a grid
            s = dict(s, items=[{"title": c["title"], "text": c.get("subtitle") or c.get("text"), "icon": c.get("icon", "auto")}
                               for c in items])
            items = s["items"]
        parts = chunks(items, cap) if items and len(items) > cap else [items]
        for j, part in enumerate(parts):
            sl = dict(s)
            if key and part is not None:
                sl[key] = part
            if len(parts) > 1:
                if j > 0:
                    sl["title"] = f"{s.get('title', '')} (continued)"
                    sl.pop("tag", None)
                    sl.pop("body", None)
                if layout == "process":
                    sl["start"] = 1 + sum(len(p) for p in parts[:j])
            sl["layout"] = layout
            sl["_src"] = idx
            sl["_why"] = why + (f", split {j + 1}/{len(parts)}" if len(parts) > 1 else "")
            out.append(sl)
    return out


def lint(slides):
    warn = []
    if STYLE == "teaching":  # full-sentence points and two-line titles are normal in this style
        LIMITS.update({"title": 90, "bullet": 150, "card": 170})
    for n, s in enumerate(slides, 1):
        where = f"slide {n} ({s['layout']})"
        def chk(text, lim, what):
            if text and len(str(text)) > lim:
                warn.append(f"{where}: {what} is {len(str(text))} chars (aim for <= {lim}): \"{str(text)[:50]}...\"")
        chk(re.sub(r"\s*\(continued\)$", "", s.get("title", "")), LIMITS["title"], "title")
        chk(s.get("body"), LIMITS["body"], "body")
        for b in (s.get("bullets") if s["layout"] == "content" else []) or []:
            chk(b, LIMITS["bullet"], "a bullet")
        for c in (s.get("cards") or []) + (s.get("items") if s["layout"] == "icon_grid" else []) or []:
            chk(c.get("text"), LIMITS["card"], f"'{c.get('title')}' text")
        for st in (s.get("steps") if s["layout"] == "process" else []) or []:
            chk(st.get("text"), LIMITS["step"], f"step '{st.get('title')}' text")
        for st in s.get("stats") or []:
            chk(st.get("label"), LIMITS["stat_label"], f"stat '{st.get('value')}' label")
        chk(s.get("quote"), LIMITS["quote"], "quote")
        chk(s.get("text") if s["layout"] == "statement" else None, LIMITS["statement"], "statement")
        qs = [s.get("question")] if s["layout"] in ("question", "think_pair_share") else s.get("questions") or []
        for q in qs:
            if q and CLOSED_Q.match(q.strip()) and s.get("closed") is not True:
                warn.append(f"{where}: \"{q[:50]}\" is a yes/no question; for discussion try How / What / Why / "
                            f"When / Where / Who / Tell about... (or set \"closed\": true)")
            chk(q, 140, "question")
        for o in s.get("options") or []:
            chk(o, 90, "an option")
        chk(s.get("story"), 450, "scenario story")
        for st in (s.get("steps") if s["layout"] == "activity" else []) or []:
            chk(st, 110, "an activity step")
        if s["layout"] == "image" and not (s.get("image") or s.get("graphic")):
            warn.append(f"{where}: needs an image path or a graphic spec")
    return warn

DECK_DIR = Path(".")
STYLE = "teaching"
FONT_SCALE = 1.0
TMP_DIR = None


def main():
    global DECK_DIR, TMP_DIR
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("deck")
    ap.add_argument("-o", "--out")
    ap.add_argument("--plan", action="store_true", help="print the layout chosen for every slide")
    a = ap.parse_args()

    deck_path = Path(a.deck).resolve()
    DECK_DIR = deck_path.parent
    deck = json.loads(deck_path.read_text())
    out = Path(a.out) if a.out else deck_path.with_suffix(".pptx")

    import tempfile
    TMP_DIR = tempfile.mkdtemp()

    prs = Presentation()
    prs.slide_width, prs.slide_height = W, H
    prs.core_properties.title = deck.get("title", "")
    prs.core_properties.author = deck.get("author", "Frontier Commons")

    global STYLE
    STYLE = deck.get("style", "teaching")
    if STYLE not in ("teaching", "workshop", "formal"):
        sys.exit('"style" must be "teaching" (default), "workshop" or "formal"')
    global FONT_SCALE
    FONT_SCALE = float(deck.get("font_scale", 0.9 if STYLE == "teaching" else 1.0))
    slides = plan(deck)
    for n, s in enumerate(slides, 1):
        fn = (EDITORIAL_OVERRIDES if STYLE == "teaching" else TEACHING_OVERRIDES if STYLE == "workshop" else {}).get(s["layout"])
        try:
            sl = (fn or LAYOUTS[s["layout"]])(prs, s, n, deck)
        except KeyError as e:
            sys.exit(f"Slide {n} ({s['layout']}, from outline item {s['_src']}): missing required field {e}")
        if s.get("notes"):
            sl.notes_slide.notes_text_frame.text = s["notes"]

    prs.save(out)
    print(f"✓ {out}  ({len(slides)} slides from {len(deck['slides'])} outline items)")
    if a.plan:
        for n, s in enumerate(slides, 1):
            label = (s.get("title") or s.get("text") or s.get("quote") or "").replace("\n", " ")[:50]
            print(f"  {n:>2}. {s['layout']:<11} {label}  [{s['_why']}]")
    issues = lint(slides) + WARNINGS
    if issues:
        print("Copy check (fix by rewriting the words; never shrink fonts):")
        for w in issues:
            print("  - " + w)


if __name__ == "__main__":
    main()
