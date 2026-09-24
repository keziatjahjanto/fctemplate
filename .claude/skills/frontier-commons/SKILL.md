---
name: frontier-commons
description: Makes anything visual for Frontier Commons in the official brand (Deep Blue, Mist, Onyx, Soft Periwinkle, Papaya Whip, Amber Gold; Inter + DM Sans; the logo; 106 brand icons). Decks (.pptx) default to the calm editorial teaching style of Frontier Commons' own Fall Fellows "Plug In" sessions (white pages, dot band, spark bullets, pillars, reflection questions, definitions, tools); a warm workshop style and a formal briefing style are also available. The builder picks each slide's layout from how many points it has and what kind, splits crowded slides, and adds icons, charts and timelines. Graphics (PNG) cover social posts, carousels, stories, quote and stat cards, event promos, banners and logo exports. Use for any Frontier Commons slides, deck, presentation, training, workshop, lesson, talk, graphic, image, post, flyer or banner.
---

# Frontier Commons

One skill, two outputs, both built from the same brand assets:

| The user wants… | Make | Script | Detailed reference |
|---|---|---|---|
| slides, a deck, a training, a workshop, a talk | `.pptx` | `scripts/build_deck.py` | `references/slides.md` |
| a teaching or discussion session | `.pptx`, teaching style | same | `references/teaching.md` (read first) |
| a post, carousel, story, flyer, banner, image | `.png` | `scripts/render.py` | `references/graphics.md` |
| the brand rules themselves | — | — | `references/brand.md`, `references/icons.md` |

Read the reference for the output before writing the outline or spec. Always read `references/brand.md` for voice, and write copy that is simple, humble, warm and jargon-free.

## Setup (once per Mac)
```bash
pip3 install --user python-pptx          # decks
<skill-dir>/scripts/install_fonts.sh      # Inter + DM Sans for PowerPoint/Keynote (restart PowerPoint after)
```
Graphics need Google Chrome; their fonts are bundled.

## Decks in 60 seconds
1. Work out the audience and purpose. **Default to `"style": "teaching"`**, which matches the Plug In sessions: copy `examples/decks/teaching-template.json` (a blank session with `[placeholders]`) and replace every bracket. Use `"workshop"` for playful, hands-on events and `"formal"` for partner or board briefings, reports and pitches.
2. Write the outline as JSON. Make most slides `{"layout": "auto", "title": "…", "points": [...]}` and add a `"kind"` when the shape isn't obvious (`steps`, `pillars`, `definition`, `principles`, `compare`, `agenda`, `question`, `activity`, `scenario`, `recap`, `think_pair_share`, `chart`). The builder chooses the layout from the number and kind of points and splits anything too long. See the tables in `references/slides.md`.
3. For teaching, follow the session shape in `references/teaching.md`: recap, pause, definition, pillars, points, reflection questions, section, application, quote, tools, deliverables. Quote scripture only as the user provides it; otherwise give the reference and a `[to be added]` note. Write speaker notes with **Say / Ask / Listen for / Transition** cues and timings.
4. Build and read the plan:
   ```bash
   python3 <skill-dir>/scripts/build_deck.py talk.json -o talk.pptx --plan
   ```
   Fix every **Copy check** line, including too-long text and yes/no discussion questions, by rewriting the words. Then rebuild.
5. Preview with `python3 <skill-dir>/scripts/preview_deck.py talk.pptx -o /tmp/fc-preview` (Quick Look; it never touches an open PowerPoint window) and look at every contact sheet.
6. Hand over the path, a one-line-per-section summary, and anything that needs the user's input (`[placeholders]`, illustrative data).

## Graphics in 30 seconds
Write a spec (`{"template", "size", "theme", "fields"}`), run `python3 <skill-dir>/scripts/render.py spec.json -o out/`, and **look at each PNG**. Templates, sizes and themes are listed in `references/graphics.md`. To support a deck, reuse its titles, points and icons so posts and slides match. A slide can also embed a graphic directly (`"graphic": {spec}` on an `image` slide).

## Non-negotiables
- Only the brand palette and its tonal steps; Amber sparingly (tags, one CTA, rules).
- Logo files come from `assets/logos/` only. They are reconstructions: tell the user to swap in the official files if they have them. Use the wordmark externally and the icon for small or repeated spots.
- Never invent statistics, quotes, names or testimonies. Use `[placeholders]` and say so.
- Use real, relevant photos, with no heavy filters and never text over busy areas.

## Folder map
```
SKILL.md                 this file
references/  brand.md · slides.md · teaching.md · graphics.md · icons.md
scripts/     build_deck.py (planner + PPTX) · preview_deck.py · render.py (PNG) · build_icons.py · install_fonts.sh
templates/   graphic templates (HTML + base.css)
assets/      logos/ · fonts/{static,web}/ · icons/{svg,deep-blue,blue,mist,onyx,periwinkle}/ + catalog.json · textures/ (dot band, dots)
examples/    decks/ (teaching-template = blank teaching session, workshop-deck, auto-deck, sample-deck = formal) · graphics/ (demo specs + output, logo export)
```
