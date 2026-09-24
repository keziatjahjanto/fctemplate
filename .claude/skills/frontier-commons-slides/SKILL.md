---
name: frontier-commons-slides
description: Builds on-brand Frontier Commons presentation decks (.pptx) from notes, a document, or an outline. It picks the right layout for each slide from how many points it has and what kind they are (numbers, steps, comparisons, features, principles), splits overloaded slides, and adds supporting visuals such as brand-coloured icons, editable charts, timelines, principle cards, photos, and branded graphics. Use whenever someone asks for Frontier Commons slides, a deck, a presentation, a pitch, a partner briefing, a training or a talk, or wants to turn content into slides for Frontier Commons.
---

# Frontier Commons Slides

You write a JSON outline and one script builds an editable, on-brand PowerPoint. You decide **what each slide says**. The builder decides **how it looks**: the layout, icons, splitting and spacing.

## Workflow

1. **Get the content.** Ask only for what's missing: the audience, the goal, and the key messages. Pull the structure from any notes or document the user gives you.
2. **Outline the story** (usually 8–15 slides): title → agenda (optional) → 2–4 sections, each opened by a `section` divider → closing. Aim for **one idea per slide**. The title states that idea; the points support it.
3. **Write each content slide as `"layout": "auto"` with `points`.** Label the points with their **kind** (next section) and let the builder choose the layout. Use an explicit layout only for `title`, `section`, `quote`, `image` and `closing`, or when you have a specific reason.
4. **Write copy in the brand voice** (`references/brand.md`): simple everyday words, no jargon or buzzwords, and a warm, collaborative tone. Put detail in `"notes"`, not on the slide.
5. **Build with the plan printed:**
   ```bash
   python3 <skill-dir>/scripts/build_deck.py talk.json -o talk.pptx --plan
   ```
   `--plan` lists the layout chosen for every slide and why, including any splits. A **Copy check** lists text that's too long. **Fix every copy-check item by shortening the words**, never by asking for smaller fonts, then rebuild.
6. **Look at the slides.** Render thumbnails with Quick Look (it doesn't disturb any open PowerPoint window):
   ```bash
   python3 - <<'EOF'
   from pptx import Presentation
   src, out = "talk.pptx", "/tmp/fc-preview"
   import os; os.makedirs(out, exist_ok=True)
   for keep in range(len(Presentation(src).slides)):
       p = Presentation(src); ids = p.slides._sldIdLst
       for i, sid in reversed(list(enumerate(list(ids)))):
           if i != keep: p.part.drop_rel(sid.rId); ids.remove(sid)
       p.save(f"{out}/s{keep+1:02d}.pptx")
   EOF
   qlmanage -t -s 1400 -o /tmp/fc-preview /tmp/fc-preview/*.pptx
   ```
   Then read the PNGs. Quick Look ignores bar order and number formats in charts; PowerPoint shows them correctly.
7. **Hand over:** give the file path, one line per section on what's in the deck, and anything that needs the user's input (placeholders, `[brackets]`, illustrative data). Remind them that Inter and DM Sans must be installed (`scripts/install_fonts.sh`, once per Mac, then restart PowerPoint).

## Deciding what goes on a slide

### How many points per slide
| Content | Per slide | What the builder does with more |
|---|---|---|
| Bullets, short (≤ 60 chars) | up to 8, shown in two columns from 6 | splits evenly: 9 → 5 + 4 |
| Bullets, normal | 3–5 | splits at 5 |
| Bullets, long (2+ over 90 chars) | 3–4 | splits at 4. Better to shorten them. |
| Features / benefits / programs (titled points) | 3–6 in an icon grid | splits at 6 |
| Principles / pillars (title + subtitle) | 2–4 cards | 5+ become an icon grid |
| Key numbers | 2–4 big stats | splits (5 → 3 + 2). For numbers that compare like with like, use `kind: "chart"` |
| Steps / journey / timeline | 3–6 | splits at 6 and keeps numbering (1–6, 7–9) |
| Agenda items | 3–8 | splits at 8 |
| One sentence | a `statement` slide | |

Splits are balanced (7 → 4 + 3, never 6 + 1), and continuation slides are titled "… (continued)". Prefer fixing the story over accepting a split: two slides with two clear ideas beat one idea across two slides.

### Word budgets (the copy check enforces these)
- Title ≤ 60 characters, about 8 words, in sentence case. Highlight one key word with `**…**`.
- Bullet ≤ 110 characters (ideally under 60).
- Card or icon-grid text ≤ 140 characters, one sentence. Step text ≤ 90.
- Stat label ≤ 60 characters. Statement ≤ 110. Quote ≤ 280. Body paragraph ≤ 320.

### Point kinds, and the layout auto picks
Give `"kind"` when the points' shape doesn't make it obvious:

| You have… | Write points as | `kind` | Auto layout |
|---|---|---|---|
| Figures ("7M students") | `"7M international students"` or `{"value": "7M", "label": "…", "detail": "…", "icon": "globe"}` | `numbers` (auto-detected) | `stats` |
| Figures in one unit, compared (regions, years) | `{"value": "34%", "label": "Asia"}` | `chart` | a bar `chart` (or write an explicit `chart` slide for other types or several series) |
| An order or sequence | `{"title", "text"}` | `steps` / `journey` | `process` with icons |
| Dated milestones | `{"date", "title", "text"}` | `timeline` (auto if dates) | `process` with dates |
| Two sides (before/after, us/them, now/next) | 2 × `{"title", "text" or "bullets"}` | `compare` | `two_column` |
| Values or principles | `{"title", "subtitle", "text"}` | `principles` | `cards` with icons |
| Features, benefits, needs, programs | `{"title", "text"}` | (default for titled points) | `icon_grid` |
| A plain list | strings | | `content` bullets |
| Short labels only (≤ 40 chars) | strings | `features` | `icon_grid` |
| Session plan | strings | `agenda` | `agenda` |
| One testimony | `{"text", "author", "role"}` | `quote` | `quote` |
| One big idea | 1 string | | `statement` |

## Supporting visuals

Every content slide should have one visual anchor that fits it. Choose on purpose:
- **Icons** are automatic on cards, icon grids and steps, picked from each point's title (see `references/icons.md` for the 106 names and their keywords). Override one with `"icon": "<name>"` or turn it off with `"icon": "none"`. On cards and steps, a row uses icons for every item or for none, so if any item has no match it falls back to numbers. In an icon grid, an unmatched item gets a generic sparkle, so set a better icon explicitly when that happens.
- **Charts**: `{"layout": "chart", "title", "chart": {"type": "column|bar|stacked|line|pie|doughnut", "categories": [...], "series": [{"name", "values"}] or "values": [...], "format": "0\"%\""}, "takeaway": "One sentence with the **so what**.", "source": "…"}`. These are native PowerPoint charts, so the user can edit the data. Always add a `takeaway` and a `source`, and **never invent data**. If the user has no numbers, use `[placeholder]` values and say so.
- **Photos**: `image` layout with `"image": "photo.jpg"`, which gets rounded corners and a centred crop. Use real, relevant photos only.
- **Branded graphics** when there's no photo: `image` layout with `"graphic": {...}`, which holds a spec from the **frontier-commons-graphics** skill (quote card, stat, list, and so on). It renders at the right size for the frame, for example `"graphic": {"template": "quote", "theme": "papaya", "fields": {"quote": "…", "author": "…"}}`.
- **Section dividers** (`section`) open each chapter with the brand's big-word-on-a-line style. Keep the word to 14 characters at most.

Don't add a visual just to fill space. A `statement` slide stays text-only.

## Explicit layouts (when not using auto)
| Layout | Fields (* required) |
|---|---|
| `title` | `title`*, `subtitle`, `date` |
| `agenda` | `title`, `items`* |
| `section` | `title`*, `kicker` |
| `content` | `title`*, `tag`, `body`, `bullets` |
| `two_column` | `title`*, `tag`, `left`*/`right`* `{heading, body, bullets}` |
| `cards` | `title`*, `tag`, `cards`* `[{title, subtitle, text, icon}]` |
| `icon_grid` | `title`*, `tag`, `items`* `[{title, text, icon}]`, `theme`: `"dark"` for a Deep Blue band |
| `stats` | `title`*, `tag`, `stats`* `[{value, label, detail, icon}]` |
| `process` | `title`*, `tag`, `steps`* `[{title, text, date, icon}]`, `numbered` |
| `chart` | `title`*, `chart`*, `takeaway`, `source` |
| `quote` | `quote`*, `author`, `role` |
| `image` | `title`*, `image` or `graphic`*, `tag`, `body`, `bullets`, `caption`, `side` |
| `statement` | `text`*, `tag` |
| `closing` | `title`, `subtitle`, `contact` |

Every slide accepts `"notes"`. Every list slide accepts `"icons": false`. The deck-level `"footer"` sets the footer text.

## Brand rules the builder enforces
- Backgrounds: Mist for content, Deep Blue for title, statement and closing slides, Periwinkle for dividers, and Papaya for journeys and quotes. Amber appears only on tags, rules and underlines.
- The wordmark goes on title and closing slides, and the icon in footers. Icons are Deep Blue (on a soft Periwinkle disc in grids), Onyx inside Periwinkle step circles, and Periwinkle on dark slides.
- Chart colours follow the palette order: Deep Blue, Periwinkle, Amber, Blue 65%, Periwinkle 86%, Onyx 40%.

To add a layout, write an `l_<name>` function in `scripts/build_deck.py` using the helpers (`bg`, `rect`, `textbox`, `bullets`, `tag`, `icon`, `icon_badge`, `footer`) and the brand constants, then register it in `LAYOUTS`, `CAPACITY` and `LIST_KEY`. Never introduce colours outside `references/brand.md`.

## Files
- `scripts/build_deck.py`: planner (auto layout, splitting, copy check) and the PPTX builder
- `scripts/build_icons.py`: re-rasterizes icons after you add one
- `scripts/install_fonts.sh`: installs Inter and DM Sans (static TTFs, SIL OFL) on macOS
- `assets/icons/`: SVGs, brand-coloured PNGs and `catalog.json` (keywords)
- `assets/`: logo files (reconstructed; replace with official ones) and fonts
- `examples/auto-deck.json`: auto layout, charts, timeline, and an embedded graphic. `examples/sample-deck.json` uses explicit layouts.
- `references/brand.md`: colours, type, logo, voice. `references/icons.md`: icon names and keywords.
