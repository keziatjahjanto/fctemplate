---
name: frontier-commons
description: Makes anything visual for Frontier Commons in the official brand (Deep Blue, Mist, Onyx, Soft Periwinkle, Papaya Whip, Amber Gold; Inter + DM Sans; the logo; 107 brand line icons). Decks (.pptx) default to the welcoming, illustrated teaching style of Frontier Commons' Fall Fellows "Plug In" sessions (recap, pause, definition, pillars, spark bullets on Papaya panels, reflection questions, section openers, quotes, tools, deliverables), starting from a blank session template. A warm workshop style and a formal briefing style are also available. The builder picks each slide's layout from how many points it has and what kind, splits crowded slides and checks copy. Graphics (PNG) cover social posts, carousels, stories, discussion prompts, quote and stat cards, event promos, banners and logo exports. Use for any Frontier Commons slides, deck, presentation, session, training, workshop, lesson, talk, graphic, image, post, flyer or banner.
---

# Frontier Commons

This single file has everything needed to make on-brand Frontier Commons **presentations** and **graphics**. The scripts in `scripts/` do the building; the `references/` files are optional deeper detail.

| The user wants… | Make | How |
|---|---|---|
| a session, training, study, workshop or talk | `.pptx` in the **teaching** style | Copy `examples/decks/teaching-template.json`, fill it in, run `scripts/build_deck.py` |
| a partner or board briefing, report or pitch | `.pptx` in the **formal** style | Same, with `"style": "formal"` |
| a playful, hands-on event | `.pptx` in the **workshop** style | Same, with `"style": "workshop"` |
| a post, carousel, story, flyer, banner or image | `.png` | Write a spec and run `scripts/render.py` |

**Setup (once per Mac):** `pip3 install --user python-pptx`, then `scripts/install_fonts.sh` (installs Inter and DM Sans; restart PowerPoint afterwards). Graphics need Google Chrome; their fonts are bundled.
**No scripts available** (for example in Claude Design or a chat without code)? Follow **section 3 · Visual spec** to recreate the same look by hand.

---

## 1 · Brand essentials

**Colours.** Only these, plus their tonal steps. Primary colours carry surfaces and structure, secondary colours support stories, and the accent is used sparingly.

| Name | HEX | Use for |
|---|---|---|
| Onyx | `#131416` | Body text, icons |
| Deep Blue | `#0D145F` | Logo, titles, dark slides |
| Mist | `#F3F4FA` | Backgrounds, text on dark |
| Soft Periwinkle | `#9C90E6` | Illustrations, data, labels, secondary buttons |
| Papaya Whip | `#FFEFD5` | Warm panels, story and community areas, section openers |
| Amber Gold | `#FABC20` | Accent only: the series line, icon discs, CTAs, tags (1–2 per slide) |

Tonal steps: Deep Blue 35% `#15219D` (highlighted words, bullets, labels) and 50% `#1F2FE0` · Periwinkle 82% `#BCB4EE`, 91% `#DCD8F6`, 95% `#ECEAFA` · Onyx 40% `#5F646D` (captions) and 85% `#D6D8DB` (hairlines). Never invent other colours. Periwinkle and Amber are never used as text on light backgrounds (their contrast is too low).

**Type.** **Inter** for titles and display (bold, tight tracking −2 to −4%). **DM Sans** for everything read (body, labels, captions). Labels are DM Sans bold in capitals with wide letter-spacing (+14–16%) in Blue 35%.

**Logo.** Use files from `assets/logos/` only (reconstructed from the guide; swap in official files if available). The **wordmark** is for external-facing work; the **icon** is for small or repeated spots. Deep Blue on light, Mist on dark. Never stretch, recolour, add effects or place it on a busy background.

**Voice.** Accessible over academic, collaborative over competitive, innovative over traditional, global over Western-centric. Use simple, warm, humble words with no jargon or buzzwords. Speak to "you" and "we". Would a first-year international student understand it?

**Never invent** statistics, quotes, testimonies, names or scripture text. Use `[placeholders]` and tell the user.

---

## 2 · Presentations

### Workflow
1. **Get the content.** Ask only for what's missing: the audience, the topic, the week, and key points or notes.
2. **Copy `examples/decks/teaching-template.json`.** It's a blank 17-slide session where every text is a `[placeholder]`. Replace every bracket with real content, delete the slides you don't need, and duplicate the ones you need more of. Set the deck-level `"kicker"` (e.g. "Week 3 · Hospitality"), which appears on every bullet slide.
3. **Build and read the plan:**
   ```bash
   python3 <skill-dir>/scripts/build_deck.py session.json -o session.pptx --plan
   ```
   `--plan` shows the layout chosen for each slide. Fix every **Copy check** line by rewriting the words (never by shrinking fonts): text that's too long, or yes/no discussion questions.
4. **Preview and look at every slide:** `python3 <skill-dir>/scripts/preview_deck.py session.pptx -o /tmp/fc-preview`, then read the `sheet-*.png` files. This uses Quick Look, so it never touches an open PowerPoint window.
5. **Hand over:** give the file path, a one-line summary per section, and anything left for the user (brackets, photos, scripture text).

### The session shape (teaching style)
Keep this rhythm: open, teach, reflect, apply, build, commit. Never go more than 3–4 teaching slides without a question, quote or pause.

| # | Slide | Layout | Content |
|---|---|---|---|
| 1 | Title | `title` | `kicker` "Frontier Commons", `title` program name, `series` (Amber italic line), `info` "Week N · Topic · Date", `org` |
| 2 | Recap | `recap` | 3–5 points from last week |
| 3 | Pause | `break` | "Let's take a\npause": quiet, prayer or settling in |
| 4 | Definition | `auto` with one point | "What [topic] is": one sentence |
| 5 | Where it shows up | `auto`, 2–4 titled points → `pillars` | e.g. books or people of scripture; `ref` for references, `note` for a footnote |
| 6 | What it looks like | `auto`, plain sentences → spark bullets | 3–5 full sentences |
| 7 | Reflection | `questions` | 2–4 open questions |
| 8 | Why it matters | `auto` → `pillars` | 3–4 reasons: LABEL plus 1–2 sentences |
| 9, 12 | Part openers | `section` | 2–5 words on two lines; the last line is highlighted |
| 10, 13 | Application / building | `auto` → bullets | 3–4 points that turn the idea into action |
| 11 | Key quote | `quote` | Quote plus a follow-up `question` |
| 14 | Three-part idea | `auto` → `pillars` | e.g. Prioritise / Decide / Aim |
| 15 | Kingdom impact | `content` with `lead` | Self-check questions |
| 16 | Tool | `tool` | `rows` of Purpose / Goal / Keep in mind, plus an example `image` |
| 17 | Deliverables | `content` | What participants produce this week |

### How the builder chooses layouts (`"layout": "auto"`, `"points": [...]`)
| Points | Becomes | Per slide (more are split evenly, 7 → 4 + 3) |
|---|---|---|
| One sentence under a title | `definition` | 1 |
| 2–4 `{title, text}` (no subtitles) | `pillars` with spot icons | 4 |
| 5–6 `{title, text}` | `icon_grid` | 6 |
| `{title, subtitle, text}` | `cards` | 4 |
| Figures ("7M students", `{value, label}`) | `stats` | 4 (use `kind: "chart"` for a bar chart) |
| Steps or `{date, …}` | `process` / timeline (`kind: "steps"`) | 6 |
| Strings ending in "?" | `questions` (1 → big `question`) | 4 |
| Plain sentences | spark bullets on a Papaya panel | 5 (8 if short) |
| Two sides | `two_column` (`kind: "compare"`) | 2 |

Other `kind`s are `agenda`, `activity`, `recap`, `scenario`, `think_pair_share` and `definition`.

Explicit layouts: `title`, `section`, `break`, `recap`, `definition`, `pillars`, `content`, `questions`, `question`, `quote`, `tool`, `think_pair_share`, `activity`, `scenario`, `agenda`, `cards`, `icon_grid`, `stats`, `process`, `chart`, `two_column`, `image`, `statement`, `closing`. Every slide takes `notes`; the teaching style also takes `lead` (a line under the title) and `icon`.

### Writing
- **Titles** are the idea in plain words, in sentence case, 60–90 characters at most. Mark one word with `**…**` to highlight it (by colour only).
- **Points** are complete, warm sentences of 60–140 characters, three or four per slide.
- **Pillar labels** are 1–3 words with 1–2 sentences under each.
- **Questions** must be open: start with *What, How, Why, When, Where, Who,* or *Tell about a time…*. Start from experience before opinion ("What is something that grieves your heart?"). Set `"closed": true` only for deliberate self-check questions.
- **Scripture:** give references (`"ref": "Revelation 21:4"`). Quote verse text only when the user supplies it; otherwise add `"note": "[Scripture references and translation to be added]"`.
- **Tie faith and practice together.** Connect design-thinking tools (empathy maps, personas, users) to the week's spiritual theme.
- **Speaker notes** for every reflection slide:
  `Ask: … / Give 1 minute of quiet. / Listen for: … / If it's quiet: … / Transition: "…"`

### Icons and illustrations
The builder picks a line icon for each slide or pillar from its words, using `assets/icons/catalog.json` (107 icons; names and keywords are in `references/icons.md`). Set `"icon": "<name>"` when the pick is wrong or ambiguous: *Job* (the book) should be `book-heart`, not a briefcase. Useful ones: `cross`, `book-open`, `book-heart`, `heart`, `hand-heart`, `users`, `eye`, `lightbulb`, `sparkles`, `shield-check`, `target`, `signpost`, `list-checks`, `clipboard-list`, `rocket`, `message-circle`, `globe`, `graduation-cap`, `house`, `handshake`, `plane-landing`, `coffee`.

### Options
- **Title slide:** set `"title_variant"` on the deck to `bold` (the default: Deep Blue with a white program name, the series in Amber italic and a Periwinkle dot texture), `illustrated` (Papaya with brand shapes) or `classic` (skewed band, optional photo).
- **Type size:** `"font_scale"` on the deck resizes all text. The default is 0.9 in the teaching style.

---

## 3 · Visual spec: the teaching style (to recreate without scripts)
Use this in Claude Design, Google Slides or by hand, on a 16:9 canvas of 1920×1080 px.

- **Title (bold):** Deep Blue `#0D145F` full bleed. A Periwinkle dot texture (2.5px dots, 22px grid, about 45% opacity) runs down the lower-left edge. On the left at x=130px: the label "FRONTIER COMMONS" (DM Sans bold, 23px, +16% tracking, Periwinkle 82%). Below it the program name in capitals, Inter bold, about 115px, line-height 0.95, white. Then the series in capitals, Inter bold *italic*, about 90px, Amber. Then the info line (DM Sans, 27px, Periwinkle 82%) and the org at the bottom left (DM Sans, 23px, Mist). No other graphics.
- **Section opener / pause:** Papaya background. A two-line statement on the left in Inter bold, about 112px (90px for longer lines), Deep Blue, with the second line in Blue 35%. On the right, an **illustration cluster** about 550px square: a Periwinkle disc (82% of the size), a Deep Blue rounded tile (58%, corner radius 20%) overlapping the disc's lower right with a Mist line icon inside, a small Amber dot at the lower left, and a navy dot texture at the top right. **Pause** slides use **ripples** instead: four concentric discs (`#ECEAFA`, `#DCD8F6`, `#BCB4EE`, `#9C90E6`) with an Amber centre.
- **Bullet slide (and recap):** white. A Papaya panel covers the left 700px (right corners rounded about 80px). In the panel: the kicker ("WEEK N · TOPIC", DM Sans bold, 22px, tracked, Blue 35%) at the top, the title about 130px further down (Inter bold, 67px for short titles), and one **Amber disc (about 200px) with a Deep Blue line icon (1.75 stroke) overlapping it**, at the same size on every slide, near the bottom. On the right: 3–5 **spark bullets** (a six-point asterisk in Blue 35%, 3 strokes) with DM Sans text at 38px in Onyx, spread evenly down the slide.
- **Pillars:** white. Title top-left (Inter bold, 47px, Deep Blue). 2–4 centred columns, each with a spot icon (a 125px Amber or Periwinkle disc with a Deep Blue line icon overlapping it), a LABEL (DM Sans bold, 36px, capitals, tracked, Blue 35%), a short 1px Periwinkle rule, text in DM Sans at 31px, and an optional bold Deep Blue reference.
- **Reflection questions:** white. A Periwinkle 91% panel down the left 640px with two overlapping speech bubbles (white and Papaya rounded rectangles with tails and Periwinkle and Amber text bars). On the right: the title, an optional "format · time" caption, and numbered rows separated by 1px `#D6D8DB` hairlines. Each row has a Periwinkle number disc and the question in Inter at 40px, Deep Blue.
- **Definition:** Mist background with a grey dot band on the left edge. Title top-left, one big sentence (Inter, 65px, Deep Blue), and an illustration cluster on the right on a Periwinkle 91% disc.
- **Quote:** white. A small Amber-disc speech icon at the top centre, the quote centred (Inter, 61px, Deep Blue), and the follow-up question in italics inside a rounded Periwinkle 91% box.
- **Tool:** white with the dot band. Labelled rows on the left (Purpose / Goal / Keep in mind, with hairlines between), and on the right a thin-bordered frame for an example image.
- **Everywhere:** no footers or page numbers, generous white space, rounded corners on shapes, highlights by colour only (never bold), and one Amber use per slide at most besides icon discs.

---

## 4 · Graphics (PNG)
Write a spec and render it:
```json
{"template": "question", "size": "square", "theme": "periwinkle", "logo": "wordmark",
 "fields": {"question": "What helped you feel **welcome**?", "format": "Pairs", "time": "4 min"}}
```
```bash
python3 <skill-dir>/scripts/render.py spec.json -o out/ [--scale 2]
```
Then **look at every PNG**.

- **Templates:** `announcement`, `quote`, `stat`, `event`, `photo`, `list`, `cards`, `question`, `divider`, `logo`
- **Sizes:** `square`, `portrait`, `story`, `landscape`, `og`, `linkedin-banner`, `email-header`, or `WxH`
- **Themes:** `deep-blue`, `mist`, `papaya`, `periwinkle`, `onyx`

List and card items get icons automatically. A slide can embed a graphic with `"graphic": {spec}` on an `image` slide. For a social set that matches a deck, reuse its titles, points and icons. Full field lists are in `references/graphics.md`.

---

## Folder map
```
SKILL.md                 everything above
scripts/     build_deck.py · preview_deck.py · render.py · build_icons.py · install_fonts.sh
templates/   graphic templates (HTML + base.css)
assets/      logos/ · fonts/{static,web}/ · icons/{svg + 5 colours, catalog.json} · textures/
examples/    decks/teaching-template.json (blank session) · workshop-deck · auto-deck · sample-deck (formal) · graphics/
references/  brand.md · slides.md · teaching.md · graphics.md · icons.md   (optional detail)
```
