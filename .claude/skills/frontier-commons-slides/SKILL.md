---
name: frontier-commons-slides
description: Builds on-brand Frontier Commons presentation decks (.pptx) from an outline, using the official palette (Deep Blue, Mist, Onyx, Soft Periwinkle, Papaya Whip, Amber Gold), Inter + DM Sans typography, the logo, and 12 ready layouts (title, agenda, section divider, content, two-column, cards, stats, process, quote, image, statement, closing). Use whenever someone asks for Frontier Commons slides, a deck, a presentation, a pitch, a partner briefing, or a talk, or wants to turn notes, a doc, or an outline into slides for Frontier Commons.
---

# Frontier Commons Slides

Turn content into a branded PowerPoint deck by writing a JSON outline and running one script. The output opens in PowerPoint, Keynote, and Google Slides (via import), and every element stays editable.

## Workflow

1. **Understand the content.** Ask only for what is missing: the audience, the goal, and the key points. If the user gives notes or a document, pull the structure from it.
2. **Plan the story** (usually 8–15 slides): title → agenda (optional) → sections with 2–4 slides each → closing. Pick the layout that fits each slide's *job* (see the table). Vary layouts; avoid three `content` slides in a row.
3. **Write copy in the brand voice** (see `references/brand.md`): simple everyday words, no jargon or buzzwords, warm and collaborative. Keep slides short:
   - Titles: at most ~8 words. Use sentence case (the title slide may use Title Case).
   - Bullets: 3–5 per slide, each under ~12 words.
   - Wrap *one* key word in a title in `**...**` to highlight it. Don't highlight everywhere.
   - Put the detail in `"notes"` (speaker notes), not on the slide.
4. **Write the JSON** to the user's working folder (e.g. `my-talk.json`). Start from `examples/sample-deck.json`.
5. **Build:**
   ```bash
   python3 <skill-dir>/scripts/build_deck.py my-talk.json -o my-talk.pptx
   ```
   (`pip3 install python-pptx` if it's missing.)
6. **Check it.** If PowerPoint is available, export a PDF and look at every slide for overflow or crowding:
   ```bash
   OUT=$HOME/Library/Containers/com.microsoft.Powerpoint/Data/preview.pdf
   osascript -e "tell application \"Microsoft PowerPoint\"
     open POSIX file \"$PWD/my-talk.pptx\"
     delay 4
     save active presentation in POSIX file \"$OUT\" as save as PDF
     close active presentation saving no
   end tell"
   ```
   Then read the PDF. If text overflows, shorten the copy rather than shrinking fonts.
7. Tell the user where the file is. Remind them that **Inter and DM Sans must be installed** for the deck to look right. Run `scripts/install_fonts.sh` once per Mac and restart PowerPoint afterwards. In Google Slides, both fonts are available from the font menu.

## Layouts

| Layout | Job | Fields (* required) |
|---|---|---|
| `title` | Opening slide, Deep Blue | `title`*, `subtitle`, `date` |
| `agenda` | Roadmap (up to 8 items) | `title`, `items`* |
| `section` | Chapter divider: big word on a line, Periwinkle | `title`* (≤ 14 chars looks best), `kicker` (e.g. "01") |
| `content` | One idea plus an explanation or bullets | `title`*, `tag`, `body`, `bullets` |
| `two_column` | Compare or pair two things | `title`*, `tag`, `left`*/`right`* each `{heading, body, bullets}` |
| `cards` | 2–4 parallel concepts (like Core Principles) | `title`*, `tag`, `cards`* `[{title, subtitle, text}]` |
| `stats` | 2–4 key numbers | `title`*, `tag`, `stats`* `[{value, label, detail}]` |
| `process` | 3–6 sequential steps or journey | `title`*, `tag`, `steps`* `[{title, text}]` |
| `quote` | Testimony or key quote, Papaya | `quote`*, `author`, `role` |
| `image` | Photo plus story | `title`*, `image`* (path relative to JSON), `tag`, `body`, `bullets`, `caption`, `side` |
| `statement` | One big sentence, Deep Blue | `text`*, `tag` |
| `closing` | Thank you plus contact, stacked wordmark | `title`, `subtitle`, `contact` |

Every slide accepts `"notes"`. The deck-level `"footer"` sets the small footer text next to page numbers.

## Brand rules the script already enforces
- Backgrounds use only Mist, Deep Blue, Periwinkle (section dividers), Papaya (community and story slides) and white (cards).
- Amber appears only on tags, underlines and rules. Don't ask for more Amber.
- The wordmark goes on the title and closing slides, and the icon in footers (the guide says external decks should show the wordmark).
- Photos get rounded corners and are centre-cropped to fill their frame.

When you need something the layouts don't cover, extend `scripts/build_deck.py` with a new `l_<name>` function that uses the same helpers (`bg`, `rect`, `textbox`, `bullets`, `tag`, `footer`) and brand constants, then register it in `LAYOUTS`. Never introduce colours outside `references/brand.md`.

## Files
- `scripts/build_deck.py`: the JSON → PPTX builder
- `scripts/install_fonts.sh`: installs Inter and DM Sans (static TTFs, SIL OFL) on macOS
- `assets/`: logo icon and wordmark PNG/SVG (reconstructed; replace with official files if available) and fonts
- `examples/sample-deck.json`: a complete example using every layout
- `references/brand.md`: colours, type, logo rules, voice and tone
