# Graphics reference

Branded graphics come from HTML templates that headless Chrome renders to PNG. Write a small JSON spec, run one script, and the PNG is ready.

## Workflow

1. **Clarify** only what's missing: the message, the platform (sets the size), and any photo. Use a sensible default for everything else.
2. **Write copy in the brand voice** (`references/brand.md`): short, simple, warm, no buzzwords. A headline should be about 8 words or fewer. Highlight *one* key word with `**...**`.
3. **Write a spec** (JSON, one object or a list for a batch or carousel) in the user's folder. Start from `examples/graphics/demo/all.json`.
4. **Render:**
   ```bash
   python3 <skill-dir>/scripts/render.py my-post.json -o out/ [--scale 2]
   ```
   Use `--scale 2` for print or extra-sharp output. `--list` shows templates and sizes.
5. **Look at every PNG** you produce (open it with the Read tool). Check for overflow, awkward line breaks, the logo colliding with content, and contrast. Fix the copy or pick a different size or theme, then re-render.
6. Give the user the file paths.

Requires Google Chrome (or Chromium, Edge, or Brave; set `CHROME_PATH` for a custom location). Fonts are bundled, so no network is needed.

## Spec format
```json
{
  "template": "announcement",
  "size": "square",
  "theme": "deep-blue",
  "logo": "wordmark",
  "name": "webinar-oct",
  "fields": { "tag": "Webinar", "headline": "Innovation in **ISM**", "cta": "Register now" }
}
```
- **size:** `square` 1080², `portrait` 1080×1350, `story` 1080×1920, `landscape` 1920×1080, `og` 1200×630 (website/X link preview), `linkedin-banner` 1584×396, `email-header` 1200×400, or any `WxH`.
- **theme:** `deep-blue` · `mist` · `papaya` · `periwinkle` · `onyx`. Each template has a default; override it for variety.
- **logo:** `wordmark` (default; the guide prefers it for external content), `icon` (busy or small graphics, or when the name already appears), `none`.
- Image fields (`image`) are paths relative to the spec file, or URLs.
- Text supports `**highlight**`, `*italic*`, and `\n` line breaks.

## Templates

| Template | Default theme | Best for | Fields (* required) |
|---|---|---|---|
| `announcement` | deep-blue | News, launches, "who we are", covers | `headline`*, `tag`, `subtext`, `cta`, `image`, `footnote` |
| `quote` | papaya | Testimonies, principles, scripture | `quote`*, `author`, `role`, `image` (round headshot) |
| `stat` | mist | "Did you know?", impact numbers | `stat`*, `label`*, `icon`, `tag`, `detail`, `source` |
| `event` | deep-blue | Webinars, gatherings, deadlines | `title`*, `tag`, `subtitle`, `date`, `time`, `location`, `cta`, `image` |
| `photo` | mist | Stories, recaps, community moments | `image`*, `headline`*, `tag`, `caption` |
| `list` | mist | Tips, steps, "4 ways…" | `title`*, `items`* `[{title, text, icon}]` (up to 5 square, 6 wide), `tag`, `footnote` |
| `cards` | mist | 2–4 principles or pillars (Papaya cards) | `title`*, `items`* `[{title, subtitle, icon, text}]` |
| `question` | periwinkle | Discussion prompts, reflection cards, "what would you do?" posts, handouts | `question`*, `tag` (default "Let's discuss"), `prompts` (≤3), `format`, `time` |
| `divider` | periwinkle | Carousel covers, section cards, the brand's big-word-on-a-line style | `word`*, `kicker` |
| `logo` | — | Export the wordmark as a transparent PNG | `layout`: `horizontal`/`stacked`; set `"transparent": true`. See `examples/graphics/export-logos.json` |

**Icons:** `list` and `cards` items get brand line icons picked from their titles automatically, the same library and rules as the slides (`icons.md`). Set `"icon": "<name>"` per item, `"icon": "none"` to drop one, or `"icons": "none"` on the spec to turn them off. A list uses icons for every item or none, and falls back to numbers when any item has no match. `stat` takes one explicit `icon`.

**Supporting a deck:** a deck can embed any of these templates in a slide (`"graphic": {spec}` on an `image` slide). To make a social or email set that matches a deck, reuse the deck's titles, points and icon names so both use the same words and pictures: a `divider` for each section, `stat` for its numbers, `list` or `cards` for its points, `quote` for its testimonies.

**Carousels:** use a list spec with the same `size` (`portrait` or `square`), open with a `divider` or `announcement`, fill the middle with `list`, `stat` or `quote`, and close with an `announcement` CTA. Name the files `-1`, `-2`, and so on.

## Brand guardrails
- Colours come only from `templates/base.css` tokens (the brand palette plus approved tonal steps). Never add new colours.
- Amber Gold is the accent: tags, buttons and rules only, so keep it to one or two uses per graphic.
- Don't place text or the logo over busy photo areas. Photos sit in their own rounded frame.
- Use real, relevant photos, with no heavy filters and nothing low-resolution. Never fabricate people's names or quotes. Use placeholders and ask for the real ones.
- Each theme applies the approved logo colour automatically: Mist on dark backgrounds, Deep Blue on light ones.

## Adding a template
Copy an existing `templates/*.html` file. Keep the `<!-- default-theme: X -->` comment, link `{{{base_css}}}`, use `class="theme-{{theme}} aspect-{{aspect}}"` on `<body>`, size everything in `var(--u)` (1% of the short side) so it scales, reserve `var(--footer)` at the bottom for the logo, and place `{{{logo}}}` in a `.logo-slot`. Placeholders: `{{field}}` (escaped plus markup), `{{{field}}}` (raw), `{{#field}}…{{/field}}` (optional block or list loop), and `{{^field}}…{{/field}}` (if empty).

## Files
- `scripts/render.py`: spec → PNG via headless Chrome
- `templates/`: `base.css` (tokens, themes, type scale) plus one HTML file per template
- `examples/graphics/demo/`: example specs and rendered output. `examples/graphics/export-logos.json` re-exports the logo PNGs.
