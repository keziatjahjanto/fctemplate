# Frontier Commons Templates

One Claude skill, [`frontier-commons`](.claude/skills/frontier-commons/SKILL.md), that produces on-brand Frontier Commons material from the *Branding Guide (January 2026)*:

- **Presentations (.pptx).** The default style is warm and built for teaching and discussion: discussion questions, think-pair-share, scenarios, activities, recaps and breaks. A formal style is available for briefings. Claude outlines the content, and the builder picks each slide's layout from how many points it has and what kind, splits crowded slides, adds brand icons, editable charts and timelines, and flags text that's too long or yes/no discussion questions.
- **Graphics (.png):** social posts, carousels, stories, discussion prompts, quote and stat cards, event promos, banners and logo exports, using the same icons as the slides.

Try asking Claude:
- "Make a 60-minute volunteer training on welcoming international students, with discussion."
- "Turn these notes into a partner briefing deck, formal style."
- "Make an Instagram post with this week's discussion question."

## Setup (once per Mac)
```bash
pip3 install --user python-pptx
.claude/skills/frontier-commons/scripts/install_fonts.sh   # Inter + DM Sans for PowerPoint/Keynote
```
Graphics need Google Chrome; their fonts are bundled.

## Using the skill
- **In this folder with Claude Code:** it loads automatically. Just ask.
- **In every project:** copy `.claude/skills/frontier-commons/` into `~/.claude/skills/`.
- **In claude.ai:** zip the `frontier-commons/` folder and upload it under Settings → Capabilities → Skills. Graphics rendering needs Chrome, which claude.ai may not have.

## Running without Claude
```bash
S=.claude/skills/frontier-commons
python3 $S/scripts/build_deck.py my-deck.json -o my-deck.pptx --plan   # --plan shows the layout picked for each slide
python3 $S/scripts/preview_deck.py my-deck.pptx -o preview/            # PNG previews via Quick Look
python3 $S/scripts/render.py my-post.json -o out/
```
Examples: `examples/decks/teaching-deck.json`, `auto-deck.json`, `sample-deck.json`, and `examples/graphics/`.

## Notes
- The logo files are **reconstructed** from the guide. Put the official files in `assets/logos/` with the same names to replace them.
- Brand rules are in `references/brand.md`; icon names and keywords in `references/icons.md`.
- Icons come from [Lucide](https://lucide.dev) (ISC licence), recoloured to the brand. Fonts: Inter, DM Sans and Montserrat (SIL OFL).
