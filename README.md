# Frontier Commons Templates

Reusable Claude skills that produce on-brand Frontier Commons material, based on the *Branding Guide (January 2026)*.

| Skill | What it makes | Try asking Claude |
|---|---|---|
| [`frontier-commons-slides`](.claude/skills/frontier-commons-slides/SKILL.md) | PowerPoint decks (.pptx) with 12 branded layouts | "Make a 10-slide Frontier Commons partner briefing from these notes" |
| [`frontier-commons-graphics`](.claude/skills/frontier-commons-graphics/SKILL.md) | PNG graphics: social posts, stories, carousels, quote and stat cards, event promos, banners | "Make an Instagram post announcing our October webinar" |

## Setup (once per Mac)
```bash
pip3 install --user python-pptx                                # slide builder
.claude/skills/frontier-commons-slides/scripts/install_fonts.sh  # Inter + DM Sans for PowerPoint/Keynote
```
Graphics need Google Chrome installed; their fonts are bundled.

## Using the skills
- **In this folder with Claude Code:** the skills load automatically. Just ask.
- **In every project:** copy or symlink the two folders into `~/.claude/skills/`.
- **In claude.ai:** zip a skill folder (e.g. `frontier-commons-slides/`) and upload it under Settings → Capabilities → Skills.

## Running without Claude
```bash
python3 .claude/skills/frontier-commons-slides/scripts/build_deck.py my-deck.json -o my-deck.pptx
python3 .claude/skills/frontier-commons-graphics/scripts/render.py my-post.json -o out/
```
See each skill's `examples/` folder for ready-made JSON.

## Notes
- The logo files are **reconstructed** from the guide. Drop the official files into each skill's `assets/` folder, keeping the same names, to replace them.
- Brand rules (colours, type, logo, voice) live in `references/brand.md` inside each skill.
