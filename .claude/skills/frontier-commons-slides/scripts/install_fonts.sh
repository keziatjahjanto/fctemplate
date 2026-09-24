#!/bin/bash
# Installs the Frontier Commons brand fonts (Inter, DM Sans — SIL Open Font License)
# so PowerPoint / Keynote / Google Slides exports display correctly on this Mac.
set -e
DIR="$(cd "$(dirname "$0")/../assets/fonts" && pwd)"
DEST="$HOME/Library/Fonts"
mkdir -p "$DEST"
cp "$DIR"/*.ttf "$DEST"/
echo "Installed $(ls "$DIR"/*.ttf | wc -l | tr -d ' ') font files to $DEST. Restart PowerPoint/Keynote if open."
