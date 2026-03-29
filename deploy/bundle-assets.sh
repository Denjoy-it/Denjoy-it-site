#!/usr/bin/env bash
# =============================================================================
# Denjoy Platform — CSS Bundle Script
# Concateneert alle portal CSS-bestanden tot één portal/css/bundle.css
# Gebruik: bash deploy/bundle-assets.sh [portal-dir]
# =============================================================================
set -euo pipefail

PORTAL_DIR="${1:-portal}"
CSS_DIR="$PORTAL_DIR/css"
OUTPUT="$CSS_DIR/bundle.css"

# Volgorde: portal.css eerst (bevat design tokens en basis), dan modules
CSS_FILES=(
  "portal.css"
  "kb.css"
  "assessment-ui.css"
  "remediate.css"
  "gebruikers.css"
  "baseline.css"
  "intune.css"
  "backup.css"
  "ca.css"
  "domains.css"
  "alerts.css"
  "exchange.css"
)

echo "Bundling CSS → $OUTPUT"
> "$OUTPUT"  # leegmaken

for FILE in "${CSS_FILES[@]}"; do
  SRC="$CSS_DIR/$FILE"
  if [[ -f "$SRC" ]]; then
    echo "/* ── $FILE ── */" >> "$OUTPUT"
    cat "$SRC" >> "$OUTPUT"
    echo "" >> "$OUTPUT"
    echo "  + $FILE"
  else
    echo "  ! SKIP: $SRC (niet gevonden)"
  fi
done

SIZE=$(wc -c < "$OUTPUT")
echo "Klaar: $OUTPUT ($SIZE bytes)"
echo ""
echo "Vervang in dashboard.html de losse <link> tags door:"
echo '  <link rel="stylesheet" href="css/bundle.css">'
