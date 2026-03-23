#!/usr/bin/env bash
# =============================================================================
# Denjoy Platform — Update script (bestaande LXC installatie bijwerken)
# Gebruik: sudo bash update.sh
# =============================================================================
set -euo pipefail

PLATFORM_DIR="/var/www/mijn-website"
APP_USER="denjoy"

RED='\033[0;31m'; GREEN='\033[0;32m'; YELLOW='\033[1;33m'; NC='\033[0m'
log()  { echo -e "${GREEN}[✓]${NC} $*"; }
warn() { echo -e "${YELLOW}[!]${NC} $*"; }

[ "$(id -u)" -eq 0 ] || { echo "Voer uit als root"; exit 1; }

log "Services stoppen..."
systemctl stop denjoy-platform denjoy-upload 2>/dev/null || true

log "Bestanden kopiëren naar ${PLATFORM_DIR}..."
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SRC_DIR="$(dirname "${SCRIPT_DIR}")"

# Kopieer alles behalve storage (database/config bewaren)
rsync -av --exclude='backend/storage/' \
          --exclude='assessment/data/' \
          --exclude='assessment/web/html/' \
          --exclude='.git/' \
          --exclude='deploy/' \
          "${SRC_DIR}/" "${PLATFORM_DIR}/"

# Rechten herstellen
chown -R "${APP_USER}:${APP_USER}" \
    "${PLATFORM_DIR}/backend/storage" \
    "${PLATFORM_DIR}/assessment/data" \
    "${PLATFORM_DIR}/assessment/web" 2>/dev/null || true

# pip update
"${PLATFORM_DIR}/.venv/bin/pip" install --quiet --upgrade flask flask-cors

log "Services herstarten..."
systemctl start denjoy-platform denjoy-upload
systemctl status denjoy-platform denjoy-upload --no-pager -l

log "Update voltooid."
