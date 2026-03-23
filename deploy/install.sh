#!/usr/bin/env bash
# =============================================================================
# Denjoy Platform — Proxmox LXC Ubuntu Server Install Script
# Getest op: Ubuntu 22.04 LTS (ook 24.04 ondersteund)
# Uitvoeren als root: sudo bash install.sh
# =============================================================================
set -euo pipefail

# ------------- Configuratie --------------------------------------------------
PLATFORM_DIR="/var/www/mijn-website"   # Installatiemap op de server
APP_USER="denjoy"                      # Systeemgebruiker voor de services
MAIN_PORT="8787"                       # app.py (hoofdbackend)
UPLOAD_PORT="8080"                     # upload_server.py (rapport-upload + KB)
NGINX_PORT="80"                        # Publieke poort (Nginx frontend)
# -----------------------------------------------------------------------------

RED='\033[0;31m'; GREEN='\033[0;32m'; YELLOW='\033[1;33m'; NC='\033[0m'
log()  { echo -e "${GREEN}[✓]${NC} $*"; }
warn() { echo -e "${YELLOW}[!]${NC} $*"; }
err()  { echo -e "${RED}[✗]${NC} $*"; exit 1; }

[ "$(id -u)" -eq 0 ] || err "Voer dit script uit als root (sudo bash install.sh)"

echo ""
echo "╔══════════════════════════════════════════════════╗"
echo "║   Denjoy Platform — LXC Ubuntu Install Script   ║"
echo "╚══════════════════════════════════════════════════╝"
echo ""

# =============================================================================
# STAP 1 — Systeem paketten
# =============================================================================
log "Systeempakketten bijwerken..."
apt-get update -qq
apt-get install -y --no-install-recommends \
    curl wget gnupg apt-transport-https ca-certificates \
    software-properties-common \
    python3 python3-pip python3-venv \
    nginx \
    sqlite3 \
    jq \
    git \
    lsb-release \
    2>/dev/null
log "Systeem up-to-date."

# =============================================================================
# STAP 2 — PowerShell Core
# =============================================================================
if ! command -v pwsh &>/dev/null; then
    log "PowerShell Core installeren..."
    UBUNTU_CODENAME=$(lsb_release -cs)
    # Microsoft package repository toevoegen
    wget -q "https://packages.microsoft.com/config/ubuntu/$(lsb_release -rs)/packages-microsoft-prod.deb" \
        -O /tmp/packages-microsoft-prod.deb
    dpkg -i /tmp/packages-microsoft-prod.deb
    rm /tmp/packages-microsoft-prod.deb
    apt-get update -qq
    apt-get install -y powershell
    log "PowerShell $(pwsh --version) geïnstalleerd."
else
    log "PowerShell al aanwezig: $(pwsh --version)"
fi

# =============================================================================
# STAP 3 — Python venv + pip pakketten
# =============================================================================
log "Python virtual environment aanmaken in ${PLATFORM_DIR}..."
python3 -m venv "${PLATFORM_DIR}/.venv"
"${PLATFORM_DIR}/.venv/bin/pip" install --quiet --upgrade pip
"${PLATFORM_DIR}/.venv/bin/pip" install --quiet flask flask-cors
log "Python dependencies geïnstalleerd (flask, flask-cors)."

# =============================================================================
# STAP 4 — Systeemgebruiker aanmaken
# =============================================================================
if ! id "${APP_USER}" &>/dev/null; then
    # PowerShell subprocessen hebben een HOME-map nodig voor module-cache en temp files
    useradd --system --create-home --home-dir /var/lib/${APP_USER} \
            --shell /bin/false "${APP_USER}"
    log "Gebruiker '${APP_USER}' aangemaakt (home: /var/lib/${APP_USER})."
else
    log "Gebruiker '${APP_USER}' bestaat al."
fi
# Zorg dat home altijd bestaat
mkdir -p /var/lib/${APP_USER}
chown "${APP_USER}:${APP_USER}" /var/lib/${APP_USER}
chmod 750 /var/lib/${APP_USER}

# =============================================================================
# STAP 5 — Mappen en rechten
# =============================================================================
log "Mappen en rechten instellen..."
mkdir -p \
    "${PLATFORM_DIR}/backend/storage/html" \
    "${PLATFORM_DIR}/backend/storage/runs" \
    "${PLATFORM_DIR}/assessment/data" \
    "${PLATFORM_DIR}/assessment/web/html"

chown -R "${APP_USER}:${APP_USER}" \
    "${PLATFORM_DIR}/backend/storage" \
    "${PLATFORM_DIR}/assessment/data" \
    "${PLATFORM_DIR}/assessment/web"

# Temp-map voor PowerShell subprocessen
mkdir -p /tmp/denjoy
chown "${APP_USER}:${APP_USER}" /tmp/denjoy

# PS module cache moet schrijfbaar zijn voor de service-user
PS_MODULE_DIR="/usr/local/share/powershell/Modules"
mkdir -p "${PS_MODULE_DIR}"
chmod 755 "${PS_MODULE_DIR}"

# PowerShell NuGet provider cache
mkdir -p /root/.local/share/powershell/Modules
mkdir -p /home/${APP_USER}/.local/share/powershell/Modules 2>/dev/null || true

log "Rechten OK."

# =============================================================================
# STAP 6 — Config.json resetten met Linux paden
# =============================================================================
CONFIG_FILE="${PLATFORM_DIR}/backend/storage/config.json"
log "Config.json instellen met Linux paden..."
cat > "${CONFIG_FILE}" <<EOF
{
  "default_run_mode": "demo",
  "assessment_ui_v1": true,
  "script_path": "${PLATFORM_DIR}/assessment/Start-M365BaselineAssessment.ps1",
  "auth_tenant_id": "",
  "auth_client_id": "",
  "auth_cert_thumbprint": "",
  "auth_client_secret": "",
  "entrafalcon_script_path": "${PLATFORM_DIR}/assessment/EntraFalcon/run_EntraFalcon.ps1",
  "entrafalcon_auth_flow": "DeviceCode",
  "entrafalcon_tenant": "",
  "entrafalcon_include_ms_apps": false,
  "entrafalcon_csv": false
}
EOF
chown "${APP_USER}:${APP_USER}" "${CONFIG_FILE}"
log "Config.json gereset."

# =============================================================================
# STAP 7 — Systemd services
# =============================================================================
log "Systemd services installeren..."

# --- Service 1: Hoofdbackend (app.py) ---
cat > /etc/systemd/system/denjoy-platform.service <<EOF
[Unit]
Description=Denjoy Platform — Hoofdbackend (app.py)
After=network.target
Wants=network.target

[Service]
Type=simple
User=${APP_USER}
Group=${APP_USER}
WorkingDirectory=${PLATFORM_DIR}/backend
ExecStart=/usr/bin/python3 ${PLATFORM_DIR}/backend/app.py
Environment=M365_LOCAL_WEBAPP_HOST=127.0.0.1
Environment=M365_LOCAL_WEBAPP_PORT=${MAIN_PORT}
Environment=M365_DATA_DIR=${PLATFORM_DIR}/backend/storage
Environment=M365_WEB_DIR=${PLATFORM_DIR}/portal
Environment=DENJOY_SESSION_HOURS=8
# Nodig voor PowerShell subprocessen die de assessment draaien
Environment=HOME=/var/lib/${APP_USER}
Environment=TMPDIR=/tmp/denjoy
Restart=always
RestartSec=5
StandardOutput=journal
StandardError=journal
SyslogIdentifier=denjoy-platform

# Security hardening
NoNewPrivileges=true
PrivateTmp=true
ProtectSystem=full
ReadWritePaths=${PLATFORM_DIR}/backend/storage ${PLATFORM_DIR}/assessment

[Install]
WantedBy=multi-user.target
EOF

# --- Service 2: Upload/KB server (upload_server.py via Flask) ---
cat > /etc/systemd/system/denjoy-upload.service <<EOF
[Unit]
Description=Denjoy Platform — Upload & KB server (upload_server.py)
After=network.target
Wants=network.target

[Service]
Type=simple
User=${APP_USER}
Group=${APP_USER}
WorkingDirectory=${PLATFORM_DIR}/assessment
ExecStart=${PLATFORM_DIR}/.venv/bin/python3 ${PLATFORM_DIR}/assessment/upload_server.py --host 127.0.0.1 --port ${UPLOAD_PORT}
# kb_api.py staat in dezelfde map als upload_server.py — PYTHONPATH zet dit goed
Environment=PYTHONPATH=${PLATFORM_DIR}/assessment
Environment=M365_DATA_DIR=${PLATFORM_DIR}/assessment/data
Environment=UPLOAD_MAX_BYTES=20971520
# Sta alle origins toe — interne tool achter Nginx, geen publiek internet
Environment=UPLOAD_ALLOWED_ORIGINS=*
Environment=HOME=/var/lib/${APP_USER}
Restart=always
RestartSec=5
StandardOutput=journal
StandardError=journal
SyslogIdentifier=denjoy-upload

NoNewPrivileges=true
PrivateTmp=true
ProtectSystem=full
ReadWritePaths=${PLATFORM_DIR}/assessment

[Install]
WantedBy=multi-user.target
EOF

systemctl daemon-reload
systemctl enable denjoy-platform denjoy-upload
log "Systemd services geregistreerd en enabled."

# =============================================================================
# STAP 8 — Nginx reverse proxy
# =============================================================================
log "Nginx configureren..."

cat > /etc/nginx/sites-available/denjoy-platform <<EOF
# Denjoy Platform — Nginx Reverse Proxy
# Gegenereerd door install.sh

upstream denjoy_main {
    server 127.0.0.1:${MAIN_PORT};
    keepalive 4;
}

upstream denjoy_upload {
    server 127.0.0.1:${UPLOAD_PORT};
    keepalive 4;
}

server {
    listen ${NGINX_PORT};
    server_name _;

    # Upload-limiet voor rapport-HTML bestanden
    client_max_body_size 25M;

    # Gzip voor snellere overdracht van grote HTML-rapporten
    gzip on;
    gzip_types text/plain text/css application/javascript application/json text/html;
    gzip_min_length 2048;

    # --- Upload server routes (Flask) ---
    location /upload-report {
        proxy_pass         http://denjoy_upload;
        proxy_http_version 1.1;
        proxy_set_header   Host \$host;
        proxy_set_header   X-Real-IP \$remote_addr;
        proxy_set_header   X-Forwarded-For \$proxy_add_x_forwarded_for;
        proxy_read_timeout 120s;
    }

    location /web/ {
        proxy_pass         http://denjoy_upload;
        proxy_http_version 1.1;
        proxy_set_header   Host \$host;
        proxy_set_header   X-Real-IP \$remote_addr;
    }

    # KB API routes die de Flask blueprint serveert
    location ~ ^/api/kb/[^/]+/(assets|vlans|pages|contacts) {
        proxy_pass         http://denjoy_upload;
        proxy_http_version 1.1;
        proxy_set_header   Host \$host;
        proxy_set_header   X-Real-IP \$remote_addr;
        proxy_set_header   Origin http://\$host;
    }

    # --- Alles overige → hoofdbackend (app.py) ---
    location / {
        proxy_pass         http://denjoy_main;
        proxy_http_version 1.1;
        proxy_set_header   Host \$host;
        proxy_set_header   X-Real-IP \$remote_addr;
        proxy_set_header   X-Forwarded-For \$proxy_add_x_forwarded_for;
        proxy_set_header   Connection "";
        proxy_read_timeout 300s;

        # Grote SSE/streaming responses voor assessment log-output
        proxy_buffering    off;
        proxy_cache        off;
    }

    # Prevent access to .git, .env etc.
    location ~ /\. {
        deny all;
        return 404;
    }
}
EOF

# Activeer site
ln -sf /etc/nginx/sites-available/denjoy-platform /etc/nginx/sites-enabled/denjoy-platform
rm -f /etc/nginx/sites-enabled/default

# Test config
nginx -t && log "Nginx config OK." || err "Nginx config fout — controleer /etc/nginx/sites-available/denjoy-platform"
systemctl enable nginx
systemctl restart nginx
log "Nginx geconfigureerd en gestart."

# =============================================================================
# STAP 9 — PowerShell modules installeren (background job)
# =============================================================================
log "PowerShell modules installeren (dit kan 5-10 minuten duren)..."
log "Voortgang volgen: journalctl -u denjoy-psmodules -f"

cat > /etc/systemd/system/denjoy-psmodules.service <<EOF
[Unit]
Description=Denjoy — PowerShell Graph/Az modules installer (eenmalig)
After=network-online.target
Wants=network-online.target

[Service]
Type=oneshot
ExecStart=/usr/bin/pwsh -NonInteractive -NoProfile -Command "\
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted; \
    Install-Module Microsoft.Graph.Authentication     -Scope AllUsers -Force -SkipPublisherCheck; \
    Install-Module Microsoft.Graph.Users              -Scope AllUsers -Force -SkipPublisherCheck; \
    Install-Module Microsoft.Graph.Groups             -Scope AllUsers -Force -SkipPublisherCheck; \
    Install-Module Microsoft.Graph.Identity.SignIns   -Scope AllUsers -Force -SkipPublisherCheck; \
    Install-Module Microsoft.Graph.Reports            -Scope AllUsers -Force -SkipPublisherCheck; \
    Install-Module Microsoft.Graph.Teams              -Scope AllUsers -Force -SkipPublisherCheck; \
    Install-Module Microsoft.Graph.Sites              -Scope AllUsers -Force -SkipPublisherCheck; \
    Install-Module Microsoft.Graph.DeviceManagement   -Scope AllUsers -Force -SkipPublisherCheck; \
    Install-Module ExchangeOnlineManagement           -Scope AllUsers -Force -SkipPublisherCheck; \
    Write-Host 'PS modules OK'; \
"
RemainAfterExit=yes
StandardOutput=journal
StandardError=journal
SyslogIdentifier=denjoy-psmodules

[Install]
WantedBy=multi-user.target
EOF

systemctl daemon-reload
systemctl start denjoy-psmodules &
PSMOD_PID=$!

# =============================================================================
# STAP 10 — Services starten
# =============================================================================
log "Services starten..."
systemctl start denjoy-platform
sleep 2
systemctl start denjoy-upload
sleep 1

# =============================================================================
# SAMENVATTING
# =============================================================================
LXC_IP=$(hostname -I | awk '{print $1}')

echo ""
echo "╔══════════════════════════════════════════════════════════════════╗"
echo "║              Denjoy Platform — Installatie voltooid             ║"
echo "╠══════════════════════════════════════════════════════════════════╣"
printf "║  Platform URL  : %-48s║\n" "http://${LXC_IP}"
printf "║  Hoofdbackend  : %-48s║\n" "http://127.0.0.1:${MAIN_PORT} (app.py)"
printf "║  Upload/KB     : %-48s║\n" "http://127.0.0.1:${UPLOAD_PORT} (Flask)"
echo "╠══════════════════════════════════════════════════════════════════╣"
echo "║  Logs volgen:                                                    ║"
echo "║    journalctl -u denjoy-platform -f                             ║"
echo "║    journalctl -u denjoy-upload -f                               ║"
echo "║    journalctl -u denjoy-psmodules -f  (PS modules install)      ║"
echo "╠══════════════════════════════════════════════════════════════════╣"
echo "║  Config bewerken:                                                ║"
printf "║    %-65s║\n" "${PLATFORM_DIR}/backend/storage/config.json"
echo "║  Na config-wijziging: systemctl restart denjoy-platform          ║"
echo "╚══════════════════════════════════════════════════════════════════╝"
echo ""
warn "PS modules worden op achtergrond geïnstalleerd."
warn "Wacht tot 'systemctl status denjoy-psmodules' Active: (exited) toont"
warn "voor je een live M365 assessment start."
echo ""
