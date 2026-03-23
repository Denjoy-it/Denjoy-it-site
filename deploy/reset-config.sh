#!/usr/bin/env bash
# =============================================================================
# Denjoy Platform — Reset config.json naar Linux defaults
# Gebruik: sudo bash reset-config.sh [--tenant-id ID] [--client-id ID] [--secret SECRET]
# =============================================================================
PLATFORM_DIR="/opt/denjoy-platform"
CONFIG="${PLATFORM_DIR}/backend/storage/config.json"
APP_USER="denjoy"

TENANT_ID=""; CLIENT_ID=""; CLIENT_SECRET=""; CERT_THUMB=""

while [[ $# -gt 0 ]]; do
    case "$1" in
        --tenant-id)    TENANT_ID="$2";    shift 2 ;;
        --client-id)    CLIENT_ID="$2";    shift 2 ;;
        --secret)       CLIENT_SECRET="$2"; shift 2 ;;
        --cert-thumb)   CERT_THUMB="$2";   shift 2 ;;
        *) echo "Onbekende optie: $1"; exit 1 ;;
    esac
done

cat > "${CONFIG}" <<EOF
{
  "default_run_mode": "demo",
  "assessment_ui_v1": true,
  "script_path": "${PLATFORM_DIR}/assessment/Start-M365BaselineAssessment.ps1",
  "auth_tenant_id": "${TENANT_ID}",
  "auth_client_id": "${CLIENT_ID}",
  "auth_cert_thumbprint": "${CERT_THUMB}",
  "auth_client_secret": "${CLIENT_SECRET}",
  "entrafalcon_script_path": "${PLATFORM_DIR}/assessment/EntraFalcon/run_EntraFalcon.ps1",
  "entrafalcon_auth_flow": "DeviceCode",
  "entrafalcon_tenant": "${TENANT_ID}",
  "entrafalcon_include_ms_apps": false,
  "entrafalcon_csv": false
}
EOF

chown "${APP_USER}:${APP_USER}" "${CONFIG}"
echo "Config gereset: ${CONFIG}"
cat "${CONFIG}"

systemctl restart denjoy-platform 2>/dev/null && echo "Service herstart." || echo "Start service handmatig: systemctl restart denjoy-platform"
