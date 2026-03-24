# Denjoy Platform — Proxmox LXC Setup

## Architectuur op de LXC

```
Browser → Nginx :80
              ├── /upload-report, /web/, /api/kb/*/assets|vlans|pages|contacts
              │        └── upload_server.py (Flask, :8080)
              │             └── kb_api.py blueprint (SQLite per tenant)
              │
              └── alles overige
                       └── app.py (stdlib ThreadingHTTPServer, :8787)
                            ├── Portal frontend (portal/)
                            ├── REST API: tenants, runs, reports, findings
                            └── PowerShell subprocessen (assessment)
                                     └── Start-M365BaselineAssessment.ps1 (pwsh)
                                          └── Microsoft.Graph modules → M365
```

---

## Proxmox LXC aanmaken

### Aanbevolen LXC specs
| Parameter     | Waarde                     |
|---------------|----------------------------|
| Template      | Ubuntu 22.04 of 24.04      |
| CPU           | 2 cores                    |
| RAM           | 2 GB (4 GB als Phase 6 actief) |
| Disk          | 20 GB (SSD-backed voor SQLite) |
| Network       | DHCP of statisch IP        |
| Nesting       | **Aan** (voor pwsh)        |

> **Nesting inschakelen in Proxmox:**
> LXC container → Options → Features → **Nesting: ja**
> Zonder nesting kan PowerShell Core crashen op bepaalde syscalls.

---

## Deployen

### 1. LXC aanmaken en Ubuntu updaten

```bash
apt-get update && apt-get upgrade -y
```

### 2. Bestanden overzetten naar de LXC

Via scp vanaf je Mac:
```bash
scp -r /pad/naar/denjoy-platform root@<LXC-IP>:/opt/denjoy-platform
```

Of via rsync (incrementeel updaten):
```bash
rsync -avz --exclude='.git/' --exclude='backend/storage/' \
    /pad/naar/denjoy-platform/ root@<LXC-IP>:/opt/denjoy-platform/
```

### 3. Install script uitvoeren

```bash
ssh root@<LXC-IP>
cd /opt/denjoy-platform/deploy
chmod +x install.sh
bash install.sh
```

### 4. Wachten op PowerShell modules

De M365 modules worden op achtergrond geïnstalleerd. Voortgang:

```bash
journalctl -u denjoy-psmodules -f
```

Gereed als je ziet: `PS modules OK` en `Active: inactive (exited)`

---

## M365 Assessment configureren

De assessment draait als `pwsh` subprocess — werkt volledig op Linux.

### Optie A: Device Code (interactief, dev/test)

Standaard instelling. Wanneer je een assessment start:
1. De device code verschijnt in de **assessment log** in de web UI
2. Ga op je laptop naar `https://microsoft.com/devicelogin`
3. Voer de code in en authenticeer

### Optie B: App Registratie (productiemethode)

Maak een App Registration in Entra ID aan met de vereiste Graph API permissies:
- User.Read.All, Group.Read.All, Directory.Read.All
- AuditLog.Read.All, Policy.Read.All
- Sites.Read.All, Team.ReadBasic.All
- SecurityEvents.Read.All, DelegatedAdminRelationship.Read.All
- DeviceManagementConfiguration.Read.All, DeviceManagementManagedDevices.Read.All

Configureer via reset-config.sh:
```bash
bash /opt/denjoy-platform/deploy/reset-config.sh \
    --tenant-id "jouw-tenant-guid" \
    --client-id "app-registratie-client-id" \
    --secret "client-secret"
```

Of met certificaat:
```bash
bash /opt/denjoy-platform/deploy/reset-config.sh \
    --tenant-id "jouw-tenant-guid" \
    --client-id "app-registratie-client-id" \
    --cert-thumb "THUMBPRINT"
```

Verander daarna `default_run_mode` van `demo` naar `script` via de platform UI of rechtstreeks in:
```
/opt/denjoy-platform/backend/storage/config.json
```

---

## Services beheren

```bash
# Status
systemctl status denjoy-platform denjoy-upload nginx

# Logs live
journalctl -u denjoy-platform -f
journalctl -u denjoy-upload -f

# Herstarten
systemctl restart denjoy-platform denjoy-upload

# Na code-update (rsync)
bash /opt/denjoy-platform/deploy/update.sh
```

---

## Firewall (ufw)

```bash
ufw allow 22/tcp      # SSH
ufw allow 80/tcp      # Nginx (platform UI)
ufw enable
```

Poorten 8787 en 8080 hoeven **niet** open — die zijn alleen intern via Nginx bereikbaar.

---

## Poortoverzicht

| Service          | Poort | Bereik       |
|------------------|-------|--------------|
| Nginx (UI)       | 80    | Extern       |
| app.py (backend) | 8787  | Alleen intern |
| upload_server.py | 8080  | Alleen intern |

---

## Troubleshooting

### PowerShell installeert niet
```bash
# Controleer Microsoft repo
cat /etc/apt/sources.list.d/microsoft-prod.list
# Retry
apt-get install -y powershell
```

### Assessment geeft "pwsh not found"
```bash
which pwsh   # moet /usr/bin/pwsh zijn
pwsh --version
```

### Upload server start niet (kb_api import error)
```bash
journalctl -u denjoy-upload --no-pager -n 30
# Controleer dat PYTHONPATH correct is in de service:
systemctl cat denjoy-upload | grep PYTHON
```

### CORS fout in browser console
De `UPLOAD_ALLOWED_ORIGINS=*` in de service lost dit op. Controleer:
```bash
systemctl cat denjoy-upload | grep ORIGINS
systemctl restart denjoy-upload
```

### Rapport uploaden mislukt (grote rapporten)
Nginx limiet verhogen in `/etc/nginx/sites-available/denjoy-platform`:
```
client_max_body_size 50M;
```
En in de service: `Environment=UPLOAD_MAX_BYTES=52428800`

### Config.json bevat nog macOS paden
```bash
bash /opt/denjoy-platform/deploy/reset-config.sh
```

---

## HTTPS toevoegen (optioneel, Let's Encrypt)

Alleen zinvol als de LXC een publiek domeinnaam heeft:

```bash
apt-get install -y certbot python3-certbot-nginx
certbot --nginx -d jouwdomein.nl
```

Voor intern gebruik met eigen CA of zelfondertekend certificaat:
```bash
openssl req -x509 -nodes -days 3650 -newkey rsa:2048 \
    -keyout /etc/ssl/private/denjoy.key \
    -out /etc/ssl/certs/denjoy.crt \
    -subj "/CN=$(hostname -I | awk '{print $1}')"
```

Dan in de nginx config `listen 80` vervangen door `listen 443 ssl` en de cert paths toevoegen.

---

## Bestandsstructuur op de LXC

```
/opt/denjoy-platform/
├── backend/
│   ├── app.py                    ← Hoofdbackend (stdlib, geen pip)
│   └── storage/
│       ├── app.db                ← SQLite database (tenants, runs, users)
│       ├── config.json           ← Platform configuratie
│       ├── html/                 ← Gegenereerde rapporten
│       └── runs/                 ← Assessment run data per tenant
├── portal/                       ← Frontend HTML/CSS/JS
├── assessment/
│   ├── Start-M365BaselineAssessment.ps1
│   ├── Modules/                  ← PS modules (Phase 1-7)
│   ├── Templates/                ← HTML rapport templates
│   ├── upload_server.py          ← Flask server
│   ├── kb_api.py                 ← KB Blueprint
│   └── data/                     ← KB SQLite per tenant
├── .venv/                        ← Python venv (flask, flask-cors)
└── deploy/                       ← Dit deployment pakket
    ├── install.sh
    ├── update.sh
    ├── reset-config.sh
    └── PROXMOX-LXC-SETUP.md
```
