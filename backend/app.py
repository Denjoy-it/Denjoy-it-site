#!/usr/bin/env python3
"""
Denjoy IT Platform — lokale backend server.

Notes:
- Serveert de portal frontend vanuit ../portal/
- Slaat lokale state op in SQLite onder backend/storage/
- Ondersteunt demo-runs en PowerShell script-runs via assessment/
"""

from __future__ import annotations

import hashlib
import json
import csv
import io
import os
import re
import secrets
import shutil
import sqlite3
import subprocess
import threading
import time
import uuid
from datetime import datetime, timedelta, timezone
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import parse_qs, unquote, urlparse
import logging
import traceback


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger(__name__)

# ============================================================
# SECURITY CONSTANTS
# ============================================================

# Content-Security-Policy voor HTML-responses
CSP_HEADER = (
    "default-src 'self'; "
    "script-src 'self' 'unsafe-inline' https://alcdn.msauth.net https://alcdn.msftauth.net; "
    "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; "
    "font-src 'self' https://fonts.gstatic.com data:; "
    "img-src 'self' data: https: blob:; "
    "connect-src 'self' https://login.microsoftonline.com https://graph.microsoft.com "
    "https://aadcdn.msauth.net https://aadcdn.msftauth.net; "
    "frame-src 'self' blob:; "
    "object-src 'none'; "
    "base-uri 'self';"
)

# In-memory rate limiter (login-pogingen per IP)
_rate_buckets: Dict[str, List[float]] = {}
_rate_lock = threading.Lock()


def _check_rate_limit(ip: str, max_attempts: int = 10, window_secs: int = 60) -> bool:
    """True = toegestaan, False = rate limit bereikt."""
    with _rate_lock:
        now = time.time()
        bucket = [t for t in _rate_buckets.get(ip, []) if now - t < window_secs]
        if len(bucket) >= max_attempts:
            _rate_buckets[ip] = bucket
            return False
        bucket.append(now)
        _rate_buckets[ip] = bucket
        return True


def _check_csrf(handler) -> bool:
    """CSRF-bescherming via Origin-header validatie."""
    origin = handler.headers.get("Origin", "")
    if not origin:
        return True  # Server-to-server (geen browser-request)
    host = handler.headers.get("Host", "")
    return (
        origin in {f"http://{host}", f"https://{host}"}
        or origin.startswith("http://localhost:")
        or origin.startswith("http://127.0.0.1:")
    )


BASE_DIR    = Path(__file__).resolve().parent          # denjoy-platform/backend/
PLATFORM_DIR = BASE_DIR.parent                          # denjoy-platform/

# Desktop/bundled deployments can override paths via environment variables:
#   M365_DATA_DIR  → writable user-data directory (SQLite, reports, runs)
#   M365_WEB_DIR   → read-only resource directory that contains the portal
_data_dir_env = os.environ.get("M365_DATA_DIR")
_web_dir_env  = os.environ.get("M365_WEB_DIR")

STORAGE_DIR         = Path(_data_dir_env) if _data_dir_env else BASE_DIR / "storage"
WEB_DIR             = Path(_web_dir_env)  if _web_dir_env  else PLATFORM_DIR / "portal"
DEFAULT_REPORTS_DIR = STORAGE_DIR / "html"
RUNS_DIR            = STORAGE_DIR / "runs"
DB_PATH             = STORAGE_DIR / "app.db"
CONFIG_PATH         = STORAGE_DIR / "config.json"


def now_iso() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat()


def ensure_dirs() -> None:
    STORAGE_DIR.mkdir(parents=True, exist_ok=True)
    RUNS_DIR.mkdir(parents=True, exist_ok=True)
    DEFAULT_REPORTS_DIR.mkdir(parents=True, exist_ok=True)


def default_config() -> Dict[str, Any]:
    return {
        "default_run_mode": "demo",
        "assessment_ui_v1": True,
        "script_path": str(PLATFORM_DIR / "assessment" / "Start-M365BaselineAssessment.ps1"),
        # App-registratie authenticatie (optioneel, laat leeg voor interactieve auth)
        "auth_tenant_id":       "",   # Azure Tenant ID (bijv. contoso.onmicrosoft.com)
        "auth_client_id":       "",   # App registratie Client ID
        "auth_cert_thumbprint": "",   # Certificate thumbprint (aanbevolen)
        "auth_client_secret":   "",   # Of client secret (minder veilig)
        # EntraFalcon instellingen
        "entrafalcon_script_path":    str(PLATFORM_DIR / "assessment" / "EntraFalcon" / "run_EntraFalcon.ps1"),
        "entrafalcon_include_ms_apps": False,         # Microsoft-eigen apps meenemen
        "entrafalcon_csv":            False,          # CSV-bestanden naast HTML genereren
    }


def load_config() -> Dict[str, Any]:
    ensure_dirs()
    if not CONFIG_PATH.exists():
        cfg = default_config()
        CONFIG_PATH.write_text(json.dumps(cfg, indent=2), encoding="utf-8")
        return cfg
    try:
        cfg = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except Exception:
        cfg = default_config()
        CONFIG_PATH.write_text(json.dumps(cfg, indent=2), encoding="utf-8")
        return cfg
    merged = default_config()
    merged.update(cfg)
    return merged


def save_config(cfg: Dict[str, Any]) -> None:
    ensure_dirs()
    CONFIG_PATH.write_text(json.dumps(cfg, indent=2), encoding="utf-8")


def get_conn() -> sqlite3.Connection:
    ensure_dirs()
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    conn = get_conn()
    cur = conn.cursor()
    cur.executescript(
        """
        CREATE TABLE IF NOT EXISTS tenants (
            id TEXT PRIMARY KEY,
            customer_name TEXT NOT NULL,
            tenant_name TEXT NOT NULL,
            tenant_guid TEXT,
            status TEXT NOT NULL DEFAULT 'active',
            owner_primary TEXT,
            owner_backup TEXT,
            tags_csv TEXT,
            risk_profile TEXT NOT NULL DEFAULT 'standard',
            notes TEXT,
            is_active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS assessment_runs (
            id TEXT PRIMARY KEY,
            tenant_id TEXT NOT NULL,
            status TEXT NOT NULL,
            run_mode TEXT NOT NULL,
            scan_type TEXT NOT NULL,
            phases_csv TEXT,
            started_by TEXT,
            started_at TEXT NOT NULL,
            completed_at TEXT,
            exit_code INTEGER,
            score_overall INTEGER,
            critical_count INTEGER DEFAULT 0,
            warning_count INTEGER DEFAULT 0,
            info_count INTEGER DEFAULT 0,
            report_path TEXT,
            snapshot_path TEXT,
            report_filename TEXT,
            is_archived INTEGER NOT NULL DEFAULT 0,
            archived_at TEXT,
            archive_reason TEXT,
            error_message TEXT,
            FOREIGN KEY (tenant_id) REFERENCES tenants(id)
        );

        CREATE TABLE IF NOT EXISTS finding_actions (
            id TEXT PRIMARY KEY,
            tenant_id TEXT NOT NULL,
            run_id TEXT,
            finding_key TEXT NOT NULL,
            title TEXT NOT NULL,
            severity TEXT NOT NULL DEFAULT 'warning',
            owner TEXT,
            status TEXT NOT NULL DEFAULT 'open',
            due_date TEXT,
            notes TEXT,
            evidence TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            closed_at TEXT,
            FOREIGN KEY (tenant_id) REFERENCES tenants(id),
            FOREIGN KEY (run_id) REFERENCES assessment_runs(id)
        );

        CREATE TABLE IF NOT EXISTS users (
            id TEXT PRIMARY KEY,
            email TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            salt TEXT NOT NULL,
            role TEXT NOT NULL DEFAULT 'klant',
            display_name TEXT,
            linked_tenant_id TEXT,
            is_active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS sessions (
            token TEXT PRIMARY KEY,
            user_id TEXT NOT NULL,
            role TEXT NOT NULL,
            email TEXT NOT NULL,
            display_name TEXT,
            created_at TEXT NOT NULL,
            expires_at TEXT NOT NULL,
            FOREIGN KEY (user_id) REFERENCES users(id)
        );

        CREATE TABLE IF NOT EXISTS audit_logs (
            id TEXT PRIMARY KEY,
            user_email TEXT,
            user_ip TEXT,
            action TEXT NOT NULL,
            resource_type TEXT,
            resource_id TEXT,
            detail TEXT,
            created_at TEXT NOT NULL
        );
        """
    )
    # Lightweight schema migration for existing local DBs.
    tenant_cols = {r[1] for r in cur.execute("PRAGMA table_info(tenants)").fetchall()}
    if "status" not in tenant_cols:
        cur.execute("ALTER TABLE tenants ADD COLUMN status TEXT NOT NULL DEFAULT 'active'")
    if "owner_primary" not in tenant_cols:
        cur.execute("ALTER TABLE tenants ADD COLUMN owner_primary TEXT")
    if "owner_backup" not in tenant_cols:
        cur.execute("ALTER TABLE tenants ADD COLUMN owner_backup TEXT")
    if "tags_csv" not in tenant_cols:
        cur.execute("ALTER TABLE tenants ADD COLUMN tags_csv TEXT")
    if "risk_profile" not in tenant_cols:
        cur.execute("ALTER TABLE tenants ADD COLUMN risk_profile TEXT NOT NULL DEFAULT 'standard'")
    run_cols = {r[1] for r in cur.execute("PRAGMA table_info(assessment_runs)").fetchall()}
    if "is_archived" not in run_cols:
        cur.execute("ALTER TABLE assessment_runs ADD COLUMN is_archived INTEGER NOT NULL DEFAULT 0")
    if "archived_at" not in run_cols:
        cur.execute("ALTER TABLE assessment_runs ADD COLUMN archived_at TEXT")
    if "archive_reason" not in run_cols:
        cur.execute("ALTER TABLE assessment_runs ADD COLUMN archive_reason TEXT")
    conn.commit()
    count = cur.execute("SELECT COUNT(*) FROM tenants").fetchone()[0]
    if count == 0:
        tenant_id = str(uuid.uuid4())
        ts = now_iso()
        cur.execute(
            """
            INSERT INTO tenants (id, customer_name, tenant_name, tenant_guid, notes, is_active, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, 1, ?, ?)
            """,
            (
                tenant_id,
                "Lokale Demo Klant",
                "Lokale Tenant",
                None,
                "Aangemaakt voor lokale MVP",
                ts,
                ts,
            ),
        )
        conn.commit()
    conn.close()


# ============================================================
# AUTH HELPERS
# ============================================================

def _hash_pw(password: str, salt: str = None):
    """Hashing met PBKDF2-SHA256. Geeft (hash_hex, salt) terug."""
    if salt is None:
        salt = secrets.token_hex(16)
    h = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), 100_000)
    return h.hex(), salt

def _verify_pw(password: str, stored_hash: str, salt: str) -> bool:
    h, _ = _hash_pw(password, salt)
    return secrets.compare_digest(h, stored_hash)

SESSION_HOURS = int(os.environ.get("DENJOY_SESSION_HOURS", "1"))


def _create_session(user_id: str, role: str, email: str, display_name: str) -> str:
    token = secrets.token_urlsafe(32)
    expires = (datetime.now(timezone.utc) + timedelta(hours=SESSION_HOURS)).astimezone().isoformat()
    db_execute(
        "INSERT INTO sessions (token,user_id,role,email,display_name,created_at,expires_at) VALUES (?,?,?,?,?,?,?)",
        (token, user_id, role, email, display_name or "", now_iso(), expires)
    )
    # Opschonen verlopen sessies
    try:
        db_execute("DELETE FROM sessions WHERE expires_at < ?", (now_iso(),))
    except Exception:
        pass
    return token

def _verify_session(token: str) -> Optional[Dict[str, Any]]:
    if not token:
        return None
    row = db_fetchone("SELECT * FROM sessions WHERE token=?", (token,))
    if not row:
        return None
    if row["expires_at"] < now_iso():
        db_execute("DELETE FROM sessions WHERE token=?", (token,))
        return None
    return dict(row)

def ensure_admin_user() -> None:
    """
    Garandeert dat het lokale admin-account altijd bestaat en up-to-date is.
    Env vars overschrijven de ingebouwde standaardwaarden (voor productie).
    """
    admin_email = os.environ.get("DENJOY_ADMIN_EMAIL", "schiphorst.d@gmail.com").strip().lower()
    admin_pw    = os.environ.get("DENJOY_ADMIN_PASSWORD", "B3@uty104").strip()
    admin_name  = os.environ.get("DENJOY_ADMIN_NAME", "Dennis Schiphorst").strip()

    pw_hash, salt = _hash_pw(admin_pw)
    existing = db_fetchone("SELECT id FROM users WHERE email=?", (admin_email,))
    if existing:
        # Altijd wachtwoord + rol bijwerken zodat inloggen gegarandeerd werkt
        db_execute(
            "UPDATE users SET password_hash=?, salt=?, role='admin', is_active=1 WHERE email=?",
            (pw_hash, salt, admin_email)
        )
    else:
        db_execute(
            "INSERT INTO users (id,email,password_hash,salt,role,display_name,is_active,created_at) "
            "VALUES (?,?,?,?,?,?,1,?)",
            (str(uuid.uuid4()), admin_email, pw_hash, salt, "admin", admin_name, now_iso())
        )
    logger.info("Admin-account gereed: %s", admin_email)

def _get_session_from_request(handler) -> Optional[Dict[str, Any]]:
    """Haal sessie op uit Authorization of Cookie header."""
    auth = handler.headers.get("Authorization", "")
    if auth.startswith("Bearer "):
        return _verify_session(auth[7:])
    cookie = handler.headers.get("Cookie", "")
    for part in cookie.split(";"):
        part = part.strip()
        if part.startswith("denjoy_session="):
            return _verify_session(part[15:])
    return None


# API-paden die geen sessie vereisen
_OPEN_API_PATHS = frozenset({
    "/api/auth/login",
    "/api/auth/logout",
    "/api/auth/microsoft",
    "/api/auth/verify",
    "/api/auth/csrf-token",
    "/api/auth/msal-config",
    "/api/health",
})


def _check_api_access(handler, path: str) -> Optional[Dict[str, Any]]:
    """
    Controleert authenticatie en autorisatie voor /api/* routes.
    Geeft sessie dict terug bij succes.
    Stuurt 401/403 en geeft None terug bij falen.
    """
    if path in _OPEN_API_PATHS or not path.startswith("/api/"):
        return {}  # Geen check nodig — leeg dict als sentinel

    sess = _get_session_from_request(handler)
    if not sess:
        handler._json(401, {"error": "Niet ingelogd."})
        return None

    # Admin-only routes: config en runs starten
    _admin_paths = {"/api/config", "/api/runs"}
    _admin_prefix = ("/api/users",)
    if path in _admin_paths or any(path.startswith(p) for p in _admin_prefix):
        if sess["role"] != "admin":
            handler._json(403, {"error": "Onvoldoende rechten."})
            return None

    # Tenant-scoped routes: klanten mogen alleen hun eigen tenant benaderen
    _tid_m = re.match(r"/api/(?:tenants|assessment|kb)/([^/]+)(?:/|$)", path)
    if _tid_m and sess["role"] != "admin":
        req_tid = _tid_m.group(1)
        user_row = db_fetchone("SELECT linked_tenant_id FROM users WHERE email=?", (sess["email"],))
        if not user_row or user_row["linked_tenant_id"] != req_tid:
            handler._json(403, {"error": "Geen toegang tot deze tenant."})
            return None

    return sess


def row_to_dict(row: sqlite3.Row) -> Dict[str, Any]:
    return {k: row[k] for k in row.keys()}


def db_fetchall(sql: str, params: Tuple[Any, ...] = ()) -> List[Dict[str, Any]]:
    conn = get_conn()
    try:
        rows = conn.execute(sql, params).fetchall()
        return [row_to_dict(r) for r in rows]
    finally:
        conn.close()


def db_fetchone(sql: str, params: Tuple[Any, ...] = ()) -> Optional[Dict[str, Any]]:
    conn = get_conn()
    try:
        row = conn.execute(sql, params).fetchone()
        return row_to_dict(row) if row else None
    finally:
        conn.close()


def db_execute(sql: str, params: Tuple[Any, ...] = ()) -> None:
    conn = get_conn()
    try:
        conn.execute(sql, params)
        conn.commit()
    finally:
        conn.close()


def db_audit(
    email: str,
    ip: str,
    action: str,
    resource_type: str = "",
    resource_id: str = "",
    detail: str = "",
) -> None:
    """Schrijft een audit-event. Mag nooit de hoofdflow blokkeren."""
    try:
        db_execute(
            "INSERT INTO audit_logs (id,user_email,user_ip,action,resource_type,resource_id,detail,created_at) "
            "VALUES (?,?,?,?,?,?,?,?)",
            (str(uuid.uuid4()), email or "", ip or "", action, resource_type or "", resource_id or "", detail or "", now_iso()),
        )
    except Exception:
        pass


def append_run_log(run_id: str, message: str) -> None:
    run_dir = RUNS_DIR / run_id
    run_dir.mkdir(parents=True, exist_ok=True)
    log_path = run_dir / "run.log"
    stamp = datetime.now().strftime("%H:%M:%S")
    with log_path.open("a", encoding="utf-8") as fh:
        fh.write(f"[{stamp}] {message}\n")


def update_run(run_id: str, **fields: Any) -> None:
    if not fields:
        return
    keys = list(fields.keys())
    sql = "UPDATE assessment_runs SET " + ", ".join([f"{k}=?" for k in keys]) + " WHERE id=?"
    vals = [fields[k] for k in keys] + [run_id]
    db_execute(sql, tuple(vals))


def phase_skip_flags(phases: List[str]) -> List[str]:
    selected = set(phases)
    flags: List[str] = []
    for i in range(1, 7):
        if f"phase{i}" not in selected:
            flags.append(f"-SkipPhase{i}")
    return flags


def find_latest_report_file(run_dir: Path) -> Optional[Path]:
    files = sorted(run_dir.glob("M365-Complete-Baseline-*.html"), key=lambda p: p.stat().st_mtime, reverse=True)
    if files:
        # Prefer a timestamped report over the convenience "latest" copy/symlink
        for f in files:
            if "latest" not in f.name.lower():
                return f
        return files[0]
    # EntraFalcon: zoek in Results_* subdirectories
    results_dirs = sorted(run_dir.glob("Results_*"), key=lambda p: p.stat().st_mtime, reverse=True)
    for d in results_dirs:
        html_files = sorted(d.glob("*.html"), key=lambda p: p.stat().st_mtime, reverse=True)
        if html_files:
            # Prefereer SecurityFindings of Summary als hoofdrapport
            for pref in ("SecurityFindings", "Findings", "Summary"):
                for f in html_files:
                    if pref.lower() in f.stem.lower():
                        return f
            return html_files[0]
    demo = sorted(run_dir.glob("*.html"), key=lambda p: p.stat().st_mtime, reverse=True)
    return demo[0] if demo else None


def list_run_html_files(run_dir: Path) -> List[Dict[str, str]]:
    """Geeft alle HTML-bestanden in de run-directory terug (inclusief subdirectories).
    Gebruikt voor de EntraFalcon multi-rapport viewer."""
    result = []
    # EntraFalcon Results_* subdirectory
    results_dirs = sorted(run_dir.glob("Results_*"), key=lambda p: p.stat().st_mtime, reverse=True)
    if results_dirs:
        d = results_dirs[0]
        for f in sorted(d.glob("*.html"), key=lambda p: p.stem):
            if "latest" not in f.name.lower():
                result.append({"name": f.stem, "path": d.name + "/" + f.name})
        return result
    # M365 baseline: één enkel rapport
    for f in sorted(run_dir.glob("M365-Complete-Baseline-*.html")):
        if "latest" not in f.name.lower():
            result.append({"name": f.stem, "path": f.name})
            break
    return result


def find_latest_summary_file(run_dir: Path) -> Optional[Path]:
    snap_dir = run_dir / "_snapshots"
    if not snap_dir.exists():
        return None
    latest = snap_dir / "M365-Complete-Baseline-latest.summary.json"
    files = sorted(snap_dir.glob("*.summary.json"), key=lambda p: p.stat().st_mtime, reverse=True)
    if files:
        for f in files:
            if "latest" not in f.name.lower():
                return f
        return files[0]
    if latest.exists():
        return latest
    return None


def extract_stats_from_summary(path: Path) -> Dict[str, Any]:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}
    totals = data.get("Totals") or {}
    metrics = data.get("Metrics") or {}
    licenses = data.get("Licenses") or []
    app_registrations = data.get("AppRegistrations") or []
    domain_dns_checks = data.get("DomainDnsChecks") or []
    licenses_total = sum(int(item.get("Total") or 0) for item in licenses if isinstance(item, dict))
    licenses_used = sum(int(item.get("Consumed") or 0) for item in licenses if isinstance(item, dict))
    return {
        "tenantName": data.get("TenantName"),
        "tenantId": data.get("TenantId"),
        "reportId": data.get("AssessmentId"),
        "reportDate": data.get("GeneratedAt"),
        "criticalIssues": totals.get("Critical", 0),
        "warnings": totals.get("Warning", 0),
        "infoItems": totals.get("Info", 0),
        "scoreOverall": totals.get("Score"),
        "mfaCoverage": metrics.get("MfaCoveragePct"),
        "usersWithoutMFA": metrics.get("MfaMissing"),
        "caPolicies": metrics.get("CAEnabled"),
        "secureScorePercentage": metrics.get("SecureScorePct"),
        "licenses": licenses,
        "licensesTotal": licenses_total,
        "licensesUsed": licenses_used,
        "appRegistrations": app_registrations,
        "domainDnsChecks": domain_dns_checks,
    }


def extract_stats_from_html(path: Path) -> Dict[str, Any]:
    try:
        html = path.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        return {}
    stats: Dict[str, Any] = {}
    patterns = {
        "totalUsers": r"Totaal Gebruikers</div></div></div><div class='stat-card'><div class='stat-number'>(\d+)</div>"  # not reliable
    }
    tenant_match = re.search(r"tenant-name['\"]>\s*([^<]+)</", html, re.I)
    if tenant_match:
        stats["tenantName"] = tenant_match.group(1).strip()
    score_pct = re.search(r"Overall Score</h3>\s*<p class=['\"]stat-value['\"]>(\d+)%</p>", html, re.I)
    if score_pct:
        stats["secureScorePercentage"] = int(score_pct.group(1))
    return stats


def parse_run_stats(run_dir: Path) -> Dict[str, Any]:
    s = find_latest_summary_file(run_dir)
    if s:
        data = extract_stats_from_summary(s)
        if data:
            return data
    r = find_latest_report_file(run_dir)
    if r:
        return extract_stats_from_html(r)
    return {}


def _parse_license_overview_from_html(path: Path) -> List[Dict[str, Any]]:
    try:
        html = path.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        return []
    pattern = re.compile(
        r'<h3 class="heading-25">([^<]+)</h3>\s*'
        r'<div class="alert alert-info alert-info-soft">\s*'
        r'<strong>Totaal:</strong>\s*(\d+)\s*&nbsp;\s*\|\s*&nbsp;\s*'
        r'<strong>Gebruikt:</strong>\s*(\d+)\s*&nbsp;\s*\|\s*&nbsp;\s*'
        r'<strong>Beschikbaar:</strong>\s*(\d+)\s*&nbsp;\s*\|\s*&nbsp;\s*'
        r'<strong>Benutting:</strong>\s*([\d.]+)%',
        re.I,
    )
    licenses: List[Dict[str, Any]] = []
    for sku, total, consumed, available, utilization in pattern.findall(html):
        licenses.append({
            "SkuPartNumber": sku.strip(),
            "Total": int(total),
            "Consumed": int(consumed),
            "Available": int(available),
            "Utilization": float(utilization),
        })
    return licenses


def _parse_app_registration_alerts_from_html(path: Path) -> List[Dict[str, Any]]:
    try:
        html = path.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        return []
    pattern = re.compile(
        r"<tr class='appreg-row[^']*'><td class='cell-pad-strong'>([^<]+)</td><td class='cell-pad-muted'>([^<]+)</td>"
        r"<td class='cell-pad'><strong>(\d+) secret\(s\)</strong><br><span class='perm-resource-title'>([^<]+)</span>(?:<br><span class='text-muted-sm2'>([^<]+)</span>)?</td>"
        r"<td class='cell-pad'><strong>(\d+) cert\(s\)</strong><br><span class='perm-resource-title'>([^<]+)</span>(?:<br><span class='text-muted-sm2'>([^<]+)</span>)?</td>",
        re.I,
    )
    items: List[Dict[str, Any]] = []
    for match in pattern.findall(html):
        display_name, created, secret_count, secret_status, secret_date, cert_count, cert_status, cert_date = match
        items.append({
            "DisplayName": display_name.strip(),
            "CreatedDateTime": created.strip(),
            "SecretCount": int(secret_count),
            "SecretExpirationStatus": secret_status.strip(),
            "SecretExpiration": secret_date.strip() if secret_date else None,
            "CertificateCount": int(cert_count),
            "CertificateExpirationStatus": cert_status.strip(),
            "CertificateExpiration": cert_date.strip() if cert_date else None,
        })
    return items


def _parse_domain_dns_checks_from_html(path: Path) -> List[Dict[str, Any]]:
    try:
        html = path.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        return []
    match = re.search(
        r"DNS Records \(SPF/DKIM/DMARC\)</h3><div class='table-container'><table[^>]*>.*?<tbody>(.*?)</tbody></table>",
        html,
        re.I | re.S,
    )
    if not match:
        return []
    body = match.group(1)
    rows = re.findall(r"<tr><td>([^<]+)</td><td>([^<]+)</td><td>([^<]+)</td><td>([^<]+)</td></tr>", body, re.I)
    return [{"Domain": d.strip(), "SPF": spf.strip(), "DMARC": dmarc.strip(), "DKIM": dkim.strip()} for d, spf, dmarc, dkim in rows]


def _parse_user_mailboxes_from_html(path: Path) -> List[Dict[str, Any]]:
    try:
        html = path.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        return []
    match = re.search(
        r"User Mailboxes \(\d+\)</h3><div class='table-search-wrap'>.*?<tbody>(.*?)</tbody></table>",
        html,
        re.I | re.S,
    )
    if not match:
        return []
    body = match.group(1)
    rows = re.findall(r"<tr><td>([^<]+)</td><td>([^<]+)</td><td>([^<]+)</td></tr>", body, re.I)
    result = []
    for email, display_name, created in rows:
        result.append({
            "PrimarySmtpAddress": email.strip(),
            "DisplayName": display_name.strip(),
            "WhenCreated": created.strip(),
        })
    return result


def _valid_domain_dns_checks(items: Any) -> bool:
    if not isinstance(items, list) or not items:
        return False
    return any(isinstance(item, dict) and (item.get("Domain") or item.get("domain")) for item in items)


def _latest_assessment_snapshot_for_tenant(tid: str) -> Dict[str, Any]:
    run = db_fetchone(
        """
        SELECT * FROM assessment_runs
        WHERE tenant_id=?
        ORDER BY is_archived ASC, COALESCE(completed_at, started_at) DESC
        LIMIT 1
        """,
        (tid,),
    )
    if not run:
        return {}
    run_dir = RUNS_DIR / run["id"]
    summary_file = find_latest_summary_file(run_dir)
    report_file = find_latest_report_file(run_dir)
    snapshot: Dict[str, Any] = {}
    if summary_file and summary_file.exists():
        try:
            snapshot = json.loads(summary_file.read_text(encoding="utf-8"))
        except Exception:
            snapshot = {}
    licenses = snapshot.get("Licenses") if isinstance(snapshot, dict) else None
    app_registrations = snapshot.get("AppRegistrations") if isinstance(snapshot, dict) else None
    domain_dns_checks = snapshot.get("DomainDnsChecks") if isinstance(snapshot, dict) else None
    user_mailboxes = snapshot.get("UserMailboxes") if isinstance(snapshot, dict) else None
    if not licenses and report_file and report_file.exists():
        licenses = _parse_license_overview_from_html(report_file)
    if not app_registrations and report_file and report_file.exists():
        app_registrations = _parse_app_registration_alerts_from_html(report_file)
    if (not _valid_domain_dns_checks(domain_dns_checks)) and report_file and report_file.exists():
        domain_dns_checks = _parse_domain_dns_checks_from_html(report_file)
    if not user_mailboxes and report_file and report_file.exists():
        user_mailboxes = _parse_user_mailboxes_from_html(report_file)
    metrics = snapshot.get("Metrics") if isinstance(snapshot, dict) else {}
    licenses = licenses or []
    app_registrations = app_registrations or []
    domain_dns_checks = domain_dns_checks or []
    user_mailboxes = user_mailboxes or []
    total = sum(int(item.get("Total") or 0) for item in licenses if isinstance(item, dict))
    used = sum(int(item.get("Consumed") or 0) for item in licenses if isinstance(item, dict))
    license_type = None
    if len(licenses) == 1 and isinstance(licenses[0], dict):
        license_type = licenses[0].get("SkuPartNumber")
    elif licenses:
        license_type = f"{len(licenses)} licentietypen"
    mfa = None
    if isinstance(metrics, dict) and metrics.get("MfaCoveragePct") is not None:
        mfa = f"{metrics.get('MfaCoveragePct')}% dekking"
    return {
        "tenant_name": snapshot.get("TenantName") if isinstance(snapshot, dict) else None,
        "tenant_id": snapshot.get("TenantId") if isinstance(snapshot, dict) else None,
        "license_type": license_type,
        "licenses_total": total or None,
        "licenses_used": used or None,
        "mfa": mfa,
        "mfa_coverage": (metrics or {}).get("MfaCoveragePct"),
        "users_without_mfa": (metrics or {}).get("MfaMissing"),
        "ca_policies": (metrics or {}).get("CAEnabled"),
        "secure_score_percentage": (metrics or {}).get("SecureScorePct"),
        "conditional_access": int((metrics or {}).get("CAEnabled") or 0) > 0,
        "assessment_generated_at": snapshot.get("GeneratedAt") if isinstance(snapshot, dict) else None,
        "assessment_report_id": snapshot.get("AssessmentId") if isinstance(snapshot, dict) else None,
        "assessment_licenses": licenses,
        "assessment_app_registrations": app_registrations,
        "assessment_domain_dns_checks": domain_dns_checks,
        "assessment_user_mailboxes": user_mailboxes,
    }


def assessment_ui_nav(tenant_id: str) -> Dict[str, Any]:
    tenant = db_fetchone("SELECT * FROM tenants WHERE id=?", (tenant_id,))
    if not tenant:
        raise ValueError("Tenant niet gevonden")
    run = db_fetchone(
        """
        SELECT * FROM assessment_runs
        WHERE tenant_id=?
        ORDER BY is_archived ASC, COALESCE(completed_at, started_at) DESC
        LIMIT 1
        """,
        (tenant_id,),
    )
    snapshot = _latest_assessment_snapshot_for_tenant(tenant_id)
    users = [u for u in (snapshot.get("assessment_user_mailboxes") or []) if isinstance(u, dict)]
    domains = [d for d in (snapshot.get("assessment_domain_dns_checks") or []) if isinstance(d, dict) and (d.get("Domain") or d.get("domain"))]
    app_regs = [a for a in (snapshot.get("assessment_app_registrations") or []) if isinstance(a, dict)]
    licenses = [l for l in (snapshot.get("assessment_licenses") or []) if isinstance(l, dict)]
    items = [
        {"key": "summary", "label": "Overzicht", "count": None},
        {"key": "users", "label": "Gebruikers", "count": len(users)} if users else None,
        {"key": "licenses", "label": "Licenties", "count": len(licenses)} if licenses else None,
        {"key": "appregs", "label": "App Registraties", "count": len(app_regs)} if app_regs else None,
        {"key": "domains_dns", "label": "Domeinen & DNS", "count": len(domains)} if domains else None,
        {"key": "mfa_ca", "label": "MFA / CA", "count": snapshot.get("users_without_mfa")} if snapshot.get("mfa_coverage") is not None or snapshot.get("ca_policies") is not None else None,
    ]
    return {
        "enabled": bool(load_config().get("assessment_ui_v1", True)),
        "tenant_name": tenant.get("tenant_name") or tenant.get("customer_name"),
        "tenant_id": tenant.get("id"),
        "latest_run_id": run.get("id") if run else None,
        "latest_report_path": run.get("report_path") if run else None,
        "generated_at": snapshot.get("assessment_generated_at"),
        "score": run.get("score_overall") if run else None,
        "critical_count": run.get("critical_count") if run else 0,
        "warning_count": run.get("warning_count") if run else 0,
        "info_count": run.get("info_count") if run else 0,
        "items": [item for item in items if item],
    }


def assessment_ui_section(tenant_id: str, section_key: str) -> Dict[str, Any]:
    bundle = assessment_ui_nav(tenant_id)
    snapshot = _latest_assessment_snapshot_for_tenant(tenant_id)
    users = [u for u in (snapshot.get("assessment_user_mailboxes") or []) if isinstance(u, dict)]
    licenses = [l for l in (snapshot.get("assessment_licenses") or []) if isinstance(l, dict)]
    app_regs = [a for a in (snapshot.get("assessment_app_registrations") or []) if isinstance(a, dict)]
    domains = [d for d in (snapshot.get("assessment_domain_dns_checks") or []) if isinstance(d, dict) and (d.get("Domain") or d.get("domain"))]
    common = {
        "tenant_name": bundle["tenant_name"],
        "generated_at": bundle["generated_at"],
        "latest_run_id": bundle["latest_run_id"],
    }
    if section_key == "summary":
        return {
            **common,
            "key": "summary",
            "title": "Assessment overzicht",
            "cards": [
                {"label": "Secure Score", "value": f"{round(snapshot['secure_score_percentage'])}%" if snapshot.get("secure_score_percentage") is not None else "—", "tone": "success"},
                {"label": "MFA Coverage", "value": f"{round(snapshot['mfa_coverage'])}%" if snapshot.get("mfa_coverage") is not None else "—", "tone": "success"},
                {"label": "Open Alerts", "value": bundle["critical_count"] + bundle["warning_count"], "tone": "warn"},
                {"label": "CA Policies", "value": snapshot.get("ca_policies") or 0, "tone": "default"},
            ],
            "bars": [
                {"label": "Gebruikers", "value": len(users), "max": max(len(users), 1)},
                {"label": "Licenties", "value": len(licenses), "max": max(len(licenses), 1)},
                {"label": "App Registraties", "value": len(app_regs), "max": max(len(app_regs), 1)},
                {"label": "Tenantdomeinen", "value": len(domains), "max": max(len(domains), 1)},
            ],
        }
    if section_key == "users":
        rows = [{"name": u.get("DisplayName"), "email": u.get("PrimarySmtpAddress"), "created": u.get("WhenCreated")} for u in users]
        return {**common, "key": "users", "title": "Gebruikers", "columns": ["Naam", "E-mail", "Aangemaakt"], "rows": rows}
    if section_key == "licenses":
        rows = [{"sku": l.get("SkuPartNumber"), "total": l.get("Total"), "used": l.get("Consumed"), "available": l.get("Available"), "utilization": f"{l.get('Utilization')}%"} for l in licenses]
        return {**common, "key": "licenses", "title": "Licenties", "columns": ["SKU", "Totaal", "Gebruikt", "Beschikbaar", "Benutting"], "rows": rows}
    if section_key == "appregs":
        rows = [{"name": a.get("DisplayName"), "secret": a.get("SecretExpirationStatus"), "secret_expiry": a.get("SecretExpiration"), "certificate": a.get("CertificateExpirationStatus"), "permission_count": a.get("PermissionCount")} for a in app_regs]
        return {**common, "key": "appregs", "title": "App Registraties", "columns": ["Naam", "Secret", "Secret verval", "Certificaat", "Permissies"], "rows": rows}
    if section_key == "domains_dns":
        rows = [{"domain": d.get("Domain") or d.get("domain"), "spf": d.get("SPF") or d.get("spf"), "dmarc": d.get("DMARC") or d.get("dmarc"), "dkim": d.get("DKIM") or d.get("dkim")} for d in domains]
        return {**common, "key": "domains_dns", "title": "Domeinen & DNS", "columns": ["Domein", "SPF", "DMARC", "DKIM"], "rows": rows}
    if section_key == "mfa_ca":
        return {
            **common,
            "key": "mfa_ca",
            "title": "MFA / Conditional Access",
            "cards": [
                {"label": "MFA Coverage", "value": f"{round(snapshot['mfa_coverage'])}%" if snapshot.get("mfa_coverage") is not None else "—", "tone": "success"},
                {"label": "Gebruikers zonder MFA", "value": snapshot.get("users_without_mfa") or 0, "tone": "warn"},
                {"label": "CA Policies", "value": snapshot.get("ca_policies") or 0, "tone": "default"},
                {"label": "Conditional Access", "value": "Actief" if snapshot.get("conditional_access") else "Niet actief", "tone": "default"},
            ],
        }
    raise ValueError("Assessment onderdeel niet gevonden")


def gather_artifacts(run_id: str, run_dir: Path) -> Dict[str, Optional[str]]:
    report = find_latest_report_file(run_dir)
    summary = find_latest_summary_file(run_dir)
    result = {"report_path": None, "snapshot_path": None, "report_filename": None}
    if report:
        rel = report.relative_to(run_dir).as_posix()
        result["report_path"] = f"/reports/{run_id}/{rel}"
        result["report_filename"] = report.name
    if summary:
        rel = summary.relative_to(run_dir).as_posix()
        result["snapshot_path"] = f"/reports/{run_id}/{rel}"
    return result


def associate_run_to_tenant_by_summary(run_id: str, stats: Dict[str, Any]) -> None:
    """Bind/merge run to tenant based on snapshot TenantId (summary JSON)."""
    parsed_tenant_id = (stats.get("tenantId") or "").strip() if isinstance(stats, dict) else ""
    parsed_tenant_name = (stats.get("tenantName") or "").strip() if isinstance(stats, dict) else ""
    if not parsed_tenant_id:
        return

    run = db_fetchone("SELECT * FROM assessment_runs WHERE id=?", (run_id,))
    if not run:
        return
    current_tenant = db_fetchone("SELECT * FROM tenants WHERE id=?", (run["tenant_id"],))
    if not current_tenant:
        return

    current_guid = (current_tenant.get("tenant_guid") or "").strip()

    # Case 1: current tenant has no GUID yet -> enrich it
    if not current_guid:
        db_execute(
            "UPDATE tenants SET tenant_guid=?, tenant_name=COALESCE(NULLIF(?, ''), tenant_name), updated_at=? WHERE id=?",
            (parsed_tenant_id, parsed_tenant_name, now_iso(), current_tenant["id"]),
        )
        return

    # Case 2: current tenant already matches parsed tenant GUID -> optional name refresh
    if current_guid.lower() == parsed_tenant_id.lower():
        if parsed_tenant_name and parsed_tenant_name != current_tenant.get("tenant_name"):
            db_execute(
                "UPDATE tenants SET tenant_name=?, updated_at=? WHERE id=?",
                (parsed_tenant_name, now_iso(), current_tenant["id"]),
            )
        return

    # Case 3: mismatch -> look for existing tenant with parsed GUID and move run there
    existing = db_fetchone("SELECT * FROM tenants WHERE lower(COALESCE(tenant_guid,''))=lower(?) LIMIT 1", (parsed_tenant_id,))
    if existing:
        db_execute("UPDATE assessment_runs SET tenant_id=? WHERE id=?", (existing["id"], run_id))
        return

    # Case 4: no matching tenant exists -> create one and move run
    new_tenant = create_tenant(
        {
            "customer_name": parsed_tenant_name or "Auto-detected tenant",
            "tenant_name": parsed_tenant_name or "Auto-detected tenant",
            "tenant_guid": parsed_tenant_id,
            "notes": "Automatisch aangemaakt op basis van gegenereerd rapport (TenantId match).",
        }
    )
    if new_tenant and new_tenant.get("id"):
        db_execute("UPDATE assessment_runs SET tenant_id=? WHERE id=?", (new_tenant["id"], run_id))


# =============================================================================
# KB (Knowledge Base) — per-tenant SQLite helpers
# =============================================================================

KB_DIR = STORAGE_DIR / "kb"


def _kb_db_path(tenant_id: str) -> Path:
    safe = os.path.basename(tenant_id.replace("..", ""))
    d = KB_DIR / safe
    d.mkdir(parents=True, exist_ok=True)
    return d / "kb.sqlite"


def _kb_conn(tenant_id: str) -> sqlite3.Connection:
    conn = sqlite3.connect(str(_kb_db_path(tenant_id)), check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def _kb_init(conn: sqlite3.Connection) -> None:
    conn.executescript("""
    CREATE TABLE IF NOT EXISTS asset_types (
        id   INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL UNIQUE,
        icon TEXT DEFAULT '🖥️'
    );
    INSERT OR IGNORE INTO asset_types (name, icon) VALUES
        ('switch','🔀'),('router','🌐'),('firewall','🛡️'),
        ('ap','📡'),('server','🖥️'),('vlan','🏷️'),
        ('subnet','🕸️'),('circuit','🔌');
    CREATE TABLE IF NOT EXISTS kb_meta (
        key   TEXT PRIMARY KEY,
        value TEXT NOT NULL
    );
    INSERT OR IGNORE INTO kb_meta (key, value) VALUES
        ('categories', '["network","security","general","procedures","hardware"]'),
        ('vlan_purposes', '[{"key":"user","label":"Gebruikers"},{"key":"server","label":"Servers"},{"key":"mgmt","label":"Management"},{"key":"guest","label":"Gasten"},{"key":"iot","label":"IoT"},{"key":"dmz","label":"DMZ"}]');
    CREATE TABLE IF NOT EXISTS assets (
        id            INTEGER PRIMARY KEY AUTOINCREMENT,
        asset_type_id INTEGER REFERENCES asset_types(id),
        name          TEXT NOT NULL,
        hostname      TEXT, ip_address TEXT, location TEXT,
        vendor TEXT, model TEXT, firmware TEXT, serial TEXT,
        notes TEXT, is_active INTEGER DEFAULT 1,
        switch_config TEXT,
        created_at TEXT DEFAULT (datetime('now')),
        updated_at TEXT DEFAULT (datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS vlans (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vlan_id INTEGER NOT NULL, name TEXT NOT NULL,
        subnet TEXT, gateway TEXT, description TEXT,
        purpose TEXT DEFAULT 'user', notes TEXT,
        created_at TEXT DEFAULT (datetime('now')),
        updated_at TEXT DEFAULT (datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS kb_pages (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        title TEXT NOT NULL, content TEXT DEFAULT '',
        category TEXT DEFAULT 'network', order_index INTEGER DEFAULT 0,
        updated_at TEXT DEFAULT (datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS contacts (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL, role TEXT, phone TEXT, email TEXT,
        is_primary_contact INTEGER DEFAULT 0, notes TEXT,
        created_at TEXT DEFAULT (datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS kb_passwords (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL, category TEXT, username TEXT,
        secret_ref TEXT, strength INTEGER DEFAULT 0,
        last_updated TEXT, notes TEXT
    );
    CREATE TABLE IF NOT EXISTS kb_software (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL, vendor TEXT, software_type TEXT,
        licenses INTEGER, cost TEXT, expiry TEXT,
        status TEXT DEFAULT 'active', ref TEXT, notes TEXT,
        unit_price REAL, total_price REAL
    );
    CREATE TABLE IF NOT EXISTS kb_domains (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        domain TEXT NOT NULL, domain_type TEXT, registrar TEXT,
        expiry TEXT, ssl_expiry TEXT, ssl_issuer TEXT,
        status TEXT DEFAULT 'active', auto_renew INTEGER DEFAULT 0,
        nameservers TEXT, notes TEXT
    );
    CREATE TABLE IF NOT EXISTS kb_changelog (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        change_date TEXT NOT NULL, user_name TEXT, action TEXT NOT NULL,
        category TEXT, ref TEXT, notes TEXT
    );
    CREATE TABLE IF NOT EXISTS kb_m365_profile (
        id INTEGER PRIMARY KEY CHECK (id = 1),
        tenant_name TEXT, tenant_id TEXT, global_admin TEXT,
        license_type TEXT, licenses_total INTEGER, licenses_used INTEGER,
        mfa TEXT, conditional_access INTEGER DEFAULT 0, mdm TEXT,
        defender INTEGER DEFAULT 0, purview INTEGER DEFAULT 0,
        hybrid INTEGER DEFAULT 0, ad_connect TEXT,
        exchange_hybrid INTEGER DEFAULT 0, shared_mailboxes INTEGER DEFAULT 0,
        guest_users INTEGER DEFAULT 0, notes TEXT,
        updated_at TEXT DEFAULT (datetime('now'))
    );
    """)
    cols = {r["name"] for r in conn.execute("PRAGMA table_info(assets)").fetchall()}
    if "switch_config" not in cols:
        conn.execute("ALTER TABLE assets ADD COLUMN switch_config TEXT")
    software_cols = {r["name"] for r in conn.execute("PRAGMA table_info(kb_software)").fetchall()}
    if "unit_price" not in software_cols:
        conn.execute("ALTER TABLE kb_software ADD COLUMN unit_price REAL")
    if "total_price" not in software_cols:
        conn.execute("ALTER TABLE kb_software ADD COLUMN total_price REAL")
    conn.commit()


def _kb_rows(conn: sqlite3.Connection, sql: str, params: tuple = ()) -> List[Dict[str, Any]]:
    return [dict(r) for r in conn.execute(sql, params).fetchall()]


def _kb_row(conn: sqlite3.Connection, sql: str, params: tuple = ()) -> Optional[Dict[str, Any]]:
    r = conn.execute(sql, params).fetchone()
    return dict(r) if r else None


def _kb_tid(path: str) -> str:
    """Extract tenant_id from /api/kb/{tenant_id}/... paths."""
    return path.split("/")[3]


def _kb_iid(path: str) -> int:
    """Extract integer item id from last path segment."""
    return int(path.split("/")[-1])


# -- asset-types ---------------------------------------------------------------

def kb_list_asset_types(tid: str) -> List[Dict[str, Any]]:
    with _kb_conn(tid) as c:
        _kb_init(c)
        return _kb_rows(c, "SELECT * FROM asset_types ORDER BY name")


def kb_create_asset_type(tid: str, data: Dict[str, Any]) -> Dict[str, Any]:
    name = (data.get("name") or "").strip()
    if not name:
        raise ValueError("name is required")
    icon = (data.get("icon") or "🖥️").strip()
    with _kb_conn(tid) as c:
        _kb_init(c)
        cur = c.execute("INSERT INTO asset_types (name, icon) VALUES (?, ?)", (name, icon))
        c.commit()
        row = _kb_row(c, "SELECT * FROM asset_types WHERE id=?", (cur.lastrowid,))
    if not row:
        raise ValueError("Insert failed")
    return row


def kb_delete_asset_type(tid: str, type_id: int) -> None:
    with _kb_conn(tid) as c:
        _kb_init(c)
        c.execute("DELETE FROM asset_types WHERE id=?", (type_id,))
        c.commit()


# -- meta (categories & vlan_purposes) -----------------------------------------

import json as _json

def kb_get_meta(tid: str) -> Dict[str, Any]:
    with _kb_conn(tid) as c:
        _kb_init(c)
        rows = _kb_rows(c, "SELECT key, value FROM kb_meta")
    result: Dict[str, Any] = {}
    for r in rows:
        try:
            result[r["key"]] = _json.loads(r["value"])
        except Exception:
            result[r["key"]] = r["value"]
    return result


def kb_put_meta(tid: str, data: Dict[str, Any]) -> Dict[str, Any]:
    allowed = {"categories", "vlan_purposes"}
    with _kb_conn(tid) as c:
        _kb_init(c)
        for key in allowed:
            if key in data:
                c.execute("INSERT OR REPLACE INTO kb_meta (key, value) VALUES (?, ?)",
                          (key, _json.dumps(data[key], ensure_ascii=False)))
        c.commit()
    return kb_get_meta(tid)


# -- assets --------------------------------------------------------------------

def kb_list_assets(tid: str, asset_type: Optional[str] = None) -> List[Dict[str, Any]]:
    with _kb_conn(tid) as c:
        _kb_init(c)
        sql = ("SELECT a.*, t.name as type_name, t.icon as type_icon "
               "FROM assets a LEFT JOIN asset_types t ON a.asset_type_id=t.id")
        params: list = []
        if asset_type:
            sql += " WHERE t.name=?"
            params.append(asset_type)
        sql += " ORDER BY a.name"
        rows = _kb_rows(c, sql, tuple(params))
    for row in rows:
        try:
            row["switch_config"] = _json.loads(row.get("switch_config") or "null")
        except Exception:
            row["switch_config"] = None
    return rows


def kb_create_asset(tid: str, data: Dict[str, Any]) -> Dict[str, Any]:
    if not data.get("name"):
        raise ValueError("name is required")
    switch_config = data.get("switch_config")
    switch_config_json = _json.dumps(switch_config, ensure_ascii=False) if switch_config is not None else None
    with _kb_conn(tid) as c:
        _kb_init(c)
        cur = c.execute(
            "INSERT INTO assets (asset_type_id,name,hostname,ip_address,location,vendor,model,firmware,serial,notes,is_active,switch_config) "
            "VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
            (data.get("asset_type_id"), data["name"], data.get("hostname"), data.get("ip_address"),
             data.get("location"), data.get("vendor"), data.get("model"), data.get("firmware"),
             data.get("serial"), data.get("notes"), int(data.get("is_active", 1)), switch_config_json),
        )
        c.commit()
        row = _kb_row(c, "SELECT a.*, t.name as type_name, t.icon as type_icon "
                         "FROM assets a LEFT JOIN asset_types t ON a.asset_type_id=t.id WHERE a.id=?",
                      (cur.lastrowid,))
    if not row:
        raise ValueError("Insert failed")
    try:
        row["switch_config"] = _json.loads(row.get("switch_config") or "null")
    except Exception:
        row["switch_config"] = None
    return row


def kb_update_asset(tid: str, asset_id: int, data: Dict[str, Any]) -> Dict[str, Any]:
    switch_config = data.get("switch_config")
    switch_config_json = _json.dumps(switch_config, ensure_ascii=False) if switch_config is not None else None
    with _kb_conn(tid) as c:
        _kb_init(c)
        c.execute(
            "UPDATE assets SET asset_type_id=?,name=?,hostname=?,ip_address=?,location=?,vendor=?,model=?,"
            "firmware=?,serial=?,notes=?,is_active=?,switch_config=?,updated_at=datetime('now') WHERE id=?",
            (data.get("asset_type_id"), data.get("name"), data.get("hostname"), data.get("ip_address"),
             data.get("location"), data.get("vendor"), data.get("model"), data.get("firmware"),
             data.get("serial"), data.get("notes"), int(data.get("is_active", 1)), switch_config_json, asset_id),
        )
        c.commit()
        row = _kb_row(c, "SELECT a.*, t.name as type_name, t.icon as type_icon "
                         "FROM assets a LEFT JOIN asset_types t ON a.asset_type_id=t.id WHERE a.id=?",
                      (asset_id,))
    if not row:
        raise ValueError("Not found")
    try:
        row["switch_config"] = _json.loads(row.get("switch_config") or "null")
    except Exception:
        row["switch_config"] = None
    return row


def kb_delete_asset(tid: str, asset_id: int) -> None:
    with _kb_conn(tid) as c:
        _kb_init(c)
        c.execute("DELETE FROM assets WHERE id=?", (asset_id,))
        c.commit()


# -- vlans ---------------------------------------------------------------------

def kb_list_vlans(tid: str) -> List[Dict[str, Any]]:
    with _kb_conn(tid) as c:
        _kb_init(c)
        return _kb_rows(c, "SELECT * FROM vlans ORDER BY vlan_id")


def kb_create_vlan(tid: str, data: Dict[str, Any]) -> Dict[str, Any]:
    if not data.get("vlan_id") or not data.get("name"):
        raise ValueError("vlan_id and name are required")
    try:
        vnum = int(data["vlan_id"])
        if not (1 <= vnum <= 4094):
            raise ValueError
    except (ValueError, TypeError):
        raise ValueError("vlan_id must be 1-4094")
    with _kb_conn(tid) as c:
        _kb_init(c)
        cur = c.execute(
            "INSERT INTO vlans (vlan_id,name,subnet,gateway,description,purpose,notes) VALUES (?,?,?,?,?,?,?)",
            (vnum, data["name"], data.get("subnet"), data.get("gateway"),
             data.get("description"), data.get("purpose", "user"), data.get("notes")),
        )
        c.commit()
        row = _kb_row(c, "SELECT * FROM vlans WHERE id=?", (cur.lastrowid,))
    if not row:
        raise ValueError("Insert failed")
    return row


def kb_update_vlan(tid: str, vlan_db_id: int, data: Dict[str, Any]) -> Dict[str, Any]:
    with _kb_conn(tid) as c:
        _kb_init(c)
        c.execute(
            "UPDATE vlans SET vlan_id=?,name=?,subnet=?,gateway=?,description=?,purpose=?,notes=?,"
            "updated_at=datetime('now') WHERE id=?",
            (data.get("vlan_id"), data.get("name"), data.get("subnet"), data.get("gateway"),
             data.get("description"), data.get("purpose", "user"), data.get("notes"), vlan_db_id),
        )
        c.commit()
        row = _kb_row(c, "SELECT * FROM vlans WHERE id=?", (vlan_db_id,))
    if not row:
        raise ValueError("Not found")
    return row


def kb_delete_vlan(tid: str, vlan_db_id: int) -> None:
    with _kb_conn(tid) as c:
        _kb_init(c)
        c.execute("DELETE FROM vlans WHERE id=?", (vlan_db_id,))
        c.commit()


# -- pages ---------------------------------------------------------------------

def kb_list_pages(tid: str) -> List[Dict[str, Any]]:
    with _kb_conn(tid) as c:
        _kb_init(c)
        return _kb_rows(c, "SELECT id,title,category,order_index,updated_at FROM kb_pages ORDER BY order_index,title")


def kb_get_page(tid: str, page_id: int) -> Optional[Dict[str, Any]]:
    with _kb_conn(tid) as c:
        _kb_init(c)
        return _kb_row(c, "SELECT * FROM kb_pages WHERE id=?", (page_id,))


def kb_create_page(tid: str, data: Dict[str, Any]) -> Dict[str, Any]:
    if not data.get("title"):
        raise ValueError("title is required")
    with _kb_conn(tid) as c:
        _kb_init(c)
        cur = c.execute(
            "INSERT INTO kb_pages (title,content,category,order_index) VALUES (?,?,?,?)",
            (data["title"], data.get("content", ""), data.get("category", "network"), data.get("order_index", 0)),
        )
        c.commit()
        row = _kb_row(c, "SELECT * FROM kb_pages WHERE id=?", (cur.lastrowid,))
    if not row:
        raise ValueError("Insert failed")
    return row


def kb_update_page(tid: str, page_id: int, data: Dict[str, Any]) -> Dict[str, Any]:
    with _kb_conn(tid) as c:
        _kb_init(c)
        c.execute(
            "UPDATE kb_pages SET title=?,content=?,category=?,order_index=?,updated_at=datetime('now') WHERE id=?",
            (data.get("title"), data.get("content", ""), data.get("category", "network"),
             data.get("order_index", 0), page_id),
        )
        c.commit()
        row = _kb_row(c, "SELECT * FROM kb_pages WHERE id=?", (page_id,))
    if not row:
        raise ValueError("Not found")
    return row


def kb_delete_page(tid: str, page_id: int) -> None:
    with _kb_conn(tid) as c:
        _kb_init(c)
        c.execute("DELETE FROM kb_pages WHERE id=?", (page_id,))
        c.commit()


# -- contacts ------------------------------------------------------------------

def kb_list_contacts(tid: str) -> List[Dict[str, Any]]:
    with _kb_conn(tid) as c:
        _kb_init(c)
        return _kb_rows(c, "SELECT * FROM contacts ORDER BY is_primary_contact DESC, name")


def kb_create_contact(tid: str, data: Dict[str, Any]) -> Dict[str, Any]:
    if not data.get("name"):
        raise ValueError("name is required")
    with _kb_conn(tid) as c:
        _kb_init(c)
        cur = c.execute(
            "INSERT INTO contacts (name,role,phone,email,is_primary_contact,notes) VALUES (?,?,?,?,?,?)",
            (data["name"], data.get("role"), data.get("phone"), data.get("email"),
             int(data.get("is_primary_contact", 0)), data.get("notes")),
        )
        c.commit()
        row = _kb_row(c, "SELECT * FROM contacts WHERE id=?", (cur.lastrowid,))
    if not row:
        raise ValueError("Insert failed")
    return row


def kb_update_contact(tid: str, contact_id: int, data: Dict[str, Any]) -> Dict[str, Any]:
    with _kb_conn(tid) as c:
        _kb_init(c)
        c.execute(
            "UPDATE contacts SET name=?,role=?,phone=?,email=?,is_primary_contact=?,notes=? WHERE id=?",
            (data.get("name"), data.get("role"), data.get("phone"), data.get("email"),
             int(data.get("is_primary_contact", 0)), data.get("notes"), contact_id),
        )
        c.commit()
        row = _kb_row(c, "SELECT * FROM contacts WHERE id=?", (contact_id,))
    if not row:
        raise ValueError("Not found")
    return row


def kb_delete_contact(tid: str, contact_id: int) -> None:
    with _kb_conn(tid) as c:
        _kb_init(c)
        c.execute("DELETE FROM contacts WHERE id=?", (contact_id,))
        c.commit()


# -- passwords -----------------------------------------------------------------

def kb_list_passwords(tid: str) -> List[Dict[str, Any]]:
    with _kb_conn(tid) as c:
        _kb_init(c)
        return _kb_rows(c, "SELECT * FROM kb_passwords ORDER BY category, name")


def kb_create_password(tid: str, data: Dict[str, Any]) -> Dict[str, Any]:
    if not data.get("name"):
        raise ValueError("name is required")
    with _kb_conn(tid) as c:
        _kb_init(c)
        cur = c.execute(
            "INSERT INTO kb_passwords (name,category,username,secret_ref,strength,last_updated,notes) VALUES (?,?,?,?,?,?,?)",
            (data["name"], data.get("category"), data.get("username"), data.get("secret_ref"),
             int(data.get("strength") or 0), data.get("last_updated"), data.get("notes")),
        )
        c.commit()
        return _kb_row(c, "SELECT * FROM kb_passwords WHERE id=?", (cur.lastrowid,))


def kb_update_password(tid: str, item_id: int, data: Dict[str, Any]) -> Dict[str, Any]:
    with _kb_conn(tid) as c:
        _kb_init(c)
        c.execute(
            "UPDATE kb_passwords SET name=?,category=?,username=?,secret_ref=?,strength=?,last_updated=?,notes=? WHERE id=?",
            (data.get("name"), data.get("category"), data.get("username"), data.get("secret_ref"),
             int(data.get("strength") or 0), data.get("last_updated"), data.get("notes"), item_id),
        )
        c.commit()
        row = _kb_row(c, "SELECT * FROM kb_passwords WHERE id=?", (item_id,))
    if not row:
        raise ValueError("Not found")
    return row


def kb_delete_password(tid: str, item_id: int) -> None:
    with _kb_conn(tid) as c:
        _kb_init(c)
        c.execute("DELETE FROM kb_passwords WHERE id=?", (item_id,))
        c.commit()


# -- software ------------------------------------------------------------------

def kb_list_software(tid: str) -> List[Dict[str, Any]]:
    with _kb_conn(tid) as c:
        _kb_init(c)
        return _kb_rows(c, "SELECT * FROM kb_software ORDER BY vendor, name")


def kb_create_software(tid: str, data: Dict[str, Any]) -> Dict[str, Any]:
    if not data.get("name"):
        raise ValueError("name is required")
    licenses = data.get("licenses")
    unit_price = data.get("unit_price")
    total_price = data.get("total_price")
    if licenses not in (None, ""):
        try:
            licenses = int(licenses)
        except (TypeError, ValueError):
            licenses = None
    else:
        licenses = None
    if unit_price not in (None, ""):
        try:
            unit_price = float(unit_price)
        except (TypeError, ValueError):
            unit_price = None
    else:
        unit_price = None
    if total_price in (None, "") and licenses is not None and unit_price is not None:
        total_price = round(licenses * unit_price, 2)
    elif total_price not in (None, ""):
        try:
            total_price = float(total_price)
        except (TypeError, ValueError):
            total_price = None
    else:
        total_price = None
    with _kb_conn(tid) as c:
        _kb_init(c)
        cur = c.execute(
            "INSERT INTO kb_software (name,vendor,software_type,licenses,cost,expiry,status,ref,notes,unit_price,total_price) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
            (data["name"], data.get("vendor"), data.get("software_type"), licenses,
             data.get("cost"), data.get("expiry"), data.get("status", "active"), data.get("ref"), data.get("notes"),
             unit_price, total_price),
        )
        c.commit()
        return _kb_row(c, "SELECT * FROM kb_software WHERE id=?", (cur.lastrowid,))


def kb_update_software(tid: str, item_id: int, data: Dict[str, Any]) -> Dict[str, Any]:
    licenses = data.get("licenses")
    unit_price = data.get("unit_price")
    total_price = data.get("total_price")
    if licenses not in (None, ""):
        try:
            licenses = int(licenses)
        except (TypeError, ValueError):
            licenses = None
    else:
        licenses = None
    if unit_price not in (None, ""):
        try:
            unit_price = float(unit_price)
        except (TypeError, ValueError):
            unit_price = None
    else:
        unit_price = None
    if total_price in (None, "") and licenses is not None and unit_price is not None:
        total_price = round(licenses * unit_price, 2)
    elif total_price not in (None, ""):
        try:
            total_price = float(total_price)
        except (TypeError, ValueError):
            total_price = None
    else:
        total_price = None
    with _kb_conn(tid) as c:
        _kb_init(c)
        c.execute(
            "UPDATE kb_software SET name=?,vendor=?,software_type=?,licenses=?,cost=?,expiry=?,status=?,ref=?,notes=?,unit_price=?,total_price=? WHERE id=?",
            (data.get("name"), data.get("vendor"), data.get("software_type"), licenses,
             data.get("cost"), data.get("expiry"), data.get("status", "active"), data.get("ref"), data.get("notes"),
             unit_price, total_price, item_id),
        )
        c.commit()
        row = _kb_row(c, "SELECT * FROM kb_software WHERE id=?", (item_id,))
    if not row:
        raise ValueError("Not found")
    return row


def kb_delete_software(tid: str, item_id: int) -> None:
    with _kb_conn(tid) as c:
        _kb_init(c)
        c.execute("DELETE FROM kb_software WHERE id=?", (item_id,))
        c.commit()


# -- domains -------------------------------------------------------------------

def kb_list_domains(tid: str) -> List[Dict[str, Any]]:
    with _kb_conn(tid) as c:
        _kb_init(c)
        return _kb_rows(c, "SELECT * FROM kb_domains ORDER BY domain")


def kb_create_domain(tid: str, data: Dict[str, Any]) -> Dict[str, Any]:
    if not data.get("domain"):
        raise ValueError("domain is required")
    with _kb_conn(tid) as c:
        _kb_init(c)
        cur = c.execute(
            "INSERT INTO kb_domains (domain,domain_type,registrar,expiry,ssl_expiry,ssl_issuer,status,auto_renew,nameservers,notes) VALUES (?,?,?,?,?,?,?,?,?,?)",
            (data["domain"], data.get("domain_type"), data.get("registrar"), data.get("expiry"),
             data.get("ssl_expiry"), data.get("ssl_issuer"), data.get("status", "active"),
             int(data.get("auto_renew", 0)), data.get("nameservers"), data.get("notes")),
        )
        c.commit()
        return _kb_row(c, "SELECT * FROM kb_domains WHERE id=?", (cur.lastrowid,))


def kb_update_domain(tid: str, item_id: int, data: Dict[str, Any]) -> Dict[str, Any]:
    with _kb_conn(tid) as c:
        _kb_init(c)
        c.execute(
            "UPDATE kb_domains SET domain=?,domain_type=?,registrar=?,expiry=?,ssl_expiry=?,ssl_issuer=?,status=?,auto_renew=?,nameservers=?,notes=? WHERE id=?",
            (data.get("domain"), data.get("domain_type"), data.get("registrar"), data.get("expiry"),
             data.get("ssl_expiry"), data.get("ssl_issuer"), data.get("status", "active"),
             int(data.get("auto_renew", 0)), data.get("nameservers"), data.get("notes"), item_id),
        )
        c.commit()
        row = _kb_row(c, "SELECT * FROM kb_domains WHERE id=?", (item_id,))
    if not row:
        raise ValueError("Not found")
    return row


def kb_delete_domain(tid: str, item_id: int) -> None:
    with _kb_conn(tid) as c:
        _kb_init(c)
        c.execute("DELETE FROM kb_domains WHERE id=?", (item_id,))
        c.commit()


# -- m365 profile --------------------------------------------------------------

def kb_get_m365_profile(tid: str) -> Dict[str, Any]:
    with _kb_conn(tid) as c:
        _kb_init(c)
        row = _kb_row(c, "SELECT * FROM kb_m365_profile WHERE id=1")
    base = row or {
        "id": 1, "tenant_name": None, "tenant_id": None, "global_admin": None,
        "license_type": None, "licenses_total": None, "licenses_used": None,
        "mfa": None, "conditional_access": 0, "mdm": None, "defender": 0,
        "purview": 0, "hybrid": 0, "ad_connect": None, "exchange_hybrid": 0,
        "shared_mailboxes": 0, "guest_users": 0, "notes": None,
    }
    assessment = _latest_assessment_snapshot_for_tenant(tid)
    if assessment:
        for key in ("tenant_name", "tenant_id", "license_type", "licenses_total", "licenses_used", "mfa"):
            if assessment.get(key) not in (None, ""):
                base[key] = assessment[key]
        if assessment.get("conditional_access") is not None:
            base["conditional_access"] = 1 if assessment["conditional_access"] else 0
        base["assessment_generated_at"] = assessment.get("assessment_generated_at")
        base["assessment_report_id"] = assessment.get("assessment_report_id")
        base["assessment_licenses"] = assessment.get("assessment_licenses") or []
        base["assessment_app_registrations"] = assessment.get("assessment_app_registrations") or []
        base["assessment_domain_dns_checks"] = assessment.get("assessment_domain_dns_checks") or []
        base["assessment_user_mailboxes"] = assessment.get("assessment_user_mailboxes") or []
    else:
        base["assessment_generated_at"] = None
        base["assessment_report_id"] = None
        base["assessment_licenses"] = []
        base["assessment_app_registrations"] = []
        base["assessment_domain_dns_checks"] = []
        base["assessment_user_mailboxes"] = []
    return base


def kb_put_m365_profile(tid: str, data: Dict[str, Any]) -> Dict[str, Any]:
    with _kb_conn(tid) as c:
        _kb_init(c)
        c.execute(
            "INSERT OR REPLACE INTO kb_m365_profile (id,tenant_name,tenant_id,global_admin,license_type,licenses_total,licenses_used,mfa,conditional_access,mdm,defender,purview,hybrid,ad_connect,exchange_hybrid,shared_mailboxes,guest_users,notes,updated_at) VALUES (1,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,datetime('now'))",
            (data.get("tenant_name"), data.get("tenant_id"), data.get("global_admin"), data.get("license_type"),
             data.get("licenses_total"), data.get("licenses_used"), data.get("mfa"), int(data.get("conditional_access", 0)),
             data.get("mdm"), int(data.get("defender", 0)), int(data.get("purview", 0)), int(data.get("hybrid", 0)),
             data.get("ad_connect"), int(data.get("exchange_hybrid", 0)), data.get("shared_mailboxes"),
             data.get("guest_users"), data.get("notes")),
        )
        c.commit()
    return kb_get_m365_profile(tid)


# -- changelog -----------------------------------------------------------------

def kb_list_changelog(tid: str) -> List[Dict[str, Any]]:
    with _kb_conn(tid) as c:
        _kb_init(c)
        return _kb_rows(c, "SELECT * FROM kb_changelog ORDER BY change_date DESC, id DESC")


def kb_create_changelog(tid: str, data: Dict[str, Any]) -> Dict[str, Any]:
    if not data.get("change_date") or not data.get("action"):
        raise ValueError("change_date and action are required")
    with _kb_conn(tid) as c:
        _kb_init(c)
        cur = c.execute(
            "INSERT INTO kb_changelog (change_date,user_name,action,category,ref,notes) VALUES (?,?,?,?,?,?)",
            (data["change_date"], data.get("user_name"), data["action"], data.get("category"), data.get("ref"), data.get("notes")),
        )
        c.commit()
        return _kb_row(c, "SELECT * FROM kb_changelog WHERE id=?", (cur.lastrowid,))


def kb_update_changelog(tid: str, item_id: int, data: Dict[str, Any]) -> Dict[str, Any]:
    with _kb_conn(tid) as c:
        _kb_init(c)
        c.execute(
            "UPDATE kb_changelog SET change_date=?,user_name=?,action=?,category=?,ref=?,notes=? WHERE id=?",
            (data.get("change_date"), data.get("user_name"), data.get("action"), data.get("category"), data.get("ref"), data.get("notes"), item_id),
        )
        c.commit()
        row = _kb_row(c, "SELECT * FROM kb_changelog WHERE id=?", (item_id,))
    if not row:
        raise ValueError("Not found")
    return row


def kb_delete_changelog(tid: str, item_id: int) -> None:
    with _kb_conn(tid) as c:
        _kb_init(c)
        c.execute("DELETE FROM kb_changelog WHERE id=?", (item_id,))
        c.commit()


# =============================================================================

class RunManager:
    def __init__(self) -> None:
        self._active_procs: Dict[str, "subprocess.Popen[str]"] = {}
        self._stop_requested: set = set()
        self._lock = threading.Lock()

    def start(self, run_id: str, phases: List[str], run_mode: str, scan_type: str = "full") -> None:
        t = threading.Thread(target=self._worker, args=(run_id, phases, run_mode, scan_type), daemon=True)
        t.start()

    def stop(self, run_id: str) -> bool:
        """Stuur SIGTERM naar het actieve proces voor deze run. Retourneert True als gevonden."""
        with self._lock:
            proc = self._active_procs.get(run_id)
            self._stop_requested.add(run_id)
        if not proc:
            return False
        try:
            proc.terminate()
            try:
                proc.wait(timeout=5)
            except subprocess.TimeoutExpired:
                proc.kill()
        except Exception:
            pass
        return True

    def _register_proc(self, run_id: str, proc: "subprocess.Popen[str]") -> None:
        with self._lock:
            self._active_procs[run_id] = proc

    def _unregister_proc(self, run_id: str) -> None:
        with self._lock:
            self._active_procs.pop(run_id, None)

    def _was_stop_requested(self, run_id: str) -> bool:
        with self._lock:
            return run_id in self._stop_requested

    def _clear_stop_flag(self, run_id: str) -> None:
        with self._lock:
            self._stop_requested.discard(run_id)

    def _worker(self, run_id: str, phases: List[str], run_mode: str, scan_type: str = "full") -> None:
        run_dir = RUNS_DIR / run_id
        (run_dir / "_snapshots").mkdir(parents=True, exist_ok=True)
        update_run(run_id, status="running")
        append_run_log(run_id, f"Run mode: {run_mode}")
        append_run_log(run_id, f"Scan type: {scan_type}")
        append_run_log(run_id, f"Phases: {', '.join(phases)}")
        try:
            if scan_type == "entrafalcon":
                self._run_entrafalcon(run_id, run_dir)
            elif run_mode == "script":
                self._run_script(run_id, phases, run_dir)
            else:
                self._run_demo(run_id, phases, run_dir)

            artifacts = gather_artifacts(run_id, run_dir)
            stats = parse_run_stats(run_dir)
            associate_run_to_tenant_by_summary(run_id, stats)
            update_run(
                run_id,
                status="completed",
                completed_at=now_iso(),
                exit_code=0,
                report_path=artifacts["report_path"],
                snapshot_path=artifacts["snapshot_path"],
                report_filename=artifacts["report_filename"],
                score_overall=stats.get("scoreOverall"),
                critical_count=stats.get("criticalIssues") or 0,
                warning_count=stats.get("warnings") or 0,
                info_count=stats.get("infoItems") or 0,
            )
            append_run_log(run_id, "Run completed.")
        except Exception as exc:
            cancelled = self._was_stop_requested(run_id)
            self._clear_stop_flag(run_id)
            status = "cancelled" if cancelled else "failed"
            update_run(
                run_id,
                status=status,
                completed_at=now_iso(),
                exit_code=1,
                error_message=str(exc),
            )
            append_run_log(run_id, f"Run {status}: {exc}")

    def _run_demo(self, run_id: str, phases: List[str], run_dir: Path) -> None:
        labels = {
            "phase1": "Users",
            "phase2": "Collaboration",
            "phase3": "Compliance",
            "phase4": "Security",
            "phase5": "Intune",
            "phase6": "Azure",
        }
        for p in phases:
            append_run_log(run_id, f"Starting {p} ({labels.get(p, p)})")
            time.sleep(0.7)
            append_run_log(run_id, f"Completed {p}")

        src_report = DEFAULT_REPORTS_DIR / "M365-Complete-Baseline-latest.html"
        src_summary = DEFAULT_REPORTS_DIR / "_snapshots" / "M365-Complete-Baseline-latest.summary.json"
        run_stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        if src_report.exists():
            unique_report_name = f"M365-Complete-Baseline-{run_stamp}-{run_id[:8]}.html"
            shutil.copy2(src_report, run_dir / unique_report_name)
            # Convenience copy for quick open/debug (kept in same run dir)
            shutil.copy2(src_report, run_dir / "M365-Complete-Baseline-latest.html")
        else:
            demo_html = (
                "<!doctype html><html><head><meta charset='utf-8'><title>Demo report</title></head>"
                "<body><div class='tenant-name'>Lokale Tenant</div><h1>Demo assessment</h1></body></html>"
            )
            unique_report_name = f"M365-Complete-Baseline-{run_stamp}-{run_id[:8]}.html"
            (run_dir / unique_report_name).write_text(demo_html, encoding="utf-8")
            (run_dir / "M365-Complete-Baseline-latest.html").write_text(demo_html, encoding="utf-8")
        if src_summary.exists():
            (run_dir / "_snapshots").mkdir(exist_ok=True)
            unique_summary_name = f"M365-Complete-Baseline-{run_stamp}-{run_id[:8]}.summary.json"
            shutil.copy2(src_summary, run_dir / "_snapshots" / unique_summary_name)
            shutil.copy2(src_summary, run_dir / "_snapshots" / "M365-Complete-Baseline-latest.summary.json")
        append_run_log(run_id, "Demo artifacts generated.")

    def _run_script(self, run_id: str, phases: List[str], run_dir: Path) -> None:
        cfg = load_config()
        script_path = Path(cfg.get("script_path") or "").expanduser().resolve()
        # Whitelist: alleen scripts binnen de assessment-map zijn toegestaan
        _allowed_dir = (PLATFORM_DIR / "assessment").resolve()
        if not str(script_path).startswith(str(_allowed_dir)):
            raise ValueError(f"Script-pad staat niet op de whitelist: {script_path}")
        if not script_path.exists():
            raise RuntimeError(f"Script not found: {script_path}")
        pwsh = shutil.which("pwsh") or shutil.which("powershell")
        if not pwsh:
            raise RuntimeError("PowerShell (pwsh/powershell) not found")

        cmd = [pwsh, "-NoLogo", "-NoProfile", "-NonInteractive", "-File", str(script_path),
               "-OutputPath", str(run_dir),
               "-ExportCsv"]        # v3.2: altijd CSV exporteren naast HTML
        cmd.extend(phase_skip_flags(phases))

        # Authenticatie parameters doorgeven indien geconfigureerd
        tenant_id   = (cfg.get("auth_tenant_id")       or "").strip()
        client_id   = (cfg.get("auth_client_id")       or "").strip()
        cert_thumb  = (cfg.get("auth_cert_thumbprint") or "").strip()
        client_sec  = (cfg.get("auth_client_secret")   or "").strip()

        # env aanmaken vóór gebruik (fix: was na de client_sec blok)
        env = os.environ.copy()
        env["M365_BASELINE_NONINTERACTIVE"] = "1"
        env["CI"] = env.get("CI", "1")

        if tenant_id:
            cmd += ["-TenantId", tenant_id]
        if client_id:
            cmd += ["-ClientId", client_id]
        if cert_thumb:
            cmd += ["-CertThumbprint", cert_thumb]
        elif client_sec:
            # Client secret via omgevingsvariabele (nooit als plaintext in cmd)
            env["M365_CLIENT_SECRET"] = client_sec
            cmd += ["-ClientSecret",
                    "(ConvertTo-SecureString $env:M365_CLIENT_SECRET -AsPlainText -Force)"]
        append_run_log(run_id, "Starting PowerShell assessment.")
        append_run_log(run_id, "Command: " + " ".join(cmd))
        proc = subprocess.Popen(
            cmd,
            cwd=str(PLATFORM_DIR / "assessment"),
            env=env,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            universal_newlines=True,
            bufsize=1,
        )
        self._register_proc(run_id, proc)
        try:
            assert proc.stdout is not None
            for line in proc.stdout:
                line = line.rstrip()
                if line:
                    append_run_log(run_id, line)
            proc.stdout.close()
        finally:
            self._unregister_proc(run_id)
        rc = proc.wait()
        append_run_log(run_id, f"PowerShell process exited with code {rc}")
        if rc != 0:
            raise RuntimeError(f"PowerShell exited with code {rc}")


    def _run_entrafalcon(self, run_id: str, run_dir: Path) -> None:
        cfg = load_config()
        script_path = Path(cfg.get("entrafalcon_script_path") or "").expanduser().resolve()
        _allowed_dir = (PLATFORM_DIR / "assessment" / "EntraFalcon").resolve()
        if not str(script_path).startswith(str(_allowed_dir)):
            raise ValueError(f"EntraFalcon script-pad staat niet op de whitelist: {script_path}")
        if not script_path.exists():
            raise RuntimeError(f"EntraFalcon script niet gevonden: {script_path}")

        pwsh = shutil.which("pwsh") or shutil.which("powershell")
        if not pwsh:
            raise RuntimeError("PowerShell (pwsh/powershell) niet gevonden")

        tenant_id  = (cfg.get("auth_tenant_id")     or "").strip()
        client_id  = (cfg.get("auth_client_id")     or "").strip()
        client_sec = (cfg.get("auth_client_secret") or "").strip()
        use_app_auth = bool(tenant_id and client_id and client_sec)

        if use_app_auth:
            cmd = [
                pwsh, "-NoLogo", "-NoProfile", "-NonInteractive",
                "-File", str(script_path),
                "-OutputFolder", str(run_dir),
                "-AuthFlow", "ClientCredentials",
                "-ClientId", client_id,
                "-AppTenantId", tenant_id,
                "-Tenant", tenant_id,
            ]
        else:
            cmd = [
                pwsh, "-NoLogo", "-NoProfile", "-NonInteractive",
                "-File", str(script_path),
                "-OutputFolder", str(run_dir),
                "-AuthFlow", "DeviceCode",
            ]

        if cfg.get("entrafalcon_include_ms_apps"):
            cmd.append("-IncludeMsApps")
        if cfg.get("entrafalcon_csv"):
            cmd.append("-Csv")

        append_run_log(run_id, "Starting Entra Security assessment.")
        if use_app_auth:
            append_run_log(run_id, f"AuthFlow: ClientCredentials (app registration: {client_id[:8]}...)")
        else:
            append_run_log(run_id, "AuthFlow: DeviceCode")
            append_run_log(run_id, "⚠ INTERACTIEF: Volg de authenticatie-instructies in de PowerShell console.")

        env = os.environ.copy()
        if use_app_auth:
            env["EF_CLIENT_SECRET"] = client_sec
        proc = subprocess.Popen(
            cmd,
            cwd=str(script_path.parent),
            env=env,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            universal_newlines=True,
            bufsize=1,
        )
        self._register_proc(run_id, proc)
        try:
            assert proc.stdout is not None
            for line in proc.stdout:
                line = line.rstrip()
                if line:
                    append_run_log(run_id, line)
            proc.stdout.close()
        finally:
            self._unregister_proc(run_id)
        rc = proc.wait()
        append_run_log(run_id, f"EntraFalcon process exited with code {rc}")
        if rc != 0:
            raise RuntimeError(f"EntraFalcon exited with code {rc}")


RUN_MANAGER = RunManager()


# ══════════════════════════════════════════════════════════════════════════════
# USER MANAGEMENT
# ══════════════════════════════════════════════════════════════════════════════

_GOD_ADMIN_EMAIL = os.environ.get("DENJOY_ADMIN_EMAIL", "schiphorst.d@gmail.com").strip().lower()


def list_users() -> List[Dict[str, Any]]:
    rows = db_fetchall(
        "SELECT id, email, role, display_name, linked_tenant_id, is_active, created_at "
        "FROM users ORDER BY role DESC, created_at ASC"
    )
    result = []
    for r in rows:
        d = dict(r)
        d["is_god_admin"] = (d["email"].lower() == _GOD_ADMIN_EMAIL)
        result.append(d)
    return result


def get_user(user_id: str) -> Optional[Dict[str, Any]]:
    row = db_fetchone(
        "SELECT id, email, role, display_name, linked_tenant_id, is_active, created_at "
        "FROM users WHERE id=?", (user_id,)
    )
    if not row:
        return None
    d = dict(row)
    d["is_god_admin"] = (d["email"].lower() == _GOD_ADMIN_EMAIL)
    return d


def create_user_account(payload: Dict[str, Any]) -> Dict[str, Any]:
    email = (payload.get("email") or "").strip().lower()
    password = (payload.get("password") or "").strip()
    role = (payload.get("role") or "klant").strip()
    display_name = (payload.get("display_name") or "").strip()
    linked_tenant_id = payload.get("linked_tenant_id") or None

    if not email:
        raise ValueError("E-mailadres is verplicht.")
    if not password:
        raise ValueError("Wachtwoord is verplicht.")
    if len(password) < 8:
        raise ValueError("Wachtwoord moet minimaal 8 tekens zijn.")
    if role not in ("admin", "klant"):
        raise ValueError("Ongeldige rol — kies 'admin' of 'klant'.")
    if db_fetchone("SELECT id FROM users WHERE lower(email)=?", (email,)):
        raise ValueError(f"E-mailadres '{email}' bestaat al.")

    pw_hash, salt = _hash_pw(password)
    uid = str(uuid.uuid4())
    db_execute(
        "INSERT INTO users (id, email, password_hash, salt, role, display_name, "
        "linked_tenant_id, is_active, created_at) VALUES (?,?,?,?,?,?,?,1,?)",
        (uid, email, pw_hash, salt, role, display_name, linked_tenant_id, now_iso())
    )
    return get_user(uid)


def update_user_account(user_id: str, payload: Dict[str, Any], requesting_email: str) -> Dict[str, Any]:
    user = get_user(user_id)
    if not user:
        raise ValueError("Gebruiker niet gevonden.")

    is_god = user["email"].lower() == _GOD_ADMIN_EMAIL
    updates: Dict[str, Any] = {}

    if "display_name" in payload:
        updates["display_name"] = (payload["display_name"] or "").strip()
    if "role" in payload:
        if is_god:
            raise ValueError("De rol van het God-Admin account kan niet worden gewijzigd.")
        if payload["role"] not in ("admin", "klant"):
            raise ValueError("Ongeldige rol.")
        updates["role"] = payload["role"]
    if "linked_tenant_id" in payload:
        updates["linked_tenant_id"] = payload["linked_tenant_id"] or None
    if "is_active" in payload:
        if is_god and not payload["is_active"]:
            raise ValueError("Het God-Admin account kan niet worden gedeactiveerd.")
        updates["is_active"] = 1 if payload["is_active"] else 0
    if "password" in payload and payload["password"]:
        pw = payload["password"].strip()
        if len(pw) < 8:
            raise ValueError("Wachtwoord moet minimaal 8 tekens zijn.")
        pw_hash, salt = _hash_pw(pw)
        updates["password_hash"] = pw_hash
        updates["salt"] = salt

    if not updates:
        return user

    set_clause = ", ".join(f"{k}=?" for k in updates)
    db_execute(f"UPDATE users SET {set_clause} WHERE id=?", list(updates.values()) + [user_id])
    return get_user(user_id)


def delete_user_account(user_id: str, requesting_email: str) -> Dict[str, Any]:
    user = get_user(user_id)
    if not user:
        raise ValueError("Gebruiker niet gevonden.")
    if user["email"].lower() == _GOD_ADMIN_EMAIL:
        raise ValueError("Het God-Admin account kan niet worden verwijderd.")
    if user["email"].lower() == requesting_email.lower():
        raise ValueError("Je kunt je eigen account niet verwijderen.")
    # Verwijder actieve sessies van deze gebruiker
    db_execute("DELETE FROM sessions WHERE user_id=?", (user_id,))
    db_execute("DELETE FROM users WHERE id=?", (user_id,))
    return {"ok": True, "deleted_id": user_id}


# ══════════════════════════════════════════════════════════════════════════════


def list_tenants() -> List[Dict[str, Any]]:
    tenants = db_fetchall("SELECT * FROM tenants WHERE is_active=1 ORDER BY customer_name, tenant_name")
    for t in tenants:
        latest = db_fetchone("SELECT * FROM assessment_runs WHERE tenant_id=? ORDER BY started_at DESC LIMIT 1", (t["id"],))
        t["latest_run"] = latest
    return tenants


def update_tenant(tenant_id: str, payload: Dict[str, Any]) -> Dict[str, Any]:
    tenant = db_fetchone("SELECT * FROM tenants WHERE id=?", (tenant_id,))
    if not tenant:
        raise ValueError("Tenant niet gevonden")

    allowed = {
        "customer_name",
        "tenant_name",
        "tenant_guid",
        "status",
        "owner_primary",
        "owner_backup",
        "tags_csv",
        "risk_profile",
        "notes",
    }
    fields: Dict[str, Any] = {}
    for k, v in payload.items():
        if k not in allowed:
            continue
        if isinstance(v, str):
            fields[k] = v.strip()
        else:
            fields[k] = v
    if "status" in fields and fields["status"] not in {"active", "onboarding", "paused", "offboarded"}:
        raise ValueError("Ongeldige status")
    if "risk_profile" in fields and fields["risk_profile"] not in {"low", "standard", "high", "critical"}:
        raise ValueError("Ongeldig risicoprofiel")
    fields["updated_at"] = now_iso()

    keys = list(fields.keys())
    if keys:
        sql = "UPDATE tenants SET " + ", ".join([f"{k}=?" for k in keys]) + " WHERE id=?"
        vals = [fields[k] for k in keys] + [tenant_id]
        db_execute(sql, tuple(vals))
    return db_fetchone("SELECT * FROM tenants WHERE id=?", (tenant_id,)) or {}


def ensure_demo_tenant_if_empty() -> None:
    row = db_fetchone("SELECT COUNT(*) AS cnt FROM tenants WHERE is_active=1")
    cnt = int((row or {}).get("cnt") or 0)
    if cnt > 0:
        return
    tenant_id = str(uuid.uuid4())
    ts = now_iso()
    db_execute(
        """
        INSERT INTO tenants
        (id, customer_name, tenant_name, tenant_guid, status, owner_primary, owner_backup, tags_csv, risk_profile, notes, is_active, created_at, updated_at)
        VALUES (?, ?, ?, ?, 'active', NULL, NULL, NULL, 'standard', ?, 1, ?, ?)
        """,
        (tenant_id, "Lokale Demo Klant", "Lokale Tenant", None, "Automatisch aangemaakt na verwijderen laatste tenant.", ts, ts),
    )


def delete_tenant(tenant_id: str, mode: str = "soft") -> Dict[str, Any]:
    tenant = db_fetchone("SELECT * FROM tenants WHERE id=?", (tenant_id,))
    if not tenant:
        raise ValueError("Tenant niet gevonden")
    mode = (mode or "soft").strip().lower()
    if mode not in {"soft", "hard"}:
        mode = "soft"

    if mode == "soft":
        db_execute(
            "UPDATE tenants SET is_active=0, status='offboarded', updated_at=? WHERE id=?",
            (now_iso(), tenant_id),
        )
        ensure_demo_tenant_if_empty()
        return {"id": tenant_id, "deleted": True, "mode": "soft"}

    # Hard delete: remove runs, artifacts and actions for this tenant.
    runs = db_fetchall("SELECT id FROM assessment_runs WHERE tenant_id=?", (tenant_id,))
    for r in runs:
        run_id = r["id"]
        db_execute("DELETE FROM finding_actions WHERE run_id=?", (run_id,))
        run_dir = RUNS_DIR / run_id
        if run_dir.exists():
            shutil.rmtree(run_dir, ignore_errors=True)
    db_execute("DELETE FROM finding_actions WHERE tenant_id=?", (tenant_id,))
    db_execute("DELETE FROM assessment_runs WHERE tenant_id=?", (tenant_id,))
    db_execute("DELETE FROM tenants WHERE id=?", (tenant_id,))
    ensure_demo_tenant_if_empty()
    return {"id": tenant_id, "deleted": True, "mode": "hard", "removed_runs": len(runs)}


def list_reports(
    tenant_id: Optional[str] = None,
    status: Optional[str] = None,
    date_from: Optional[str] = None,
    date_to: Optional[str] = None,
    q: Optional[str] = None,
    archived: str = "exclude",
    limit: int = 300,
) -> List[Dict[str, Any]]:
    sql = """
    SELECT r.*, t.customer_name, t.tenant_name, t.tenant_guid, t.status AS tenant_status
    FROM assessment_runs r
    JOIN tenants t ON t.id = r.tenant_id
    WHERE r.report_path IS NOT NULL
    """
    params: List[Any] = []
    if tenant_id:
        sql += " AND r.tenant_id=?"
        params.append(tenant_id)
    if status:
        sql += " AND r.status=?"
        params.append(status)
    if date_from:
        sql += " AND COALESCE(r.completed_at, r.started_at) >= ?"
        params.append(date_from)
    if date_to:
        sql += " AND COALESCE(r.completed_at, r.started_at) <= ?"
        params.append(date_to)
    if q:
        like = f"%{q}%"
        sql += " AND (t.tenant_name LIKE ? OR t.customer_name LIKE ? OR r.id LIKE ? OR COALESCE(r.report_filename,'') LIKE ?)"
        params.extend([like, like, like, like])
    if archived == "only":
        sql += " AND COALESCE(r.is_archived, 0)=1"
    elif archived == "include":
        pass
    else:
        sql += " AND COALESCE(r.is_archived, 0)=0"
    sql += " ORDER BY COALESCE(r.completed_at, r.started_at) DESC LIMIT ?"
    params.append(limit)

    rows = db_fetchall(sql, tuple(params))
    for r in rows:
        r["phases"] = [p for p in (r.get("phases_csv") or "").split(",") if p]
    return rows


def archive_run(run_id: str, reason: Optional[str] = None) -> Dict[str, Any]:
    row = db_fetchone("SELECT * FROM assessment_runs WHERE id=?", (run_id,))
    if not row:
        raise ValueError("Run niet gevonden")
    db_execute(
        "UPDATE assessment_runs SET is_archived=1, archived_at=?, archive_reason=? WHERE id=?",
        (now_iso(), (reason or "Handmatig gearchiveerd").strip(), run_id),
    )
    return get_run(run_id) or {}


def restore_run(run_id: str) -> Dict[str, Any]:
    row = db_fetchone("SELECT * FROM assessment_runs WHERE id=?", (run_id,))
    if not row:
        raise ValueError("Run niet gevonden")
    db_execute(
        "UPDATE assessment_runs SET is_archived=0, archived_at=NULL, archive_reason=NULL WHERE id=?",
        (run_id,),
    )
    return get_run(run_id) or {}


def apply_retention_policy(tenant_id: Optional[str], keep_latest: int, keep_days: int) -> Dict[str, Any]:
    keep_latest = max(0, int(keep_latest))
    keep_days = max(0, int(keep_days))
    rows = list_reports(tenant_id=tenant_id, archived="exclude", limit=5000)
    now_ts = datetime.now(timezone.utc)
    threshold_sec = keep_days * 86400
    to_archive: List[str] = []

    by_tenant: Dict[str, List[Dict[str, Any]]] = {}
    for r in rows:
        by_tenant.setdefault(r["tenant_id"], []).append(r)

    for _, tenant_rows in by_tenant.items():
        for idx, r in enumerate(tenant_rows):
            ts_raw = r.get("completed_at") or r.get("started_at")
            age_match = False
            if ts_raw:
                try:
                    ts = datetime.fromisoformat(ts_raw)
                    age_sec = (now_ts - ts).total_seconds()
                    age_match = threshold_sec > 0 and age_sec >= threshold_sec
                except Exception:
                    age_match = False
            index_match = keep_latest > 0 and idx >= keep_latest
            if (keep_latest == 0 or index_match) or age_match:
                to_archive.append(r["id"])

    archived_count = 0
    for run_id in sorted(set(to_archive)):
        archive_run(run_id, reason=f"Retention policy: keep_latest={keep_latest}, keep_days={keep_days}")
        archived_count += 1

    return {
        "scanned": len(rows),
        "archived": archived_count,
        "keep_latest": keep_latest,
        "keep_days": keep_days,
        "tenant_id": tenant_id,
    }


def reports_csv(rows: List[Dict[str, Any]]) -> str:
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(
        [
            "run_id",
            "tenant_name",
            "customer_name",
            "tenant_guid",
            "tenant_status",
            "run_status",
            "run_mode",
            "started_at",
            "completed_at",
            "score_overall",
            "critical_count",
            "warning_count",
            "info_count",
            "phases",
            "report_path",
            "report_filename",
        ]
    )
    for r in rows:
        writer.writerow(
            [
                r.get("id"),
                r.get("tenant_name"),
                r.get("customer_name"),
                r.get("tenant_guid"),
                r.get("tenant_status"),
                r.get("status"),
                r.get("run_mode"),
                r.get("started_at"),
                r.get("completed_at"),
                r.get("score_overall"),
                r.get("critical_count"),
                r.get("warning_count"),
                r.get("info_count"),
                ",".join(r.get("phases") or []),
                r.get("report_path"),
                r.get("report_filename"),
            ]
        )
    return output.getvalue()


def run_diff_for_tenant(tenant_id: str, from_run_id: Optional[str], to_run_id: Optional[str]) -> Dict[str, Any]:
    if from_run_id and to_run_id:
        older = get_run(from_run_id)
        newer = get_run(to_run_id)
        if not older or not newer:
            raise ValueError("Run(s) niet gevonden")
    else:
        recent = list_reports(tenant_id=tenant_id, limit=2)
        if len(recent) < 2:
            return {"hasDiff": False}
        newer, older = recent[0], recent[1]

    if older.get("tenant_id") != tenant_id or newer.get("tenant_id") != tenant_id:
        raise ValueError("Runs horen niet bij deze tenant")

    def n(v: Any) -> int:
        return int(v or 0)

    delta_score = n(newer.get("score_overall")) - n(older.get("score_overall"))
    delta_critical = n(newer.get("critical_count")) - n(older.get("critical_count"))
    delta_warning = n(newer.get("warning_count")) - n(older.get("warning_count"))
    delta_info = n(newer.get("info_count")) - n(older.get("info_count"))

    trend = "stable"
    if delta_score > 0 or delta_critical < 0:
        trend = "improved"
    elif delta_score < 0 or delta_critical > 0:
        trend = "worsened"

    return {
        "hasDiff": True,
        "trend": trend,
        "from": {
            "run_id": older.get("id"),
            "completed_at": older.get("completed_at") or older.get("started_at"),
            "score_overall": older.get("score_overall"),
            "critical_count": older.get("critical_count"),
            "warning_count": older.get("warning_count"),
            "info_count": older.get("info_count"),
        },
        "to": {
            "run_id": newer.get("id"),
            "completed_at": newer.get("completed_at") or newer.get("started_at"),
            "score_overall": newer.get("score_overall"),
            "critical_count": newer.get("critical_count"),
            "warning_count": newer.get("warning_count"),
            "info_count": newer.get("info_count"),
        },
        "delta": {
            "score_overall": delta_score,
            "critical_count": delta_critical,
            "warning_count": delta_warning,
            "info_count": delta_info,
        },
    }


def list_actions(tenant_id: str, status: Optional[str] = None) -> List[Dict[str, Any]]:
    sql = "SELECT * FROM finding_actions WHERE tenant_id=?"
    params: List[Any] = [tenant_id]
    if status and status != "all":
        sql += " AND status=?"
        params.append(status)
    sql += " ORDER BY CASE status WHEN 'open' THEN 0 WHEN 'in_progress' THEN 1 WHEN 'done' THEN 2 ELSE 3 END, due_date IS NULL, due_date, updated_at DESC"
    return db_fetchall(sql, tuple(params))


def create_action(payload: Dict[str, Any]) -> Dict[str, Any]:
    tenant_id = (payload.get("tenant_id") or "").strip()
    if not tenant_id:
        raise ValueError("tenant_id is verplicht")
    if not db_fetchone("SELECT id FROM tenants WHERE id=?", (tenant_id,)):
        raise ValueError("Tenant niet gevonden")

    action_id = str(uuid.uuid4())
    ts = now_iso()
    status = (payload.get("status") or "open").strip()
    if status not in {"open", "in_progress", "done", "accepted"}:
        status = "open"
    severity = (payload.get("severity") or "warning").strip()
    if severity not in {"critical", "warning", "info"}:
        severity = "warning"

    db_execute(
        """
        INSERT INTO finding_actions
        (id, tenant_id, run_id, finding_key, title, severity, owner, status, due_date, notes, evidence, created_at, updated_at, closed_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            action_id,
            tenant_id,
            (payload.get("run_id") or "").strip() or None,
            (payload.get("finding_key") or "").strip() or f"manual-{action_id[:8]}",
            (payload.get("title") or "").strip() or "Nieuwe actie",
            severity,
            (payload.get("owner") or "").strip() or None,
            status,
            (payload.get("due_date") or "").strip() or None,
            (payload.get("notes") or "").strip() or None,
            (payload.get("evidence") or "").strip() or None,
            ts,
            ts,
            ts if status == "done" else None,
        ),
    )
    return db_fetchone("SELECT * FROM finding_actions WHERE id=?", (action_id,)) or {}


def update_action(action_id: str, payload: Dict[str, Any]) -> Dict[str, Any]:
    row = db_fetchone("SELECT * FROM finding_actions WHERE id=?", (action_id,))
    if not row:
        raise ValueError("Actie niet gevonden")
    allowed = {"owner", "status", "due_date", "notes", "evidence", "title", "severity"}
    fields: Dict[str, Any] = {}
    for k, v in payload.items():
        if k not in allowed:
            continue
        fields[k] = v.strip() if isinstance(v, str) else v
    if "status" in fields and fields["status"] not in {"open", "in_progress", "done", "accepted"}:
        raise ValueError("Ongeldige status")
    if "severity" in fields and fields["severity"] not in {"critical", "warning", "info"}:
        raise ValueError("Ongeldige severity")
    if fields.get("status") == "done":
        fields["closed_at"] = now_iso()
    elif "status" in fields:
        fields["closed_at"] = None
    fields["updated_at"] = now_iso()

    keys = list(fields.keys())
    if keys:
        sql = "UPDATE finding_actions SET " + ", ".join([f"{k}=?" for k in keys]) + " WHERE id=?"
        vals = [fields[k] for k in keys] + [action_id]
        db_execute(sql, tuple(vals))
    return db_fetchone("SELECT * FROM finding_actions WHERE id=?", (action_id,)) or {}


def list_runs(tenant_id: Optional[str] = None, limit: int = 100) -> List[Dict[str, Any]]:
    if tenant_id:
        sql = """
        SELECT r.*, t.customer_name, t.tenant_name
        FROM assessment_runs r
        JOIN tenants t ON t.id = r.tenant_id
        WHERE tenant_id=?
        ORDER BY started_at DESC LIMIT ?
        """
        rows = db_fetchall(sql, (tenant_id, limit))
    else:
        sql = """
        SELECT r.*, t.customer_name, t.tenant_name
        FROM assessment_runs r
        JOIN tenants t ON t.id = r.tenant_id
        ORDER BY started_at DESC LIMIT ?
        """
        rows = db_fetchall(sql, (limit,))
    for r in rows:
        r["phases"] = [p for p in (r.get("phases_csv") or "").split(",") if p]
    return rows


def get_run(run_id: str) -> Optional[Dict[str, Any]]:
    row = db_fetchone(
        """
        SELECT r.*, t.customer_name, t.tenant_name
        FROM assessment_runs r
        JOIN tenants t ON t.id = r.tenant_id
        WHERE r.id=?
        """,
        (run_id,),
    )
    if row:
        row["phases"] = [p for p in (row.get("phases_csv") or "").split(",") if p]
    return row


def tenant_overview(tenant_id: str) -> Dict[str, Any]:
    tenant = db_fetchone("SELECT * FROM tenants WHERE id=?", (tenant_id,))
    if not tenant:
        return {"hasData": False}
    latest = db_fetchone("SELECT * FROM assessment_runs WHERE tenant_id=? ORDER BY started_at DESC LIMIT 1", (tenant_id,))
    if not latest:
        return {"hasData": False, "tenantName": tenant["tenant_name"], "tenantId": tenant.get("tenant_guid") or tenant["id"]}
    run_dir = RUNS_DIR / latest["id"]
    stats = parse_run_stats(run_dir)
    return {
        "hasData": True,
        "tenantName": stats.get("tenantName") or tenant["tenant_name"],
        "tenantId": stats.get("tenantId") or tenant.get("tenant_guid") or tenant["id"],
        "reportDate": stats.get("reportDate") or latest.get("completed_at") or latest.get("started_at"),
        "reportId": stats.get("reportId") or latest["id"],
        "criticalIssues": latest.get("critical_count") or stats.get("criticalIssues") or 0,
        "warnings": latest.get("warning_count") or stats.get("warnings") or 0,
        "infoItems": latest.get("info_count") or stats.get("infoItems") or 0,
        "mfaCoverage": stats.get("mfaCoverage"),
        "usersWithoutMFA": stats.get("usersWithoutMFA"),
        "caPolicies": stats.get("caPolicies"),
        "secureScorePercentage": stats.get("secureScorePercentage"),
        "scoreOverall": latest.get("score_overall"),
        "reportPath": latest.get("report_path"),
        "latestRunStatus": latest.get("status"),
        "secureScoreCurrent": stats.get("secureScorePercentage"),
        "secureScoreMax": 100 if stats.get("secureScorePercentage") is not None else None,
    }


def create_tenant(payload: Dict[str, Any]) -> Dict[str, Any]:
    tenant_id = str(uuid.uuid4())
    ts = now_iso()
    customer_name = (payload.get("customer_name") or payload.get("customerName") or "").strip() or "Lokale Klant"
    tenant_name = (payload.get("tenant_name") or payload.get("tenantName") or "").strip() or customer_name
    tenant_guid = (payload.get("tenant_guid") or payload.get("tenantGuid") or "").strip() or None
    status = (payload.get("status") or "active").strip()
    if status not in {"active", "onboarding", "paused", "offboarded"}:
        status = "active"
    owner_primary = (payload.get("owner_primary") or "").strip() or None
    owner_backup = (payload.get("owner_backup") or "").strip() or None
    tags_csv = (payload.get("tags_csv") or "").strip() or None
    risk_profile = (payload.get("risk_profile") or "standard").strip()
    if risk_profile not in {"low", "standard", "high", "critical"}:
        risk_profile = "standard"
    notes = (payload.get("notes") or "").strip() or None
    db_execute(
        "INSERT INTO tenants (id, customer_name, tenant_name, tenant_guid, status, owner_primary, owner_backup, tags_csv, risk_profile, notes, is_active, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?, ?)",
        (tenant_id, customer_name, tenant_name, tenant_guid, status, owner_primary, owner_backup, tags_csv, risk_profile, notes, ts, ts),
    )
    return db_fetchone("SELECT * FROM tenants WHERE id=?", (tenant_id,)) or {}


def create_run(payload: Dict[str, Any]) -> Dict[str, Any]:
    tenant_id = payload.get("tenant_id") or payload.get("tenantId")
    if not tenant_id:
        raise ValueError("tenant_id is verplicht")
    if not db_fetchone("SELECT id FROM tenants WHERE id=?", (tenant_id,)):
        raise ValueError("Tenant niet gevonden")
    # Voorkom concurrent runs voor dezelfde tenant
    active = db_fetchone(
        "SELECT id FROM assessment_runs WHERE tenant_id=? AND status IN ('queued','running') LIMIT 1",
        (tenant_id,),
    )
    if active:
        raise ValueError(f"Er loopt al een actief assessment voor deze tenant (run: {active['id'][:8]}…)")

    scan_type = str(payload.get("scan_type") or "full")

    if scan_type == "entrafalcon":
        # EntraFalcon heeft geen M365-fasen; altijd script-mode
        phases = ["entrafalcon"]
        run_mode = "script"
    else:
        phases = payload.get("phases") or [f"phase{i}" for i in range(1, 7)]
        if not isinstance(phases, list):
            raise ValueError("phases moet een array zijn")
        phases = [str(p) for p in phases if re.fullmatch(r"phase[1-6]", str(p))]
        if not phases:
            raise ValueError("Geen geldige phases opgegeven")
        cfg = load_config()
        run_mode = str(payload.get("run_mode") or payload.get("runMode") or cfg.get("default_run_mode") or "demo")
        if run_mode not in {"demo", "script"}:
            run_mode = "demo"

    run_id = str(uuid.uuid4())
    db_execute(
        """
        INSERT INTO assessment_runs (id, tenant_id, status, run_mode, scan_type, phases_csv, started_by, started_at)
        VALUES (?, ?, 'queued', ?, ?, ?, ?, ?)
        """,
        (
            run_id,
            tenant_id,
            run_mode,
            scan_type,
            ",".join(phases),
            str(payload.get("started_by") or "local-user"),
            now_iso(),
        ),
    )
    append_run_log(run_id, "Run queued.")
    RUN_MANAGER.start(run_id, phases, run_mode, scan_type)
    return get_run(run_id) or {"id": run_id}


def delete_run(run_id: str) -> Dict[str, Any]:
    run = db_fetchone("SELECT * FROM assessment_runs WHERE id=?", (run_id,))
    if not run:
        raise ValueError("Run niet gevonden")
    db_execute("DELETE FROM finding_actions WHERE run_id=?", (run_id,))
    db_execute("DELETE FROM assessment_runs WHERE id=?", (run_id,))
    run_dir = RUNS_DIR / run_id
    if run_dir.exists():
        shutil.rmtree(run_dir, ignore_errors=True)
    return {"id": run_id, "deleted": True}


class Handler(BaseHTTPRequestHandler):
    server_version = "M365BaselineLocal/0.1"

    def log_message(self, fmt: str, *args: Any) -> None:
        print(f"[{self.log_date_time_string()}] {fmt % args}")

    def _json(self, status: int, payload: Any) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("X-Content-Type-Options", "nosniff")
        self.end_headers()
        self.wfile.write(body)

    def _read_json(self) -> Dict[str, Any]:
        size = int(self.headers.get("Content-Length", "0") or "0")
        data = self.rfile.read(size) if size else b"{}"
        return json.loads(data.decode("utf-8")) if data else {}

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        path = parsed.path
        qs = parse_qs(parsed.query)
        try:
            # ── Sessie / autorisatie-check ──
            if path.startswith("/api/") and path not in _OPEN_API_PATHS:
                _sess = _check_api_access(self, path)
                if _sess is None:
                    return  # 401/403 al verzonden

            # Auth endpoints
            if path == "/api/auth/verify":
                sess = _get_session_from_request(self)
                if sess:
                    return self._json(200, {"ok": True, "role": sess["role"], "email": sess["email"], "display_name": sess["display_name"]})
                return self._json(401, {"ok": False, "error": "Niet ingelogd"})
            if path == "/api/health":
                return self._json(200, {"ok": True, "projectDir": str(PLATFORM_DIR)})
            if path == "/api/auth/msal-config":
                # Publiek endpoint — geeft alleen de niet-geheime MSAL-instellingen terug
                cfg = load_config()
                return self._json(200, {
                    "auth_client_id": cfg.get("auth_client_id", ""),
                    "auth_tenant_id": cfg.get("auth_tenant_id", ""),
                })
            if path == "/api/auth/csrf-token":
                # Geeft een sessie-gebonden CSRF-token terug (voor toekomstige header-validatie)
                sess = _get_session_from_request(self)
                token = sess["token"] if sess else secrets.token_urlsafe(16)
                return self._json(200, {"csrf_token": token})
            if path == "/api/config":
                return self._json(200, load_config())
            # ── Gebruikersbeheer ──
            if path == "/api/users":
                return self._json(200, {"items": list_users()})
            if re.fullmatch(r"/api/users/[^/]+", path):
                u = get_user(path.split("/")[3])
                if not u:
                    return self._json(404, {"error": "Gebruiker niet gevonden"})
                return self._json(200, u)
            if path == "/api/tenants":
                return self._json(200, {"items": list_tenants()})
            if re.fullmatch(r"/api/tenants/[^/]+", path):
                tenant = db_fetchone("SELECT * FROM tenants WHERE id=?", (path.split("/")[3],))
                if not tenant:
                    return self._json(404, {"error": "Tenant niet gevonden"})
                return self._json(200, tenant)
            if re.fullmatch(r"/api/tenants/[^/]+/overview", path):
                return self._json(200, tenant_overview(path.split("/")[3]))
            if re.fullmatch(r"/api/assessment/[^/]+/nav", path):
                tenant_id = path.split("/")[3]
                return self._json(200, assessment_ui_nav(tenant_id))
            if re.fullmatch(r"/api/assessment/[^/]+/section/[^/]+", path):
                tenant_id = path.split("/")[3]
                section_key = path.split("/")[5]
                return self._json(200, assessment_ui_section(tenant_id, section_key))
            if re.fullmatch(r"/api/tenants/[^/]+/runs", path):
                return self._json(200, {"items": list_runs(path.split("/")[3], 200)})
            if re.fullmatch(r"/api/tenants/[^/]+/runs/diff", path):
                tenant_id = path.split("/")[3]
                from_run_id = qs.get("from_run_id", [None])[0]
                to_run_id = qs.get("to_run_id", [None])[0]
                return self._json(200, run_diff_for_tenant(tenant_id, from_run_id, to_run_id))
            if re.fullmatch(r"/api/tenants/[^/]+/actions", path):
                tenant_id = path.split("/")[3]
                status = qs.get("status", [None])[0]
                return self._json(200, {"items": list_actions(tenant_id, status)})
            if path == "/api/runs":
                tenant_id = qs.get("tenant_id", [None])[0]
                limit = int(qs.get("limit", ["100"])[0])
                return self._json(200, {"items": list_runs(tenant_id, limit)})
            if re.fullmatch(r"/api/runs/[^/]+", path):
                run = get_run(path.split("/")[3])
                if not run:
                    return self._json(404, {"error": "Run niet gevonden"})
                return self._json(200, run)
            if re.fullmatch(r"/api/runs/[^/]+/logs", path):
                run_id = path.split("/")[3]
                log_path = RUNS_DIR / run_id / "run.log"
                if not log_path.exists():
                    return self._json(200, {"text": "", "lines": []})
                text = log_path.read_text(encoding="utf-8", errors="ignore")
                lines = text.splitlines()[-400:]
                return self._json(200, {"text": "\n".join(lines), "lines": lines})
            if re.fullmatch(r"/api/runs/[^/]+/files", path):
                run_id = path.split("/")[3]
                run_dir = RUNS_DIR / run_id
                if not run_dir.exists():
                    return self._json(404, {"error": "Run niet gevonden"})
                files = list_run_html_files(run_dir)
                return self._json(200, {"items": files, "run_id": run_id})
            if path == "/api/reports/list":
                runs = [r for r in list_runs(None, 200) if r.get("report_path")]
                return self._json(
                    200,
                    [
                        {
                            "id": r["id"],
                            "tenantId": r["tenant_id"],
                            "tenantName": r["tenant_name"],
                            "path": r["report_path"],
                            "createdDisplay": r.get("completed_at") or r.get("started_at"),
                            "sizeDisplay": "-",
                        }
                        for r in runs
                    ],
                )
            if path == "/api/reports":
                tenant_id = qs.get("tenant_id", [None])[0]
                status = qs.get("status", [None])[0]
                date_from = qs.get("from", [None])[0]
                date_to = qs.get("to", [None])[0]
                q = qs.get("q", [None])[0]
                archived = qs.get("archived", ["exclude"])[0]
                limit = int(qs.get("limit", ["300"])[0])
                return self._json(200, {"items": list_reports(tenant_id, status, date_from, date_to, q, archived, limit)})
            if path == "/api/reports/export.csv":
                tenant_id = qs.get("tenant_id", [None])[0]
                status = qs.get("status", [None])[0]
                date_from = qs.get("from", [None])[0]
                date_to = qs.get("to", [None])[0]
                q = qs.get("q", [None])[0]
                archived = qs.get("archived", ["exclude"])[0]
                rows = list_reports(tenant_id, status, date_from, date_to, q, archived, 1000)
                csv_text = reports_csv(rows)
                body = csv_text.encode("utf-8")
                self.send_response(HTTPStatus.OK)
                self.send_header("Content-Type", "text/csv; charset=utf-8")
                self.send_header("Content-Disposition", "attachment; filename=reports-export.csv")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
                return
            if path == "/api/reports/stats":
                tenant_id = qs.get("tenant_id", [None])[0]
                if tenant_id:
                    return self._json(200, tenant_overview(tenant_id))
                tenants = list_tenants()
                if not tenants:
                    return self._json(200, {"hasData": False})
                return self._json(200, tenant_overview(tenants[0]["id"]))
            # KB routes (GET)
            if re.fullmatch(r"/api/kb/[^/]+/asset-types", path):
                return self._json(200, kb_list_asset_types(_kb_tid(path)))
            if re.fullmatch(r"/api/kb/[^/]+/meta", path):
                return self._json(200, kb_get_meta(_kb_tid(path)))
            if re.fullmatch(r"/api/kb/[^/]+/assets", path):
                return self._json(200, kb_list_assets(_kb_tid(path), qs.get("type", [None])[0]))
            if re.fullmatch(r"/api/kb/[^/]+/assets/\d+", path):
                tid = _kb_tid(path); iid = _kb_iid(path)
                rows = kb_list_assets(tid)
                item = next((r for r in rows if r["id"] == iid), None)
                if not item:
                    return self._json(404, {"error": "Not found"})
                return self._json(200, item)
            if re.fullmatch(r"/api/kb/[^/]+/vlans", path):
                return self._json(200, kb_list_vlans(_kb_tid(path)))
            if re.fullmatch(r"/api/kb/[^/]+/pages", path):
                return self._json(200, kb_list_pages(_kb_tid(path)))
            if re.fullmatch(r"/api/kb/[^/]+/pages/\d+", path):
                row = kb_get_page(_kb_tid(path), _kb_iid(path))
                if not row:
                    return self._json(404, {"error": "Not found"})
                return self._json(200, row)
            if re.fullmatch(r"/api/kb/[^/]+/contacts", path):
                return self._json(200, kb_list_contacts(_kb_tid(path)))
            if re.fullmatch(r"/api/kb/[^/]+/passwords", path):
                return self._json(200, kb_list_passwords(_kb_tid(path)))
            if re.fullmatch(r"/api/kb/[^/]+/software", path):
                return self._json(200, kb_list_software(_kb_tid(path)))
            if re.fullmatch(r"/api/kb/[^/]+/domains", path):
                return self._json(200, kb_list_domains(_kb_tid(path)))
            if re.fullmatch(r"/api/kb/[^/]+/m365", path):
                return self._json(200, kb_get_m365_profile(_kb_tid(path)))
            if re.fullmatch(r"/api/kb/[^/]+/changelog", path):
                return self._json(200, kb_list_changelog(_kb_tid(path)))
            if path.startswith("/reports/Templates/"):
                fp = (PLATFORM_DIR / "assessment" / "Templates" / path[len("/reports/Templates/") :]).resolve()
                if not str(fp).startswith(str((PLATFORM_DIR / "assessment" / "Templates").resolve())) or not fp.exists() or not fp.is_file():
                    self.send_error(404)
                    return
                return self._serve_file(fp)
            if path.startswith("/reports/"):
                return self._serve_report(path)
            # /site/ → serveert de hoofdwebsite (PLATFORM_DIR root)
            if path.startswith("/site/") or path in ("/site", "/site/index.html"):
                return self._serve_site(path)
            return self._serve_web(path)
        except Exception as exc:
            logger.error("500 in GET %s: %s", path, traceback.format_exc())
            return self._json(500, {"error": "Interne serverfout."})

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        path = parsed.path
        # CSRF-check voor alle state-muterende endpoints buiten initiële auth
        if path not in ("/api/auth/login", "/api/auth/microsoft", "/api/auth/logout"):
            if not _check_csrf(self):
                return self._json(403, {"error": "CSRF validatie mislukt."})
        # ── Sessie / autorisatie-check ──
        if path.startswith("/api/") and path not in _OPEN_API_PATHS:
            _sess = _check_api_access(self, path)
            if _sess is None:
                return  # 401/403 al verzonden
        try:
            # ── Auth endpoints ──
            if path == "/api/auth/login":
                body = self._read_json()
                email = (body.get("email") or "").strip().lower()
                password = body.get("password") or ""
                ip = self.client_address[0]
                # Rate limiting
                if not _check_rate_limit(ip, max_attempts=10, window_secs=60):
                    db_audit("", ip, "login_rate_limited", detail=email)
                    return self._json(429, {"ok": False, "error": "Te veel inlogpogingen. Probeer het later opnieuw."})
                # Input validatie
                if not email or not re.match(r'^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$', email):
                    return self._json(400, {"ok": False, "error": "Ongeldig e-mailadres."})
                if len(password) < 1 or len(password) > 512:
                    return self._json(400, {"ok": False, "error": "Ongeldig wachtwoord."})
                user = db_fetchone("SELECT * FROM users WHERE lower(email)=? AND is_active=1", (email,))
                if not user or not _verify_pw(password, user["password_hash"], user["salt"]):
                    db_audit(email, ip, "login_failed")
                    return self._json(401, {"ok": False, "error": "Onjuist e-mailadres of wachtwoord."})
                token = _create_session(user["id"], user["role"], user["email"], user["display_name"])
                db_audit(email, ip, "login_success", "user", user["id"])
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Set-Cookie", f"denjoy_session={token}; Path=/; HttpOnly; SameSite=Strict; Max-Age=3600")
                body_bytes = json.dumps({"ok": True, "token": token, "role": user["role"], "email": user["email"], "display_name": user["display_name"]}).encode()
                self.send_header("Content-Length", str(len(body_bytes)))
                self.end_headers()
                self.wfile.write(body_bytes)
                return
            if path == "/api/auth/logout":
                sess = _get_session_from_request(self)
                if sess:
                    db_execute("DELETE FROM sessions WHERE token=?", (sess["token"],))
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Set-Cookie", "denjoy_session=; Path=/; Max-Age=0")
                body_bytes = b'{"ok":true}'
                self.send_header("Content-Length", str(len(body_bytes)))
                self.end_headers()
                self.wfile.write(body_bytes)
                return
            if path == "/api/auth/microsoft":
                # Klant ingelogd via Microsoft — registreer/update en geef sessie-token
                body = self._read_json()
                email = (body.get("email") or "").strip().lower()
                name  = body.get("name") or email
                tenant_id_ms = body.get("tenant_id") or ""
                if not email:
                    return self._json(400, {"ok": False, "error": "email vereist"})
                user = db_fetchone("SELECT * FROM users WHERE lower(email)=?", (email,))
                if not user:
                    # Automatisch aanmaken als klant (read-only)
                    uid = str(uuid.uuid4())
                    pw_hash, salt = _hash_pw(secrets.token_hex(16))  # random onbruikbaar wachtwoord
                    db_execute(
                        "INSERT INTO users (id,email,password_hash,salt,role,display_name,is_active,created_at) VALUES (?,?,?,?,?,?,1,?)",
                        (uid, email, pw_hash, salt, "klant", name, now_iso())
                    )
                    user = db_fetchone("SELECT * FROM users WHERE id=?", (uid,))
                token = _create_session(user["id"], user["role"], user["email"], user["display_name"] or name)
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Set-Cookie", f"denjoy_session={token}; Path=/; HttpOnly; SameSite=Strict; Max-Age=3600")
                body_bytes = json.dumps({"ok": True, "token": token, "role": user["role"], "email": user["email"], "display_name": user["display_name"] or name}).encode()
                self.send_header("Content-Length", str(len(body_bytes)))
                self.end_headers()
                self.wfile.write(body_bytes)
                return
            if path == "/api/tenants":
                return self._json(201, create_tenant(self._read_json()))
            if re.fullmatch(r"/api/tenants/[^/]+/delete", path):
                tenant_id = path.split("/")[3]
                payload = self._read_json()
                mode = payload.get("mode") or parse_qs(parsed.query).get("mode", ["soft"])[0]
                return self._json(200, delete_tenant(tenant_id, mode))
            if path == "/api/actions":
                return self._json(201, create_action(self._read_json()))
            if re.fullmatch(r"/api/runs/[^/]+/stop", path):
                run_id = path.split("/")[3]
                ok = RUN_MANAGER.stop(run_id)
                if ok:
                    append_run_log(run_id, "⏹ Stop-verzoek ontvangen via API.")
                    return self._json(200, {"ok": True})
                return self._json(404, {"error": "Geen actief proces gevonden voor deze run"})
            if re.fullmatch(r"/api/runs/[^/]+/delete", path):
                run_id = path.split("/")[3]
                return self._json(200, delete_run(run_id))
            if re.fullmatch(r"/api/reports/[^/]+/archive", path):
                run_id = path.split("/")[3]
                payload = self._read_json()
                return self._json(200, archive_run(run_id, payload.get("reason")))
            if re.fullmatch(r"/api/reports/[^/]+/restore", path):
                run_id = path.split("/")[3]
                return self._json(200, restore_run(run_id))
            if path == "/api/reports/retention/apply":
                payload = self._read_json()
                tenant_id = payload.get("tenant_id") or None
                keep_latest = payload.get("keep_latest", 10)
                keep_days = payload.get("keep_days", 90)
                return self._json(200, apply_retention_policy(tenant_id, keep_latest, keep_days))
            if path == "/api/runs":
                return self._json(201, create_run(self._read_json()))
            # ── Gebruikersbeheer ──
            if path == "/api/users":
                return self._json(201, create_user_account(self._read_json()))
            if re.fullmatch(r"/api/users/[^/]+/reset-password", path):
                uid = path.split("/")[3]
                pwd = (self._read_json().get("password") or "").strip()
                return self._json(200, update_user_account(uid, {"password": pwd}, _sess.get("email", "")))
            if path == "/api/config":
                payload = self._read_json()
                cfg = load_config()
                for k in ("default_run_mode", "script_path",
                          "auth_tenant_id", "auth_client_id",
                          "auth_cert_thumbprint", "auth_client_secret",
                          "assessment_ui_v1",
                          "entrafalcon_script_path",
                          "entrafalcon_include_ms_apps", "entrafalcon_csv"):
                    if k in payload:
                        cfg[k] = payload[k]
                save_config(cfg)
                # Geef config terug zonder geheimen
                safe = {k: v for k, v in cfg.items() if k not in ("auth_client_secret",)}
                return self._json(200, safe)
            # KB routes (POST)
            if re.fullmatch(r"/api/kb/[^/]+/asset-types", path):
                return self._json(201, kb_create_asset_type(_kb_tid(path), self._read_json()))
            if re.fullmatch(r"/api/kb/[^/]+/assets", path):
                return self._json(201, kb_create_asset(_kb_tid(path), self._read_json()))
            if re.fullmatch(r"/api/kb/[^/]+/vlans", path):
                return self._json(201, kb_create_vlan(_kb_tid(path), self._read_json()))
            if re.fullmatch(r"/api/kb/[^/]+/pages", path):
                return self._json(201, kb_create_page(_kb_tid(path), self._read_json()))
            if re.fullmatch(r"/api/kb/[^/]+/contacts", path):
                return self._json(201, kb_create_contact(_kb_tid(path), self._read_json()))
            if re.fullmatch(r"/api/kb/[^/]+/passwords", path):
                return self._json(201, kb_create_password(_kb_tid(path), self._read_json()))
            if re.fullmatch(r"/api/kb/[^/]+/software", path):
                return self._json(201, kb_create_software(_kb_tid(path), self._read_json()))
            if re.fullmatch(r"/api/kb/[^/]+/domains", path):
                return self._json(201, kb_create_domain(_kb_tid(path), self._read_json()))
            if re.fullmatch(r"/api/kb/[^/]+/changelog", path):
                return self._json(201, kb_create_changelog(_kb_tid(path), self._read_json()))
            return self._json(404, {"error": "Niet gevonden"})
        except ValueError as exc:
            return self._json(400, {"error": str(exc)})
        except Exception as exc:
            logger.error("500 in POST %s: %s", path, traceback.format_exc())
            return self._json(500, {"error": "Interne serverfout."})

    def do_DELETE(self) -> None:
        parsed = urlparse(self.path)
        path = parsed.path
        if not _check_csrf(self):
            return self._json(403, {"error": "CSRF validatie mislukt."})
        _sess = _check_api_access(self, path)
        if _sess is None:
            return
        try:
            if re.fullmatch(r"/api/tenants/[^/]+", path):
                tenant_id = path.split("/")[3]
                mode = parse_qs(parsed.query).get("mode", ["soft"])[0]
                return self._json(200, delete_tenant(tenant_id, mode))
            # ── Gebruikersbeheer ──
            if re.fullmatch(r"/api/users/[^/]+", path):
                uid = path.split("/")[3]
                return self._json(200, delete_user_account(uid, _sess.get("email", "")))
            if re.fullmatch(r"/api/runs/[^/]+", path):
                run_id = path.split("/")[3]
                return self._json(200, delete_run(run_id))
            # KB routes (DELETE)
            if re.fullmatch(r"/api/kb/[^/]+/asset-types/\d+", path):
                kb_delete_asset_type(_kb_tid(path), _kb_iid(path))
                return self._json(200, {"ok": True})
            if re.fullmatch(r"/api/kb/[^/]+/assets/\d+", path):
                kb_delete_asset(_kb_tid(path), _kb_iid(path))
                return self._json(200, {"ok": True})
            if re.fullmatch(r"/api/kb/[^/]+/vlans/\d+", path):
                kb_delete_vlan(_kb_tid(path), _kb_iid(path))
                return self._json(200, {"ok": True})
            if re.fullmatch(r"/api/kb/[^/]+/pages/\d+", path):
                kb_delete_page(_kb_tid(path), _kb_iid(path))
                return self._json(200, {"ok": True})
            if re.fullmatch(r"/api/kb/[^/]+/contacts/\d+", path):
                kb_delete_contact(_kb_tid(path), _kb_iid(path))
                return self._json(200, {"ok": True})
            if re.fullmatch(r"/api/kb/[^/]+/passwords/\d+", path):
                kb_delete_password(_kb_tid(path), _kb_iid(path))
                return self._json(200, {"ok": True})
            if re.fullmatch(r"/api/kb/[^/]+/software/\d+", path):
                kb_delete_software(_kb_tid(path), _kb_iid(path))
                return self._json(200, {"ok": True})
            if re.fullmatch(r"/api/kb/[^/]+/domains/\d+", path):
                kb_delete_domain(_kb_tid(path), _kb_iid(path))
                return self._json(200, {"ok": True})
            if re.fullmatch(r"/api/kb/[^/]+/changelog/\d+", path):
                kb_delete_changelog(_kb_tid(path), _kb_iid(path))
                return self._json(200, {"ok": True})
            return self._json(404, {"error": "Niet gevonden"})
        except ValueError as exc:
            return self._json(400, {"error": str(exc)})
        except Exception as exc:
            logger.error("500 in DELETE %s: %s", path, traceback.format_exc())
            return self._json(500, {"error": "Interne serverfout."})

    def do_PATCH(self) -> None:
        parsed = urlparse(self.path)
        path = parsed.path
        if not _check_csrf(self):
            return self._json(403, {"error": "CSRF validatie mislukt."})
        _sess = _check_api_access(self, path)
        if _sess is None:
            return
        try:
            if re.fullmatch(r"/api/tenants/[^/]+", path):
                tenant_id = path.split("/")[3]
                return self._json(200, update_tenant(tenant_id, self._read_json()))
            # ── Gebruikersbeheer ──
            if re.fullmatch(r"/api/users/[^/]+", path):
                uid = path.split("/")[3]
                return self._json(200, update_user_account(uid, self._read_json(), _sess.get("email", "")))
            if re.fullmatch(r"/api/actions/[^/]+", path):
                action_id = path.split("/")[3]
                return self._json(200, update_action(action_id, self._read_json()))
            return self._json(404, {"error": "Niet gevonden"})
        except ValueError as exc:
            return self._json(400, {"error": str(exc)})
        except Exception as exc:
            logger.error("500 in PATCH %s: %s", path, traceback.format_exc())
            return self._json(500, {"error": "Interne serverfout."})

    def do_PUT(self) -> None:
        parsed = urlparse(self.path)
        path = parsed.path
        if not _check_csrf(self):
            return self._json(403, {"error": "CSRF validatie mislukt."})
        _sess = _check_api_access(self, path)
        if _sess is None:
            return
        try:
            # KB routes (PUT)
            if re.fullmatch(r"/api/kb/[^/]+/meta", path):
                return self._json(200, kb_put_meta(_kb_tid(path), self._read_json()))
            if re.fullmatch(r"/api/kb/[^/]+/assets/\d+", path):
                return self._json(200, kb_update_asset(_kb_tid(path), _kb_iid(path), self._read_json()))
            if re.fullmatch(r"/api/kb/[^/]+/vlans/\d+", path):
                return self._json(200, kb_update_vlan(_kb_tid(path), _kb_iid(path), self._read_json()))
            if re.fullmatch(r"/api/kb/[^/]+/pages/\d+", path):
                return self._json(200, kb_update_page(_kb_tid(path), _kb_iid(path), self._read_json()))
            if re.fullmatch(r"/api/kb/[^/]+/contacts/\d+", path):
                return self._json(200, kb_update_contact(_kb_tid(path), _kb_iid(path), self._read_json()))
            if re.fullmatch(r"/api/kb/[^/]+/passwords/\d+", path):
                return self._json(200, kb_update_password(_kb_tid(path), _kb_iid(path), self._read_json()))
            if re.fullmatch(r"/api/kb/[^/]+/software/\d+", path):
                return self._json(200, kb_update_software(_kb_tid(path), _kb_iid(path), self._read_json()))
            if re.fullmatch(r"/api/kb/[^/]+/domains/\d+", path):
                return self._json(200, kb_update_domain(_kb_tid(path), _kb_iid(path), self._read_json()))
            if re.fullmatch(r"/api/kb/[^/]+/m365", path):
                return self._json(200, kb_put_m365_profile(_kb_tid(path), self._read_json()))
            if re.fullmatch(r"/api/kb/[^/]+/changelog/\d+", path):
                return self._json(200, kb_update_changelog(_kb_tid(path), _kb_iid(path), self._read_json()))
            return self._json(404, {"error": "Niet gevonden"})
        except ValueError as exc:
            return self._json(400, {"error": str(exc)})
        except Exception as exc:
            logger.error("500 in PUT %s: %s", path, traceback.format_exc())
            return self._json(500, {"error": "Interne serverfout."})

    def _serve_web(self, path: str) -> None:
        # /portal/... → portalbestanden (WEB_DIR)
        if path.startswith("/portal/") or path in ("/portal", "/portal/"):
            if path in ("/portal", "/portal/"):
                fp = WEB_DIR / "index.html"
            else:
                rel = unquote(path[len("/portal/"):])
                if ".." in Path(rel).parts:
                    self.send_error(403)
                    return
                fp = WEB_DIR / rel
                if fp.is_dir():
                    fp = fp / "index.html"
        else:
            # / en overige statische bestanden → hoofdwebsite (PLATFORM_DIR)
            if path in ("", "/"):
                fp = PLATFORM_DIR / "index.html"
            else:
                rel = unquote(path.lstrip("/"))
                if ".." in Path(rel).parts:
                    self.send_error(403)
                    return
                fp = PLATFORM_DIR / rel
                if fp.is_dir():
                    fp = fp / "index.html"
        if not fp.exists() or not fp.is_file():
            self.send_error(404)
            return
        self._serve_file(fp)

    def _serve_site(self, path: str) -> None:
        """Serveert de hoofdwebsite-bestanden vanuit PLATFORM_DIR."""
        # Verwijder het /site/ prefix
        rel = unquote(path[len("/site/"):].lstrip("/")) if path.startswith("/site/") else "index.html"
        if not rel:
            rel = "index.html"
        # Veiligheidscheck — geen directory traversal
        if ".." in Path(rel).parts:
            self.send_error(403)
            return
        # Portal-bestanden via /site/portal/ → redirect naar de echte portal
        if rel.startswith("portal/"):
            portal_rel = rel[len("portal/"):]
            self.send_response(302)
            self.send_header("Location", f"/{portal_rel}")
            self.end_headers()
            return
        fp = (PLATFORM_DIR / rel).resolve()
        if not str(fp).startswith(str(PLATFORM_DIR.resolve())):
            self.send_error(403)
            return
        if fp.is_dir():
            fp = fp / "index.html"
        if not fp.exists() or not fp.is_file():
            self.send_error(404)
            return
        self._serve_file(fp)

    def _serve_report(self, path: str) -> None:
        rel = path[len("/reports/") :]
        parts = [p for p in rel.split("/") if p]
        if len(parts) < 2:
            self.send_error(404)
            return
        run_id = parts[0]
        fp = (RUNS_DIR / run_id).joinpath(*parts[1:])
        if not fp.exists() or not fp.is_file():
            self.send_error(404)
            return
        self._serve_file(fp)

    def _serve_file(self, fp: Path) -> None:
        data = fp.read_bytes()
        mime = {
            ".html": "text/html; charset=utf-8",
            ".css": "text/css; charset=utf-8",
            ".js": "application/javascript; charset=utf-8",
            ".json": "application/json; charset=utf-8",
            ".png": "image/png",
            ".jpg": "image/jpeg",
            ".jpeg": "image/jpeg",
            ".svg": "image/svg+xml",
        }.get(fp.suffix.lower(), "application/octet-stream")
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", mime)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("X-Content-Type-Options", "nosniff")
        if mime.startswith("text/html"):
            self.send_header("Content-Security-Policy", CSP_HEADER)
        self.end_headers()
        self.wfile.write(data)


def run(host: str = "127.0.0.1", port: int = 8787) -> None:
    ensure_dirs()
    init_db()
    ensure_admin_user()
    if not WEB_DIR.exists():
        raise SystemExit(f"Web folder not found: {WEB_DIR}")
    print(f"Platform dir: {PLATFORM_DIR}")
    print(f"Web dir     : {WEB_DIR}")
    print(f"Storage dir : {STORAGE_DIR}")
    print(f"Open        : http://{host}:{port}")
    server = ThreadingHTTPServer((host, port), Handler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")
    finally:
        server.server_close()


if __name__ == "__main__":
    run(
        host=os.environ.get("M365_LOCAL_WEBAPP_HOST", "127.0.0.1"),
        port=int(os.environ.get("M365_LOCAL_WEBAPP_PORT", "8787")),
    )
