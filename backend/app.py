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
import html as html_lib
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
SKU_FRIENDLY_MAP_PATH = PLATFORM_DIR / "shared" / "m365-sku-friendly-names.json"
CAPABILITY_MATRIX_PATH = PLATFORM_DIR / "shared" / "denjoy-capability-matrix.json"


def now_iso() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat()


_sku_friendly_map_cache: Optional[Dict[str, str]] = None
_capability_matrix_cache: Optional[Dict[str, Any]] = None


def load_sku_friendly_map() -> Dict[str, str]:
    global _sku_friendly_map_cache
    if _sku_friendly_map_cache is not None:
        return _sku_friendly_map_cache
    try:
        data = json.loads(SKU_FRIENDLY_MAP_PATH.read_text(encoding="utf-8"))
        if isinstance(data, dict):
            _sku_friendly_map_cache = {str(k).upper(): str(v) for k, v in data.items()}
            return _sku_friendly_map_cache
    except Exception:
        pass
    _sku_friendly_map_cache = {}
    return _sku_friendly_map_cache


def get_sku_friendly_name(sku: str) -> str:
    key = (sku or "").strip()
    if not key:
        return "Onbekende licentie"
    friendly = load_sku_friendly_map().get(key.upper())
    if friendly:
        return friendly
    return _friendly_license_name(key) or key


def load_capability_matrix() -> Dict[str, Any]:
    global _capability_matrix_cache
    if _capability_matrix_cache is not None:
        return _capability_matrix_cache
    try:
        data = json.loads(CAPABILITY_MATRIX_PATH.read_text(encoding="utf-8"))
        if isinstance(data, dict):
            _capability_matrix_cache = data
            return data
    except Exception:
        pass
    _capability_matrix_cache = {"version": 1, "modules": [], "defaults": {}}
    return _capability_matrix_cache


def _find_capability_module(section: str) -> Optional[Dict[str, Any]]:
    section_key = (section or "").strip().lower()
    modules = load_capability_matrix().get("modules") or []
    for module in modules:
        if isinstance(module, dict) and str(module.get("section") or "").lower() == section_key:
            return module
    return None


_PORTAL_TAB_ALIASES: Dict[str, str] = {
    "overzicht": "summary",
    "apparaten": "devices",
    "configuratie": "config",
    "geschiedenis": "history",
    "regels": "mailbox-rules",
    "mailboxen": "mailboxes",
    "forwarding": "forwarding",
}


def _find_capability_subsection(section: str, subsection: str) -> Optional[Dict[str, Any]]:
    module = _find_capability_module(section)
    if not module:
        return None
    subsection_key = (subsection or "").strip().lower()
    for item in module.get("subsections") or []:
        if isinstance(item, dict) and str(item.get("key") or "").lower() == subsection_key:
            return item
    # Try Dutch portal tab name → English matrix key alias
    alias_key = _PORTAL_TAB_ALIASES.get(subsection_key)
    if alias_key:
        for item in module.get("subsections") or []:
            if isinstance(item, dict) and str(item.get("key") or "").lower() == alias_key:
                return item
    return None


def _has_auth_profile_config(tenant_id: str) -> bool:
    profile = get_tenant_auth_profile(tenant_id, include_secret=True)
    cfg = load_config()
    client_id = (profile.get("auth_client_id") or cfg.get("auth_client_id") or "").strip()
    tenant_auth_id = (profile.get("auth_tenant_id") or cfg.get("auth_tenant_id") or "").strip()
    cert_thumb = (profile.get("auth_cert_thumbprint") or cfg.get("auth_cert_thumbprint") or "").strip()
    client_secret = (profile.get("auth_client_secret") or cfg.get("auth_client_secret") or "").strip()
    return bool(client_id and tenant_auth_id and (cert_thumb or client_secret))


def _connector_available_for_section(section: str) -> bool:
    return (section or "").strip().lower() in {
        "gebruikers", "identity", "apps", "ca", "alerts",
        "intune", "exchange", "teams", "sharepoint",
        "backup", "domains", "compliance", "hybrid",
    }


def _build_capability_status(tenant_id: str, section: str, subsection: str) -> Dict[str, Any]:
    module = _find_capability_module(section)
    sub = _find_capability_subsection(section, subsection)
    if not module or not sub:
        assessment_snapshot = _latest_assessment_snapshot_for_tenant(tenant_id)
        return {
            "section": section, "section_label": section.capitalize(),
            "subsection": subsection, "subsection_label": subsection.capitalize(),
            "engine": "unknown", "live_source": None, "access_method": "unknown",
            "overview_supported": False, "supports_live": False,
            "supports_snapshot": True, "assessment_fallback": True, "backend_only": True,
            "gdap_required": False, "gdap_sufficient": False,
            "extra_roles": [], "extra_consent": [], "cache_minutes": 0, "write_supported": False,
            "connector_available": False, "app_registration_ready": False,
            "assessment_available": bool(assessment_snapshot),
            "assessment_generated_at": assessment_snapshot.get("assessment_generated_at") if assessment_snapshot else None,
            "status": "snapshot_only", "status_label": "Snapshot",
            "status_reason": "Capability definitie niet gevonden in matrix — snapshot-only modus.",
        }

    supports_live = bool(sub.get("supports_live"))
    connector_available = _connector_available_for_section(section)
    auth_ready = _has_auth_profile_config(tenant_id)
    assessment_snapshot = _latest_assessment_snapshot_for_tenant(tenant_id)
    assessment_available = bool(assessment_snapshot)
    access_method = str(module.get("access_method") or "")
    gdap_required = bool(sub.get("gdap_required"))
    gdap_sufficient = bool(sub.get("gdap_sufficient"))

    status = "ready"
    status_label = "Live beschikbaar"
    status_reason = "Connector en basisconfiguratie zijn aanwezig."

    if not supports_live:
        status = "snapshot_only"
        status_label = "Live niet ondersteund"
        status_reason = "Dit onderdeel is bedoeld als historie of portaldata en heeft geen eigen live connector."
    elif not connector_available:
        status = "not_implemented"
        status_label = "Connector nog niet beschikbaar"
        status_reason = "Voor dit onderdeel is de control-plane richting vastgelegd, maar de live connector is nog niet in de huidige portal aangesloten."
    elif access_method == "azure_lighthouse":
        status = "not_implemented"
        status_label = "Azure engine nog niet actief"
        status_reason = "Deze capability hoort bij de Azure-engine en vereist Azure Lighthouse + Azure APIs in een volgende bouwstap."
    elif not auth_ready:
        status = "config_required"
        status_label = "App-configuratie vereist"
        status_reason = "De tenant heeft nog geen complete app-registratieconfiguratie met client-id en certificaat/secret."
    elif gdap_required and gdap_sufficient:
        status = "validation_required"
        status_label = "Live via GDAP"
        status_reason = "Basisconfiguratie is aanwezig. Valideer nog wel of de juiste GDAP-relatie, security group en roltoewijzing actief zijn."
    elif access_method == "customer_app_consent_first":
        status = "validation_required"
        status_label = "Live via App Consent"
        status_reason = "Basisconfiguratie is aanwezig. Dit onderdeel werkt het best met customer app consent of een tenant-specifieke app-registratie."
    elif access_method == "hybrid_gdap_or_customer_app":
        status = "validation_required"
        status_label = "Live via hybride toegang"
        status_reason = "Basisconfiguratie is aanwezig. Afhankelijk van workload en tenant zijn GDAP, extra rollen of customer app consent nodig."

    return {
        "section": module.get("section"),
        "section_label": module.get("label"),
        "subsection": sub.get("key"),
        "subsection_label": sub.get("label"),
        "engine": module.get("engine"),
        "live_source": module.get("live_source"),
        "access_method": access_method,
        "overview_supported": bool(module.get("overview_supported")),
        "supports_live": supports_live,
        "supports_snapshot": bool(load_capability_matrix().get("defaults", {}).get("supports_snapshot", True)),
        "assessment_fallback": bool(load_capability_matrix().get("defaults", {}).get("assessment_fallback", True)),
        "backend_only": bool(load_capability_matrix().get("defaults", {}).get("backend_only", True)),
        "gdap_required": gdap_required,
        "gdap_sufficient": gdap_sufficient,
        "extra_roles": list(sub.get("extra_roles") or []),
        "extra_consent": list(sub.get("extra_consent") or []),
        "cache_minutes": int(sub.get("cache_minutes") or 0),
        "write_supported": bool(sub.get("write_supported")),
        "connector_available": connector_available,
        "app_registration_ready": auth_ready,
        "assessment_available": assessment_available,
        "assessment_generated_at": assessment_snapshot.get("assessment_generated_at") if assessment_snapshot else None,
        "status": status,
        "status_label": status_label,
        "status_reason": status_reason,
    }


def get_tenant_capabilities(tenant_id: str) -> Dict[str, Any]:
    modules_out: List[Dict[str, Any]] = []
    for module in load_capability_matrix().get("modules") or []:
        if not isinstance(module, dict):
            continue
        subs = []
        for item in module.get("subsections") or []:
            if not isinstance(item, dict):
                continue
            try:
                subs.append(_build_capability_status(tenant_id, str(module.get("section") or ""), str(item.get("key") or "")))
            except ValueError:
                continue
        mod_out = dict(module)
        mod_out["subsections"] = subs
        modules_out.append(mod_out)
    return {"ok": True, "tenant_id": tenant_id, "modules": modules_out}


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
        # Tenant-specifieke app registratie-profielen (key = tenant_id)
        # Elke entry: {auth_tenant_id, auth_client_id, auth_cert_thumbprint, auth_client_secret}
        "tenant_auth_profiles": {},
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
    # Env var overrides — secrets should not live in config.json
    _env_secret = os.environ.get("DENJOY_CLIENT_SECRET", "").strip()
    if _env_secret:
        merged["auth_client_secret"] = _env_secret
    _env_tenant = os.environ.get("DENJOY_AUTH_TENANT_ID", "").strip()
    if _env_tenant:
        merged["auth_tenant_id"] = _env_tenant
    _env_client = os.environ.get("DENJOY_AUTH_CLIENT_ID", "").strip()
    if _env_client:
        merged["auth_client_id"] = _env_client
    return merged


def save_config(cfg: Dict[str, Any]) -> None:
    ensure_dirs()
    CONFIG_PATH.write_text(json.dumps(cfg, indent=2), encoding="utf-8")


def get_tenant_auth_profile(tenant_id: str, include_secret: bool = False) -> Dict[str, Any]:
    cfg = load_config()
    profiles = cfg.get("tenant_auth_profiles") if isinstance(cfg.get("tenant_auth_profiles"), dict) else {}
    profile = profiles.get(tenant_id) if isinstance(profiles, dict) else None
    profile = profile if isinstance(profile, dict) else {}
    result = {
        "auth_tenant_id": (profile.get("auth_tenant_id") or "").strip(),
        "auth_client_id": (profile.get("auth_client_id") or "").strip(),
        "auth_cert_thumbprint": (profile.get("auth_cert_thumbprint") or "").strip(),
    }
    if include_secret:
        result["auth_client_secret"] = (profile.get("auth_client_secret") or "").strip()
    return result


def save_tenant_auth_profile(tenant_id: str, payload: Dict[str, Any]) -> Dict[str, Any]:
    tenant = db_fetchone("SELECT id, tenant_guid FROM tenants WHERE id=?", (tenant_id,))
    if not tenant:
        raise ValueError("Tenant niet gevonden")

    cfg = load_config()
    profiles = cfg.get("tenant_auth_profiles") if isinstance(cfg.get("tenant_auth_profiles"), dict) else {}
    profile = profiles.get(tenant_id) if isinstance(profiles.get(tenant_id), dict) else {}

    auth_tenant_id = (payload.get("auth_tenant_id") or "").strip()
    auth_client_id = (payload.get("auth_client_id") or "").strip()
    auth_cert_thumbprint = (payload.get("auth_cert_thumbprint") or "").strip()

    tenant_guid = (tenant.get("tenant_guid") or "").strip()
    if tenant_guid and auth_tenant_id and tenant_guid.lower() != auth_tenant_id.lower():
        raise ValueError("App-registratie Tenant ID moet overeenkomen met de tenant GUID van de geselecteerde tenant.")

    # Secret alleen vervangen als expliciet meegegeven; leeg veld betekent 'ongewijzigd laten'.
    if "auth_client_secret" in payload:
        incoming_secret = (payload.get("auth_client_secret") or "").strip()
        if incoming_secret:
            profile["auth_client_secret"] = incoming_secret

    profile["auth_tenant_id"] = auth_tenant_id
    profile["auth_client_id"] = auth_client_id
    profile["auth_cert_thumbprint"] = auth_cert_thumbprint

    profiles[tenant_id] = profile
    cfg["tenant_auth_profiles"] = profiles
    save_config(cfg)
    return get_tenant_auth_profile(tenant_id, include_secret=False)


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

        CREATE TABLE IF NOT EXISTS remediation_history (
            id TEXT PRIMARY KEY,
            tenant_id TEXT NOT NULL,
            remediation_id TEXT NOT NULL,
            title TEXT NOT NULL,
            executed_by TEXT,
            executed_at TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'success',
            dry_run INTEGER NOT NULL DEFAULT 0,
            params_json TEXT,
            result_json TEXT,
            error_message TEXT,
            FOREIGN KEY (tenant_id) REFERENCES tenants(id)
        );

        CREATE TABLE IF NOT EXISTS provisioning_history (
            id TEXT PRIMARY KEY,
            tenant_id TEXT NOT NULL,
            action TEXT NOT NULL,
            target_upn TEXT,
            target_display_name TEXT,
            executed_by TEXT,
            executed_at TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'success',
            dry_run INTEGER NOT NULL DEFAULT 0,
            params_json TEXT,
            result_json TEXT,
            error_message TEXT,
            FOREIGN KEY (tenant_id) REFERENCES tenants(id)
        );

        CREATE TABLE IF NOT EXISTS baselines (
            id TEXT PRIMARY KEY,
            name TEXT NOT NULL,
            description TEXT,
            source_tenant_id TEXT,
            source_tenant_name TEXT,
            config_json TEXT NOT NULL,
            created_by TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS baseline_assignments (
            id TEXT PRIMARY KEY,
            baseline_id TEXT NOT NULL,
            tenant_id TEXT NOT NULL,
            assigned_by TEXT,
            assigned_at TEXT NOT NULL,
            last_checked_at TEXT,
            last_applied_at TEXT,
            compliance_score INTEGER,
            compliance_json TEXT,
            status TEXT NOT NULL DEFAULT 'assigned',
            FOREIGN KEY (baseline_id) REFERENCES baselines(id),
            FOREIGN KEY (tenant_id) REFERENCES tenants(id),
            UNIQUE(baseline_id, tenant_id)
        );

        CREATE TABLE IF NOT EXISTS baseline_history (
            id TEXT PRIMARY KEY,
            baseline_id TEXT NOT NULL,
            tenant_id TEXT NOT NULL,
            action TEXT NOT NULL,
            executed_by TEXT,
            executed_at TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'success',
            dry_run INTEGER NOT NULL DEFAULT 0,
            result_json TEXT,
            error_message TEXT,
            FOREIGN KEY (baseline_id) REFERENCES baselines(id),
            FOREIGN KEY (tenant_id) REFERENCES tenants(id)
        );
        CREATE TABLE IF NOT EXISTS intune_scan_history (
            id TEXT PRIMARY KEY,
            tenant_id TEXT NOT NULL,
            action TEXT NOT NULL,
            executed_by TEXT,
            executed_at TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'success',
            dry_run INTEGER NOT NULL DEFAULT 0,
            result_json TEXT,
            error_message TEXT,
            FOREIGN KEY (tenant_id) REFERENCES tenants(id)
        );
        CREATE TABLE IF NOT EXISTS backup_history (
            id TEXT PRIMARY KEY,
            tenant_id TEXT NOT NULL,
            action TEXT NOT NULL,
            executed_by TEXT,
            executed_at TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'success',
            result_json TEXT,
            error_message TEXT,
            FOREIGN KEY (tenant_id) REFERENCES tenants(id)
        );
        CREATE TABLE IF NOT EXISTS ca_history (
            id TEXT PRIMARY KEY,
            tenant_id TEXT NOT NULL,
            action TEXT NOT NULL,
            policy_id TEXT,
            executed_by TEXT,
            executed_at TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'success',
            result_json TEXT,
            error_message TEXT,
            FOREIGN KEY (tenant_id) REFERENCES tenants(id)
        );
        CREATE TABLE IF NOT EXISTS alert_config (
            id TEXT PRIMARY KEY,
            tenant_id TEXT NOT NULL UNIQUE,
            webhook_url TEXT,
            webhook_type TEXT NOT NULL DEFAULT 'teams',
            email_addr TEXT,
            updated_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS scan_findings (
            id TEXT PRIMARY KEY,
            tenant_id TEXT NOT NULL,
            domain TEXT NOT NULL,
            control TEXT NOT NULL,
            title TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'info',
            finding TEXT,
            impact TEXT NOT NULL DEFAULT 'low',
            recommendation TEXT,
            service TEXT,
            metric_value REAL,
            raw_json TEXT,
            scanned_at TEXT NOT NULL,
            FOREIGN KEY (tenant_id) REFERENCES tenants(id)
        );

        -- ── Fase 3: MSP control plane tabellen ─────────────────────────────────
        CREATE TABLE IF NOT EXISTS customers (
            id                    TEXT PRIMARY KEY,
            name                  TEXT NOT NULL,
            status                TEXT NOT NULL DEFAULT 'active',
            primary_contact_name  TEXT,
            primary_contact_email TEXT,
            notes                 TEXT,
            created_at            TEXT NOT NULL,
            updated_at            TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS customer_services (
            id           TEXT PRIMARY KEY,
            customer_id  TEXT NOT NULL,
            service_key  TEXT NOT NULL,
            is_enabled   INTEGER NOT NULL DEFAULT 1,
            onboarded_at TEXT,
            notes        TEXT,
            FOREIGN KEY (customer_id) REFERENCES customers(id),
            UNIQUE(customer_id, service_key)
        );

        CREATE TABLE IF NOT EXISTS integrations (
            id                         TEXT PRIMARY KEY,
            tenant_id                  TEXT,
            integration_type           TEXT NOT NULL,
            status                     TEXT NOT NULL DEFAULT 'unknown',
            auth_mode                  TEXT,
            gdap_status                TEXT,
            lighthouse_status          TEXT,
            app_registration_status    TEXT,
            certificate_status         TEXT,
            last_validated_at          TEXT,
            details_json               TEXT,
            created_at                 TEXT NOT NULL,
            updated_at                 TEXT NOT NULL,
            FOREIGN KEY (tenant_id) REFERENCES tenants(id)
        );

        CREATE TABLE IF NOT EXISTS m365_snapshots (
            id                TEXT PRIMARY KEY,
            tenant_id         TEXT NOT NULL,
            section           TEXT NOT NULL,
            subsection        TEXT NOT NULL,
            source_type       TEXT NOT NULL DEFAULT 'assessment',
            generated_at      TEXT NOT NULL,
            stale_after_at    TEXT,
            data_json         TEXT,
            summary_json      TEXT,
            assessment_run_id TEXT,
            FOREIGN KEY (tenant_id) REFERENCES tenants(id),
            FOREIGN KEY (assessment_run_id) REFERENCES assessment_runs(id)
        );

        CREATE TABLE IF NOT EXISTS action_logs (
            id             TEXT PRIMARY KEY,
            portal_user_id TEXT,
            tenant_id      TEXT,
            engine         TEXT,
            section        TEXT,
            subsection     TEXT,
            action_type    TEXT NOT NULL,
            target_id      TEXT,
            result         TEXT NOT NULL DEFAULT 'success',
            error_message  TEXT,
            metadata_json  TEXT,
            created_at     TEXT NOT NULL,
            FOREIGN KEY (tenant_id) REFERENCES tenants(id)
        );

        CREATE TABLE IF NOT EXISTS approvals (
            id              TEXT PRIMARY KEY,
            action_log_id   TEXT NOT NULL,
            approval_status TEXT NOT NULL DEFAULT 'pending',
            requested_by    TEXT,
            approved_by     TEXT,
            requested_at    TEXT NOT NULL,
            approved_at     TEXT,
            reason          TEXT,
            FOREIGN KEY (action_log_id) REFERENCES action_logs(id)
        );

        -- ── Fase 6: Rollen en klant-toegangsmodel ─────────────────────────────
        CREATE TABLE IF NOT EXISTS portal_roles (
            id          TEXT PRIMARY KEY,
            role_key    TEXT NOT NULL UNIQUE,
            label       TEXT NOT NULL,
            description TEXT,
            created_at  TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS user_customer_access (
            id             TEXT PRIMARY KEY,
            portal_user_id TEXT NOT NULL,
            customer_id    TEXT NOT NULL,
            portal_role_id TEXT NOT NULL,
            scope          TEXT,
            granted_by     TEXT,
            granted_at     TEXT NOT NULL,
            expires_at     TEXT,
            FOREIGN KEY (portal_user_id) REFERENCES users(id),
            FOREIGN KEY (customer_id) REFERENCES customers(id),
            FOREIGN KEY (portal_role_id) REFERENCES portal_roles(id),
            UNIQUE(portal_user_id, customer_id)
        );

        -- ── Fase 4: Azure subscriptions registry ──────────────────────────────
        CREATE TABLE IF NOT EXISTS subscriptions (
            id                    TEXT PRIMARY KEY,
            tenant_id             TEXT NOT NULL,
            azure_subscription_id TEXT NOT NULL,
            display_name          TEXT,
            state                 TEXT NOT NULL DEFAULT 'active',
            lighthouse_onboarded  INTEGER NOT NULL DEFAULT 0,
            management_group      TEXT,
            created_at            TEXT NOT NULL,
            FOREIGN KEY (tenant_id) REFERENCES tenants(id),
            UNIQUE(tenant_id, azure_subscription_id)
        );

        -- ── Fase 4: Azure snapshot tabellen ───────────────────────────────────
        CREATE TABLE IF NOT EXISTS azure_resource_snapshots (
            id              TEXT PRIMARY KEY,
            tenant_id       TEXT NOT NULL,
            subscription_id TEXT,
            section         TEXT NOT NULL,
            subsection      TEXT NOT NULL,
            generated_at    TEXT NOT NULL,
            stale_after_at  TEXT,
            data_json       TEXT,
            summary_json    TEXT,
            FOREIGN KEY (tenant_id) REFERENCES tenants(id)
        );

        CREATE TABLE IF NOT EXISTS alert_snapshots (
            id           TEXT PRIMARY KEY,
            tenant_id    TEXT NOT NULL,
            alert_type   TEXT NOT NULL,
            generated_at TEXT NOT NULL,
            data_json    TEXT,
            summary_json TEXT,
            FOREIGN KEY (tenant_id) REFERENCES tenants(id)
        );

        CREATE TABLE IF NOT EXISTS cost_snapshots (
            id              TEXT PRIMARY KEY,
            tenant_id       TEXT NOT NULL,
            subscription_id TEXT,
            period_start    TEXT NOT NULL,
            period_end      TEXT NOT NULL,
            generated_at    TEXT NOT NULL,
            data_json       TEXT,
            summary_json    TEXT,
            FOREIGN KEY (tenant_id) REFERENCES tenants(id)
        );

        -- ── Fase 7: Job queue voor assessment en live retrieval ────────────────
        CREATE TABLE IF NOT EXISTS job_queue (
            id            TEXT PRIMARY KEY,
            job_type      TEXT NOT NULL,
            tenant_id     TEXT,
            payload_json  TEXT,
            status        TEXT NOT NULL DEFAULT 'pending',
            priority      INTEGER NOT NULL DEFAULT 5,
            attempt_count INTEGER NOT NULL DEFAULT 0,
            max_attempts  INTEGER NOT NULL DEFAULT 3,
            scheduled_at  TEXT NOT NULL,
            started_at    TEXT,
            completed_at  TEXT,
            error_message TEXT,
            result_json   TEXT,
            created_at    TEXT NOT NULL,
            FOREIGN KEY (tenant_id) REFERENCES tenants(id)
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
    action_cols = {r[1] for r in cur.execute("PRAGMA table_info(finding_actions)").fetchall()}
    if "kb_asset_id" not in action_cols:
        cur.execute("ALTER TABLE finding_actions ADD COLUMN kb_asset_id INTEGER")
    if "kb_asset_name" not in action_cols:
        cur.execute("ALTER TABLE finding_actions ADD COLUMN kb_asset_name TEXT")
    audit_cols = {r[1] for r in cur.execute("PRAGMA table_info(audit_logs)").fetchall()}
    if "tenant_id" not in audit_cols:
        cur.execute("ALTER TABLE audit_logs ADD COLUMN tenant_id TEXT")
    # Fase 3 — customer_id op tenants (optionele koppeling aan customers tabel)
    tenant_cols_v2 = {r[1] for r in cur.execute("PRAGMA table_info(tenants)").fetchall()}
    if "customer_id" not in tenant_cols_v2:
        cur.execute("ALTER TABLE tenants ADD COLUMN customer_id TEXT REFERENCES customers(id)")
    # Fase 6 — extra kolommen op users tabel
    user_cols = {r[1] for r in cur.execute("PRAGMA table_info(users)").fetchall()}
    if "last_login_at" not in user_cols:
        cur.execute("ALTER TABLE users ADD COLUMN last_login_at TEXT")
    if "entra_object_id" not in user_cols:
        cur.execute("ALTER TABLE users ADD COLUMN entra_object_id TEXT")
    # Fase 5 migration — backup_history tabel aanmaken als die nog niet bestaat
    existing_tables = {r[0] for r in cur.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()}
    if "backup_history" not in existing_tables:
        cur.execute("""
            CREATE TABLE backup_history (
                id TEXT PRIMARY KEY,
                tenant_id TEXT NOT NULL,
                action TEXT NOT NULL,
                executed_by TEXT,
                executed_at TEXT NOT NULL,
                status TEXT NOT NULL DEFAULT 'success',
                result_json TEXT,
                error_message TEXT,
                FOREIGN KEY (tenant_id) REFERENCES tenants(id)
            )
        """)
    # ── Performance indexes (idempotent) ──────────────────────────────────────
    cur.executescript("""
        CREATE INDEX IF NOT EXISTS idx_runs_tenant_status
            ON assessment_runs(tenant_id, status);
        CREATE INDEX IF NOT EXISTS idx_runs_tenant_completed
            ON assessment_runs(tenant_id, completed_at DESC);
        CREATE INDEX IF NOT EXISTS idx_actions_run
            ON finding_actions(run_id);
        CREATE INDEX IF NOT EXISTS idx_actions_tenant
            ON finding_actions(tenant_id);
        CREATE INDEX IF NOT EXISTS idx_remediation_tenant
            ON remediation_history(tenant_id);
        CREATE INDEX IF NOT EXISTS idx_ca_history_tenant
            ON ca_history(tenant_id);
        CREATE INDEX IF NOT EXISTS idx_backup_history_tenant
            ON backup_history(tenant_id);
        CREATE INDEX IF NOT EXISTS idx_audit_logs_tenant
            ON audit_logs(tenant_id);
        CREATE INDEX IF NOT EXISTS idx_audit_logs_ts
            ON audit_logs(created_at DESC);
        CREATE INDEX IF NOT EXISTS idx_sessions_user
            ON sessions(user_id);
        CREATE INDEX IF NOT EXISTS idx_sessions_expires
            ON sessions(expires_at);
        CREATE INDEX IF NOT EXISTS idx_baseline_assignments_tenant
            ON baseline_assignments(tenant_id);
        CREATE INDEX IF NOT EXISTS idx_baseline_assignments_baseline
            ON baseline_assignments(baseline_id);
        CREATE INDEX IF NOT EXISTS idx_scan_findings_tenant_at
            ON scan_findings(tenant_id, scanned_at DESC);
        CREATE INDEX IF NOT EXISTS idx_scan_findings_domain
            ON scan_findings(domain, status);
        CREATE INDEX IF NOT EXISTS idx_scan_findings_control
            ON scan_findings(tenant_id, domain, control);
        CREATE INDEX IF NOT EXISTS idx_m365_snapshots_tenant_section
            ON m365_snapshots(tenant_id, section, subsection);
        CREATE INDEX IF NOT EXISTS idx_m365_snapshots_generated
            ON m365_snapshots(tenant_id, generated_at DESC);
        CREATE INDEX IF NOT EXISTS idx_action_logs_tenant
            ON action_logs(tenant_id, created_at DESC);
        CREATE INDEX IF NOT EXISTS idx_integrations_tenant
            ON integrations(tenant_id);
        CREATE INDEX IF NOT EXISTS idx_customers_status
            ON customers(status);
        CREATE INDEX IF NOT EXISTS idx_user_customer_access_user
            ON user_customer_access(portal_user_id);
        CREATE INDEX IF NOT EXISTS idx_user_customer_access_customer
            ON user_customer_access(customer_id);
        CREATE INDEX IF NOT EXISTS idx_subscriptions_tenant
            ON subscriptions(tenant_id);
        CREATE INDEX IF NOT EXISTS idx_azure_snapshots_tenant
            ON azure_resource_snapshots(tenant_id, section, subsection);
        CREATE INDEX IF NOT EXISTS idx_alert_snapshots_tenant
            ON alert_snapshots(tenant_id, alert_type);
        CREATE INDEX IF NOT EXISTS idx_cost_snapshots_tenant
            ON cost_snapshots(tenant_id, period_start DESC);
        CREATE INDEX IF NOT EXISTS idx_job_queue_status
            ON job_queue(status, scheduled_at);
        CREATE INDEX IF NOT EXISTS idx_job_queue_tenant
            ON job_queue(tenant_id, status);
    """)
    conn.commit()
    # Seed standaard portal_roles als die nog niet bestaan
    _default_roles = [
        ("msp_super_admin", "MSP Super Admin", "Volledige platformtoegang"),
        ("engineer",        "Engineer",         "Operationele toegang, acties uitvoeren"),
        ("monitoring_operator", "Monitoring Operator", "Lezen en monitoring, geen schrijftoegang"),
        ("billing_analyst", "Billing Analyst",  "Toegang tot kosten- en licentiedata"),
        ("read_only",       "Alleen lezen",      "Read-only toegang tot alle modules"),
    ]
    for rkey, rlabel, rdesc in _default_roles:
        existing = cur.execute("SELECT id FROM portal_roles WHERE role_key=?", (rkey,)).fetchone()
        if not existing:
            cur.execute(
                "INSERT INTO portal_roles (id, role_key, label, description, created_at) VALUES (?,?,?,?,?)",
                (str(uuid.uuid4()), rkey, rlabel, rdesc, now_iso()),
            )
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
    "/api/upload-report",   # PowerShell scripts uploaden rapporten zonder sessie-token
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
        handler._json(401, {"error": "Niet ingelogd.", "error_code": "unauthorized"})
        return None

    # Admin-only routes: config en runs starten
    _admin_paths = {"/api/config", "/api/runs"}
    _admin_prefix = ("/api/users", "/api/remediate", "/api/m365", "/api/baselines", "/api/intune", "/api/backup",
                     "/api/ca", "/api/domains", "/api/alerts", "/api/exchange", "/api/identity",
                     "/api/apps", "/api/collaboration", "/api/capabilities",
                     "/api/customers", "/api/audit", "/api/approvals", "/api/integrations", "/api/portal-roles",
                     "/api/jobs")
    if path in _admin_paths or any(path.startswith(p) for p in _admin_prefix):
        if sess["role"] != "admin":
            handler._json(403, {"error": "Onvoldoende rechten.", "error_code": "forbidden"})
            return None

    # Tenant-scoped routes: klanten mogen alleen hun eigen tenant benaderen
    _tid_m = re.match(r"/api/(?:tenants|assessment|kb|identity|apps|collaboration)/([^/]+)(?:/|$)", path)
    if _tid_m and sess["role"] != "admin":
        req_tid = _tid_m.group(1)
        user_row = db_fetchone("SELECT linked_tenant_id FROM users WHERE email=?", (sess["email"],))
        if not user_row or user_row["linked_tenant_id"] != req_tid:
            handler._json(403, {"error": "Geen toegang tot deze tenant.", "error_code": "forbidden"})
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


def db_execute(sql: str, params: Tuple[Any, ...] = ()) -> int:
    """Voert een write-query uit en retourneert het aantal gewijzigde rijen."""
    conn = get_conn()
    try:
        cur = conn.execute(sql, params)
        conn.commit()
        return cur.rowcount
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


def list_audit_logs(
    tenant_id: Optional[str] = None,
    user_email: Optional[str] = None,
    action: Optional[str] = None,
    date_from: Optional[str] = None,
    date_to: Optional[str] = None,
    limit: int = 200,
) -> List[Dict[str, Any]]:
    """Retourneert audit_logs met optionele filters."""
    clauses: List[str] = []
    params: List[Any] = []
    if tenant_id:
        clauses.append("tenant_id=?")
        params.append(tenant_id)
    if user_email:
        clauses.append("lower(user_email) LIKE ?")
        params.append(f"%{user_email.lower()}%")
    if action:
        clauses.append("action LIKE ?")
        params.append(f"%{action}%")
    if date_from:
        clauses.append("created_at >= ?")
        params.append(date_from)
    if date_to:
        clauses.append("created_at <= ?")
        params.append(date_to)
    where = ("WHERE " + " AND ".join(clauses)) if clauses else ""
    limit = min(max(1, limit), 1000)
    return db_fetchall(
        f"SELECT * FROM audit_logs {where} ORDER BY created_at DESC LIMIT ?",
        tuple(params) + (limit,),
    )


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
    demo = sorted(run_dir.glob("*.html"), key=lambda p: p.stat().st_mtime, reverse=True)
    return demo[0] if demo else None


def list_run_html_files(run_dir: Path) -> List[Dict[str, str]]:
    """Geeft HTML-bestanden in de run-directory terug."""
    result = []
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


def find_latest_json_manifest_file(run_dir: Path) -> Optional[Path]:
    json_dir = run_dir / "json"
    manifest = json_dir / "manifest.json"
    if manifest.exists():
        return manifest
    return None


def _safe_json_load(path: Path) -> Dict[str, Any]:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _load_assessment_json_payloads(run_dir: Path) -> Dict[Tuple[str, str], Dict[str, Any]]:
    manifest_file = find_latest_json_manifest_file(run_dir)
    if not manifest_file:
        return {}
    manifest = _safe_json_load(manifest_file)
    payloads: Dict[Tuple[str, str], Dict[str, Any]] = {}
    for item in manifest.get("files") or []:
        if not isinstance(item, dict):
            continue
        section = str(item.get("section") or "").strip().lower()
        subsection = str(item.get("subsection") or "").strip().lower()
        relative = str(item.get("relative") or "").strip()
        if not section or not subsection or not relative:
            continue
        payload_file = (manifest_file.parent / relative).resolve()
        if not payload_file.exists():
            continue
        payload = _safe_json_load(payload_file)
        if payload:
            payloads[(section, subsection)] = payload
    return payloads


def _assessment_json_payload(snapshot: Dict[str, Any], section: str, subsection: str) -> Optional[Dict[str, Any]]:
    payloads = snapshot.get("assessment_json_payloads") or {}
    return payloads.get(((section or "").strip().lower(), (subsection or "").strip().lower()))


def _payload_value(item: Dict[str, Any], *keys: str, default: Any = None) -> Any:
    if not isinstance(item, dict):
        return default
    for key in keys:
        if key in item and item.get(key) not in (None, ""):
            return item.get(key)
    return default


JSON_PHASE_DEFINITIONS: List[Dict[str, Any]] = [
    {
        "number": 1,
        "id": "phase1",
        "nav_label": "Identiteit",
        "title": "Phase 1: Users, Licensing & Security Basics",
        "pairs": [("gebruikers", "users"), ("gebruikers", "licenses"), ("identity", "mfa")],
    },
    {
        "number": 2,
        "id": "phase2",
        "nav_label": "Samenwerking",
        "title": "Phase 2: Collaboration & Storage",
        "pairs": [("teams", "teams"), ("sharepoint", "sharepoint-sites"), ("sharepoint", "sharepoint-settings"), ("backup", "onedrive"), ("exchange", "mailboxes")],
    },
    {
        "number": 3,
        "id": "phase3",
        "nav_label": "Compliance",
        "title": "Phase 3: Compliance & Security Policies",
        "pairs": [("ca", "policies"), ("apps", "registrations")],
    },
    {
        "number": 4,
        "id": "phase4",
        "nav_label": "Advanced Security",
        "title": "Phase 4: Advanced Security & Compliance",
        "pairs": [("alerts", "secure-score"), ("alerts", "audit-logs"), ("identity", "admin-roles")],
    },
    {
        "number": 5,
        "id": "phase5",
        "nav_label": "Intune",
        "title": "Phase 5: Intune Configuration",
        "pairs": [("intune", "summary"), ("intune", "devices"), ("intune", "compliance"), ("intune", "config")],
    },
    {
        "number": 6,
        "id": "phase6",
        "nav_label": "Azure",
        "title": "Phase 6: Azure Infrastructure",
        "pairs": [("azure", "subscriptions"), ("azure", "resources"), ("azure", "alerts")],
    },
]


def _run_json_manifest_path(run_dir: Path) -> Optional[str]:
    manifest = find_latest_json_manifest_file(run_dir)
    if not manifest:
        return None
    rel = manifest.relative_to(run_dir).as_posix()
    return f"/reports/{run_dir.name}/{rel}"


def _build_json_phase_summary(payloads: List[Dict[str, Any]]) -> str:
    if not payloads:
        return "Geen JSON-payloads beschikbaar."
    labels = [str(p.get("label") or f"{p.get('section')}/{p.get('subsection')}") for p in payloads if isinstance(p, dict)]
    if len(labels) == 1:
        return f"1 onderdeel beschikbaar: {labels[0]}."
    return f"{len(labels)} onderdelen beschikbaar: {', '.join(labels[:3])}{' …' if len(labels) > 3 else ''}."


def _assessment_json_report_for_run(run_id: str) -> Dict[str, Any]:
    run = get_run(run_id)
    if not run:
        raise ValueError("Run niet gevonden")
    run_dir = RUNS_DIR / run_id
    payloads = _load_assessment_json_payloads(run_dir)
    if not payloads:
        return {"ok": False, "error": "Geen assessment JSON beschikbaar voor deze run"}
    manifest_path = _run_json_manifest_path(run_dir)
    snapshot = _latest_assessment_snapshot_for_tenant(run["tenant_id"])
    generated_at = None
    if manifest_path:
        manifest_file = find_latest_json_manifest_file(run_dir)
        manifest = _safe_json_load(manifest_file) if manifest_file else {}
        generated_at = manifest.get("generated_at")
    phases = []
    for phase_def in JSON_PHASE_DEFINITIONS:
        phase_payloads = []
        for section, subsection in phase_def["pairs"]:
            payload = payloads.get((section, subsection))
            if payload:
                phase_payloads.append(payload)
        if not phase_payloads:
            continue
        score = None
        if phase_def["number"] == 1:
            score = snapshot.get("mfa_coverage")
        elif phase_def["number"] == 4:
            score = snapshot.get("secure_score_percentage")
        phase_items = 0
        for payload in phase_payloads:
            phase_items += len(payload.get("items") or [])
        phases.append({
            "id": phase_def["id"],
            "number": phase_def["number"],
            "navLabel": phase_def["nav_label"],
            "renderLabel": phase_def["title"],
            "summary": _build_json_phase_summary(phase_payloads),
            "score": score,
            "critical": 0,
            "warning": 0,
            "info": phase_items,
            "payloads": phase_payloads,
        })
    return {
        "ok": True,
        "run_id": run_id,
        "tenant_id": run.get("tenant_id"),
        "tenant_name": run.get("tenant_name"),
        "customer_name": run.get("customer_name"),
        "generated_at": generated_at or snapshot.get("assessment_generated_at") or run.get("completed_at") or run.get("started_at"),
        "manifest_path": manifest_path,
        "phases": phases,
    }


def _assessment_nav_sort_key(section: str, subsection: str) -> Tuple[int, int, str]:
    for phase in JSON_PHASE_DEFINITIONS:
        for idx, pair in enumerate(phase["pairs"]):
            if pair == (section, subsection):
                return (phase["number"], idx, f"{section}:{subsection}")
    return (999, 999, f"{section}:{subsection}")


def _assessment_portal_target(section: str, subsection: str) -> Optional[Dict[str, Any]]:
    section_key = (section or "").strip().lower()
    subsection_key = (subsection or "").strip().lower()

    if section_key == "gebruikers":
        tab_map = {
            "users": ("gbTab", "gebruikers"),
            "licenses": ("gbTab", "licenties"),
            "history": ("gbTab", "geschiedenis"),
        }
        mapped = tab_map.get(subsection_key)
        if mapped:
            return {"section": "gebruikers", "tab_type": mapped[0], "tab_key": mapped[1]}
        return {"section": "gebruikers"}

    if section_key == "ca":
        tab_map = {
            "policies": ("caTab", "policies"),
            "named-locations": ("caTab", "locations"),
            "history": ("caTab", "geschiedenis"),
        }
        mapped = tab_map.get(subsection_key)
        if mapped:
            return {"section": "ca", "tab_type": mapped[0], "tab_key": mapped[1]}
        return {"section": "ca"}

    if section_key in {"teams", "sharepoint", "identity", "apps"}:
        return {"section": section_key, "tab_type": "liveTab", "tab_key": subsection_key}

    if section_key == "domains":
        tab_map = {
            "domains-list": "domains-list",
            "domains-analyse": "domains-analyse",
        }
        if subsection_key in tab_map:
            return {"section": "domains", "tab_type": "liveTab", "tab_key": tab_map[subsection_key]}
        return {"section": "domains"}

    if section_key == "exchange":
        tab_map = {
            "mailboxes": "mailboxen",
            "forwarding": "forwarding",
            "mailbox-rules": "regels",
        }
        if subsection_key in tab_map:
            return {"section": "exchange", "tab_type": "liveTab", "tab_key": tab_map[subsection_key]}
        return {"section": "exchange"}

    if section_key == "intune":
        tab_map = {
            "summary": "overzicht",
            "devices": "apparaten",
            "compliance": "compliance",
            "config": "configuratie",
            "history": "geschiedenis",
        }
        if subsection_key in tab_map:
            return {"section": "intune", "tab_type": "liveTab", "tab_key": tab_map[subsection_key]}
        return {"section": "intune"}

    if section_key == "backup":
        tab_map = {
            "summary": "overzicht",
            "onedrive": "onedrive",
            "exchange": "exchange",
            "sharepoint": "sharepoint",
            "history": "geschiedenis",
        }
        if subsection_key in tab_map:
            return {"section": "backup", "tab_type": "liveTab", "tab_key": tab_map[subsection_key]}
        return {"section": "backup"}

    if section_key == "alerts":
        tab_map = {
            "audit-logs": "auditlog",
            "secure-score": "securescr",
            "sign-ins": "signins",
            "notifications": "config",
        }
        if subsection_key in tab_map:
            return {"section": "alerts", "tab_type": "liveTab", "tab_key": tab_map[subsection_key]}
        return {"section": "alerts"}

    return None


def _assessment_item_coverage(tenant_id: str, section: str, subsection: str) -> Dict[str, Any]:
    target = _assessment_portal_target(section, subsection)
    capability = None
    try:
        capability = _build_capability_status(tenant_id, section, subsection)
    except Exception:
        capability = None

    if capability:
        status = str(capability.get("status") or "unknown")
        if target and status in {"ready", "validation_required", "config_required"}:
            bucket = "live_workspace"
            bucket_label = "Live in portal"
            detail = capability.get("status_reason") or "Deze dataset heeft een portal-workspace met live connector."
        elif target and status == "snapshot_only":
            bucket = "snapshot_workspace"
            bucket_label = "Snapshot in portal"
            detail = capability.get("status_reason") or "Dit onderdeel gebruikt assessmentdata in plaats van live data."
        elif not target and bool(capability.get("supports_live")):
            bucket = "live_backend_only"
            bucket_label = "Live connector, workspace ontbreekt"
            detail = "De backend/capability is aanwezig, maar er is nog geen aparte workspace in de portal."
        elif status == "snapshot_only":
            bucket = "snapshot_only"
            bucket_label = "Alleen snapshot"
            detail = capability.get("status_reason") or "Dit onderdeel is nu alleen beschikbaar vanuit assessmentdata."
        else:
            bucket = "not_available"
            bucket_label = capability.get("status_label") or "Nog niet gekoppeld"
            detail = capability.get("status_reason") or "Dit onderdeel vraagt nog extra bouw of configuratie."
        return {
            "bucket": bucket,
            "bucket_label": bucket_label,
            "detail": detail,
            "workspace_available": bool(target),
            "open_target": target,
            "capability": capability,
        }

    return {
        "bucket": "report_only",
        "bucket_label": "Alleen rapport",
        "detail": "Dit onderdeel is wel in de assessment-output gevonden, maar nog niet gekoppeld aan een capability-profiel of workspace.",
        "workspace_available": bool(target),
        "open_target": target,
        "capability": None,
    }


def _format_assessment_json_cell(value: Any) -> str:
    if value in (None, ""):
        return "—"
    if isinstance(value, list):
        if not value:
            return "—"
        parts = []
        for item in value:
            if isinstance(item, dict):
                label = item.get("displayName") or item.get("DisplayName") or item.get("userPrincipalName") or item.get("UserPrincipalName")
                parts.append(str(label or json.dumps(item, ensure_ascii=False)))
            else:
                parts.append(str(item))
        return ", ".join(parts)
    if isinstance(value, dict):
        label = value.get("displayName") or value.get("DisplayName") or value.get("userPrincipalName") or value.get("UserPrincipalName")
        return str(label or json.dumps(value, ensure_ascii=False))
    return str(value)


def _rows_from_json_payload(payload: Dict[str, Any]) -> Tuple[List[str], List[Dict[str, str]]]:
    items = payload.get("items") or []
    if not isinstance(items, list) or not items:
        return ([], [])
    columns: List[str] = []
    for item in items:
        if isinstance(item, dict):
            for key in item.keys():
                if key not in columns:
                    columns.append(str(key))
    rows: List[Dict[str, str]] = []
    for item in items:
        if not isinstance(item, dict):
            continue
        row = {str(column): _format_assessment_json_cell(item.get(column)) for column in columns}
        rows.append(row)
    return (columns, rows)


def _cards_from_json_summary(payload: Dict[str, Any]) -> List[Dict[str, Any]]:
    summary = payload.get("summary") or {}
    if not isinstance(summary, dict):
        return []
    cards = []
    for key, value in summary.items():
        if value in (None, ""):
            continue
        cards.append({"label": str(key), "value": _format_assessment_json_cell(value), "tone": "default"})
    return cards[:8]


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


def _parse_license_assignments_from_html(path: Path) -> Dict[str, List[Dict[str, str]]]:
    try:
        html = path.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        return {}

    section_match = re.search(
        r"<h2 class=\"section-title\">.*?Licentie Overzicht</h2>(.*?)</div>\s*<div class=\"section section-advice-panel\">",
        html,
        re.I | re.S,
    )
    if not section_match:
        return {}
    section_html = section_match.group(1)

    assignments: Dict[str, List[Dict[str, str]]] = {}
    block_pattern = re.compile(
        r'<h3 class="heading-25">([^<]+)</h3>\s*'
        r'<div class="alert alert-info alert-info-soft">.*?</div>\s*'
        r'(?:<div class="table-container">\s*<table>.*?<tbody>(.*?)</tbody>\s*</table></div>|<p class=[\'"]text-muted-italic mb-20[\'"]>Geen gebruikers toegewezen aan deze licentie\.</p>)',
        re.I | re.S,
    )
    for sku_name, rows_html in block_pattern.findall(section_html):
        rows: List[Dict[str, str]] = []
        for upn, display_name in re.findall(r"<tr><td>([^<]+)</td><td>([^<]+)</td></tr>", rows_html or "", re.I):
            rows.append({
                "UserPrincipalName": upn.strip(),
                "DisplayName": display_name.strip(),
            })
        for alias in _license_key_aliases(sku_name.strip()):
            assignments[alias] = rows
    return assignments


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


def _parse_teams_from_html(path: Path) -> List[Dict[str, Any]]:
    try:
        html = path.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        return []
    match = re.search(
        r"Microsoft Teams \(\d+\)</h3>.*?<tbody>(.*?)</tbody></table>",
        html,
        re.I | re.S,
    )
    if not match:
        return []
    body = match.group(1)
    result: List[Dict[str, Any]] = []
    for row_html in re.findall(r"<tr>(.*?)</tr>", body, re.I | re.S):
        cells = re.findall(r"<td[^>]*>(.*?)</td>", row_html, re.I | re.S)
        if len(cells) < 4:
            continue
        mail = _strip_html_fragment(cells[0])
        display_name = _strip_html_fragment(cells[1])
        member_count_raw = _strip_html_fragment(cells[2])
        created = _strip_html_fragment(cells[3])
        try:
            member_count = int(re.sub(r"[^\d]", "", member_count_raw) or "0")
        except Exception:
            member_count = 0
        result.append({
            "id": mail or display_name,
            "mail": mail,
            "displayName": display_name or mail,
            "memberCount": member_count,
            "createdAt": created or None,
            "visibility": None,
            "ownerCount": 0,
            "isDynamic": False,
        })
    return result


def _parse_sharepoint_sites_from_html(path: Path) -> List[Dict[str, Any]]:
    try:
        html = path.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        return []
    match = re.search(
        r"Top 10 Grootste Sites</h4>.*?<tbody>(.*?)</tbody></table>",
        html,
        re.I | re.S,
    )
    if not match:
        return []
    body = match.group(1)
    result: List[Dict[str, Any]] = []
    for row_html in re.findall(r"<tr>(.*?)</tr>", body, re.I | re.S):
        cells = re.findall(r"<td[^>]*>(.*?)</td>", row_html, re.I | re.S)
        if len(cells) < 4:
            continue
        site_html, storage_html, status_html, modified_html = cells[:4]
        url_match = re.search(r"href=['\"]([^'\"]+)['\"]", site_html, re.I)
        result.append({
            "id": url_match.group(1).strip() if url_match else _strip_html_fragment(site_html),
            "displayName": _strip_html_fragment(site_html),
            "webUrl": url_match.group(1).strip() if url_match else None,
            "storageUsed": None,
            "storageLabel": f"{_strip_html_fragment(storage_html)} GB" if _strip_html_fragment(storage_html) else "—",
            "lastModified": _strip_html_fragment(modified_html) or None,
            "status": _strip_html_fragment(status_html) or None,
        })
    return result


def _parse_onedrive_from_html(path: Path) -> List[Dict[str, Any]]:
    try:
        html = path.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        return []
    match = re.search(
        r"Top 5 Grootste OneDrive Sites</h4>.*?<tbody>(.*?)</tbody></table>",
        html,
        re.I | re.S,
    )
    if not match:
        return []
    body = match.group(1)
    result: List[Dict[str, Any]] = []
    for row_html in re.findall(r"<tr>(.*?)</tr>", body, re.I | re.S):
        cells = re.findall(r"<td[^>]*>(.*?)</td>", row_html, re.I | re.S)
        if len(cells) < 3:
            continue
        owner_html, storage_html, modified_html = cells[:3]
        result.append({
            "driveId": _strip_html_fragment(owner_html),
            "ownerName": _strip_html_fragment(owner_html),
            "status": "assessment_snapshot",
            "storageGB": _strip_html_fragment(storage_html),
            "modified": _strip_html_fragment(modified_html) or None,
        })
    return result


def _parse_sharepoint_settings_from_html(path: Path) -> Dict[str, Any]:
    try:
        html = path.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        return {}
    match = re.search(
        r"SharePoint Tenant Sharing Settings</h4>.*?<tbody>(.*?)</tbody></table>",
        html,
        re.I | re.S,
    )
    if not match:
        return {}
    settings_rows = re.findall(r"<tr><td>(.*?)</td><td>(.*?)</td></tr>", match.group(1), re.I | re.S)
    mapped: Dict[str, Any] = {}
    for key_html, value_html in settings_rows:
        key = _strip_html_fragment(key_html).lower()
        value = _strip_html_fragment(value_html)
        if "external sharing capability" in key:
            mapped["sharingCapability"] = value
            mapped["guestSharingEnabled"] = value.lower() not in {"disabled", "uitgeschakeld", "nee"}
        elif "default link permission" in key:
            mapped["defaultLinkPermission"] = value
        elif "loop default sharing scope" in key:
            mapped["defaultSharingLinkType"] = value
    return mapped


def _strip_html_fragment(value: str) -> str:
    text = re.sub(r"<[^>]+>", "", value or "")
    return html_lib.unescape(text).strip()


def _friendly_license_name(value: str) -> str:
    text = (value or "").strip()
    if not text:
        return ""
    text = re.sub(r"^MICROSOFT_", "", text, flags=re.I)
    text = re.sub(r"^STANDARD_", "", text, flags=re.I)
    text = re.sub(r"^PREMIUM_", "", text, flags=re.I)
    text = re.sub(r"_+", " ", text).strip().lower()
    return " ".join(part.capitalize() for part in text.split())


def _license_key_aliases(value: str) -> List[str]:
    raw = (value or "").strip()
    aliases: List[str] = []
    friendly = get_sku_friendly_name(raw)
    for candidate in [raw, raw.upper(), friendly, friendly.upper(), _friendly_license_name(raw), _friendly_license_name(raw).upper()]:
        if candidate and candidate not in aliases:
            aliases.append(candidate)
    return aliases


def _parse_user_overview_counts_from_html(path: Path) -> Dict[str, int]:
    try:
        html = path.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        return {}
    counts: Dict[str, int] = {}
    pattern = re.compile(
        r"stat-number[\"']>(\d+)</div>\s*<div class=['\"]stat-label['\"]>([^<]+)</div>",
        re.I,
    )
    label_map = {
        "totaal gebruikers": "total",
        "actieve gebruikers": "active",
        "uitgeschakelde gebruikers": "disabled",
        "guest gebruikers": "guest",
    }
    for number, label in pattern.findall(html):
        normalized = label.strip().lower()
        key = label_map.get(normalized)
        if key:
            counts[key] = int(number)
    return counts


def _parse_assessment_users_from_html(path: Path) -> List[Dict[str, Any]]:
    try:
        html = path.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        return []

    overview_match = re.search(
        r"<h2 class=\"section-title\">.*?Gebruikers Overzicht</h2>(.*?)</div>\s*<!-- End Overview Section -->",
        html,
        re.I | re.S,
    )
    if not overview_match:
        return []
    overview_html = overview_match.group(1)

    users: List[Dict[str, Any]] = []

    def _extract_rows(section_label: str, enabled: bool) -> None:
        match = re.search(
            rf"{section_label}\s*\(\d+\)</h3>.*?<tbody>(.*?)</tbody>",
            overview_html,
            re.I | re.S,
        )
        if not match:
            return
        body = match.group(1)
        for row_html in re.findall(r"<tr>(.*?)</tr>", body, re.I | re.S):
            cells = re.findall(r"<td[^>]*>(.*?)</td>", row_html, re.I | re.S)
            if len(cells) < 2:
                continue
            upn = _strip_html_fragment(cells[0])
            display_name = _strip_html_fragment(cells[1]) or upn
            last_sign_in = _strip_html_fragment(cells[3]) if len(cells) > 3 else None
            users.append({
                "id": upn or display_name,
                "displayName": display_name,
                "userPrincipalName": upn,
                "mail": upn,
                "accountEnabled": enabled,
                "createdDateTime": None,
                "department": None,
                "jobTitle": None,
                "lastSignIn": last_sign_in if last_sign_in and last_sign_in != "—" else None,
                "userType": "Guest" if "#ext#" in (upn or "").lower() else "Member",
            })

    _extract_rows("Actieve gebruikers", True)
    _extract_rows("Uitgeschakelde gebruikers", False)
    return users


def _valid_domain_dns_checks(items: Any) -> bool:
    if not isinstance(items, list) or not items:
        return False
    return any(isinstance(item, dict) and (item.get("Domain") or item.get("domain")) for item in items)


def _latest_completed_run_for_tenant(tid: str) -> Optional[Dict[str, Any]]:
    return db_fetchone(
        """
        SELECT * FROM assessment_runs
        WHERE tenant_id=? AND status='completed'
        ORDER BY is_archived ASC, COALESCE(completed_at, started_at) DESC
        LIMIT 1
        """,
        (tid,),
    )


def _latest_assessment_snapshot_for_tenant(tid: str) -> Dict[str, Any]:
    run = _latest_completed_run_for_tenant(tid)
    if not run:
        return {}
    run_dir = RUNS_DIR / run["id"]
    summary_file = find_latest_summary_file(run_dir)
    report_file = find_latest_report_file(run_dir)
    assessment_json_payloads = _load_assessment_json_payloads(run_dir)
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
    teams = snapshot.get("Teams") if isinstance(snapshot, dict) else None
    assessment_users = None
    user_overview_counts = None
    license_assignments = None
    if not licenses and report_file and report_file.exists():
        licenses = _parse_license_overview_from_html(report_file)
    if not app_registrations and report_file and report_file.exists():
        app_registrations = _parse_app_registration_alerts_from_html(report_file)
    if (not _valid_domain_dns_checks(domain_dns_checks)) and report_file and report_file.exists():
        domain_dns_checks = _parse_domain_dns_checks_from_html(report_file)
    if not user_mailboxes and report_file and report_file.exists():
        user_mailboxes = _parse_user_mailboxes_from_html(report_file)
    if not teams and report_file and report_file.exists():
        teams = _parse_teams_from_html(report_file)
    if report_file and report_file.exists():
        assessment_users = _parse_assessment_users_from_html(report_file)
        user_overview_counts = _parse_user_overview_counts_from_html(report_file)
        license_assignments = _parse_license_assignments_from_html(report_file)
    users_payload = _assessment_json_payload({"assessment_json_payloads": assessment_json_payloads}, "gebruikers", "users")
    licenses_payload = _assessment_json_payload({"assessment_json_payloads": assessment_json_payloads}, "gebruikers", "licenses")
    teams_payload = _assessment_json_payload({"assessment_json_payloads": assessment_json_payloads}, "teams", "teams")
    sharepoint_sites_payload = _assessment_json_payload({"assessment_json_payloads": assessment_json_payloads}, "sharepoint", "sharepoint-sites")
    sharepoint_settings_payload = _assessment_json_payload({"assessment_json_payloads": assessment_json_payloads}, "sharepoint", "sharepoint-settings")
    onedrive_payload = _assessment_json_payload({"assessment_json_payloads": assessment_json_payloads}, "backup", "onedrive")
    ca_payload = _assessment_json_payload({"assessment_json_payloads": assessment_json_payloads}, "ca", "policies")
    apps_payload = _assessment_json_payload({"assessment_json_payloads": assessment_json_payloads}, "apps", "registrations")
    intune_summary_payload = _assessment_json_payload({"assessment_json_payloads": assessment_json_payloads}, "intune", "summary")
    intune_devices_payload = _assessment_json_payload({"assessment_json_payloads": assessment_json_payloads}, "intune", "devices")
    intune_compliance_payload = _assessment_json_payload({"assessment_json_payloads": assessment_json_payloads}, "intune", "compliance")
    intune_config_payload = _assessment_json_payload({"assessment_json_payloads": assessment_json_payloads}, "intune", "config")
    exchange_payload = _assessment_json_payload({"assessment_json_payloads": assessment_json_payloads}, "exchange", "mailboxes")
    alerts_score_payload = _assessment_json_payload({"assessment_json_payloads": assessment_json_payloads}, "alerts", "secure-score")
    alerts_audit_payload = _assessment_json_payload({"assessment_json_payloads": assessment_json_payloads}, "alerts", "audit-logs")
    identity_mfa_payload = _assessment_json_payload({"assessment_json_payloads": assessment_json_payloads}, "identity", "mfa")
    identity_admin_payload = _assessment_json_payload({"assessment_json_payloads": assessment_json_payloads}, "identity", "admin-roles")
    if not assessment_users and isinstance(users_payload, dict):
        assessment_users = users_payload.get("items") or []
    if not user_overview_counts and isinstance(users_payload, dict):
        user_overview_counts = users_payload.get("summary") or {}
    if not licenses and isinstance(licenses_payload, dict):
        licenses = licenses_payload.get("items") or []
    if not teams and isinstance(teams_payload, dict):
        teams = teams_payload.get("items") or []
    if not snapshot.get("SharePointSites") and isinstance(sharepoint_sites_payload, dict):
        snapshot["SharePointSites"] = sharepoint_sites_payload.get("items") or []
    if not snapshot.get("SharePointTenantSettings") and isinstance(sharepoint_settings_payload, dict):
        snapshot["SharePointTenantSettings"] = sharepoint_settings_payload.get("summary") or {}
    if not snapshot.get("Top5OneDriveBySize") and isinstance(onedrive_payload, dict):
        snapshot["Top5OneDriveBySize"] = onedrive_payload.get("items") or []
    if not snapshot.get("CAPolicies") and isinstance(ca_payload, dict):
        snapshot["CAPolicies"] = ca_payload.get("items") or []
    if not snapshot.get("AppRegistrations") and isinstance(apps_payload, dict):
        snapshot["AppRegistrations"] = apps_payload.get("items") or []
    if not snapshot.get("IntuneDevices") and isinstance(intune_devices_payload, dict):
        snapshot["IntuneDevices"] = intune_devices_payload.get("items") or []
    if not snapshot.get("IntuneCompliance") and isinstance(intune_compliance_payload, dict):
        snapshot["IntuneCompliance"] = intune_compliance_payload.get("items") or []
    if not snapshot.get("IntuneConfigProfiles") and isinstance(intune_config_payload, dict):
        snapshot["IntuneConfigProfiles"] = intune_config_payload.get("items") or []
    if not snapshot.get("UserMailboxes") and isinstance(exchange_payload, dict):
        snapshot["UserMailboxes"] = exchange_payload.get("items") or []
    metrics = snapshot.get("Metrics") if isinstance(snapshot, dict) else {}
    if isinstance(metrics, dict):
        if isinstance(identity_mfa_payload, dict):
            mfa_summary = identity_mfa_payload.get("summary") or {}
            if metrics.get("MfaCoveragePct") is None and mfa_summary.get("mfaCoveragePct") is not None:
                metrics["MfaCoveragePct"] = mfa_summary.get("mfaCoveragePct")
            if metrics.get("MfaMissing") is None and mfa_summary.get("usersWithoutMfa") is not None:
                metrics["MfaMissing"] = mfa_summary.get("usersWithoutMfa")
        if isinstance(ca_payload, dict):
            ca_summary = ca_payload.get("summary") or {}
            if metrics.get("CAEnabled") is None and ca_summary.get("enabled") is not None:
                metrics["CAEnabled"] = ca_summary.get("enabled")
        if isinstance(alerts_score_payload, dict):
            score_summary = alerts_score_payload.get("summary") or {}
            if metrics.get("SecureScorePct") is None and score_summary.get("percentage") is not None:
                metrics["SecureScorePct"] = score_summary.get("percentage")
    licenses = licenses or []
    app_registrations = app_registrations or []
    domain_dns_checks = domain_dns_checks or []
    user_mailboxes = user_mailboxes or []
    teams = teams or []
    license_assignments = license_assignments or {}
    for item in licenses:
        if not isinstance(item, dict):
            continue
        sku = (item.get("SkuPartNumber") or "").strip()
        assigned_users: List[Dict[str, str]] = []
        for alias in _license_key_aliases(sku):
            if alias in license_assignments:
                assigned_users = license_assignments.get(alias, [])
                break
        item["AssignedUsers"] = assigned_users
    total = sum(int(item.get("Total") or 0) for item in licenses if isinstance(item, dict))
    used = sum(int(item.get("Consumed") or 0) for item in licenses if isinstance(item, dict))
    license_type = None
    if len(licenses) == 1 and isinstance(licenses[0], dict):
        license_type = licenses[0].get("displayName") or licenses[0].get("SkuPartNumber")
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
        "assessment_teams": teams,
        "assessment_users": assessment_users or [],
        "assessment_user_counts": user_overview_counts or {},
        "assessment_license_assignments": license_assignments,
        "assessment_json_payloads": assessment_json_payloads,
        "assessment_json_identity_mfa": identity_mfa_payload or {},
        "assessment_json_identity_admin_roles": identity_admin_payload or {},
        "assessment_json_alerts_audit_logs": alerts_audit_payload or {},
    }


def _snapshot_raw(tid: str) -> Dict[str, Any]:
    """Returns the full raw snapshot dict for the latest completed assessment run."""
    run = _latest_completed_run_for_tenant(tid)
    if not run:
        return {}
    run_dir = RUNS_DIR / run["id"]
    s = find_latest_summary_file(run_dir)
    data = _safe_json_load(s) if s and s.exists() else {}
    data["_assessment_json_payloads"] = _load_assessment_json_payloads(run_dir)
    return data if isinstance(data, dict) else {}


def _snapshot_raw_metrics(tid: str) -> Dict[str, Any]:
    """Returns the Metrics dict from the latest assessment snapshot, or {}."""
    return _snapshot_raw(tid).get("Metrics") or {}


def _sharepoint_storage_to_gb(value: Any) -> float:
    """Normaliseert storagewaarden uit live bytes of snapshot-GB naar GB."""
    if value in (None, "", "—"):
        return 0.0
    try:
        number = float(value)
    except (TypeError, ValueError):
        return 0.0
    # Live Graph-data komt in bytes terug; snapshotdata meestal al in GB.
    if number > 1024 * 1024 * 1024:
        return round(number / (1024 ** 3), 2)
    return round(number, 2)


def _build_sharepoint_capacity_summary(tid: str, sites: Optional[List[Dict[str, Any]]] = None) -> Dict[str, Any]:
    """Bouwt dezelfde quota/capaciteitssamenvatting als in het HTML-rapport."""
    snap = _latest_assessment_snapshot_for_tenant(tid)
    payload = _assessment_json_payload(snap, "sharepoint", "sharepoint-sites")
    payload_summary = payload.get("summary") if isinstance(payload, dict) else {}
    payload_summary = payload_summary if isinstance(payload_summary, dict) else {}
    site_items = [item for item in (sites or []) if isinstance(item, dict)]

    total_sites = int(payload_summary.get("totalSites") or len(site_items) or 0)
    inactive_sites = int(
        payload_summary.get("inactiveSites")
        or sum(1 for item in site_items if bool(item.get("isInactive")) or str(item.get("status") or "").lower() == "inactief")
    )
    sites_with_storage = int(
        payload_summary.get("sitesWithStorage")
        or sum(1 for item in site_items if _sharepoint_storage_to_gb(item.get("storageUsed")) > 0)
    )

    total_storage_used_gb = payload_summary.get("totalStorageUsedGB")
    if total_storage_used_gb in (None, ""):
        total_storage_used_gb = sum(_sharepoint_storage_to_gb(item.get("storageUsed")) for item in site_items)
    total_storage_used_gb = round(float(total_storage_used_gb or 0), 2)

    # Sum all consumed licenses across all SKUs (matches PowerShell report formula)
    licenses_total = int(snap.get("licenses_total") or sum(
        int(lic.get("Consumed") or 0) for lic in (snap.get("Licenses") or [])
    ) or 0)
    base_storage_gb = 1024
    storage_per_license_gb = 10
    bonus_storage_gb = 0
    total_capacity_gb = round(base_storage_gb + (licenses_total * storage_per_license_gb) + bonus_storage_gb, 2)
    storage_remaining_gb = round(total_capacity_gb - total_storage_used_gb, 2)
    storage_used_pct = round((total_storage_used_gb / total_capacity_gb) * 100, 1) if total_capacity_gb > 0 else 0.0
    avg_per_site_gb = round((total_storage_used_gb / sites_with_storage), 2) if sites_with_storage > 0 else 0.0

    capacity_label = f"{base_storage_gb} GB base + {licenses_total} licenses x {storage_per_license_gb} GB"
    if bonus_storage_gb > 0:
        capacity_label += f" + {round(bonus_storage_gb, 0)} GB bonus"

    return {
        "totalSites": total_sites,
        "inactiveSites": inactive_sites,
        "sitesWithStorage": sites_with_storage,
        "totalStorageUsedGB": total_storage_used_gb,
        "totalCapacityGB": total_capacity_gb,
        "storageRemainingGB": storage_remaining_gb,
        "storageUsedPct": storage_used_pct,
        "avgStoragePerSiteGB": avg_per_site_gb,
        "storageCapacityLabel": capacity_label,
        "storageQuotaFormula": "1 TB + (licenses x 10 GB)",
        "licenseUnitsForQuota": licenses_total,
    }


def _parse_iso_dt(value: Optional[str]) -> Optional[datetime]:
    if not value:
        return None
    try:
        return datetime.fromisoformat(str(value).replace("Z", "+00:00"))
    except Exception:
        return None


def _attach_source_meta(payload: Dict[str, Any], source: str = "live", generated_at: Optional[str] = None, tenant_id: Optional[str] = None) -> Dict[str, Any]:
    item = dict(payload or {})
    item["_source"] = source
    if generated_at is None and source == "assessment_snapshot" and tenant_id:
        generated_at = _latest_assessment_snapshot_for_tenant(tenant_id).get("assessment_generated_at")
    if generated_at is None and source == "live":
        generated_at = now_iso()
    if generated_at:
        item["_generated_at"] = generated_at
    if source == "assessment_snapshot":
        dt = _parse_iso_dt(generated_at)
        if dt is not None:
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=timezone.utc)
            age = datetime.now(timezone.utc) - dt.astimezone(timezone.utc)
            item["_stale"] = age > timedelta(minutes=30)
    else:
        item["_stale"] = False
    return item


def _api_error(code: str, message: str, http_status: int = 400) -> Tuple[int, Dict[str, Any]]:
    """Gestandaardiseerde foutrespons met error_code.
    Codes: unauthorized, forbidden, not_found, validation_error,
           config_required, not_implemented, connector_unavailable,
           assessment_only, external_api_error, internal_error
    """
    return http_status, {"ok": False, "error": message, "error_code": code}


def _snapshot_as_intune_summary(tid: str) -> Optional[Dict[str, Any]]:
    """Returns Intune summary from snapshot IntuneSummary + DevicesByOS, or None."""
    snap = _latest_assessment_snapshot_for_tenant(tid)
    payload = _assessment_json_payload(snap, "intune", "summary")
    if isinstance(payload, dict):
        summary = payload.get("summary") or {}
        by_os = {}
        for entry in payload.get("items") or []:
            if not isinstance(entry, dict):
                continue
            name = _payload_value(entry, "Name", "name", default="Unknown")
            count = int(_payload_value(entry, "Count", "count", default=0) or 0)
            by_os[str(name)] = {"total": count, "compliant": 0}
        return {
            "ok": True,
            "score": round(float(summary.get("compliancePercentage") or 0)),
            "total": int(summary.get("totalDevices") or 0),
            "compliantCount": int(summary.get("compliantDevices") or 0),
            "byOs": by_os,
            "_source": "assessment_snapshot",
        }
    raw = _snapshot_raw(tid)
    summary = raw.get("IntuneSummary") or {}
    if not summary:
        metrics = raw.get("Metrics") or {}
        pct = metrics.get("IntuneCompliancePct")
        if pct is None:
            return None
        return {"ok": True, "score": round(float(pct)), "total": 0, "compliantCount": 0, "byOs": {}, "_source": "assessment_snapshot"}
    total = int(summary.get("TotalDevices") or 0)
    compliant = int(summary.get("CompliantDevices") or 0)
    score = int(summary.get("CompliancePercentage") or 0)
    by_os_raw = raw.get("IntuneDevicesByOS") or []
    by_os = {}
    for entry in by_os_raw:
        if isinstance(entry, dict):
            name = entry.get("Name") or entry.get("name") or "Unknown"
            count = int(entry.get("Count") or entry.get("count") or 0)
            by_os[name] = {"total": count, "compliant": 0}
    return {"ok": True, "score": score, "total": total, "compliantCount": compliant, "byOs": by_os, "_source": "assessment_snapshot"}


def _snapshot_as_intune_devices(tid: str) -> List[Dict[str, Any]]:
    """Returns Intune device list from snapshot."""
    snap = _latest_assessment_snapshot_for_tenant(tid)
    payload = _assessment_json_payload(snap, "intune", "devices")
    if isinstance(payload, dict):
        result = []
        for d in payload.get("items") or []:
            if not isinstance(d, dict):
                continue
            result.append({
                "id": _payload_value(d, "Id", "id", "DeviceName", "deviceName", default=""),
                "deviceName": _payload_value(d, "DeviceName", "deviceName", default=""),
                "operatingSystem": _payload_value(d, "OperatingSystem", "operatingSystem", default=""),
                "osVersion": _payload_value(d, "OsVersion", "osVersion", default=""),
                "complianceState": _payload_value(d, "ComplianceState", "complianceState", default="unknown"),
                "userPrincipalName": _payload_value(d, "UserPrincipalName", "userPrincipalName", default=""),
                "userDisplayName": _payload_value(d, "UserDisplayName", "userDisplayName", default=""),
                "lastSyncDateTime": _payload_value(d, "LastSyncDateTime", "lastSyncDateTime"),
                "enrolledDateTime": _payload_value(d, "EnrolledDateTime", "enrolledDateTime"),
                "manufacturer": _payload_value(d, "Manufacturer", "manufacturer", default=""),
                "model": _payload_value(d, "Model", "model", default=""),
            })
        return result
    raw = _snapshot_raw(tid)
    devices = raw.get("IntuneDevices") or []
    result = []
    for d in devices:
        if not isinstance(d, dict):
            continue
        result.append({
            "id": d.get("Id") or d.get("id") or "",
            "deviceName": d.get("DeviceName") or d.get("deviceName") or "",
            "operatingSystem": d.get("OperatingSystem") or d.get("operatingSystem") or "",
            "osVersion": d.get("OsVersion") or d.get("osVersion") or "",
            "complianceState": d.get("ComplianceState") or d.get("complianceState") or "unknown",
            "userPrincipalName": d.get("UserPrincipalName") or d.get("userPrincipalName") or "",
            "userDisplayName": d.get("UserDisplayName") or d.get("userDisplayName") or "",
            "lastSyncDateTime": d.get("LastSyncDateTime") or d.get("lastSyncDateTime"),
            "enrolledDateTime": d.get("EnrolledDateTime") or d.get("enrolledDateTime"),
            "manufacturer": d.get("Manufacturer") or d.get("manufacturer") or "",
            "model": d.get("Model") or d.get("model") or "",
        })
    return result


def _snapshot_as_intune_compliance(tid: str) -> List[Dict[str, Any]]:
    """Returns compliance policies from snapshot."""
    snap = _latest_assessment_snapshot_for_tenant(tid)
    payload = _assessment_json_payload(snap, "intune", "compliance")
    if isinstance(payload, dict):
        result = []
        for p in payload.get("items") or []:
            if not isinstance(p, dict):
                continue
            result.append({
                "id": _payload_value(p, "Id", "id", "DisplayName", "displayName", default=""),
                "displayName": _payload_value(p, "DisplayName", "displayName", default=""),
                "platform": _payload_value(p, "Platform", "platform", default=""),
                "createdDateTime": _payload_value(p, "CreatedDateTime", "createdDateTime"),
                "lastModifiedDateTime": _payload_value(p, "LastModifiedDateTime", "lastModifiedDateTime"),
            })
        return result
    raw = _snapshot_raw(tid)
    items = raw.get("IntuneCompliance") or []
    result = []
    for p in items:
        if not isinstance(p, dict):
            continue
        result.append({
            "id": p.get("Id") or p.get("id") or "",
            "displayName": p.get("DisplayName") or p.get("displayName") or "",
            "platform": p.get("Platform") or p.get("platform") or "",
            "createdDateTime": p.get("CreatedDateTime") or p.get("createdDateTime"),
            "lastModifiedDateTime": p.get("LastModifiedDateTime") or p.get("lastModifiedDateTime"),
        })
    return result


def _snapshot_as_intune_config(tid: str) -> List[Dict[str, Any]]:
    """Returns config profiles from snapshot."""
    snap = _latest_assessment_snapshot_for_tenant(tid)
    payload = _assessment_json_payload(snap, "intune", "config")
    if isinstance(payload, dict):
        result = []
        for p in payload.get("items") or []:
            if not isinstance(p, dict):
                continue
            plat = _payload_value(p, "Platform", "platform", default="")
            result.append({
                "id": _payload_value(p, "Id", "id", "DisplayName", "displayName", default=""),
                "displayName": _payload_value(p, "DisplayName", "displayName", default=""),
                "platform": plat,
                "platforms": plat,
                "createdDateTime": _payload_value(p, "CreatedDateTime", "createdDateTime"),
                "lastModifiedDateTime": _payload_value(p, "LastModifiedDateTime", "lastModifiedDateTime"),
                "isAssigned": False,
                "type": "legacy",
            })
        return result
    raw = _snapshot_raw(tid)
    items = raw.get("IntuneConfigProfiles") or []
    result = []
    for p in items:
        if not isinstance(p, dict):
            continue
        plat = p.get("Platform") or p.get("platform") or ""
        result.append({
            "id": p.get("Id") or p.get("id") or "",
            "displayName": p.get("DisplayName") or p.get("displayName") or "",
            "platform": plat,
            "platforms": plat,
            "createdDateTime": p.get("CreatedDateTime") or p.get("createdDateTime"),
            "lastModifiedDateTime": p.get("LastModifiedDateTime") or p.get("lastModifiedDateTime"),
            "isAssigned": False,
            "type": "legacy",
        })
    return result


def _snapshot_as_users(tid: str) -> List[Dict[str, Any]]:
    """Maps snapshot assessment users → user objects expected by gebruikers.js."""
    snap = _latest_assessment_snapshot_for_tenant(tid)
    users_payload = _assessment_json_payload(snap, "gebruikers", "users")
    licenses_payload = _assessment_json_payload(snap, "gebruikers", "licenses")
    mfa_payload = _assessment_json_payload(snap, "identity", "mfa")

    # Build MFA lookup: UPN → True (registered) / False (not registered)
    mfa_registered_upns: set = set()
    if isinstance(mfa_payload, dict):
        for item in mfa_payload.get("items") or []:
            if not isinstance(item, dict):
                continue
            if _payload_value(item, "mfaRegistered", "MfaRegistered", default=False):
                upn = str(_payload_value(item, "userPrincipalName", "UPN", default="") or "").strip().lower()
                if upn:
                    mfa_registered_upns.add(upn)

    if isinstance(users_payload, dict):
        license_map: Dict[str, List[str]] = {}
        license_sku_map: Dict[str, List[str]] = {}
        if isinstance(licenses_payload, dict):
            for lic in licenses_payload.get("items") or []:
                if not isinstance(lic, dict):
                    continue
                sku = str(_payload_value(lic, "skuPartNumber", "SkuPartNumber", default="") or "").strip()
                display_name = str(_payload_value(lic, "displayName", "DisplayName", default=sku) or sku)
                for assigned_user in lic.get("assignedUsers") or lic.get("AssignedUsers") or []:
                    if not isinstance(assigned_user, dict):
                        continue
                    upn = str(_payload_value(assigned_user, "userPrincipalName", "UserPrincipalName", default="") or "").strip()
                    if not upn:
                        continue
                    if display_name:
                        license_map.setdefault(upn, []).append(display_name)
                    if sku:
                        license_sku_map.setdefault(upn, []).append(sku)
        result = []
        for user in users_payload.get("items") or []:
            if not isinstance(user, dict):
                continue
            upn = str(_payload_value(user, "userPrincipalName", "UserPrincipalName", "mail", "Mail", default="") or "").strip()
            licenses = license_map.get(upn, [])
            sku_ids = license_sku_map.get(upn, [])
            on_prem = _payload_value(user, "onPremisesSyncEnabled", "OnPremisesSyncEnabled", default=False)
            mfa_methods = ["MFA geregistreerd (snapshot)"] if upn.lower() in mfa_registered_upns else []
            result.append({
                "id": _payload_value(user, "id", "Id", "userPrincipalName", "UserPrincipalName", default=upn),
                "displayName": _payload_value(user, "displayName", "DisplayName", default=upn),
                "userPrincipalName": upn,
                "mail": _payload_value(user, "mail", "Mail", default=upn),
                "accountEnabled": bool(_payload_value(user, "accountEnabled", "AccountEnabled", default=True)),
                "userType": _payload_value(user, "userType", "UserType", default="Member"),
                "createdDateTime": _payload_value(user, "createdDateTime", "CreatedDateTime"),
                "department": _payload_value(user, "department", "Department"),
                "jobTitle": _payload_value(user, "jobTitle", "JobTitle"),
                "officeLocation": _payload_value(user, "officeLocation", "OfficeLocation"),
                "preferredLanguage": _payload_value(user, "preferredLanguage", "PreferredLanguage"),
                "onPremisesSyncEnabled": bool(on_prem) if on_prem is not None else False,
                "licenses": licenses,
                "licenseSkuIds": sku_ids,
                "licenseCount": int(_payload_value(user, "licenseCount", "LicenseCount", default=len(licenses)) or 0),
                "mfaMethods": mfa_methods,
                "groups": [],
            })
        if result:
            return result
    licenses = snap.get("assessment_licenses") or []
    sku_to_users: Dict[str, List[Dict[str, str]]] = {}
    upn_to_licenses: Dict[str, List[str]] = {}
    upn_to_sku_ids: Dict[str, List[str]] = {}
    for lic in licenses:
        if not isinstance(lic, dict):
            continue
        sku = (lic.get("SkuPartNumber") or "").strip()
        assigned_users = [u for u in (lic.get("AssignedUsers") or []) if isinstance(u, dict)]
        if not sku:
            continue
        sku_to_users[sku] = assigned_users
        for user in assigned_users:
            upn = (user.get("UserPrincipalName") or "").strip()
            if not upn:
                continue
            upn_to_licenses.setdefault(upn, []).append(sku)
            upn_to_sku_ids.setdefault(upn, []).append(sku)
    assessment_users = snap.get("assessment_users") or []
    if isinstance(assessment_users, list) and assessment_users:
        enriched = []
        for user in assessment_users:
            if not isinstance(user, dict):
                continue
            upn = (user.get("userPrincipalName") or "").strip()
            item = dict(user)
            item["licenses"] = [get_sku_friendly_name(sku) for sku in upn_to_licenses.get(upn, [])]
            item["licenseSkuIds"] = upn_to_sku_ids.get(upn, [])
            item["licenseCount"] = len(item["licenses"])
            if "mfaMethods" not in item:
                item["mfaMethods"] = ["MFA geregistreerd (snapshot)"] if upn.lower() in mfa_registered_upns else []
            if "groups" not in item:
                item["groups"] = []
            if "onPremisesSyncEnabled" not in item:
                item["onPremisesSyncEnabled"] = False
            enriched.append(item)
        return enriched
    users = []
    for m in (snap.get("assessment_user_mailboxes") or []):
        if not isinstance(m, dict):
            continue
        upn = m.get("PrimarySmtpAddress") or ""
        users.append({
            "id": upn,
            "displayName": m.get("DisplayName") or upn,
            "userPrincipalName": upn,
            "mail": upn,
            "accountEnabled": True,
            "createdDateTime": m.get("WhenCreated"),
            "department": None,
            "jobTitle": None,
            "licenses": [get_sku_friendly_name(sku) for sku in upn_to_licenses.get(upn, [])],
            "licenseSkuIds": upn_to_sku_ids.get(upn, []),
            "licenseCount": len(upn_to_licenses.get(upn, [])),
        })
    return users


def _snapshot_as_licenses(tid: str) -> List[Dict[str, Any]]:
    """Maps snapshot Licenses → license objects expected by gebruikers.js."""
    snap = _latest_assessment_snapshot_for_tenant(tid)
    payload = _assessment_json_payload(snap, "gebruikers", "licenses")
    if isinstance(payload, dict):
        licenses = []
        for l in payload.get("items") or []:
            if not isinstance(l, dict):
                continue
            sku = str(_payload_value(l, "skuPartNumber", "SkuPartNumber", "skuId", "SkuId", default="") or "")
            assigned_users = []
            for user in l.get("assignedUsers") or l.get("AssignedUsers") or []:
                if not isinstance(user, dict):
                    continue
                assigned_users.append({
                    "displayName": _payload_value(user, "displayName", "DisplayName", default=""),
                    "userPrincipalName": _payload_value(user, "userPrincipalName", "UserPrincipalName", default=""),
                })
            licenses.append({
                "skuId": sku,
                "skuPartNumber": sku,
                "displayName": _payload_value(l, "displayName", "DisplayName", default=get_sku_friendly_name(sku)),
                "enabled": int(_payload_value(l, "total", "Total", default=0) or 0),
                "consumed": int(_payload_value(l, "consumed", "Consumed", default=0) or 0),
                "available": int(_payload_value(l, "available", "Available", default=0) or 0),
                "utilization": _payload_value(l, "utilization", "Utilization"),
                "assignedUsers": assigned_users,
            })
        if licenses:
            return licenses
    licenses = []
    for l in (snap.get("assessment_licenses") or []):
        if not isinstance(l, dict):
            continue
        sku = l.get("SkuPartNumber") or ""
        licenses.append({
            "skuId": sku,
            "skuPartNumber": sku,
            "displayName": get_sku_friendly_name(sku),
            "enabled": int(l.get("Total") or 0),
            "consumed": int(l.get("Consumed") or 0),
            "available": int(l.get("Available") or 0),
        })
    return licenses


def _normalize_user_license_payload(user: Dict[str, Any]) -> Dict[str, Any]:
    item = dict(user or {})
    raw_licenses = item.get("licenses") or []
    if isinstance(raw_licenses, list):
        item["licenses"] = [get_sku_friendly_name(str(lic)) for lic in raw_licenses]
    raw_sku = item.get("licenseSkuIds") or []
    if isinstance(raw_sku, list):
        item["licenseSkuIds"] = [str(sku) for sku in raw_sku]
    if "licenseCount" not in item:
        item["licenseCount"] = len(item.get("licenseSkuIds") or item.get("licenses") or [])
    return item


def _normalize_license_payload(license_item: Dict[str, Any]) -> Dict[str, Any]:
    item = dict(license_item or {})
    sku = str(item.get("skuPartNumber") or item.get("skuId") or "").strip()
    item["displayName"] = get_sku_friendly_name(sku)
    return item


def _snapshot_as_mailboxes(tid: str) -> List[Dict[str, Any]]:
    """Maps snapshot UserMailboxes → basic Exchange mailbox objects."""
    snap = _latest_assessment_snapshot_for_tenant(tid)
    payload = _assessment_json_payload(snap, "exchange", "mailboxes")
    if isinstance(payload, dict):
        result = []
        for m in payload.get("items") or []:
            if not isinstance(m, dict):
                continue
            smtp = _payload_value(m, "PrimarySmtpAddress", "primarySmtpAddress", "Mail", "mail", default="")
            result.append({
                "id": smtp or _payload_value(m, "DisplayName", "displayName", default=""),
                "displayName": _payload_value(m, "DisplayName", "displayName", default=smtp),
                "primarySmtpAddress": smtp,
                "recipientTypeDetails": _payload_value(m, "RecipientTypeDetails", "recipientTypeDetails", default="UserMailbox"),
                "whenCreated": _payload_value(m, "WhenCreated", "whenCreated", "CreatedDateTime", "createdDateTime"),
            })
        if result:
            return result
    result = []
    for m in (snap.get("assessment_user_mailboxes") or []):
        if not isinstance(m, dict):
            continue
        result.append({
            "id": m.get("PrimarySmtpAddress") or "",
            "displayName": m.get("DisplayName") or "",
            "primarySmtpAddress": m.get("PrimarySmtpAddress") or "",
            "recipientTypeDetails": "UserMailbox",
            "whenCreated": m.get("WhenCreated"),
        })
    return result


def _snapshot_as_mailbox_detail(tid: str, uid: str) -> Optional[Dict[str, Any]]:
    """Looks up a single mailbox from snapshot by id or email, returns detail-shape or None."""
    snap = _latest_assessment_snapshot_for_tenant(tid)
    for m in (snap.get("assessment_user_mailboxes") or []):
        if not isinstance(m, dict):
            continue
        smtp = m.get("PrimarySmtpAddress") or ""
        name = m.get("DisplayName") or ""
        if smtp.lower() == uid.lower() or name.lower() == uid.lower():
            return {
                "ok": True,
                "id": smtp,
                "displayName": name,
                "mail": smtp,
                "upn": smtp,
                "department": None,
                "jobTitle": None,
                "office": None,
                "mobile": None,
                "timezone": None,
                "language": None,
                "autoReply": {"status": "disabled"},
                "forwarding": {"enabled": False, "address": None},
                "_source": "assessment_snapshot",
            }
    return None


def _snapshot_as_ca_policies(tid: str) -> List[Dict[str, Any]]:
    """Maps snapshot CAPolicies → CA policy objects expected by ca.js."""
    snap = _latest_assessment_snapshot_for_tenant(tid)
    payload = _assessment_json_payload(snap, "ca", "policies")
    if isinstance(payload, dict):
        result = []
        for p in payload.get("items") or []:
            if not isinstance(p, dict):
                continue
            state = str(_payload_value(p, "State", "state", default="unknown") or "unknown").lower()
            result.append({
                "id": _payload_value(p, "Id", "id", default=""),
                "displayName": _payload_value(p, "DisplayName", "displayName", default=""),
                "state": state,
                "createdAt": _payload_value(p, "CreatedDateTime", "createdDateTime", "CreatedAt", "createdAt"),
                "modifiedAt": _payload_value(p, "ModifiedDateTime", "modifiedDateTime", "ModifiedAt", "modifiedAt"),
                "userScope": "—",
                "appScope": "—",
                "grantControl": "Geen",
                "sessionCtrl": "Nee",
            })
        if result:
            return result
    raw = _snapshot_raw(tid)
    items = raw.get("CAPolicies") or []
    result = []
    for p in items:
        if not isinstance(p, dict):
            continue
        result.append({
            "id": p.get("Id") or "",
            "displayName": p.get("DisplayName") or "",
            "state": p.get("State") or "unknown",
            "createdAt": p.get("CreatedAt"),
            "modifiedAt": p.get("ModifiedAt"),
            "userScope": p.get("UserScope") or "—",
            "appScope": p.get("AppScope") or "—",
            "grantControl": p.get("GrantControl") or "Geen",
            "sessionCtrl": p.get("SessionCtrl") or "Nee",
        })
    return result


def _snapshot_as_domains(tid: str) -> List[Dict[str, Any]]:
    """Maps snapshot DomainDnsChecks → basic domain list expected by domains.js."""
    raw = _snapshot_raw(tid)
    items = raw.get("DomainDnsChecks") or []
    result = []
    for d in items:
        if not isinstance(d, dict):
            continue
        domain_name = d.get("Domain") or d.get("domain")
        if not domain_name:
            continue
        result.append({
            "id": domain_name,
            "isDefault": False,
            "isVerified": True,
            "isInitial": domain_name.lower().endswith(".onmicrosoft.com"),
            "supportedServices": [],
        })
    return result


def _snapshot_as_cis_data(tid: str) -> Optional[Dict[str, Any]]:
    """Reads CIS benchmark results from assessment JSON payload."""
    snap = _latest_assessment_snapshot_for_tenant(tid)
    return _assessment_json_payload(snap, "compliance", "cis")


_ZT_SCRIPT = PLATFORM_DIR / "assessment" / "Invoke-ZeroTrustAssessment.ps1"


def _run_zerotrust_ps(tenant_id: str, action: str, output_folder: str = "") -> Dict[str, Any]:
    """Roept de Zero Trust Assessment PS wrapper aan."""
    profile = get_tenant_auth_profile(tenant_id, include_secret=True)
    ps_script = _ZT_SCRIPT.resolve()
    if not ps_script.exists():
        raise FileNotFoundError(f"ZeroTrust script niet gevonden: {ps_script}")
    pwsh = shutil.which("pwsh") or shutil.which("powershell")
    if not pwsh:
        raise RuntimeError("PowerShell niet gevonden")
    cmd = [pwsh, "-NonInteractive", "-NoProfile", "-File", str(ps_script), "-Action", action]
    if profile.get("auth_tenant_id"):
        cmd += ["-TenantId", profile["auth_tenant_id"]]
    if profile.get("auth_client_id"):
        cmd += ["-ClientId", profile["auth_client_id"]]
    if profile.get("auth_cert_thumbprint"):
        cmd += ["-CertThumbprint", profile["auth_cert_thumbprint"]]
    if output_folder:
        cmd += ["-OutputFolder", output_folder]
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=86400)
        output = (proc.stdout or "") + (proc.stderr or "")
        if "##RESULT##" in output:
            return json.loads(output.split("##RESULT##")[-1].strip().split("\n")[0])
    except subprocess.TimeoutExpired:
        return {"ok": False, "error": "Timeout: assessment duurt te lang"}
    except Exception as e:
        return {"ok": False, "error": str(e)}
    return {"ok": False, "error": "Geen output van PS script"}


def _zt_output_folder(tenant_id: str) -> str:
    """Standaard outputpad voor ZT-rapporten per tenant."""
    base = PLATFORM_DIR / "portal" / "zerotrust_reports" / tenant_id
    base.mkdir(parents=True, exist_ok=True)
    return str(base)


def _snapshot_as_hybrid_sync(tid: str) -> Optional[Dict[str, Any]]:
    """Reads Hybrid Identity sync data from assessment JSON payload."""
    snap = _latest_assessment_snapshot_for_tenant(tid)
    return _assessment_json_payload(snap, "hybrid", "sync")


def _snapshot_as_teams(tid: str) -> List[Dict[str, Any]]:
    """Maps snapshot Teams → basic team list expected by the Teams workspace."""
    snap = _latest_assessment_snapshot_for_tenant(tid)
    payload = _assessment_json_payload(snap, "teams", "teams")
    if isinstance(payload, dict):
        result = []
        for item in payload.get("items") or []:
            if not isinstance(item, dict):
                continue
            mail = _payload_value(item, "Mail", "mail", default="")
            result.append({
                "id": _payload_value(item, "Id", "id", "Mail", "mail", "DisplayName", "displayName", default=""),
                "mail": mail,
                "displayName": _payload_value(item, "DisplayName", "displayName", default=mail),
                "memberCount": int(_payload_value(item, "MemberCount", "memberCount", default=0) or 0),
                "createdAt": _payload_value(item, "CreatedDateTime", "createdDateTime", "CreatedAt", "createdAt"),
                "visibility": _payload_value(item, "Visibility", "visibility"),
                "ownerCount": int(_payload_value(item, "OwnerCount", "ownerCount", default=0) or 0),
                "isDynamic": bool(_payload_value(item, "IsDynamic", "isDynamic", default=False)),
            })
        if result:
            return result
    result = []
    for item in (snap.get("assessment_teams") or []):
        if not isinstance(item, dict):
            continue
        result.append({
            "id": item.get("id") or item.get("mail") or item.get("displayName") or "",
            "mail": item.get("mail") or "",
            "displayName": item.get("displayName") or item.get("mail") or "",
            "memberCount": int(item.get("memberCount") or 0),
            "createdAt": item.get("createdAt"),
            "visibility": item.get("visibility"),
            "ownerCount": int(item.get("ownerCount") or 0),
            "isDynamic": bool(item.get("isDynamic")) if item.get("isDynamic") is not None else False,
        })
    return result


def _snapshot_as_sharepoint_sites(tid: str) -> List[Dict[str, Any]]:
    """Maps snapshot SharePoint data → site list expected by the SharePoint workspace."""
    snap = _latest_assessment_snapshot_for_tenant(tid)
    payload = _assessment_json_payload(snap, "sharepoint", "sharepoint-sites")
    if isinstance(payload, dict):
        result = []
        for item in payload.get("items") or []:
            if not isinstance(item, dict):
                continue
            storage_gb = _payload_value(item, "StorageUsedGB", "storageUsedGB")
            storage_label = "—"
            if storage_gb not in (None, ""):
                storage_label = f"{storage_gb} GB"
            result.append({
                "id": _payload_value(item, "Id", "id", "WebUrl", "webUrl", "DisplayName", "displayName", default=""),
                "displayName": _payload_value(item, "DisplayName", "displayName", default=""),
                "webUrl": _payload_value(item, "WebUrl", "webUrl"),
                "createdAt": _payload_value(item, "CreatedDateTime", "createdDateTime"),
                "lastModified": _payload_value(item, "LastModifiedDateTime", "lastModifiedDateTime"),
                "isRootSite": bool(_payload_value(item, "IsRootSite", "isRootSite", default=False)),
                "storageUsed": storage_gb,
                "storageLabel": storage_label,
                "status": "Inactief" if bool(_payload_value(item, "IsInactive", "isInactive", default=False)) else "Actief",
            })
        if result:
            return result
    raw = _snapshot_raw(tid)
    items = raw.get("SharePointSites") or []
    result = []
    for item in items:
        if not isinstance(item, dict):
            continue
        storage_gb = item.get("StorageUsedGB")
        storage_label = "—"
        if storage_gb not in (None, ""):
            storage_label = f"{storage_gb} GB"
        result.append({
            "id": item.get("Id") or item.get("id") or item.get("WebUrl") or item.get("DisplayName") or "",
            "displayName": item.get("DisplayName") or item.get("displayName") or "",
            "webUrl": item.get("WebUrl") or item.get("webUrl"),
            "createdAt": item.get("CreatedDateTime") or item.get("createdDateTime"),
            "lastModified": item.get("LastModifiedDateTime") or item.get("lastModifiedDateTime"),
            "isRootSite": bool(item.get("IsRootSite")) if item.get("IsRootSite") is not None else False,
            "storageUsed": storage_gb,
            "storageLabel": storage_label,
            "status": "Inactief" if item.get("IsInactive") else "Actief",
        })
    if result:
        return result
    run = db_fetchone(
        "SELECT * FROM assessment_runs WHERE tenant_id=? ORDER BY is_archived ASC, COALESCE(completed_at, started_at) DESC LIMIT 1",
        (tid,),
    )
    if not run:
        return []
    report_file = find_latest_report_file(RUNS_DIR / run["id"])
    if not report_file or not report_file.exists():
        return []
    return _parse_sharepoint_sites_from_html(report_file)


def _snapshot_as_sharepoint_settings(tid: str) -> Optional[Dict[str, Any]]:
    snap = _latest_assessment_snapshot_for_tenant(tid)
    payload = _assessment_json_payload(snap, "sharepoint", "sharepoint-settings")
    if isinstance(payload, dict):
        summary = payload.get("summary") or {}
        notes = payload.get("meta", {}).get("notes") or []
        note_text = notes[0] if notes else ""
        return {
            "ok": True,
            "sharingCapability": note_text or summary.get("sharingCapability"),
            "defaultLinkPermission": _payload_value(summary, "defaultLinkPermission", "DefaultLinkPermission"),
            "defaultSharingLinkType": _payload_value(summary, "defaultSharingLinkType", "DefaultSharingLinkType"),
            "guestSharingEnabled": bool(summary.get("tenantSettingsAvailable")),
            "_source": "assessment_snapshot",
        }
    raw = _snapshot_raw(tid)
    settings = raw.get("SharePointTenantSettings")
    if isinstance(settings, dict):
        return {
            "ok": True,
            "sharingCapability": settings.get("ExternalSharing") or settings.get("sharingCapability"),
            "defaultLinkPermission": settings.get("DefaultLinkPermission") or settings.get("defaultLinkPermission"),
            "defaultSharingLinkType": settings.get("LoopDefaultSharingLinkScope") or settings.get("defaultSharingLinkType"),
            "guestSharingEnabled": settings.get("ExternalSharing", "").lower() not in {"disabled", "uitgeschakeld", "nee"},
            "_source": "assessment_snapshot",
        }
    run = db_fetchone(
        "SELECT * FROM assessment_runs WHERE tenant_id=? ORDER BY is_archived ASC, COALESCE(completed_at, started_at) DESC LIMIT 1",
        (tid,),
    )
    if not run:
        return None
    report_file = find_latest_report_file(RUNS_DIR / run["id"])
    if not report_file or not report_file.exists():
        return None
    parsed = _parse_sharepoint_settings_from_html(report_file)
    if not parsed:
        return None
    parsed["ok"] = True
    parsed["_source"] = "assessment_snapshot"
    return parsed


def _snapshot_as_sharepoint_backup(tid: str) -> Dict[str, Any]:
    sites = _snapshot_as_sharepoint_sites(tid)
    if not sites:
        return {"ok": True, "policies": [], "count": 0, "note": "Geen SharePoint assessmentdata beschikbaar.", "_source": "assessment_snapshot"}
    return {
        "ok": True,
        "policies": [{
            "id": "assessment-sharepoint",
            "displayName": "Assessment SharePoint sites",
            "status": "assessment_snapshot",
            "createdAt": _latest_assessment_snapshot_for_tenant(tid).get("assessment_generated_at"),
            "retentionPeriodInDays": 0,
            "siteCount": len(sites),
            "sites": [
                {
                    "siteId": site.get("id"),
                    "siteName": site.get("displayName"),
                    "siteUrl": site.get("webUrl"),
                    "status": site.get("status") or "assessment_snapshot",
                }
                for site in sites
            ],
        }],
        "count": 1,
        "_source": "assessment_snapshot",
        "note": "Gegevens uit laatste assessment; geen live M365 Backup-policydata.",
    }


def _snapshot_as_onedrive_backup(tid: str) -> Dict[str, Any]:
    snap = _latest_assessment_snapshot_for_tenant(tid)
    payload = _assessment_json_payload(snap, "backup", "onedrive")
    if isinstance(payload, dict):
        drives = []
        for item in payload.get("items") or []:
            if not isinstance(item, dict):
                continue
            owner_name = _payload_value(item, "OwnerDisplayName", "ownerDisplayName", "OwnerPrincipalName", "ownerPrincipalName", default="")
            drives.append({
                "driveId": _payload_value(item, "OwnerPrincipalName", "ownerPrincipalName", "Url", "url", default=""),
                "ownerName": owner_name,
                "status": "assessment_snapshot",
                "storageGB": _payload_value(item, "StorageUsedGB", "storageUsedGB", default=0) or 0,
                "modified": _payload_value(item, "LastModifiedDateTime", "lastModifiedDateTime"),
            })
        if drives:
            return {
                "ok": True,
                "policies": [{
                    "id": "assessment-onedrive",
                    "displayName": "Assessment OneDrive sites",
                    "status": "assessment_snapshot",
                    "createdAt": snap.get("assessment_generated_at"),
                    "retentionPeriodInDays": 0,
                    "driveCount": len(drives),
                    "drives": drives,
                }],
                "count": 1,
                "_source": "assessment_snapshot",
                "note": "Gegevens uit laatste assessment; geen live M365 Backup-policydata.",
            }
    raw = _snapshot_raw(tid)
    drives_raw = raw.get("Top5OneDriveBySize") or []
    drives = []
    for item in drives_raw:
        if not isinstance(item, dict):
            continue
        drives.append({
            "driveId": item.get("Owner") or item.get("owner") or "",
            "ownerName": item.get("Owner") or item.get("owner") or "",
            "status": "assessment_snapshot",
            "storageGB": item.get("StorageUsedGB") or item.get("storageGB") or 0,
            "modified": item.get("LastModifiedDateTime") or item.get("lastModifiedDateTime"),
        })
    if not drives:
        run = db_fetchone(
            "SELECT * FROM assessment_runs WHERE tenant_id=? ORDER BY is_archived ASC, COALESCE(completed_at, started_at) DESC LIMIT 1",
            (tid,),
        )
        report_file = find_latest_report_file(RUNS_DIR / run["id"]) if run else None
        if report_file and report_file.exists():
            drives = _parse_onedrive_from_html(report_file)
    if not drives:
        return {"ok": True, "policies": [], "count": 0, "note": "Geen OneDrive assessmentdata beschikbaar.", "_source": "assessment_snapshot"}
    return {
        "ok": True,
        "policies": [{
            "id": "assessment-onedrive",
            "displayName": "Assessment OneDrive sites",
            "status": "assessment_snapshot",
            "createdAt": _latest_assessment_snapshot_for_tenant(tid).get("assessment_generated_at"),
            "retentionPeriodInDays": 0,
            "driveCount": len(drives),
            "drives": drives,
        }],
        "count": 1,
        "_source": "assessment_snapshot",
        "note": "Gegevens uit laatste assessment; geen live M365 Backup-policydata.",
    }


def assessment_ui_nav(tenant_id: str) -> Dict[str, Any]:
    tenant = db_fetchone("SELECT * FROM tenants WHERE id=?", (tenant_id,))
    if not tenant:
        raise ValueError("Tenant niet gevonden")
    run = _latest_completed_run_for_tenant(tenant_id)
    snapshot = _latest_assessment_snapshot_for_tenant(tenant_id)
    json_payloads = snapshot.get("assessment_json_payloads") or {}
    users = [u for u in (snapshot.get("assessment_user_mailboxes") or []) if isinstance(u, dict)]
    domains = [d for d in (snapshot.get("assessment_domain_dns_checks") or []) if isinstance(d, dict) and (d.get("Domain") or d.get("domain"))]
    app_regs = [a for a in (snapshot.get("assessment_app_registrations") or []) if isinstance(a, dict)]
    licenses = [l for l in (snapshot.get("assessment_licenses") or []) if isinstance(l, dict)]
    if json_payloads:
        dynamic_items = [{"key": "summary", "label": "Overzicht", "count": None}]
        ordered = sorted(json_payloads.items(), key=lambda kv: _assessment_nav_sort_key(kv[0][0], kv[0][1]))
        for (section, subsection), payload in ordered:
            if not isinstance(payload, dict):
                continue
            coverage = _assessment_item_coverage(tenant_id, section, subsection)
            dynamic_items.append({
                "key": f"{section}:{subsection}",
                "label": payload.get("label") or f"{section} / {subsection}",
                "count": len(payload.get("items") or []),
                "coverage": coverage,
            })
        items = dynamic_items
    else:
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
    json_payloads = snapshot.get("assessment_json_payloads") or {}
    users = [u for u in (snapshot.get("assessment_user_mailboxes") or []) if isinstance(u, dict)]
    licenses = [l for l in (snapshot.get("assessment_licenses") or []) if isinstance(l, dict)]
    app_regs = [a for a in (snapshot.get("assessment_app_registrations") or []) if isinstance(a, dict)]
    domains = [d for d in (snapshot.get("assessment_domain_dns_checks") or []) if isinstance(d, dict) and (d.get("Domain") or d.get("domain"))]
    common = {
        "tenant_name": bundle["tenant_name"],
        "generated_at": bundle["generated_at"],
        "latest_run_id": bundle["latest_run_id"],
    }
    if section_key == "summary" and json_payloads:
        bars = []
        ordered = sorted(json_payloads.items(), key=lambda kv: _assessment_nav_sort_key(kv[0][0], kv[0][1]))
        for (_section, _subsection), payload in ordered:
            if not isinstance(payload, dict):
                continue
            bars.append({
                "label": payload.get("label") or f"{_section}:{_subsection}",
                "value": len(payload.get("items") or []),
                "max": max(len(payload.get("items") or []), 1),
            })
        return {
            **common,
            "key": "summary",
            "title": "Assessment overzicht",
            "cards": [
                {"label": "JSON onderdelen", "value": len(json_payloads), "tone": "default"},
                {"label": "Secure Score", "value": f"{round(snapshot['secure_score_percentage'])}%" if snapshot.get("secure_score_percentage") is not None else "—", "tone": "success"},
                {"label": "MFA Coverage", "value": f"{round(snapshot['mfa_coverage'])}%" if snapshot.get("mfa_coverage") is not None else "—", "tone": "success"},
                {"label": "Open Alerts", "value": bundle["critical_count"] + bundle["warning_count"], "tone": "warn"},
            ],
            "bars": bars[:12],
        }
    if ":" in section_key and json_payloads:
        section, subsection = section_key.split(":", 1)
        payload = json_payloads.get((section, subsection))
        if isinstance(payload, dict):
            columns, rows = _rows_from_json_payload(payload)
            coverage = _assessment_item_coverage(tenant_id, section, subsection)
            return {
                **common,
                "key": section_key,
                "title": payload.get("label") or f"{section} / {subsection}",
                "cards": _cards_from_json_summary(payload),
                "columns": columns,
                "rows": rows,
                "coverage": coverage,
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
    result = {"report_path": None, "snapshot_path": None, "report_filename": None, "json_manifest_path": None}
    if report:
        rel = report.relative_to(run_dir).as_posix()
        result["report_path"] = f"/reports/{run_id}/{rel}"
        result["report_filename"] = report.name
    if summary:
        rel = summary.relative_to(run_dir).as_posix()
        result["snapshot_path"] = f"/reports/{run_id}/{rel}"
    result["json_manifest_path"] = _run_json_manifest_path(run_dir)
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


def import_run_snapshots_to_db(run_id: str) -> int:
    """After a completed run, import all portal JSON payloads into m365_snapshots.

    Returns the number of snapshots written.
    """
    run = db_fetchone("SELECT tenant_id, completed_at FROM assessment_runs WHERE id=?", (run_id,))
    if not run or not run.get("tenant_id"):
        return 0
    tenant_id = run["tenant_id"]
    generated_at = run.get("completed_at") or now_iso()
    run_dir = RUNS_DIR / run_id
    payloads = _load_assessment_json_payloads(run_dir)
    if not payloads:
        return 0
    written = 0
    for (section, subsection), payload in payloads.items():
        snap_id = str(uuid.uuid4())
        # Build a lightweight summary: top-level string/int values only
        summary: Dict[str, Any] = {
            k: v for k, v in payload.items()
            if isinstance(v, (str, int, float, bool)) and k not in ("_source", "_generated_at", "_stale")
        }
        db_execute(
            "INSERT OR REPLACE INTO m365_snapshots "
            "(id, tenant_id, section, subsection, source_type, generated_at, data_json, summary_json, assessment_run_id) "
            "VALUES (?, ?, ?, ?, 'assessment', ?, ?, ?, ?)",
            (
                snap_id, tenant_id, section, subsection,
                generated_at,
                json.dumps(payload, ensure_ascii=False),
                json.dumps(summary, ensure_ascii=False),
                run_id,
            ),
        )
        written += 1
    logger.info("import_run_snapshots_to_db: %d snapshots written for run %s (tenant %s)", written, run_id, tenant_id)
    return written


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
    domain_cols = {r["name"] for r in conn.execute("PRAGMA table_info(kb_domains)").fetchall()}
    for _col, _def in [("source", "TEXT DEFAULT 'manual'"), ("spf", "TEXT"), ("dmarc", "TEXT"), ("dkim", "TEXT")]:
        if _col not in domain_cols:
            conn.execute(f"ALTER TABLE kb_domains ADD COLUMN {_col} {_def}")
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
        manual = _kb_rows(c, "SELECT * FROM kb_domains ORDER BY domain")
    snap = _latest_assessment_snapshot_for_tenant(tid) or {}
    assessment_checks = [
        ch for ch in (snap.get("assessment_domain_dns_checks") or [])
        if isinstance(ch, dict) and (ch.get("Domain") or ch.get("domain"))
    ]
    if not assessment_checks:
        return [dict(r, source=r.get("source") or "manual") for r in manual]
    manual_domains = {(r.get("domain") or "").strip().lower() for r in manual}
    # Enrich manual domains with assessment signals
    result: List[Dict[str, Any]] = []
    for row in manual:
        row = dict(row)
        row.setdefault("source", "manual")
        row_domain = (row.get("domain") or "").strip().lower()
        for ch in assessment_checks:
            if (ch.get("Domain") or ch.get("domain") or "").strip().lower() == row_domain:
                row["spf"] = str(ch.get("SPF") or ch.get("spf") or row.get("spf") or "")
                row["dmarc"] = str(ch.get("DMARC") or ch.get("dmarc") or row.get("dmarc") or "")
                row["dkim"] = str(ch.get("DKIM") or ch.get("dkim") or row.get("dkim") or "")
                break
        result.append(row)
    # Append assessment-only domains (not yet in manual list) as read-only rows
    pseudo_id = -1
    for ch in assessment_checks:
        ch_domain = (ch.get("Domain") or ch.get("domain") or "").strip().lower()
        if not ch_domain or ch_domain in manual_domains:
            continue
        result.append({
            "id": pseudo_id,
            "domain": ch.get("Domain") or ch.get("domain"),
            "domain_type": "M365",
            "registrar": None, "expiry": None, "ssl_expiry": None,
            "ssl_issuer": None, "status": "active", "auto_renew": 0,
            "nameservers": None, "notes": None,
            "source": "assessment",
            "spf": str(ch.get("SPF") or ch.get("spf") or ""),
            "dmarc": str(ch.get("DMARC") or ch.get("dmarc") or ""),
            "dkim": str(ch.get("DKIM") or ch.get("dkim") or ""),
        })
        pseudo_id -= 1
    return result


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
        if base.get("license_type"):
            base["license_type"] = get_sku_friendly_name(str(base["license_type"]))
        if assessment.get("conditional_access") is not None:
            base["conditional_access"] = 1 if assessment["conditional_access"] else 0
        base["assessment_generated_at"] = assessment.get("assessment_generated_at")
        base["assessment_report_id"] = assessment.get("assessment_report_id")
        raw_licenses = assessment.get("assessment_licenses") or []
        normalized_licenses = []
        for item in raw_licenses:
            if not isinstance(item, dict):
                continue
            license_item = dict(item)
            sku = str(license_item.get("SkuPartNumber") or license_item.get("sku_part_number") or "").strip()
            if sku:
                license_item["displayName"] = get_sku_friendly_name(sku)
            normalized_licenses.append(license_item)
        base["assessment_licenses"] = normalized_licenses
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

# ══════════════════════════════════════════════════════════════════════════════
# REMEDIATION — catalogus, uitvoering en geschiedenis
# ══════════════════════════════════════════════════════════════════════════════

REMEDIATION_CATALOG: List[Dict[str, Any]] = [
    {
        "id": "enable-security-defaults",
        "title": "Security Defaults inschakelen",
        "description": "Schakelt Microsoft Security Defaults in voor de tenant. Vereist MFA voor alle gebruikers en blokkeert legacy authenticatie. Incompatibel met bestaande Conditional Access policies.",
        "category": "identity",
        "category_label": "Identiteit & Toegang",
        "severity": "critical",
        "risk": "low",
        "risk_label": "Laag risico",
        "dry_run_supported": True,
        "params_schema": [],
        "graph_endpoint": "PATCH /policies/identitySecurityDefaultsEnforcementPolicy",
        "permissions_required": ["Policy.ReadWrite.ConditionalAccess"],
        "tags": ["mfa", "baseline", "security-defaults"],
    },
    {
        "id": "block-legacy-auth",
        "title": "Legacy authenticatie blokkeren (CA Policy)",
        "description": "Maakt een Conditional Access policy aan die SMTP, POP3, IMAP en andere legacy protocollen blokkeert voor alle gebruikers. Controleer eerst of geen kritieke systemen hiervan afhankelijk zijn.",
        "category": "identity",
        "category_label": "Identiteit & Toegang",
        "severity": "critical",
        "risk": "medium",
        "risk_label": "Middel risico",
        "dry_run_supported": True,
        "params_schema": [],
        "graph_endpoint": "POST /identity/conditionalAccess/policies",
        "permissions_required": ["Policy.ReadWrite.ConditionalAccess"],
        "tags": ["legacy-auth", "ca-policy"],
    },
    {
        "id": "require-mfa-all-users",
        "title": "MFA vereisen voor alle gebruikers (CA Policy)",
        "description": "Maakt een Conditional Access policy aan die multifactorauthenticatie verplicht stelt voor alle gebruikers bij alle cloud-apps. Zorg dat alle gebruikers MFA al hebben ingesteld voor activatie.",
        "category": "identity",
        "category_label": "Identiteit & Toegang",
        "severity": "critical",
        "risk": "medium",
        "risk_label": "Middel risico",
        "dry_run_supported": True,
        "params_schema": [],
        "graph_endpoint": "POST /identity/conditionalAccess/policies",
        "permissions_required": ["Policy.ReadWrite.ConditionalAccess"],
        "tags": ["mfa", "ca-policy", "all-users"],
    },
    {
        "id": "revoke-user-sessions",
        "title": "Alle sessies intrekken voor gebruiker",
        "description": "Forceert uitloggen van alle actieve sessies voor een opgegeven gebruiker. Gebruik bij verdachte activiteit, gecompromitteerd account of offboarding.",
        "category": "identity",
        "category_label": "Identiteit & Toegang",
        "severity": "warning",
        "risk": "low",
        "risk_label": "Laag risico",
        "dry_run_supported": True,
        "params_schema": [
            {"name": "user_upn", "label": "Gebruiker (UPN/e-mail)", "type": "text", "required": True, "placeholder": "gebruiker@bedrijf.nl"},
        ],
        "graph_endpoint": "POST /users/{id}/revokeSignInSessions",
        "permissions_required": ["User.ReadWrite.All"],
        "tags": ["sessions", "offboarding", "incident-response"],
    },
    {
        "id": "disable-user",
        "title": "Gebruikersaccount blokkeren",
        "description": "Blokkeert het opgegeven account zodat de gebruiker niet meer kan inloggen. Het account blijft intact voor auditing en mailbox-delegatie.",
        "category": "identity",
        "category_label": "Identiteit & Toegang",
        "severity": "warning",
        "risk": "medium",
        "risk_label": "Middel risico",
        "dry_run_supported": True,
        "params_schema": [
            {"name": "user_upn", "label": "Gebruiker (UPN/e-mail)", "type": "text", "required": True, "placeholder": "gebruiker@bedrijf.nl"},
        ],
        "graph_endpoint": "PATCH /users/{id}",
        "permissions_required": ["User.ReadWrite.All"],
        "tags": ["account", "offboarding", "incident-response"],
    },
    {
        "id": "enable-modern-auth",
        "title": "Modern authenticatie inschakelen (Exchange)",
        "description": "Schakelt moderne authenticatie in voor Exchange Online. Vereist voor MFA-ondersteuning in oudere Outlook-clients. Vereist Exchange Online PowerShell — zie instructies.",
        "category": "mail",
        "category_label": "E-mail & Beveiliging",
        "severity": "warning",
        "risk": "low",
        "risk_label": "Laag risico",
        "dry_run_supported": True,
        "params_schema": [],
        "graph_endpoint": "Exchange Online PowerShell",
        "permissions_required": ["Exchange.ManageAsApp"],
        "tags": ["email", "modern-auth", "exchange"],
    },
    {
        "id": "set-outbound-spam-filter",
        "title": "Uitgaand spamfilter aanscherpen",
        "description": "Configureert het uitgaande spamfilter in Exchange Online om mailmisbruik te detecteren. Vereist Exchange Online PowerShell — zie instructies.",
        "category": "mail",
        "category_label": "E-mail & Beveiliging",
        "severity": "warning",
        "risk": "low",
        "risk_label": "Laag risico",
        "dry_run_supported": True,
        "params_schema": [],
        "graph_endpoint": "Exchange Online PowerShell",
        "permissions_required": ["Exchange.ManageAsApp"],
        "tags": ["email", "spam", "exchange"],
    },
    {
        "id": "restrict-guest-invitations",
        "title": "Gastuitnodigingen beperken tot admins",
        "description": "Past het autorisatiebeleid aan zodat alleen beheerders externe gastgebruikers kunnen uitnodigen. Voorkomt dat medewerkers onbeheerd externe toegang verlenen.",
        "category": "identity",
        "category_label": "Identiteit & Toegang",
        "severity": "warning",
        "risk": "low",
        "risk_label": "Laag risico",
        "dry_run_supported": True,
        "params_schema": [],
        "graph_endpoint": "PATCH /policies/authorizationPolicy",
        "permissions_required": ["Policy.ReadWrite.Authorization"],
        "tags": ["guests", "external-access", "governance"],
    },
    {
        "id": "enable-sspr",
        "title": "Self-Service Password Reset (SSPR) inschakelen",
        "description": "Schakelt Self-Service Password Reset in voor alle gebruikers via het authenticatiemethodenbeleid. Vermindert helpdeskbelasting en geeft gebruikers controle over hun eigen wachtwoord.",
        "category": "identity",
        "category_label": "Identiteit & Toegang",
        "severity": "info",
        "risk": "low",
        "risk_label": "Laag risico",
        "dry_run_supported": True,
        "params_schema": [],
        "graph_endpoint": "PATCH /policies/authenticationMethodsPolicy",
        "permissions_required": ["Policy.ReadWrite.AuthenticationMethod"],
        "tags": ["sspr", "password-reset", "self-service"],
    },
    {
        "id": "require-mfa-admins",
        "title": "MFA vereisen voor beheerders (CA Policy)",
        "description": "Maakt een Conditional Access policy aan die MFA verplicht stelt voor alle gebruikers met een beheerdersrol. Minder ingrijpend dan MFA voor alle gebruikers — ideale eerste stap.",
        "category": "identity",
        "category_label": "Identiteit & Toegang",
        "severity": "critical",
        "risk": "low",
        "risk_label": "Laag risico",
        "dry_run_supported": True,
        "params_schema": [],
        "graph_endpoint": "POST /identity/conditionalAccess/policies",
        "permissions_required": ["Policy.ReadWrite.ConditionalAccess"],
        "tags": ["mfa", "admins", "ca-policy", "privileged"],
    },
]

_REMEDIATION_BY_ID: Dict[str, Dict[str, Any]] = {r["id"]: r for r in REMEDIATION_CATALOG}


def get_remediation_catalog(category: Optional[str] = None) -> List[Dict[str, Any]]:
    if category:
        return [r for r in REMEDIATION_CATALOG if r.get("category") == category]
    return list(REMEDIATION_CATALOG)


def list_remediation_history(tenant_id: str, limit: int = 100) -> List[Dict[str, Any]]:
    rows = db_fetchall(
        "SELECT * FROM remediation_history WHERE tenant_id=? ORDER BY executed_at DESC LIMIT ?",
        (tenant_id, limit),
    )
    for r in rows:
        try:
            r["result"] = json.loads(r.get("result_json") or "{}")
        except Exception:
            r["result"] = {}
    return rows


def execute_remediation(
    tenant_id: str,
    remediation_id: str,
    params: Dict[str, Any],
    dry_run: bool,
    executed_by: str,
) -> Dict[str, Any]:
    """
    Voert een remediation uit via PowerShell/Graph API.
    Logt het resultaat in remediation_history.
    """
    remediation = _REMEDIATION_BY_ID.get(remediation_id)
    if not remediation:
        raise ValueError(f"Onbekende remediation-ID: {remediation_id}")
    if dry_run and not remediation.get("dry_run_supported"):
        raise ValueError(f"Dry-run wordt niet ondersteund voor: {remediation_id}")

    tenant = db_fetchone("SELECT * FROM tenants WHERE id=?", (tenant_id,))
    if not tenant:
        raise ValueError("Tenant niet gevonden")

    tenant_guid = (tenant.get("tenant_guid") or "").strip()
    if not tenant_guid:
        raise ValueError(
            "Tenant GUID niet geconfigureerd. "
            "Vul de Tenant GUID in bij Admin > Tenants voordat je remediations uitvoert."
        )

    profile = get_tenant_auth_profile(tenant_id, include_secret=True)
    cfg = load_config()
    client_id   = (profile.get("auth_client_id") or cfg.get("auth_client_id") or "").strip()
    cert_thumb  = (profile.get("auth_cert_thumbprint") or cfg.get("auth_cert_thumbprint") or "").strip()
    client_sec  = (profile.get("auth_client_secret") or cfg.get("auth_client_secret") or "").strip()

    if not client_id:
        raise ValueError(
            "App-registratie (Client ID) niet geconfigureerd. "
            "Stel dit in via Admin > Tenant-instellingen."
        )

    ps_script = (PLATFORM_DIR / "assessment" / "Invoke-DenjoyRemediation.ps1").resolve()
    if not ps_script.exists():
        raise FileNotFoundError(f"Remediation-script niet gevonden: {ps_script}")

    pwsh = shutil.which("pwsh") or shutil.which("powershell")
    if not pwsh:
        raise RuntimeError("PowerShell niet gevonden op dit systeem.")

    params_json_str = json.dumps(params, ensure_ascii=False)

    cmd = [
        pwsh, "-NoLogo", "-NoProfile", "-NonInteractive",
        "-File", str(ps_script),
        "-RemediationId", remediation_id,
        "-TenantId", tenant_guid,
        "-ClientId", client_id,
        "-ParamsJson", params_json_str,
    ]
    if cert_thumb:
        cmd += ["-CertThumbprint", cert_thumb]
    if dry_run:
        cmd.append("-DryRun")

    env = os.environ.copy()
    if client_sec and not cert_thumb:
        env["M365_CLIENT_SECRET"] = client_sec

    result_data: Dict[str, Any] = {}
    error_message: Optional[str] = None
    status = "success"

    try:
        proc = subprocess.run(
            cmd,
            env=env,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=120,
        )
        output = (proc.stdout or "").strip()
        if proc.returncode != 0:
            status = "failed"
            error_message = output[-2000:] if output else f"PowerShell exit code {proc.returncode}"
        else:
            marker = "##RESULT##"
            if marker in output:
                json_part = output[output.rfind(marker) + len(marker):].strip()
                try:
                    result_data = json.loads(json_part)
                except Exception:
                    result_data = {"raw_output": output[-500:]}
            else:
                result_data = {"output": output[-500:]}
    except subprocess.TimeoutExpired:
        status = "failed"
        error_message = "Remediation timed out (120s)"
    except Exception as exc:
        status = "failed"
        error_message = str(exc)

    final_status = status
    if dry_run and status == "success":
        final_status = "dry_run"

    history_id = str(uuid.uuid4())
    db_execute(
        """
        INSERT INTO remediation_history
        (id, tenant_id, remediation_id, title, executed_by, executed_at,
         status, dry_run, params_json, result_json, error_message)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            history_id, tenant_id, remediation_id, remediation["title"],
            executed_by, now_iso(), final_status, 1 if dry_run else 0,
            params_json_str,
            json.dumps(result_data, ensure_ascii=False) if result_data else None,
            error_message,
        ),
    )

    db_audit(
        executed_by, "", "remediation_executed",
        "tenant", tenant_id,
        f"remediation_id={remediation_id} dry_run={dry_run} status={final_status}",
    )

    ok = (status == "success")
    msg = (result_data.get("message") if isinstance(result_data, dict) else None) \
          or (error_message if not ok else f"{remediation['title']} uitgevoerd.")

    return {
        "ok": ok,
        "message": msg,
        "history_id": history_id,
        "result": result_data if ok else {},
    }


# ══════════════════════════════════════════════════════════════════════════════
# USER MANAGEMENT — gebruikers beheer via Graph API (Fase 2)
# ══════════════════════════════════════════════════════════════════════════════

_USER_MGMT_SCRIPT = PLATFORM_DIR / "assessment" / "Invoke-DenjoyUserManagement.ps1"


def _run_user_mgmt(
    tenant_id: str,
    action: str,
    params: Dict[str, Any],
    dry_run: bool = False,
    executed_by: str = "admin",
) -> Dict[str, Any]:
    """Voert een user-management actie uit via PowerShell en logt het resultaat."""
    tenant = db_fetchone("SELECT * FROM tenants WHERE id=?", (tenant_id,))
    if not tenant:
        raise ValueError("Tenant niet gevonden")

    tenant_guid = (tenant.get("tenant_guid") or "").strip()
    if not tenant_guid:
        raise ValueError(
            "Tenant GUID niet geconfigureerd. "
            "Vul de Tenant GUID in bij Admin > Tenants."
        )

    profile = get_tenant_auth_profile(tenant_id, include_secret=True)
    cfg = load_config()
    client_id  = (profile.get("auth_client_id") or cfg.get("auth_client_id") or "").strip()
    cert_thumb = (profile.get("auth_cert_thumbprint") or cfg.get("auth_cert_thumbprint") or "").strip()
    client_sec = (profile.get("auth_client_secret") or cfg.get("auth_client_secret") or "").strip()

    if not client_id:
        raise ValueError(
            "App-registratie (Client ID) niet geconfigureerd. "
            "Stel dit in via Admin > Tenant-instellingen."
        )

    ps_script = _USER_MGMT_SCRIPT.resolve()
    if not ps_script.exists():
        raise FileNotFoundError(f"User-management script niet gevonden: {ps_script}")

    pwsh = shutil.which("pwsh") or shutil.which("powershell")
    if not pwsh:
        raise RuntimeError("PowerShell niet gevonden op dit systeem.")

    params_json_str = json.dumps(params, ensure_ascii=False)

    cmd = [
        pwsh, "-NoLogo", "-NoProfile", "-NonInteractive",
        "-File", str(ps_script),
        "-Action", action,
        "-TenantId", tenant_guid,
        "-ClientId", client_id,
        "-ParamsJson", params_json_str,
    ]
    if cert_thumb:
        cmd += ["-CertThumbprint", cert_thumb]
    if client_sec and not cert_thumb:
        cmd += ["-ClientSecret", client_sec]
    if dry_run:
        cmd.append("-DryRun")

    result_data: Dict[str, Any] = {}
    error_message: Optional[str] = None
    status = "success"

    try:
        proc = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=120,
        )
        output = (proc.stdout or "").strip()
        if proc.returncode != 0:
            status = "failed"
            error_message = output[-2000:] if output else f"PowerShell exit code {proc.returncode}"
        else:
            marker = "##RESULT##"
            if marker in output:
                json_part = output[output.rfind(marker) + len(marker):].strip()
                try:
                    result_data = json.loads(json_part)
                except Exception:
                    result_data = {"raw_output": output[-500:]}
            else:
                result_data = {"output": output[-500:]}
    except subprocess.TimeoutExpired:
        status = "failed"
        error_message = "Actie timed out (120s)"
    except Exception as exc:
        status = "failed"
        error_message = str(exc)

    # Schrijf/muteer-acties loggen in provisioning_history
    if action in ("create-user", "offboard-user"):
        final_status = "dry_run" if (dry_run and status == "success") else status
        target_upn = params.get("userPrincipalName") or params.get("user_id") or ""
        target_display = params.get("displayName") or params.get("display_name") or target_upn
        db_execute(
            """
            INSERT INTO provisioning_history
            (id, tenant_id, action, target_upn, target_display_name, executed_by, executed_at,
             status, dry_run, params_json, result_json, error_message)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                str(uuid.uuid4()), tenant_id, action, target_upn, target_display,
                executed_by, now_iso(), final_status, 1 if dry_run else 0,
                params_json_str,
                json.dumps(result_data, ensure_ascii=False) if result_data else None,
                error_message,
            ),
        )
        db_audit(
            executed_by, "", f"user_mgmt_{action}",
            "tenant", tenant_id,
            f"action={action} target={target_upn} dry_run={dry_run} status={final_status}",
        )

    ok = (status == "success")
    return {
        "ok": ok,
        "result": result_data if ok else {},
        "error": error_message if not ok else None,
    }


def list_provisioning_history(tenant_id: str, limit: int = 100) -> List[Dict[str, Any]]:
    rows = db_fetchall(
        """SELECT * FROM provisioning_history
           WHERE tenant_id=? ORDER BY executed_at DESC LIMIT ?""",
        (tenant_id, limit),
    )
    for r in rows:
        try:
            r["result"] = json.loads(r.get("result_json") or "{}")
        except Exception:
            r["result"] = {}
    return rows


# ══════════════════════════════════════════════════════════════════════════════
# BASELINE & GOLD TENANT — Desired State Engine (Fase 3)
# ══════════════════════════════════════════════════════════════════════════════

_BASELINE_SCRIPT = PLATFORM_DIR / "assessment" / "Invoke-DenjoyBaseline.ps1"


def _run_baseline_ps(
    tenant_id: str,
    action: str,
    params: Dict[str, Any],
    dry_run: bool = False,
) -> Dict[str, Any]:
    """Voert een baseline-actie uit via PowerShell."""
    tenant = db_fetchone("SELECT * FROM tenants WHERE id=?", (tenant_id,))
    if not tenant:
        raise ValueError("Tenant niet gevonden")
    tenant_guid = (tenant.get("tenant_guid") or "").strip()
    if not tenant_guid:
        raise ValueError("Tenant GUID niet geconfigureerd.")

    profile = get_tenant_auth_profile(tenant_id, include_secret=True)
    cfg = load_config()
    client_id  = (profile.get("auth_client_id") or cfg.get("auth_client_id") or "").strip()
    cert_thumb = (profile.get("auth_cert_thumbprint") or cfg.get("auth_cert_thumbprint") or "").strip()
    client_sec = (profile.get("auth_client_secret") or cfg.get("auth_client_secret") or "").strip()

    if not client_id:
        raise ValueError("App-registratie (Client ID) niet geconfigureerd.")

    ps_script = _BASELINE_SCRIPT.resolve()
    if not ps_script.exists():
        raise FileNotFoundError(f"Baseline-script niet gevonden: {ps_script}")

    pwsh = shutil.which("pwsh") or shutil.which("powershell")
    if not pwsh:
        raise RuntimeError("PowerShell niet gevonden.")

    params_json_str = json.dumps(params, ensure_ascii=False)
    cmd = [
        pwsh, "-NoLogo", "-NoProfile", "-NonInteractive",
        "-File", str(ps_script),
        "-Action", action,
        "-TenantId", tenant_guid,
        "-ClientId", client_id,
        "-ParamsJson", params_json_str,
    ]
    if cert_thumb:
        cmd += ["-CertThumbprint", cert_thumb]
    if client_sec and not cert_thumb:
        cmd += ["-ClientSecret", client_sec]
    if dry_run:
        cmd.append("-DryRun")

    result_data: Dict[str, Any] = {}
    error_message: Optional[str] = None
    status = "success"

    try:
        proc = subprocess.run(
            cmd,
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding="utf-8", errors="replace", timeout=180,
        )
        output = (proc.stdout or "").strip()
        if proc.returncode != 0:
            status = "failed"
            error_message = output[-2000:] if output else f"Exit code {proc.returncode}"
        else:
            marker = "##RESULT##"
            if marker in output:
                json_part = output[output.rfind(marker) + len(marker):].strip()
                try:
                    result_data = json.loads(json_part)
                except Exception:
                    result_data = {"raw_output": output[-500:]}
            else:
                result_data = {"output": output[-500:]}
    except subprocess.TimeoutExpired:
        status = "failed"
        error_message = "Baseline actie timed out (180s)"
    except Exception as exc:
        status = "failed"
        error_message = str(exc)

    return {"ok": status == "success", "result": result_data, "error": error_message}


# ── CRUD voor baselines ───────────────────────────────────────────────────────

def list_baselines() -> List[Dict[str, Any]]:
    rows = db_fetchall("SELECT * FROM baselines ORDER BY created_at DESC")
    for r in rows:
        try:
            cfg = json.loads(r.get("config_json") or "{}")
            cats = list(cfg.get("categories", {}).keys())
            r["categories"] = cats
            r["category_count"] = len(cats)
        except Exception:
            r["categories"] = []
            r["category_count"] = 0
        r.pop("config_json", None)   # Niet meesturen in lijstoverzicht
    return rows


def get_baseline(baseline_id: str) -> Optional[Dict[str, Any]]:
    row = db_fetchone("SELECT * FROM baselines WHERE id=?", (baseline_id,))
    if not row:
        return None
    try:
        row["config"] = json.loads(row.get("config_json") or "{}")
    except Exception:
        row["config"] = {}
    return row


def create_baseline(
    name: str,
    description: str,
    config: Dict[str, Any],
    source_tenant_id: Optional[str],
    source_tenant_name: Optional[str],
    created_by: str,
) -> Dict[str, Any]:
    if not name.strip():
        raise ValueError("Naam is verplicht")
    bid = str(uuid.uuid4())
    now = now_iso()
    db_execute(
        """INSERT INTO baselines
           (id, name, description, source_tenant_id, source_tenant_name,
            config_json, created_by, created_at, updated_at)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
        (bid, name.strip(), description or "", source_tenant_id, source_tenant_name,
         json.dumps(config, ensure_ascii=False), created_by, now, now),
    )
    return get_baseline(bid)


def update_baseline(baseline_id: str, payload: Dict[str, Any]) -> Dict[str, Any]:
    row = db_fetchone("SELECT id FROM baselines WHERE id=?", (baseline_id,))
    if not row:
        raise ValueError("Baseline niet gevonden")
    fields, vals = [], []
    if "name" in payload:
        fields.append("name=?"); vals.append(payload["name"].strip())
    if "description" in payload:
        fields.append("description=?"); vals.append(payload.get("description") or "")
    if "config" in payload:
        fields.append("config_json=?"); vals.append(json.dumps(payload["config"], ensure_ascii=False))
    if not fields:
        return get_baseline(baseline_id)
    fields.append("updated_at=?"); vals.append(now_iso())
    vals.append(baseline_id)
    db_execute(f"UPDATE baselines SET {', '.join(fields)} WHERE id=?", tuple(vals))
    return get_baseline(baseline_id)


def delete_baseline(baseline_id: str) -> Dict[str, Any]:
    row = db_fetchone("SELECT id FROM baselines WHERE id=?", (baseline_id,))
    if not row:
        raise ValueError("Baseline niet gevonden")
    db_execute("DELETE FROM baseline_assignments WHERE baseline_id=?", (baseline_id,))
    db_execute("DELETE FROM baseline_history WHERE baseline_id=?", (baseline_id,))
    db_execute("DELETE FROM baselines WHERE id=?", (baseline_id,))
    return {"ok": True}


# ── Assignments ───────────────────────────────────────────────────────────────

def list_assignments(baseline_id: Optional[str] = None, tenant_id: Optional[str] = None) -> List[Dict[str, Any]]:
    if baseline_id:
        rows = db_fetchall(
            """SELECT ba.*, b.name as baseline_name, t.customer_name as tenant_name
               FROM baseline_assignments ba
               JOIN baselines b ON ba.baseline_id = b.id
               JOIN tenants t ON ba.tenant_id = t.id
               WHERE ba.baseline_id=? ORDER BY ba.assigned_at DESC""",
            (baseline_id,)
        )
    elif tenant_id:
        rows = db_fetchall(
            """SELECT ba.*, b.name as baseline_name, t.customer_name as tenant_name
               FROM baseline_assignments ba
               JOIN baselines b ON ba.baseline_id = b.id
               JOIN tenants t ON ba.tenant_id = t.id
               WHERE ba.tenant_id=? ORDER BY ba.assigned_at DESC""",
            (tenant_id,)
        )
    else:
        rows = db_fetchall(
            """SELECT ba.*, b.name as baseline_name, t.customer_name as tenant_name
               FROM baseline_assignments ba
               JOIN baselines b ON ba.baseline_id = b.id
               JOIN tenants t ON ba.tenant_id = t.id
               ORDER BY ba.assigned_at DESC"""
        )
    for r in rows:
        try:
            r["compliance"] = json.loads(r.get("compliance_json") or "{}")
        except Exception:
            r["compliance"] = {}
    return rows


def assign_baseline(baseline_id: str, tenant_id: str, assigned_by: str) -> Dict[str, Any]:
    if not db_fetchone("SELECT id FROM baselines WHERE id=?", (baseline_id,)):
        raise ValueError("Baseline niet gevonden")
    if not db_fetchone("SELECT id FROM tenants WHERE id=?", (tenant_id,)):
        raise ValueError("Tenant niet gevonden")
    existing = db_fetchone(
        "SELECT id FROM baseline_assignments WHERE baseline_id=? AND tenant_id=?",
        (baseline_id, tenant_id)
    )
    if existing:
        raise ValueError("Baseline is al gekoppeld aan deze tenant")
    aid = str(uuid.uuid4())
    db_execute(
        """INSERT INTO baseline_assignments
           (id, baseline_id, tenant_id, assigned_by, assigned_at, status)
           VALUES (?, ?, ?, ?, ?, 'assigned')""",
        (aid, baseline_id, tenant_id, assigned_by, now_iso()),
    )
    return db_fetchone("SELECT * FROM baseline_assignments WHERE id=?", (aid,))


def unassign_baseline(baseline_id: str, tenant_id: str) -> Dict[str, Any]:
    db_execute(
        "DELETE FROM baseline_assignments WHERE baseline_id=? AND tenant_id=?",
        (baseline_id, tenant_id)
    )
    return {"ok": True}


def check_baseline_compliance(baseline_id: str, tenant_id: str, executed_by: str) -> Dict[str, Any]:
    baseline = get_baseline(baseline_id)
    if not baseline:
        raise ValueError("Baseline niet gevonden")

    config = baseline.get("config") or {}
    result = _run_baseline_ps(tenant_id, "compare-baseline", {"baseline_json": json.dumps(config, ensure_ascii=False)})

    compliance_data = result.get("result", {}) if result["ok"] else {}
    score = compliance_data.get("score", 0) if result["ok"] else 0
    status = "compliant" if score == 100 else ("non_compliant" if score < 80 else "partial")

    now = now_iso()
    db_execute(
        """UPDATE baseline_assignments
           SET last_checked_at=?, compliance_score=?, compliance_json=?, status=?
           WHERE baseline_id=? AND tenant_id=?""",
        (now, score, json.dumps(compliance_data, ensure_ascii=False), status, baseline_id, tenant_id),
    )
    db_execute(
        """INSERT INTO baseline_history
           (id, baseline_id, tenant_id, action, executed_by, executed_at, status, dry_run, result_json, error_message)
           VALUES (?, ?, ?, 'check', ?, ?, ?, 0, ?, ?)""",
        (str(uuid.uuid4()), baseline_id, tenant_id, executed_by, now,
         "success" if result["ok"] else "failed",
         json.dumps(compliance_data, ensure_ascii=False),
         result.get("error")),
    )
    db_audit(executed_by, "", "baseline_check", "tenant", tenant_id,
             f"baseline_id={baseline_id} score={score}")
    return {"ok": result["ok"], "score": score, "status": status, "compliance": compliance_data, "error": result.get("error")}


def apply_baseline_to_tenant(baseline_id: str, tenant_id: str, dry_run: bool, executed_by: str) -> Dict[str, Any]:
    baseline = get_baseline(baseline_id)
    if not baseline:
        raise ValueError("Baseline niet gevonden")

    config = baseline.get("config") or {}
    result = _run_baseline_ps(tenant_id, "apply-baseline", {"baseline_json": json.dumps(config, ensure_ascii=False)}, dry_run)

    result_data = result.get("result", {})
    final_status = "dry_run" if dry_run else ("success" if result["ok"] else "failed")

    now = now_iso()
    if not dry_run and result["ok"]:
        db_execute(
            "UPDATE baseline_assignments SET last_applied_at=?, status='applied' WHERE baseline_id=? AND tenant_id=?",
            (now, baseline_id, tenant_id),
        )
    db_execute(
        """INSERT INTO baseline_history
           (id, baseline_id, tenant_id, action, executed_by, executed_at, status, dry_run, result_json, error_message)
           VALUES (?, ?, ?, 'apply', ?, ?, ?, ?, ?, ?)""",
        (str(uuid.uuid4()), baseline_id, tenant_id, executed_by, now,
         final_status, 1 if dry_run else 0,
         json.dumps(result_data, ensure_ascii=False),
         result.get("error")),
    )
    db_audit(executed_by, "", "baseline_apply", "tenant", tenant_id,
             f"baseline_id={baseline_id} dry_run={dry_run} status={final_status}")
    return {"ok": result["ok"], "result": result_data, "error": result.get("error")}


def list_baseline_history(baseline_id: Optional[str] = None, tenant_id: Optional[str] = None, limit: int = 100) -> List[Dict[str, Any]]:
    if baseline_id and tenant_id:
        rows = db_fetchall(
            "SELECT * FROM baseline_history WHERE baseline_id=? AND tenant_id=? ORDER BY executed_at DESC LIMIT ?",
            (baseline_id, tenant_id, limit)
        )
    elif baseline_id:
        rows = db_fetchall(
            "SELECT * FROM baseline_history WHERE baseline_id=? ORDER BY executed_at DESC LIMIT ?",
            (baseline_id, limit)
        )
    else:
        rows = db_fetchall(
            "SELECT * FROM baseline_history ORDER BY executed_at DESC LIMIT ?",
            (limit,)
        )
    for r in rows:
        try:
            r["result"] = json.loads(r.get("result_json") or "{}")
        except Exception:
            r["result"] = {}
    return rows


# ── Intune / Device Management (Fase 4) ──────────────────────────────────────

_INTUNE_SCRIPT = PLATFORM_DIR / "assessment" / "Invoke-DenjoyIntune.ps1"


def _run_intune_ps(tenant_id: str, action: str, params: Dict[str, Any], dry_run: bool = False, executed_by: str = "system") -> Dict[str, Any]:
    """Voer een Intune PS-actie uit en log naar intune_scan_history."""
    profile = get_tenant_auth_profile(tenant_id, include_secret=True)
    ps_script = _INTUNE_SCRIPT.resolve()
    if not ps_script.exists():
        raise FileNotFoundError(f"Intune script niet gevonden: {ps_script}")

    cmd = [
        "pwsh", "-NonInteractive", "-NoProfile", "-File", str(ps_script),
        "-Action", action,
        "-TenantId",  profile["auth_tenant_id"],
        "-ClientId",  profile["auth_client_id"],
        "-ParamsJson", json.dumps(params),
    ]
    if profile.get("auth_cert_thumbprint"):
        cmd += ["-CertThumbprint", profile["auth_cert_thumbprint"]]
    elif profile.get("auth_client_secret"):
        cmd += ["-ClientSecret", profile["auth_client_secret"]]
    if dry_run:
        cmd.append("-DryRun")

    proc = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
    output = proc.stdout + proc.stderr
    logger.info("[Intune] action=%s tenant=%s dry_run=%s exit=%s", action, tenant_id, dry_run, proc.returncode)

    result: Dict[str, Any] = {}
    if "##RESULT##" in output:
        try:
            result = json.loads(output.split("##RESULT##")[-1].strip().split("\n")[0])
        except Exception:
            result = {"ok": False, "error": "Kon resultaat niet parsen"}
    else:
        result = {"ok": False, "error": output[-500:] if output else "Geen output"}

    # Log in history
    final_status = "dry_run" if dry_run else ("success" if result.get("ok") else "failed")
    err_msg = result.get("error") if not result.get("ok") else None
    db_execute(
        "INSERT INTO intune_scan_history (id,tenant_id,action,executed_by,executed_at,status,dry_run,result_json,error_message) "
        "VALUES (?,?,?,?,?,?,?,?,?)",
        (str(uuid.uuid4()), tenant_id, action, executed_by, now_iso(), final_status, int(dry_run),
         json.dumps(result), err_msg)
    )
    return result


def list_intune_history(tenant_id: str, limit: int = 50) -> List[Dict[str, Any]]:
    return db_fetchall(
        "SELECT * FROM intune_scan_history WHERE tenant_id=? ORDER BY executed_at DESC LIMIT ?",
        (tenant_id, limit)
    )


# ── Backup Module (Fase 5) ────────────────────────────────────────────────────

_BACKUP_SCRIPT = PLATFORM_DIR / "assessment" / "Invoke-DenjoyBackup.ps1"


def _run_backup_ps(tenant_id: str, action: str, executed_by: str = "system") -> Dict[str, Any]:
    """Voer een Backup PS-actie uit en log naar backup_history."""
    profile = get_tenant_auth_profile(tenant_id, include_secret=True)
    ps_script = _BACKUP_SCRIPT.resolve()
    if not ps_script.exists():
        raise FileNotFoundError(f"Backup script niet gevonden: {ps_script}")

    cmd = [
        "pwsh", "-NonInteractive", "-NoProfile", "-File", str(ps_script),
        "-Action", action,
        "-TenantId",  profile["auth_tenant_id"],
        "-ClientId",  profile["auth_client_id"],
        "-ParamsJson", "{}",
    ]
    if profile.get("auth_cert_thumbprint"):
        cmd += ["-CertThumbprint", profile["auth_cert_thumbprint"]]
    elif profile.get("auth_client_secret"):
        cmd += ["-ClientSecret", profile["auth_client_secret"]]

    proc = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
    output = proc.stdout + proc.stderr
    logger.info("[Backup] action=%s tenant=%s exit=%s", action, tenant_id, proc.returncode)

    result: Dict[str, Any] = {}
    if "##RESULT##" in output:
        try:
            result = json.loads(output.split("##RESULT##")[-1].strip().split("\n")[0])
        except Exception:
            result = {"ok": False, "error": "Kon resultaat niet parsen"}
    else:
        result = {"ok": False, "error": output[-500:] if output else "Geen output"}

    final_status = "success" if result.get("ok") else "failed"
    err_msg = result.get("error") if not result.get("ok") else None
    db_execute(
        "INSERT INTO backup_history (id,tenant_id,action,executed_by,executed_at,status,result_json,error_message) "
        "VALUES (?,?,?,?,?,?,?,?)",
        (str(uuid.uuid4()), tenant_id, action, executed_by, now_iso(), final_status,
         json.dumps(result), err_msg)
    )
    return result


def list_backup_history(tenant_id: str, limit: int = 50) -> List[Dict[str, Any]]:
    return db_fetchall(
        "SELECT * FROM backup_history WHERE tenant_id=? ORDER BY executed_at DESC LIMIT ?",
        (tenant_id, limit)
    )


# ── Conditional Access (Fase 6) ───────────────────────────────────────────────

_CA_SCRIPT = PLATFORM_DIR / "assessment" / "Invoke-DenjoyCa.ps1"


def _run_ca_ps(tenant_id: str, action: str, params: Dict[str, Any], dry_run: bool = False, executed_by: str = "system") -> Dict[str, Any]:
    profile = get_tenant_auth_profile(tenant_id, include_secret=True)
    ps_script = _CA_SCRIPT.resolve()
    if not ps_script.exists():
        raise FileNotFoundError(f"CA script niet gevonden: {ps_script}")
    cmd = [
        "pwsh", "-NonInteractive", "-NoProfile", "-File", str(ps_script),
        "-Action", action,
        "-TenantId", profile["auth_tenant_id"],
        "-ClientId", profile["auth_client_id"],
        "-ParamsJson", json.dumps(params),
    ]
    if profile.get("auth_cert_thumbprint"):
        cmd += ["-CertThumbprint", profile["auth_cert_thumbprint"]]
    elif profile.get("auth_client_secret"):
        cmd += ["-ClientSecret", profile["auth_client_secret"]]
    if dry_run:
        cmd.append("-DryRun")
    proc = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
    output = proc.stdout + proc.stderr
    logger.info("[CA] action=%s tenant=%s dry_run=%s exit=%s", action, tenant_id, dry_run, proc.returncode)
    result: Dict[str, Any] = {}
    if "##RESULT##" in output:
        try:
            result = json.loads(output.split("##RESULT##")[-1].strip().split("\n")[0])
            if action == "list-policies" and not dry_run:
                threading.Thread(target=_persist_live_findings, args=(tenant_id, "ca", action, result), daemon=True).start()
        except Exception:
            result = {"ok": False, "error": "Parse fout"}
    else:
        result = {"ok": False, "error": output[-500:] if output else "Geen output"}
    final_status = "dry_run" if dry_run else ("success" if result.get("ok") else "failed")
    policy_id = params.get("policy_id")
    db_execute(
        "INSERT INTO ca_history (id,tenant_id,action,policy_id,executed_by,executed_at,status,result_json,error_message) "
        "VALUES (?,?,?,?,?,?,?,?,?)",
        (str(uuid.uuid4()), tenant_id, action, policy_id, executed_by, now_iso(), final_status,
         json.dumps(result), result.get("error") if not result.get("ok") else None)
    )
    return result


def list_ca_history(tenant_id: str, limit: int = 50) -> List[Dict[str, Any]]:
    return db_fetchall(
        "SELECT * FROM ca_history WHERE tenant_id=? ORDER BY executed_at DESC LIMIT ?",
        (tenant_id, limit)
    )


# ── Domains Analyser (Fase 7) ─────────────────────────────────────────────────

_DOMAINS_SCRIPT = PLATFORM_DIR / "assessment" / "Invoke-DenjoyDomains.ps1"


def _run_domains_ps(tenant_id: str, action: str, params: Dict[str, Any]) -> Dict[str, Any]:
    profile = get_tenant_auth_profile(tenant_id, include_secret=True)
    ps_script = _DOMAINS_SCRIPT.resolve()
    if not ps_script.exists():
        raise FileNotFoundError(f"Domains script niet gevonden: {ps_script}")
    cmd = [
        "pwsh", "-NonInteractive", "-NoProfile", "-File", str(ps_script),
        "-Action", action,
        "-TenantId", profile["auth_tenant_id"],
        "-ClientId", profile["auth_client_id"],
        "-ParamsJson", json.dumps(params),
    ]
    if profile.get("auth_cert_thumbprint"):
        cmd += ["-CertThumbprint", profile["auth_cert_thumbprint"]]
    elif profile.get("auth_client_secret"):
        cmd += ["-ClientSecret", profile["auth_client_secret"]]
    proc = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
    output = proc.stdout + proc.stderr
    logger.info("[Domains] action=%s tenant=%s exit=%s", action, tenant_id, proc.returncode)
    if "##RESULT##" in output:
        try:
            return json.loads(output.split("##RESULT##")[-1].strip().split("\n")[0])
        except Exception:
            return {"ok": False, "error": "Parse fout"}
    return {"ok": False, "error": output[-500:] if output else "Geen output"}


# ── Alerts & Audit Logs (Fase 8) ──────────────────────────────────────────────

_ALERTS_SCRIPT = PLATFORM_DIR / "assessment" / "Invoke-DenjoyAlerts.ps1"


def _run_alerts_ps(tenant_id: str, action: str, params: Dict[str, Any]) -> Dict[str, Any]:
    profile = get_tenant_auth_profile(tenant_id, include_secret=True)
    ps_script = _ALERTS_SCRIPT.resolve()
    if not ps_script.exists():
        raise FileNotFoundError(f"Alerts script niet gevonden: {ps_script}")
    cmd = [
        "pwsh", "-NonInteractive", "-NoProfile", "-File", str(ps_script),
        "-Action", action,
        "-TenantId", profile["auth_tenant_id"],
        "-ClientId", profile["auth_client_id"],
        "-ParamsJson", json.dumps(params),
    ]
    if profile.get("auth_cert_thumbprint"):
        cmd += ["-CertThumbprint", profile["auth_cert_thumbprint"]]
    elif profile.get("auth_client_secret"):
        cmd += ["-ClientSecret", profile["auth_client_secret"]]
    proc = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
    output = proc.stdout + proc.stderr
    logger.info("[Alerts] action=%s tenant=%s exit=%s", action, tenant_id, proc.returncode)
    if "##RESULT##" in output:
        try:
            return json.loads(output.split("##RESULT##")[-1].strip().split("\n")[0])
        except Exception:
            return {"ok": False, "error": "Parse fout"}
    return {"ok": False, "error": output[-500:] if output else "Geen output"}


def get_alert_config(tenant_id: str) -> Dict[str, Any]:
    rows = db_fetchall("SELECT * FROM alert_config WHERE tenant_id=?", (tenant_id,))
    return rows[0] if rows else {}


def upsert_alert_config(tenant_id: str, webhook_url: str, webhook_type: str, email_addr: str) -> None:
    existing = get_alert_config(tenant_id)
    if existing:
        db_execute(
            "UPDATE alert_config SET webhook_url=?,webhook_type=?,email_addr=?,updated_at=? WHERE tenant_id=?",
            (webhook_url, webhook_type, email_addr, now_iso(), tenant_id)
        )
    else:
        db_execute(
            "INSERT INTO alert_config (id,tenant_id,webhook_url,webhook_type,email_addr,updated_at) VALUES (?,?,?,?,?,?)",
            (str(uuid.uuid4()), tenant_id, webhook_url, webhook_type, email_addr, now_iso())
        )


def _fire_webhook_for_tenant(tenant_id: str, event: str, payload_data: Dict[str, Any]) -> None:
    """Stuurt een webhook naar alle geconfigureerde kanalen voor deze tenant (fire-and-forget)."""
    import threading
    cfg = get_alert_config(tenant_id)
    if not cfg or not cfg.get("webhook_url"):
        return
    webhook_url  = cfg["webhook_url"]
    webhook_type = cfg.get("webhook_type", "teams")

    def _send() -> None:
        try:
            import urllib.request
            if webhook_type == "teams":
                body_dict: Dict[str, Any] = {
                    "@type": "MessageCard", "@context": "http://schema.org/extensions",
                    "themeColor": "FF6B2B",
                    "summary": payload_data.get("title", event),
                    "title": f"🔔 Denjoy IT — {payload_data.get('title', event)}",
                    "text": payload_data.get("message", ""),
                    "sections": [{"facts": [{"name": k, "value": str(v)} for k, v in payload_data.get("facts", {}).items()]}],
                }
            elif webhook_type == "slack":
                facts_text = "\n".join(f"*{k}:* {v}" for k, v in payload_data.get("facts", {}).items())
                body_dict = {"text": f"🔔 *Denjoy IT — {payload_data.get('title', event)}*\n{payload_data.get('message', '')}\n{facts_text}"}
            else:
                body_dict = {"source": "denjoy", "event": event, **payload_data}
            body = json.dumps(body_dict).encode()
            req = urllib.request.Request(webhook_url, data=body, headers={"Content-Type": "application/json"}, method="POST")
            with urllib.request.urlopen(req, timeout=10):
                pass
        except Exception as exc:
            logger.warning("Webhook fire mislukt voor tenant %s event %s: %s", tenant_id, event, exc)

    threading.Thread(target=_send, daemon=True).start()


def send_test_webhook(webhook_url: str, webhook_type: str) -> Dict[str, Any]:
    import urllib.request
    payload: Dict[str, Any] = {}
    if webhook_type == "teams":
        payload = {"@type": "MessageCard", "@context": "http://schema.org/extensions",
                   "summary": "Denjoy Test", "themeColor": "0078D7",
                   "title": "✅ Denjoy IT Platform — Test melding",
                   "text": "Webhook verbinding succesvol geconfigureerd."}
    elif webhook_type == "slack":
        payload = {"text": "✅ *Denjoy IT Platform* — Test melding\nWebhook verbinding succesvol geconfigureerd."}
    else:
        payload = {"source": "denjoy", "event": "test", "message": "Webhook verbinding succesvol."}
    body = json.dumps(payload).encode()
    req = urllib.request.Request(webhook_url, data=body, headers={"Content-Type": "application/json"}, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            return {"ok": True, "status": resp.status}
    except Exception as exc:
        return {"ok": False, "error": str(exc)}


# ── Exchange & Email (Fase 9) ─────────────────────────────────────────────────

_EXCHANGE_SCRIPT = PLATFORM_DIR / "assessment" / "Invoke-DenjoyExchange.ps1"


def _run_exchange_ps(tenant_id: str, action: str, params: Dict[str, Any]) -> Dict[str, Any]:
    profile = get_tenant_auth_profile(tenant_id, include_secret=True)
    ps_script = _EXCHANGE_SCRIPT.resolve()
    if not ps_script.exists():
        raise FileNotFoundError(f"Exchange script niet gevonden: {ps_script}")
    cmd = [
        "pwsh", "-NonInteractive", "-NoProfile", "-File", str(ps_script),
        "-Action", action,
        "-TenantId", profile["auth_tenant_id"],
        "-ClientId", profile["auth_client_id"],
        "-ParamsJson", json.dumps(params),
    ]
    if profile.get("auth_cert_thumbprint"):
        cmd += ["-CertThumbprint", profile["auth_cert_thumbprint"]]
    elif profile.get("auth_client_secret"):
        cmd += ["-ClientSecret", profile["auth_client_secret"]]
    proc = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
    output = proc.stdout + proc.stderr
    logger.info("[Exchange] action=%s tenant=%s exit=%s", action, tenant_id, proc.returncode)
    if "##RESULT##" in output:
        try:
            data = json.loads(output.split("##RESULT##")[-1].strip().split("\n")[0])
            threading.Thread(target=_persist_live_findings, args=(tenant_id, "exchange", action, data), daemon=True).start()
            return data
        except Exception:
            return {"ok": False, "error": "Parse fout"}
    return {"ok": False, "error": output[-500:] if output else "Geen output"}


# ── Identiteit & Toegang ──────────────────────────────────────────────────────

# ═══════════════════════════════════════════════════════════════
# SCAN FINDINGS — extractors en persistentie
# ═══════════════════════════════════════════════════════════════

def _status_from_pct(pct: float, ok_threshold: float = 95.0, warn_threshold: float = 75.0) -> str:
    if pct >= ok_threshold:
        return "ok"
    if pct >= warn_threshold:
        return "warning"
    return "critical"

def _impact_from_status(status: str) -> str:
    return {"ok": "low", "warning": "high", "critical": "critical", "info": "low"}.get(status, "medium")

def _extract_identity_mfa(data: Dict[str, Any]) -> List[Dict[str, Any]]:
    findings: List[Dict[str, Any]] = []
    total = int(data.get("total") or 0)
    registered = int(data.get("mfaRegistered") or 0)
    pct = float(data.get("mfaPercentage") or 0)
    status = _status_from_pct(pct)
    recs = {
        "ok": "MFA-dekking is goed. Overweeg passwordless uitrol (Microsoft Authenticator + FIDO2).",
        "warning": "Verhoog MFA-dekking. Gebruik een CA-policy om MFA te vereisen voor alle gebruikers.",
        "critical": "Kritiek: implementeer direct een CA-policy die MFA vereist. Alle accounts zijn kwetsbaar.",
    }
    findings.append({"control": "mfa-coverage", "title": "MFA-registratie gebruikers",
        "status": status, "impact": _impact_from_status(status),
        "finding": f"{registered}/{total} gebruikers MFA-geregistreerd ({pct}%)",
        "recommendation": recs[status], "service": "Identity Beheer", "metric_value": pct})
    users = data.get("users") or []
    admin_no_mfa = [u for u in users if u.get("isAdmin") and not u.get("isMfaRegistered")]
    if admin_no_mfa:
        findings.append({"control": "admin-mfa", "title": f"Admins zonder MFA ({len(admin_no_mfa)})",
            "status": "critical", "impact": "critical",
            "finding": f"{len(admin_no_mfa)} beheerdersaccount(s) zonder MFA-registratie gedetecteerd",
            "recommendation": "Vereis direct MFA voor alle beheerdersaccounts via CA-policy.",
            "service": "Identity Beheer", "metric_value": float(len(admin_no_mfa))})
    elif total > 0:
        findings.append({"control": "admin-mfa", "title": "Admin MFA-dekking",
            "status": "ok", "impact": "low",
            "finding": "Alle beheerdersaccounts hebben MFA geregistreerd",
            "recommendation": "Handhaven.", "service": "Identity Beheer", "metric_value": 0.0})
    return findings

def _extract_identity_guests(data: Dict[str, Any]) -> List[Dict[str, Any]]:
    guests = data.get("guests") or []
    total = int(data.get("count") or len(guests))
    if total == 0:
        return [{"control": "guest-accounts", "title": "Gastaccounts",
            "status": "ok", "impact": "low",
            "finding": "Geen gastgebruikers gevonden in de tenant",
            "recommendation": "Geen actie vereist.", "service": "Identity Beheer", "metric_value": 0.0}]
    disabled = sum(1 for g in guests if not g.get("accountEnabled"))
    status = "ok" if total <= 20 else ("warning" if total <= 100 else "critical")
    return [{"control": "guest-accounts", "title": f"Gastaccounts ({total})",
        "status": status, "impact": _impact_from_status(status),
        "finding": f"{total} gastgebruikers aanwezig, {disabled} uitgeschakeld",
        "recommendation": "Review gastaccounts regelmatig. Verwijder inactieve gasten of gebruik Azure AD Access Reviews.",
        "service": "Identity Beheer", "metric_value": float(total)}]

def _extract_identity_security_defaults(data: Dict[str, Any]) -> List[Dict[str, Any]]:
    enabled = data.get("securityDefaultsEnabled")
    ca_count = int(data.get("caEnabledPolicies") or 0)
    rec = data.get("recommendation") or ""
    if enabled is True and ca_count == 0:
        status, finding = "ok", "Security Defaults ingeschakeld (geen conflicterende CA-policies)"
    elif enabled is False and ca_count >= 3:
        status, finding = "ok", f"Security Defaults uitgeschakeld — {ca_count} CA-policies actief (correct)"
    elif enabled is False and ca_count == 0:
        status, finding = "critical", "Security Defaults uitgeschakeld én geen CA-policies actief"
    else:
        status, finding = "warning", f"Security Defaults: {'aan' if enabled else 'uit'}, {ca_count} CA-policies"
    return [{"control": "security-defaults", "title": "Security Defaults status",
        "status": status, "impact": _impact_from_status(status), "finding": finding,
        "recommendation": rec or "Zorg dat óf Security Defaults óf CA-policies actief zijn — niet beide.",
        "service": "Zero Trust Baseline", "metric_value": float(ca_count)}]

def _extract_identity_legacy_auth(data: Dict[str, Any]) -> List[Dict[str, Any]]:
    users = data.get("users") or []
    count = int(data.get("affectedUsers") or len(users))
    if count == 0:
        return [{"control": "legacy-auth", "title": "Legacy authenticatie",
            "status": "ok", "impact": "low",
            "finding": "Geen legacy-auth activiteit gevonden (afgelopen 30 dagen)",
            "recommendation": "Handhaven. Overweeg een CA-policy om legacy auth expliciet te blokkeren.",
            "service": "Zero Trust Baseline", "metric_value": 0.0}]
    return [{"control": "legacy-auth", "title": f"Legacy auth actief ({count} gebruikers)",
        "status": "critical", "impact": "critical",
        "finding": f"{count} gebruiker(s) met legacy-auth activiteit in de afgelopen 30 dagen",
        "recommendation": "Blokkeer legacy authenticatie via CA-policy. Legacy clients omzeilen MFA.",
        "service": "Zero Trust Baseline", "metric_value": float(count)}]

def _extract_identity_admin_roles(data: Dict[str, Any]) -> List[Dict[str, Any]]:
    roles = data.get("roles") or []
    total_admins = int(data.get("totalAdmins") or 0)
    role_count = int(data.get("roleCount") or len(roles))
    if total_admins < 2:
        status = "critical"
    elif total_admins > 6:
        status = "warning"
    else:
        status = "ok"
    return [{"control": "admin-roles", "title": f"Beheerdersrollen ({role_count} rollen, {total_admins} admins)",
        "status": status, "impact": _impact_from_status(status),
        "finding": f"{role_count} actieve directoryrolls met in totaal {total_admins} unieke admins",
        "recommendation": "Houd het aantal Global Admins beperkt (2–4). Gebruik privileged roles zo minimaal mogelijk.",
        "service": "Identity Beheer", "metric_value": float(total_admins)}]

def _extract_appregs(data: Dict[str, Any]) -> List[Dict[str, Any]]:
    findings: List[Dict[str, Any]] = []
    apps = data.get("apps") or []
    total = int(data.get("total") or len(apps))
    expired = int(data.get("expired") or 0)
    critical_n = int(data.get("critical") or 0)
    warning_n = int(data.get("warning") or 0)
    if expired > 0:
        findings.append({"control": "appregs-expired", "title": f"Verlopen secrets/certs ({expired})",
            "status": "critical", "impact": "critical",
            "finding": f"{expired} app-registratie(s) met verlopen secret of certificaat",
            "recommendation": "Vernieuw direct alle verlopen secrets en certificaten om uitval te voorkomen.",
            "service": "App Registraties", "metric_value": float(expired)})
    if critical_n > 0:
        findings.append({"control": "appregs-expiring-soon", "title": f"Secrets verlopen binnen 14 dagen ({critical_n})",
            "status": "critical", "impact": "high",
            "finding": f"{critical_n} app-registratie(s) met secret/cert dat binnen 14 dagen verloopt",
            "recommendation": "Vernieuw deze secrets/certs op zeer korte termijn.",
            "service": "App Registraties", "metric_value": float(critical_n)})
    if warning_n > 0:
        findings.append({"control": "appregs-expiring-warning", "title": f"Secrets verlopen binnen 30 dagen ({warning_n})",
            "status": "warning", "impact": "medium",
            "finding": f"{warning_n} app-registratie(s) met secret/cert dat binnen 30 dagen verloopt",
            "recommendation": "Plan het vernieuwen van deze secrets/certs.",
            "service": "App Registraties", "metric_value": float(warning_n)})
    if not findings:
        findings.append({"control": "appregs-status", "title": f"App registraties ({total})",
            "status": "ok", "impact": "low",
            "finding": f"{total} app-registratie(s) — geen verlopen of bijna-verlopen secrets",
            "recommendation": "Blijf secrets en certificaten monitoren voor verval.",
            "service": "App Registraties", "metric_value": float(total)})
    return findings

def _extract_exchange_forwarding(data: Dict[str, Any]) -> List[Dict[str, Any]]:
    items = data.get("forwarding") or data.get("items") or []
    count = len(items)
    if count == 0:
        return [{"control": "exchange-forwarding", "title": "Externe e-mail forwarding",
            "status": "ok", "impact": "low",
            "finding": "Geen actieve externe e-mail forwardings gevonden",
            "recommendation": "Handhaven. Monitor forwarding rules regelmatig.",
            "service": "Exchange", "metric_value": 0.0}]
    return [{"control": "exchange-forwarding", "title": f"Externe forwarding actief ({count})",
        "status": "critical", "impact": "critical",
        "finding": f"{count} mailbox(en) stuurt e-mail extern door — potentieel dataverlies",
        "recommendation": "Blokkeer automatisch extern doorsturen via transport rule. Review deze forwardings direct.",
        "service": "Exchange", "metric_value": float(count)}]

def _extract_exchange_rules(data: Dict[str, Any]) -> List[Dict[str, Any]]:
    suspicious = int(data.get("suspicious") or 0)
    total = int(data.get("total") or 0)
    if suspicious == 0:
        return [{"control": "exchange-inbox-rules", "title": "Verdachte inboxregels",
            "status": "ok", "impact": "low",
            "finding": f"Geen verdachte inboxregels gevonden ({total} regels gecontroleerd)",
            "recommendation": "Blijf inboxregels periodiek monitoren.",
            "service": "Exchange", "metric_value": 0.0}]
    return [{"control": "exchange-inbox-rules", "title": f"Verdachte inboxregels ({suspicious})",
        "status": "critical", "impact": "critical",
        "finding": f"{suspicious} verdachte inboxregel(s) gedetecteerd van {total} totaal",
        "recommendation": "Onderzoek en verwijder verdachte inboxregels. Dit kan wijzen op een gehackt account.",
        "service": "Exchange", "metric_value": float(suspicious)}]

def _extract_ca_policies(data: Dict[str, Any]) -> List[Dict[str, Any]]:
    policies = data.get("policies") or []
    enabled = sum(1 for p in policies if (p.get("state") or "").lower() == "enabled")
    report_only = sum(1 for p in policies if (p.get("state") or "").lower() in ("enabledforreportingbutnotenforcingleway", "reportonly"))
    total = len(policies)
    if enabled >= 3:
        status = "ok"
    elif enabled >= 1:
        status = "warning"
    else:
        status = "critical"
    return [{"control": "ca-policies", "title": f"Conditional Access policies ({enabled} actief)",
        "status": status, "impact": _impact_from_status(status),
        "finding": f"{enabled} actieve CA-policies van {total} totaal ({report_only} report-only)",
        "recommendation": (
            "Goede CA-coverage. Controleer of MFA, legacy-auth blokkering en admin-bescherming zijn opgenomen." if status == "ok"
            else "Breid CA-policies uit. Implementeer minimaal: MFA voor alle users, legacy auth blokkering, admin bescherming."
            if status == "warning"
            else "Geen CA-policies actief! Implementeer direct minimale CA-bescherming."
        ),
        "service": "Zero Trust Baseline", "metric_value": float(enabled)}]

def _extract_teams(data: Dict[str, Any]) -> List[Dict[str, Any]]:
    teams = data.get("teams") or []
    total = int(data.get("count") or len(teams))
    public = int(data.get("publicCount") or sum(1 for t in teams if (t.get("visibility") or "").lower() == "public"))
    pct_public = (public / total * 100) if total > 0 else 0
    status = "ok" if pct_public < 20 else ("warning" if pct_public < 50 else "critical")
    return [{"control": "teams-public", "title": f"Teams ({total} totaal, {public} publiek)",
        "status": status, "impact": _impact_from_status(status),
        "finding": f"{public}/{total} Teams zijn publiek zichtbaar ({pct_public:.0f}%)",
        "recommendation": "Gebruik bij voorkeur Private Teams. Review publieke Teams op gevoelige inhoud.",
        "service": "Samenwerking", "metric_value": float(public)}]

def _extract_sharepoint_settings(data: Dict[str, Any]) -> List[Dict[str, Any]]:
    sharing = str(data.get("sharingCapability") or "unknown").lower()
    if sharing in ("disabled", "existingexternalusersharingonly"):
        status = "ok"
    elif sharing in ("externalusersharingonly",):
        status = "warning"
    elif sharing in ("externaluserandguestsharing",):
        status = "critical"
    else:
        status = "info"
    return [{"control": "sharepoint-sharing", "title": "SharePoint externe sharing",
        "status": status, "impact": _impact_from_status(status),
        "finding": f"Sharingsniveau: {data.get('sharingCapability') or 'Onbekend'}",
        "recommendation": "Beperk extern delen tot 'Existing external users' of 'Disabled'. Voorkom anonieme links.",
        "service": "Samenwerking", "metric_value": None}]

# Map van (domain, action) → extractorfunctie
_FINDING_EXTRACTORS: Dict[Tuple[str, str], Any] = {
    ("identity", "list-mfa"):               _extract_identity_mfa,
    ("identity", "list-guests"):            _extract_identity_guests,
    ("identity", "get-security-defaults"):  _extract_identity_security_defaults,
    ("identity", "list-legacy-auth"):       _extract_identity_legacy_auth,
    ("identity", "list-admin-roles"):       _extract_identity_admin_roles,
    ("apps", "list-appregs"):              _extract_appregs,
    ("exchange", "list-forwarding"):        _extract_exchange_forwarding,
    ("exchange", "list-mailbox-rules"):     _extract_exchange_rules,
    ("ca", "list-policies"):               _extract_ca_policies,
    ("collaboration", "list-teams"):        _extract_teams,
    ("collaboration", "get-sharepoint-settings"): _extract_sharepoint_settings,
}


def _persist_live_findings(tenant_id: str, domain: str, action: str, data: Dict[str, Any]) -> None:
    """Sla gestructureerde bevindingen op in scan_findings na een succesvolle PS-run."""
    if not data or data.get("ok") is False:
        return
    extractor = _FINDING_EXTRACTORS.get((domain, action))
    if not extractor:
        return
    try:
        findings = extractor(data)
    except Exception as e:
        logger.warning("[findings] Extractor %s/%s fout: %s", domain, action, e)
        return
    if not findings:
        return
    ts = now_iso()
    raw_json = json.dumps(data, ensure_ascii=False)[:8000]  # cap op 8KB
    conn = get_conn()
    try:
        with conn:
            conn.executemany(
                """
                INSERT INTO scan_findings
                    (id, tenant_id, domain, control, title, status, finding,
                     impact, recommendation, service, metric_value, raw_json, scanned_at)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
                """,
                [
                    (
                        str(uuid.uuid4()), tenant_id, domain,
                        f["control"], f["title"], f["status"], f.get("finding"),
                        f.get("impact", "low"), f.get("recommendation"),
                        f.get("service"), f.get("metric_value"),
                        raw_json if i == 0 else None, ts,
                    )
                    for i, f in enumerate(findings)
                ],
            )
    except Exception as e:
        logger.warning("[findings] DB write fout tenant=%s domain=%s: %s", tenant_id, domain, e)
    finally:
        conn.close()


def _get_tenant_health_score(tenant_id: str) -> Dict[str, Any]:
    """Bereken health score op basis van meest recente findings per control."""
    conn = get_conn()
    try:
        rows = conn.execute(
            """
            SELECT f.domain, f.control, f.status, f.impact, f.title, f.scanned_at
            FROM scan_findings f
            INNER JOIN (
                SELECT domain, control, MAX(scanned_at) AS max_at
                FROM scan_findings WHERE tenant_id=?
                GROUP BY domain, control
            ) latest ON f.domain=latest.domain AND f.control=latest.control AND f.scanned_at=latest.max_at
            WHERE f.tenant_id=?
            ORDER BY f.domain, f.control
            """,
            (tenant_id, tenant_id),
        ).fetchall()
        findings = [dict(r) for r in rows]
        total = len(findings)
        ok_count = sum(1 for f in findings if f["status"] == "ok")
        warn_count = sum(1 for f in findings if f["status"] == "warning")
        crit_count = sum(1 for f in findings if f["status"] == "critical")
        score = round((ok_count * 1.0 + warn_count * 0.5) / total * 100) if total else None
        return {
            "ok": True, "tenant_id": tenant_id,
            "score": score,
            "total": total, "ok_count": ok_count,
            "warning_count": warn_count, "critical_count": crit_count,
            "findings": findings,
        }
    finally:
        conn.close()


def _get_findings_trend(tenant_id: str, days: int = 30) -> List[Dict[str, Any]]:
    """Dagelijks gemiddelde health score voor trends-grafiek."""
    cutoff = (datetime.now(timezone.utc) - timedelta(days=days)).isoformat()
    conn = get_conn()
    try:
        rows = conn.execute(
            """
            SELECT
                substr(scanned_at, 1, 10) AS day,
                COUNT(*) AS total,
                SUM(CASE WHEN status='ok' THEN 1 ELSE 0 END) AS ok_count,
                SUM(CASE WHEN status='warning' THEN 1 ELSE 0 END) AS warn_count,
                SUM(CASE WHEN status='critical' THEN 1 ELSE 0 END) AS crit_count
            FROM scan_findings
            WHERE tenant_id=? AND scanned_at >= ?
            GROUP BY day
            ORDER BY day ASC
            """,
            (tenant_id, cutoff),
        ).fetchall()
        result = []
        for r in rows:
            total = r["total"] or 1
            score = round((r["ok_count"] * 1.0 + r["warn_count"] * 0.5) / total * 100)
            result.append({"date": r["day"], "score": score,
                "ok": r["ok_count"], "warning": r["warn_count"], "critical": r["crit_count"]})
        return result
    finally:
        conn.close()


_IDENTITY_SCRIPT = PLATFORM_DIR / "assessment" / "Invoke-DenjoyIdentity.ps1"


def _run_identity_ps(tenant_id: str, action: str, params: Dict[str, Any]) -> Dict[str, Any]:
    profile = get_tenant_auth_profile(tenant_id, include_secret=True)
    ps_script = _IDENTITY_SCRIPT.resolve()
    if not ps_script.exists():
        raise FileNotFoundError(f"Identity script niet gevonden: {ps_script}")
    cmd = [
        "pwsh", "-NonInteractive", "-NoProfile", "-File", str(ps_script),
        "-Action", action,
        "-TenantId", profile["auth_tenant_id"],
        "-ClientId", profile["auth_client_id"],
        "-ParamsJson", json.dumps(params),
    ]
    if profile.get("auth_cert_thumbprint"):
        cmd += ["-CertThumbprint", profile["auth_cert_thumbprint"]]
    elif profile.get("auth_client_secret"):
        cmd += ["-ClientSecret", profile["auth_client_secret"]]
    proc = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
    output = proc.stdout + proc.stderr
    logger.info("[Identity] action=%s tenant=%s exit=%s", action, tenant_id, proc.returncode)
    if "##RESULT##" in output:
        try:
            data = json.loads(output.split("##RESULT##")[-1].strip().split("\n")[0])
            threading.Thread(target=_persist_live_findings, args=(tenant_id, "identity", action, data), daemon=True).start()
            return data
        except Exception:
            return {"ok": False, "error": "Parse fout"}
    return {"ok": False, "error": output[-500:] if output else "Geen output"}


# ── App Registraties ──────────────────────────────────────────────────────────

_APPREGS_SCRIPT = PLATFORM_DIR / "assessment" / "Invoke-DenjoyApps.ps1"


def _run_appregs_ps(tenant_id: str, action: str, params: Dict[str, Any]) -> Dict[str, Any]:
    profile = get_tenant_auth_profile(tenant_id, include_secret=True)
    ps_script = _APPREGS_SCRIPT.resolve()
    if not ps_script.exists():
        raise FileNotFoundError(f"AppRegs script niet gevonden: {ps_script}")
    cmd = [
        "pwsh", "-NonInteractive", "-NoProfile", "-File", str(ps_script),
        "-Action", action,
        "-TenantId", profile["auth_tenant_id"],
        "-ClientId", profile["auth_client_id"],
        "-ParamsJson", json.dumps(params),
    ]
    if profile.get("auth_cert_thumbprint"):
        cmd += ["-CertThumbprint", profile["auth_cert_thumbprint"]]
    elif profile.get("auth_client_secret"):
        cmd += ["-ClientSecret", profile["auth_client_secret"]]
    proc = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
    output = proc.stdout + proc.stderr
    logger.info("[AppRegs] action=%s tenant=%s exit=%s", action, tenant_id, proc.returncode)
    if "##RESULT##" in output:
        try:
            data = json.loads(output.split("##RESULT##")[-1].strip().split("\n")[0])
            threading.Thread(target=_persist_live_findings, args=(tenant_id, "apps", action, data), daemon=True).start()
            return data
        except Exception:
            return {"ok": False, "error": "Parse fout"}
    return {"ok": False, "error": output[-500:] if output else "Geen output"}


# ── Samenwerking: SharePoint & Teams ─────────────────────────────────────────

_COLLAB_SCRIPT = PLATFORM_DIR / "assessment" / "Invoke-DenjoyCollaboration.ps1"


def _run_collab_ps(tenant_id: str, action: str, params: Dict[str, Any]) -> Dict[str, Any]:
    profile = get_tenant_auth_profile(tenant_id, include_secret=True)
    ps_script = _COLLAB_SCRIPT.resolve()
    if not ps_script.exists():
        raise FileNotFoundError(f"Collaboration script niet gevonden: {ps_script}")
    cmd = [
        "pwsh", "-NonInteractive", "-NoProfile", "-File", str(ps_script),
        "-Action", action,
        "-TenantId", profile["auth_tenant_id"],
        "-ClientId", profile["auth_client_id"],
        "-ParamsJson", json.dumps(params),
    ]
    if profile.get("auth_cert_thumbprint"):
        cmd += ["-CertThumbprint", profile["auth_cert_thumbprint"]]
    elif profile.get("auth_client_secret"):
        cmd += ["-ClientSecret", profile["auth_client_secret"]]
    proc = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
    output = proc.stdout + proc.stderr
    logger.info("[Collab] action=%s tenant=%s exit=%s", action, tenant_id, proc.returncode)
    if "##RESULT##" in output:
        try:
            data = json.loads(output.split("##RESULT##")[-1].strip().split("\n")[0])
            threading.Thread(target=_persist_live_findings, args=(tenant_id, "collaboration", action, data), daemon=True).start()
            return data
        except Exception:
            return {"ok": False, "error": "Parse fout"}
    return {"ok": False, "error": output[-500:] if output else "Geen output"}


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
        self._run_disconnect_cleanup(run_id)
        return True

    def _run_disconnect_cleanup(self, run_id: str) -> None:
        """Best-effort cleanup for terminated runs where script finally may not complete."""
        try:
            stop_script = (PLATFORM_DIR / "assessment" / "Stop-M365BaselineAssessment.ps1").resolve()
            if not stop_script.exists():
                append_run_log(run_id, "Stop cleanup script not found; skipping disconnect cleanup.")
                return
            pwsh = shutil.which("pwsh") or shutil.which("powershell")
            if not pwsh:
                append_run_log(run_id, "PowerShell not found; skipping disconnect cleanup.")
                return

            cmd = [pwsh, "-NoLogo", "-NoProfile", "-NonInteractive", "-File", str(stop_script)]
            append_run_log(run_id, "Running forced disconnect cleanup...")
            proc = subprocess.run(
                cmd,
                cwd=str(PLATFORM_DIR / "assessment"),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=45,
            )
            output = (proc.stdout or "").strip()
            if output:
                for line in output.splitlines()[-80:]:
                    append_run_log(run_id, f"[cleanup] {line}")
            append_run_log(run_id, f"Forced disconnect cleanup exit code: {proc.returncode}")
        except subprocess.TimeoutExpired:
            append_run_log(run_id, "Forced disconnect cleanup timed out.")
        except Exception as exc:
            append_run_log(run_id, f"Forced disconnect cleanup failed: {exc}")

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
            if run_mode == "script":
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
            try:
                snap_count = import_run_snapshots_to_db(run_id)
                if snap_count:
                    append_run_log(run_id, f"{snap_count} portal JSON snapshots opgeslagen in database.")
            except Exception as snap_exc:
                logger.warning("Snapshot import mislukt voor run %s: %s", run_id, snap_exc)
            # Webhook notificatie na voltooide assessment
            try:
                run_meta = db_fetchone("SELECT tenant_id, score_overall, critical_count FROM assessment_runs WHERE id=?", (run_id,))
                if run_meta and run_meta.get("tenant_id"):
                    _fire_webhook_for_tenant(
                        run_meta["tenant_id"],
                        "assessment_completed",
                        {
                            "title": "Assessment voltooid",
                            "message": f"De M365-assessment is succesvol afgerond.",
                            "facts": {
                                "Tenant": run_meta["tenant_id"],
                                "Score": f"{run_meta.get('score_overall') or '—'}%",
                                "Kritiek": str(run_meta.get("critical_count") or 0),
                                "Tijdstip": now_iso(),
                            },
                        },
                    )
            except Exception as wh_exc:
                logger.warning("Webhook na assessment mislukt: %s", wh_exc)
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
               "-ExportCsv", "-ExportJson"]        # assessment exports voor CSV + portal JSON
        cmd.extend(phase_skip_flags(phases))

        # Authenticatie parameters doorgeven: tenant-specifiek profiel eerst, daarna legacy global fallback.
        run = db_fetchone("SELECT tenant_id FROM assessment_runs WHERE id=?", (run_id,)) or {}
        tenant_row = db_fetchone("SELECT customer_name, tenant_name, tenant_guid FROM tenants WHERE id=?", (run.get("tenant_id") or "",)) or {}
        tenant_guid_from_selection = (tenant_row.get("tenant_guid") or "").strip()
        tenant_id_for_profile = (run.get("tenant_id") or "").strip()
        tenant_profile = get_tenant_auth_profile(tenant_id_for_profile, include_secret=True) if tenant_id_for_profile else {}
        tenant_id_cfg = (cfg.get("auth_tenant_id") or "").strip()
        tenant_id_profile = (tenant_profile.get("auth_tenant_id") or "").strip()
        if tenant_guid_from_selection and tenant_id_profile and tenant_guid_from_selection.lower() != tenant_id_profile.lower():
            raise RuntimeError(
                "Tenant-profiel auth_tenant_id komt niet overeen met de geselecteerde tenant GUID. "
                "Werk de app-registratie in Admin > tenant-instellingen bij."
            )
        if tenant_guid_from_selection and tenant_id_cfg and tenant_id_profile == "" and tenant_guid_from_selection.lower() != tenant_id_cfg.lower():
            raise RuntimeError(
                "Globale auth_tenant_id komt niet overeen met de geselecteerde tenant. "
                "Stel een tenant-specifieke app-registratie in voor deze tenant."
            )

        # Voorkeur: tenant profiel -> tenant GUID -> legacy globale config
        tenant_id = tenant_id_profile or tenant_guid_from_selection or tenant_id_cfg
        if not tenant_id:
            raise RuntimeError("Geen TenantId beschikbaar voor script-authenticatie (tenant_guid/auth_tenant_id ontbreekt).")
        client_id   = (tenant_profile.get("auth_client_id") or cfg.get("auth_client_id") or "").strip()
        cert_thumb  = (tenant_profile.get("auth_cert_thumbprint") or cfg.get("auth_cert_thumbprint") or "").strip()
        client_sec  = (tenant_profile.get("auth_client_secret") or cfg.get("auth_client_secret") or "").strip()

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
            # Client secret alleen via omgevingsvariabele (niet als command-arg).
            # Start-M365BaselineAssessment.ps1 converteert dit naar SecureString.
            env["M365_CLIENT_SECRET"] = client_sec
        append_run_log(run_id, "Starting PowerShell assessment.")
        append_run_log(run_id, "Command: " + " ".join(cmd))
        proc = subprocess.Popen(
            cmd,
            cwd=str(PLATFORM_DIR / "assessment"),
            env=env,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
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


RUN_MANAGER = RunManager()


# ══════════════════════════════════════════════════════════════════════════════
# JOB QUEUE DISPATCHER
# ══════════════════════════════════════════════════════════════════════════════

class JobDispatcher:
    """Achtergrondthread die pending jobs uit job_queue oppakt en uitvoert.

    Ondersteunde job_types:
      - assessment_run : start een assessment voor tenant_id (payload: phases, run_mode)
      - snapshot_import: importeer portal JSON van een run_id in m365_snapshots
    """

    _POLL_INTERVAL = 15  # seconden tussen polls

    def __init__(self) -> None:
        self._stop = threading.Event()
        self._thread: Optional[threading.Thread] = None

    def start(self) -> None:
        if self._thread and self._thread.is_alive():
            return
        self._stop.clear()
        self._thread = threading.Thread(target=self._loop, daemon=True, name="job-dispatcher")
        self._thread.start()
        logger.info("JobDispatcher gestart.")

    def stop(self) -> None:
        self._stop.set()

    def _loop(self) -> None:
        while not self._stop.wait(self._POLL_INTERVAL):
            try:
                self._poll()
            except Exception:
                logger.error("JobDispatcher poll-fout: %s", traceback.format_exc())

    def _poll(self) -> None:
        now = now_iso()
        rows = db_fetchall(
            "SELECT * FROM job_queue WHERE status='pending' AND scheduled_at<=? "
            "ORDER BY priority ASC, scheduled_at ASC LIMIT 5",
            (now,),
        )
        for row in rows:
            job_id = row["id"]
            # Claim de job (optimistic locking via status update)
            updated = db_execute(
                "UPDATE job_queue SET status='running', started_at=?, attempt_count=attempt_count+1 "
                "WHERE id=? AND status='pending'",
                (now, job_id),
            )
            if not updated:
                continue  # al geclaimd door andere thread
            logger.info("JobDispatcher: job %s (%s) gestart.", job_id, row["job_type"])
            threading.Thread(
                target=self._run_job,
                args=(dict(row),),
                daemon=True,
                name=f"job-{job_id[:8]}",
            ).start()

    def _run_job(self, row: Dict[str, Any]) -> None:
        job_id = row["id"]
        job_type = row["job_type"]
        tenant_id = row.get("tenant_id")
        payload: Dict[str, Any] = {}
        try:
            payload = json.loads(row.get("payload_json") or "{}")
        except Exception:
            pass
        try:
            result = self._dispatch(job_type, tenant_id, payload)
            db_execute(
                "UPDATE job_queue SET status='completed', completed_at=?, result_json=? WHERE id=?",
                (now_iso(), json.dumps(result, ensure_ascii=False), job_id),
            )
            logger.info("JobDispatcher: job %s voltooid.", job_id)
        except Exception as exc:
            attempt = int(db_fetchone("SELECT attempt_count FROM job_queue WHERE id=?", (job_id,))["attempt_count"] or 1)
            max_att = int(row.get("max_attempts") or 3)
            next_status = "pending" if attempt < max_att else "failed"
            next_scheduled = now_iso() if next_status == "pending" else None
            db_execute(
                "UPDATE job_queue SET status=?, error_message=?, completed_at=?, scheduled_at=COALESCE(?,scheduled_at) WHERE id=?",
                (next_status, str(exc), now_iso() if next_status == "failed" else None, next_scheduled, job_id),
            )
            logger.error("JobDispatcher: job %s mislukt (poging %d/%d): %s", job_id, attempt, max_att, exc)

    def _dispatch(self, job_type: str, tenant_id: Optional[str], payload: Dict[str, Any]) -> Dict[str, Any]:
        if job_type == "assessment_run":
            if not tenant_id:
                raise ValueError("tenant_id vereist voor assessment_run job")
            phases = payload.get("phases") or ["phase1", "phase2", "phase3", "phase4", "phase5", "phase6"]
            run_mode = payload.get("run_mode") or load_config().get("default_run_mode") or "demo"
            scan_type = payload.get("scan_type") or "full"
            run_id = str(uuid.uuid4())
            ts = now_iso()
            db_execute(
                "INSERT INTO assessment_runs (id, tenant_id, status, run_mode, created_at, updated_at) "
                "VALUES (?, ?, 'queued', ?, ?, ?)",
                (run_id, tenant_id, run_mode, ts, ts),
            )
            RUN_MANAGER.start(run_id, phases, run_mode, scan_type)
            return {"run_id": run_id, "status": "started"}

        if job_type == "snapshot_import":
            run_id = payload.get("run_id")
            if not run_id:
                raise ValueError("run_id vereist voor snapshot_import job")
            count = import_run_snapshots_to_db(run_id)
            return {"snapshots_written": count}

        raise ValueError(f"Onbekend job_type: {job_type!r}")


JOB_DISPATCHER = JobDispatcher()


# ══════════════════════════════════════════════════════════════════════════════
# USER MANAGEMENT
# ══════════════════════════════════════════════════════════════════════════════

_GOD_ADMIN_EMAIL = os.environ.get("DENJOY_ADMIN_EMAIL", "schiphorst.d@gmail.com").strip().lower()


def list_users() -> List[Dict[str, Any]]:
    rows = db_fetchall(
        "SELECT id, email, role, display_name, linked_tenant_id, is_active, created_at, last_login_at "
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
    if role not in ("admin", "klant", "security"):
        raise ValueError("Ongeldige rol — kies 'admin', 'security' of 'klant'.")
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
        if payload["role"] not in ("admin", "klant", "security"):
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


# ── Customer CRUD (Fase 3) ────────────────────────────────────────────────────

def list_customers(status: Optional[str] = None) -> List[Dict[str, Any]]:
    """Lijst alle klanten, optioneel gefilterd op status."""
    if status:
        rows = db_fetchall("SELECT * FROM customers WHERE status=? ORDER BY name", (status,))
    else:
        rows = db_fetchall("SELECT * FROM customers ORDER BY name")
    for c in rows:
        cnt = db_fetchone(
            "SELECT COUNT(*) AS cnt FROM tenants WHERE customer_id=? AND is_active=1", (c["id"],)
        )
        c["tenant_count"] = int((cnt or {}).get("cnt") or 0)
    return rows


def get_customer(customer_id: str) -> Optional[Dict[str, Any]]:
    """Haalt één klant op inclusief gekoppelde tenants en services."""
    c = db_fetchone("SELECT * FROM customers WHERE id=?", (customer_id,))
    if not c:
        return None
    c["tenants"] = db_fetchall(
        "SELECT id, tenant_name, tenant_guid, status FROM tenants WHERE customer_id=? AND is_active=1",
        (customer_id,),
    )
    c["services"] = db_fetchall(
        "SELECT * FROM customer_services WHERE customer_id=? ORDER BY service_key",
        (customer_id,),
    )
    return c


def create_customer(payload: Dict[str, Any]) -> Dict[str, Any]:
    """Maakt een nieuwe klantkaart aan."""
    name = (payload.get("name") or "").strip()
    if not name:
        raise ValueError("name is verplicht")
    status = (payload.get("status") or "active").strip()
    if status not in {"active", "onboarding", "paused", "offboarded"}:
        status = "active"
    cid = str(uuid.uuid4())
    ts = now_iso()
    db_execute(
        "INSERT INTO customers (id, name, status, primary_contact_name, primary_contact_email, notes, created_at, updated_at) "
        "VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
        (
            cid, name, status,
            (payload.get("primary_contact_name") or "").strip() or None,
            (payload.get("primary_contact_email") or "").strip() or None,
            (payload.get("notes") or "").strip() or None,
            ts, ts,
        ),
    )
    return get_customer(cid) or {}


def update_customer(customer_id: str, payload: Dict[str, Any]) -> Dict[str, Any]:
    """Werkt een bestaande klantkaart bij."""
    if not db_fetchone("SELECT id FROM customers WHERE id=?", (customer_id,)):
        raise ValueError("Klant niet gevonden")
    allowed = {"name", "status", "primary_contact_name", "primary_contact_email", "notes"}
    fields: Dict[str, Any] = {}
    for k, v in payload.items():
        if k not in allowed:
            continue
        fields[k] = v.strip() if isinstance(v, str) else v
    if "status" in fields and fields["status"] not in {"active", "onboarding", "paused", "offboarded"}:
        raise ValueError("Ongeldige status")
    if not fields:
        return get_customer(customer_id) or {}
    fields["updated_at"] = now_iso()
    sql = "UPDATE customers SET " + ", ".join(f"{k}=?" for k in fields) + " WHERE id=?"
    db_execute(sql, tuple(fields.values()) + (customer_id,))
    return get_customer(customer_id) or {}


def delete_customer(customer_id: str) -> Dict[str, Any]:
    """Verwijdert een klant. Mislukt als er nog actieve tenants gekoppeld zijn."""
    if not db_fetchone("SELECT id FROM customers WHERE id=?", (customer_id,)):
        raise ValueError("Klant niet gevonden")
    linked = db_fetchone(
        "SELECT COUNT(*) AS cnt FROM tenants WHERE customer_id=? AND is_active=1", (customer_id,)
    )
    if int((linked or {}).get("cnt") or 0) > 0:
        raise ValueError("Klant heeft nog actieve tenants — ontkoppel of deactiveer tenants eerst")
    db_execute("DELETE FROM customer_services WHERE customer_id=?", (customer_id,))
    db_execute("DELETE FROM customers WHERE id=?", (customer_id,))
    return {"id": customer_id, "deleted": True}


# ── Portal Roles ──────────────────────────────────────────────────────────────

def list_portal_roles() -> List[Dict[str, Any]]:
    return db_fetchall("SELECT * FROM portal_roles ORDER BY role_key")


# ── User Customer Access ──────────────────────────────────────────────────────

def list_user_customer_access(customer_id: Optional[str] = None,
                               user_id: Optional[str] = None) -> List[Dict[str, Any]]:
    if customer_id and user_id:
        rows = db_fetchall("SELECT * FROM user_customer_access WHERE customer_id=? AND portal_user_id=?",
                           (customer_id, user_id))
    elif customer_id:
        rows = db_fetchall("SELECT * FROM user_customer_access WHERE customer_id=?", (customer_id,))
    elif user_id:
        rows = db_fetchall("SELECT * FROM user_customer_access WHERE portal_user_id=?", (user_id,))
    else:
        rows = db_fetchall("SELECT * FROM user_customer_access")
    return rows


def grant_customer_access(customer_id: str, user_id: str, role_key: str,
                           granted_by: str = "", expires_at: Optional[str] = None,
                           scope: Optional[str] = None) -> Dict[str, Any]:
    role = db_fetchone("SELECT id FROM portal_roles WHERE role_key=?", (role_key,))
    if not role:
        raise ValueError(f"Onbekende rol: {role_key}")
    if not db_fetchone("SELECT id FROM customers WHERE id=?", (customer_id,)):
        raise ValueError("Klant niet gevonden")
    if not db_fetchone("SELECT id FROM users WHERE id=?", (user_id,)):
        raise ValueError("Gebruiker niet gevonden")
    aid = str(uuid.uuid4())
    db_execute(
        "INSERT OR REPLACE INTO user_customer_access "
        "(id, portal_user_id, customer_id, portal_role_id, scope, granted_by, granted_at, expires_at) "
        "VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
        (aid, user_id, customer_id, role["id"], scope, granted_by, now_iso(), expires_at),
    )
    return db_fetchone("SELECT * FROM user_customer_access WHERE id=?", (aid,)) or {}


def revoke_customer_access(customer_id: str, user_id: str) -> Dict[str, Any]:
    row = db_fetchone(
        "SELECT id FROM user_customer_access WHERE customer_id=? AND portal_user_id=?",
        (customer_id, user_id),
    )
    if not row:
        raise ValueError("Toegangstoewijzing niet gevonden")
    db_execute("DELETE FROM user_customer_access WHERE id=?", (row["id"],))
    return {"ok": True, "deleted": True, "customer_id": customer_id, "user_id": user_id}


# ── Integrations CRUD ─────────────────────────────────────────────────────────

def list_integrations(tenant_id: Optional[str] = None) -> List[Dict[str, Any]]:
    if tenant_id:
        return db_fetchall("SELECT * FROM integrations WHERE tenant_id=? ORDER BY integration_type",
                           (tenant_id,))
    return db_fetchall("SELECT * FROM integrations ORDER BY tenant_id, integration_type")


def get_integration(integration_id: str) -> Optional[Dict[str, Any]]:
    return db_fetchone("SELECT * FROM integrations WHERE id=?", (integration_id,))


def upsert_integration(tenant_id: str, integration_type: str,
                       payload: Dict[str, Any]) -> Dict[str, Any]:
    existing = db_fetchone(
        "SELECT id FROM integrations WHERE tenant_id=? AND integration_type=?",
        (tenant_id, integration_type),
    )
    ts = now_iso()
    allowed = {"status", "auth_mode", "gdap_status", "lighthouse_status",
               "app_registration_status", "certificate_status", "last_validated_at", "details_json"}
    fields: Dict[str, Any] = {k: v for k, v in payload.items() if k in allowed}
    if existing:
        fields["updated_at"] = ts
        sql = "UPDATE integrations SET " + ", ".join(f"{k}=?" for k in fields) + " WHERE id=?"
        db_execute(sql, tuple(fields.values()) + (existing["id"],))
        return db_fetchone("SELECT * FROM integrations WHERE id=?", (existing["id"],)) or {}
    iid = str(uuid.uuid4())
    db_execute(
        "INSERT INTO integrations (id, tenant_id, integration_type, status, auth_mode, "
        "gdap_status, lighthouse_status, app_registration_status, certificate_status, "
        "last_validated_at, details_json, created_at, updated_at) "
        "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
        (iid, tenant_id, integration_type,
         fields.get("status", "unknown"), fields.get("auth_mode"),
         fields.get("gdap_status"), fields.get("lighthouse_status"),
         fields.get("app_registration_status"), fields.get("certificate_status"),
         fields.get("last_validated_at"), fields.get("details_json"), ts, ts),
    )
    return db_fetchone("SELECT * FROM integrations WHERE id=?", (iid,)) or {}


# ── Subscriptions (Azure) ─────────────────────────────────────────────────────

def list_subscriptions(tenant_id: str) -> List[Dict[str, Any]]:
    return db_fetchall(
        "SELECT * FROM subscriptions WHERE tenant_id=? ORDER BY display_name", (tenant_id,)
    )


def upsert_subscription(tenant_id: str, payload: Dict[str, Any]) -> Dict[str, Any]:
    azure_sub_id = (payload.get("azure_subscription_id") or "").strip()
    if not azure_sub_id:
        raise ValueError("azure_subscription_id is verplicht")
    existing = db_fetchone(
        "SELECT id FROM subscriptions WHERE tenant_id=? AND azure_subscription_id=?",
        (tenant_id, azure_sub_id),
    )
    ts = now_iso()
    if existing:
        db_execute(
            "UPDATE subscriptions SET display_name=?, state=?, lighthouse_onboarded=?, "
            "management_group=? WHERE id=?",
            (payload.get("display_name"), payload.get("state", "active"),
             1 if payload.get("lighthouse_onboarded") else 0,
             payload.get("management_group"), existing["id"]),
        )
        return db_fetchone("SELECT * FROM subscriptions WHERE id=?", (existing["id"],)) or {}
    sid = str(uuid.uuid4())
    db_execute(
        "INSERT INTO subscriptions (id, tenant_id, azure_subscription_id, display_name, "
        "state, lighthouse_onboarded, management_group, created_at) VALUES (?,?,?,?,?,?,?,?)",
        (sid, tenant_id, azure_sub_id, payload.get("display_name"),
         payload.get("state", "active"),
         1 if payload.get("lighthouse_onboarded") else 0,
         payload.get("management_group"), ts),
    )
    return db_fetchone("SELECT * FROM subscriptions WHERE id=?", (sid,)) or {}


# ── Approval Workflow ─────────────────────────────────────────────────────────

def create_approval(action_log_id: str, requested_by: str,
                    reason: Optional[str] = None) -> Dict[str, Any]:
    if not db_fetchone("SELECT id FROM action_logs WHERE id=?", (action_log_id,)):
        raise ValueError("action_log niet gevonden")
    existing = db_fetchone("SELECT id FROM approvals WHERE action_log_id=?", (action_log_id,))
    if existing:
        raise ValueError("Er is al een approval-verzoek voor deze actie")
    aid = str(uuid.uuid4())
    db_execute(
        "INSERT INTO approvals (id, action_log_id, approval_status, requested_by, requested_at, reason) "
        "VALUES (?, ?, 'pending', ?, ?, ?)",
        (aid, action_log_id, requested_by, now_iso(), reason),
    )
    return db_fetchone("SELECT * FROM approvals WHERE id=?", (aid,)) or {}


def list_approvals(status: Optional[str] = None, limit: int = 100) -> List[Dict[str, Any]]:
    if status:
        return db_fetchall(
            "SELECT ap.*, al.tenant_id, al.section, al.action_type "
            "FROM approvals ap LEFT JOIN action_logs al ON al.id=ap.action_log_id "
            "WHERE ap.approval_status=? ORDER BY ap.requested_at DESC LIMIT ?",
            (status, min(limit, 500)),
        )
    return db_fetchall(
        "SELECT ap.*, al.tenant_id, al.section, al.action_type "
        "FROM approvals ap LEFT JOIN action_logs al ON al.id=ap.action_log_id "
        "ORDER BY ap.requested_at DESC LIMIT ?",
        (min(limit, 500),),
    )


def decide_approval(approval_id: str, decision: str,
                    decided_by: str, reason: Optional[str] = None) -> Dict[str, Any]:
    ap = db_fetchone("SELECT * FROM approvals WHERE id=?", (approval_id,))
    if not ap:
        raise ValueError("Approval niet gevonden")
    if ap["approval_status"] != "pending":
        raise ValueError(f"Approval is al afgehandeld: {ap['approval_status']}")
    if decision not in ("approved", "rejected"):
        raise ValueError("decision moet 'approved' of 'rejected' zijn")
    db_execute(
        "UPDATE approvals SET approval_status=?, approved_by=?, approved_at=?, reason=? WHERE id=?",
        (decision, decided_by, now_iso(), reason or ap["reason"], approval_id),
    )
    return db_fetchone("SELECT * FROM approvals WHERE id=?", (approval_id,)) or {}


# ── Customer Health ───────────────────────────────────────────────────────────

def get_customer_health(customer_id: str) -> Dict[str, Any]:
    """Samengevat health-overzicht per klant op basis van beschikbare data."""
    c = db_fetchone("SELECT * FROM customers WHERE id=?", (customer_id,))
    if not c:
        raise ValueError("Klant niet gevonden")
    tenants = db_fetchall(
        "SELECT id, tenant_name, tenant_guid, status FROM tenants WHERE customer_id=? AND is_active=1",
        (customer_id,),
    )
    health: Dict[str, Any] = {
        "customer_id": customer_id,
        "customer_name": c["name"],
        "status": c["status"],
        "tenant_count": len(tenants),
        "tenants": [],
        "_generated_at": now_iso(),
    }
    for t in tenants:
        tid = t["id"]
        last_run = db_fetchone(
            "SELECT completed_at, score_overall, critical_count, warning_count "
            "FROM assessment_runs WHERE tenant_id=? AND status='completed' "
            "ORDER BY completed_at DESC LIMIT 1",
            (tid,),
        )
        integrations = db_fetchall(
            "SELECT integration_type, status, gdap_status FROM integrations WHERE tenant_id=?",
            (tid,),
        )
        health["tenants"].append({
            "tenant_id": tid,
            "tenant_name": t["tenant_name"],
            "status": t["status"],
            "last_assessment": last_run,
            "integrations": integrations,
        })
    return health


# ── Tenant Onboarding Status ──────────────────────────────────────────────────

def get_tenant_onboarding_status(tenant_id: str) -> Dict[str, Any]:
    """Bepaalt de onboarding-voortgang voor een tenant op basis van bekende data."""
    t = db_fetchone("SELECT * FROM tenants WHERE id=?", (tenant_id,))
    if not t:
        raise ValueError("Tenant niet gevonden")
    integrations = {
        r["integration_type"]: r
        for r in db_fetchall("SELECT * FROM integrations WHERE tenant_id=?", (tenant_id,))
    }
    has_gdap = integrations.get("gdap", {}).get("gdap_status") == "active"
    has_app  = integrations.get("customer_app", {}).get("app_registration_status") == "active"
    has_lighthouse = integrations.get("lighthouse", {}).get("lighthouse_status") == "active"
    last_run = db_fetchone(
        "SELECT id, completed_at FROM assessment_runs WHERE tenant_id=? AND status='completed' "
        "ORDER BY completed_at DESC LIMIT 1",
        (tenant_id,),
    )
    steps = [
        {"key": "tenant_registered",  "label": "Tenant geregistreerd",         "done": bool(t.get("tenant_guid"))},
        {"key": "gdap_configured",    "label": "GDAP-relatie geconfigureerd",   "done": has_gdap},
        {"key": "app_registered",     "label": "App-registratie geconfigureerd","done": has_app},
        {"key": "assessment_run",     "label": "Eerste assessment uitgevoerd",  "done": bool(last_run)},
        {"key": "lighthouse_onboarded","label": "Azure Lighthouse geconfigureerd","done": has_lighthouse},
    ]
    done_count = sum(1 for s in steps if s["done"])
    return {
        "tenant_id": tenant_id,
        "tenant_name": t["tenant_name"],
        "completion_pct": round((done_count / len(steps)) * 100),
        "steps": steps,
        "last_assessment_at": last_run["completed_at"] if last_run else None,
        "_generated_at": now_iso(),
    }


# ── Azure Resource Snapshots ──────────────────────────────────────────────────
def list_azure_snapshots(tenant_id: str, subscription_id: Optional[str] = None) -> List[Dict[str, Any]]:
    if subscription_id:
        return db_fetchall(
            "SELECT id, tenant_id, subscription_id, section, subsection, generated_at, summary_json "
            "FROM azure_resource_snapshots WHERE tenant_id=? AND subscription_id=? ORDER BY generated_at DESC",
            (tenant_id, subscription_id),
        )
    return db_fetchall(
        "SELECT id, tenant_id, subscription_id, section, subsection, generated_at, summary_json "
        "FROM azure_resource_snapshots WHERE tenant_id=? ORDER BY generated_at DESC LIMIT 200",
        (tenant_id,),
    )


def upsert_azure_snapshot(tenant_id: str, section: str, subsection: str, payload: Dict[str, Any]) -> Dict[str, Any]:
    subscription_id = (payload.get("subscription_id") or "").strip() or None
    snap_id = str(uuid.uuid4())
    summary = {k: v for k, v in payload.items() if isinstance(v, (str, int, float, bool)) and k != "data"}
    db_execute(
        "INSERT OR REPLACE INTO azure_resource_snapshots "
        "(id, tenant_id, subscription_id, section, subsection, generated_at, stale_after_at, data_json, summary_json) "
        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
        (
            snap_id, tenant_id, subscription_id,
            section.lower(), subsection.lower(),
            now_iso(),
            payload.get("stale_after_at") or None,
            json.dumps(payload.get("data") or payload, ensure_ascii=False),
            json.dumps(summary, ensure_ascii=False),
        ),
    )
    return db_fetchone("SELECT * FROM azure_resource_snapshots WHERE id=?", (snap_id,)) or {}


# ── Alert Snapshots ───────────────────────────────────────────────────────────
def list_alert_snapshots(tenant_id: str, alert_type: Optional[str] = None) -> List[Dict[str, Any]]:
    if alert_type:
        return db_fetchall(
            "SELECT id, tenant_id, alert_type, generated_at, summary_json "
            "FROM alert_snapshots WHERE tenant_id=? AND alert_type=? ORDER BY generated_at DESC LIMIT 100",
            (tenant_id, alert_type),
        )
    return db_fetchall(
        "SELECT id, tenant_id, alert_type, generated_at, summary_json "
        "FROM alert_snapshots WHERE tenant_id=? ORDER BY generated_at DESC LIMIT 200",
        (tenant_id,),
    )


def upsert_alert_snapshot(tenant_id: str, alert_type: str, payload: Dict[str, Any]) -> Dict[str, Any]:
    snap_id = str(uuid.uuid4())
    summary = {k: v for k, v in payload.items() if isinstance(v, (str, int, float, bool)) and k != "data"}
    db_execute(
        "INSERT OR REPLACE INTO alert_snapshots "
        "(id, tenant_id, alert_type, generated_at, data_json, summary_json) "
        "VALUES (?, ?, ?, ?, ?, ?)",
        (
            snap_id, tenant_id, alert_type.lower(),
            now_iso(),
            json.dumps(payload.get("data") or payload, ensure_ascii=False),
            json.dumps(summary, ensure_ascii=False),
        ),
    )
    return db_fetchone("SELECT * FROM alert_snapshots WHERE id=?", (snap_id,)) or {}


# ── Cost Snapshots ────────────────────────────────────────────────────────────
def list_cost_snapshots(tenant_id: str, subscription_id: Optional[str] = None) -> List[Dict[str, Any]]:
    if subscription_id:
        return db_fetchall(
            "SELECT id, tenant_id, subscription_id, period_start, period_end, generated_at, summary_json "
            "FROM cost_snapshots WHERE tenant_id=? AND subscription_id=? ORDER BY period_start DESC LIMIT 24",
            (tenant_id, subscription_id),
        )
    return db_fetchall(
        "SELECT id, tenant_id, subscription_id, period_start, period_end, generated_at, summary_json "
        "FROM cost_snapshots WHERE tenant_id=? ORDER BY period_start DESC LIMIT 48",
        (tenant_id,),
    )


def upsert_cost_snapshot(tenant_id: str, payload: Dict[str, Any]) -> Dict[str, Any]:
    period_start = (payload.get("period_start") or "").strip()
    period_end = (payload.get("period_end") or "").strip()
    if not period_start or not period_end:
        raise ValueError("period_start en period_end zijn verplicht")
    subscription_id = (payload.get("subscription_id") or "").strip() or None
    snap_id = str(uuid.uuid4())
    summary = {k: v for k, v in payload.items() if isinstance(v, (str, int, float, bool)) and k not in ("data", "period_start", "period_end")}
    db_execute(
        "INSERT OR REPLACE INTO cost_snapshots "
        "(id, tenant_id, subscription_id, period_start, period_end, generated_at, data_json, summary_json) "
        "VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
        (
            snap_id, tenant_id, subscription_id,
            period_start, period_end,
            now_iso(),
            json.dumps(payload.get("data") or payload, ensure_ascii=False),
            json.dumps(summary, ensure_ascii=False),
        ),
    )
    return db_fetchone("SELECT * FROM cost_snapshots WHERE id=?", (snap_id,)) or {}


# ── Job Queue ─────────────────────────────────────────────────────────────────
def enqueue_job(job_type: str, tenant_id: Optional[str] = None,
                payload: Optional[Dict[str, Any]] = None, priority: int = 5,
                scheduled_at: Optional[str] = None) -> Dict[str, Any]:
    job_id = str(uuid.uuid4())
    db_execute(
        "INSERT INTO job_queue (id, job_type, tenant_id, payload_json, status, priority, "
        "attempt_count, max_attempts, scheduled_at, created_at) VALUES (?,?,?,?,?,?,0,3,?,?)",
        (
            job_id, job_type, tenant_id,
            json.dumps(payload or {}, ensure_ascii=False),
            "pending", priority,
            scheduled_at or now_iso(),
            now_iso(),
        ),
    )
    return db_fetchone("SELECT * FROM job_queue WHERE id=?", (job_id,)) or {}


def list_jobs(tenant_id: Optional[str] = None, status: Optional[str] = None,
              limit: int = 100) -> List[Dict[str, Any]]:
    where: List[str] = []
    params: List[Any] = []
    if tenant_id:
        where.append("tenant_id=?"); params.append(tenant_id)
    if status:
        where.append("status=?"); params.append(status)
    clause = f"WHERE {' AND '.join(where)}" if where else ""
    return db_fetchall(
        f"SELECT id, job_type, tenant_id, status, priority, attempt_count, "
        f"scheduled_at, started_at, completed_at, error_message, created_at "
        f"FROM job_queue {clause} ORDER BY priority ASC, scheduled_at ASC LIMIT ?",
        tuple(params) + (limit,),
    )


def cancel_job(job_id: str) -> Dict[str, Any]:
    row = db_fetchone("SELECT * FROM job_queue WHERE id=?", (job_id,))
    if not row:
        raise ValueError("Job niet gevonden")
    if row["status"] not in ("pending", "failed"):
        raise ValueError(f"Job kan niet worden geannuleerd met status '{row['status']}'")
    db_execute("UPDATE job_queue SET status='cancelled', completed_at=? WHERE id=?",
               (now_iso(), job_id))
    return db_fetchone("SELECT * FROM job_queue WHERE id=?", (job_id,)) or {}


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

    kb_asset_id = payload.get("kb_asset_id") or None
    kb_asset_name = (payload.get("kb_asset_name") or "").strip() or None
    db_execute(
        """
        INSERT INTO finding_actions
        (id, tenant_id, run_id, finding_key, title, severity, owner, status, due_date, notes, evidence, kb_asset_id, kb_asset_name, created_at, updated_at, closed_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
            int(kb_asset_id) if kb_asset_id is not None else None,
            kb_asset_name,
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
    allowed = {"owner", "status", "due_date", "notes", "evidence", "title", "severity", "kb_asset_id", "kb_asset_name"}
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


def list_actions_for_asset(tenant_id: str, asset_id: int) -> List[Dict[str, Any]]:
    """Geef alle bevindingen terug die gekoppeld zijn aan een specifiek KB-asset."""
    return db_fetchall(
        "SELECT * FROM finding_actions WHERE tenant_id=? AND kb_asset_id=? "
        "ORDER BY CASE status WHEN 'open' THEN 0 WHEN 'in_progress' THEN 1 WHEN 'done' THEN 2 ELSE 3 END, updated_at DESC",
        (tenant_id, asset_id),
    )


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
        r["json_manifest_path"] = _run_json_manifest_path(RUNS_DIR / r["id"])
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
        row["json_manifest_path"] = _run_json_manifest_path(RUNS_DIR / row["id"])
    return row


def tenant_overview(tenant_id: str) -> Dict[str, Any]:
    tenant = db_fetchone("SELECT * FROM tenants WHERE id=?", (tenant_id,))
    if not tenant:
        return {"hasData": False}
    latest = _latest_completed_run_for_tenant(tenant_id)
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
    tenant = db_fetchone("SELECT id, customer_name, tenant_name, tenant_guid FROM tenants WHERE id=?", (tenant_id,))
    if not tenant:
        raise ValueError("Tenant niet gevonden")
    # Voorkom concurrent runs voor dezelfde tenant
    active = db_fetchone(
        "SELECT id FROM assessment_runs WHERE tenant_id=? AND status IN ('queued','running') LIMIT 1",
        (tenant_id,),
    )
    if active:
        raise ValueError(f"Er loopt al een actief assessment voor deze tenant (run: {active['id'][:8]}…)")

    scan_type = str(payload.get("scan_type") or "full")
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

    # Tenant-veiligheid: voorkom dat een run start met context van een andere tenant
    selected_tenant_guid = (tenant.get("tenant_guid") or "").strip().lower()
    request_auth_tenant = str(payload.get("auth_tenant_id") or payload.get("authTenantId") or "").strip().lower()
    if run_mode == "script":
        if not selected_tenant_guid:
            raise ValueError(
                "Voor script-runs is tenant_guid verplicht op de geselecteerde tenant. "
                "Vul de Tenant GUID in bij Tenant-instellingen voordat je de assessment start."
            )
        if request_auth_tenant and request_auth_tenant != selected_tenant_guid:
            raise ValueError(
                "Je bent aangemeld op een andere tenant dan de geselecteerde tenant. "
                "Wissel Microsoft-tenant en start daarna opnieuw."
            )

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
            if re.fullmatch(r"/api/capabilities/[^/]+", path):
                tenant_id = path.split("/")[3]
                return self._json(200, get_tenant_capabilities(tenant_id))
            if re.fullmatch(r"/api/capabilities/[^/]+/[^/]+/[^/]+", path):
                parts = path.split("/")
                tenant_id, section, subsection = parts[3], parts[4], parts[5]
                return self._json(200, {"ok": True, "capability": _build_capability_status(tenant_id, section, subsection)})
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
                    return self._json(404, {"error": "Tenant niet gevonden", "error_code": "not_found"})
                return self._json(200, tenant)
            if re.fullmatch(r"/api/tenants/[^/]+/auth-config", path):
                if _sess.get("role") != "admin":
                    return self._json(403, {"error": "Onvoldoende rechten."})
                tenant_id = path.split("/")[3]
                if not db_fetchone("SELECT id FROM tenants WHERE id=?", (tenant_id,)):
                    return self._json(404, {"error": "Tenant niet gevonden", "error_code": "not_found"})
                return self._json(200, get_tenant_auth_profile(tenant_id, include_secret=False))
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
            if re.fullmatch(r"/api/tenants/[^/]+/onboarding", path):
                tenant_id = path.split("/")[3]
                if not db_fetchone("SELECT id FROM tenants WHERE id=?", (tenant_id,)):
                    return self._json(404, {"error": "Tenant niet gevonden", "error_code": "not_found"})
                return self._json(200, get_tenant_onboarding_status(tenant_id))
            if re.fullmatch(r"/api/tenants/[^/]+/subscriptions", path):
                tenant_id = path.split("/")[3]
                if not db_fetchone("SELECT id FROM tenants WHERE id=?", (tenant_id,)):
                    return self._json(404, {"error": "Tenant niet gevonden", "error_code": "not_found"})
                return self._json(200, {"items": list_subscriptions(tenant_id), "tenant_id": tenant_id})
            if re.fullmatch(r"/api/tenants/[^/]+/integrations", path):
                tenant_id = path.split("/")[3]
                if not db_fetchone("SELECT id FROM tenants WHERE id=?", (tenant_id,)):
                    return self._json(404, {"error": "Tenant niet gevonden", "error_code": "not_found"})
                return self._json(200, {"items": list_integrations(tenant_id=tenant_id), "tenant_id": tenant_id})
            if re.fullmatch(r"/api/tenants/[^/]+/snapshots", path):
                tenant_id = path.split("/")[3]
                if not db_fetchone("SELECT id FROM tenants WHERE id=?", (tenant_id,)):
                    return self._json(404, {"error": "Tenant niet gevonden", "error_code": "not_found"})
                section_f = qs.get("section", [None])[0]
                subsection_f = qs.get("subsection", [None])[0]
                where = ["tenant_id=?"]
                params: List[Any] = [tenant_id]
                if section_f:
                    where.append("section=?")
                    params.append(section_f.lower())
                if subsection_f:
                    where.append("subsection=?")
                    params.append(subsection_f.lower())
                rows = db_fetchall(
                    f"SELECT id, tenant_id, section, subsection, source_type, generated_at, summary_json, assessment_run_id "
                    f"FROM m365_snapshots WHERE {' AND '.join(where)} ORDER BY generated_at DESC LIMIT 200",
                    tuple(params),
                )
                return self._json(200, {"items": rows, "tenant_id": tenant_id})
            if re.fullmatch(r"/api/tenants/[^/]+/snapshots/[^/]+/[^/]+", path):
                parts = path.split("/")
                tenant_id = parts[3]
                section_p = parts[5].lower()
                subsection_p = parts[6].lower()
                row = db_fetchone(
                    "SELECT * FROM m365_snapshots WHERE tenant_id=? AND section=? AND subsection=? "
                    "ORDER BY generated_at DESC LIMIT 1",
                    (tenant_id, section_p, subsection_p),
                )
                if not row:
                    return self._json(404, {"error": "Snapshot niet gevonden", "error_code": "not_found"})
                # Parse data_json inline for convenience
                if row.get("data_json"):
                    try:
                        row = dict(row)
                        row["data"] = json.loads(row.pop("data_json"))
                    except Exception:
                        pass
                return self._json(200, row)
            # ── Azure Resource Snapshots (GET) ────────────────────────────────
            if re.fullmatch(r"/api/tenants/[^/]+/azure-snapshots", path):
                tid = path.split("/")[3]
                sub_f = qs.get("subscription_id", [None])[0]
                return self._json(200, {"items": list_azure_snapshots(tid, sub_f), "tenant_id": tid})
            # ── Alert Snapshots (GET) ─────────────────────────────────────────
            if re.fullmatch(r"/api/tenants/[^/]+/alert-snapshots", path):
                tid = path.split("/")[3]
                atype_f = qs.get("alert_type", [None])[0]
                return self._json(200, {"items": list_alert_snapshots(tid, atype_f), "tenant_id": tid})
            # ── Cost Snapshots (GET) ──────────────────────────────────────────
            if re.fullmatch(r"/api/tenants/[^/]+/cost-snapshots", path):
                tid = path.split("/")[3]
                sub_f = qs.get("subscription_id", [None])[0]
                return self._json(200, {"items": list_cost_snapshots(tid, sub_f), "tenant_id": tid})
            # ── Geplande assessments (GET) ────────────────────────────────────
            if path == "/api/scheduled-runs":
                items = db_fetchall(
                    "SELECT id, job_type, tenant_id, status, scheduled_at, created_at, payload_json "
                    "FROM job_queue WHERE job_type='assessment' ORDER BY scheduled_at ASC LIMIT 100"
                )
                for item in items:
                    try:
                        item["payload"] = json.loads(item.pop("payload_json") or "{}")
                    except Exception:
                        item["payload"] = {}
                return self._json(200, {"items": items})
            # ── Job Queue (GET) ───────────────────────────────────────────────
            if path == "/api/jobs":
                tid_f = qs.get("tenant_id", [None])[0]
                status_f = qs.get("status", [None])[0]
                limit_f = int(qs.get("limit", ["100"])[0])
                return self._json(200, {"items": list_jobs(tid_f, status_f, limit_f)})
            if re.fullmatch(r"/api/tenants/[^/]+/jobs", path):
                tid = path.split("/")[3]
                status_f = qs.get("status", [None])[0]
                return self._json(200, {"items": list_jobs(tid, status_f)})
            # ── Customer API (GET) ─────────────────────────────────────────────
            if path == "/api/customers":
                status_filter = qs.get("status", [None])[0]
                return self._json(200, {"items": list_customers(status_filter)})
            if re.fullmatch(r"/api/customers/[^/]+/tenants", path):
                cid = path.split("/")[3]
                if not db_fetchone("SELECT id FROM customers WHERE id=?", (cid,)):
                    http_s, pl = _api_error("not_found", "Klant niet gevonden", 404)
                    return self._json(http_s, pl)
                tenants = db_fetchall(
                    "SELECT * FROM tenants WHERE customer_id=? AND is_active=1 ORDER BY tenant_name", (cid,)
                )
                return self._json(200, {"items": tenants, "customer_id": cid})
            if re.fullmatch(r"/api/customers/[^/]+", path):
                c = get_customer(path.split("/")[3])
                if not c:
                    http_s, pl = _api_error("not_found", "Klant niet gevonden", 404)
                    return self._json(http_s, pl)
                return self._json(200, c)
            if re.fullmatch(r"/api/customers/[^/]+/services", path):
                cid = path.split("/")[3]
                if not db_fetchone("SELECT id FROM customers WHERE id=?", (cid,)):
                    http_s, pl = _api_error("not_found", "Klant niet gevonden", 404)
                    return self._json(http_s, pl)
                svcs = db_fetchall(
                    "SELECT * FROM customer_services WHERE customer_id=? ORDER BY service_key", (cid,)
                )
                return self._json(200, {"items": svcs, "customer_id": cid})
            if re.fullmatch(r"/api/customers/[^/]+/health", path):
                cid = path.split("/")[3]
                h = get_customer_health(cid)
                if not h:
                    http_s, pl = _api_error("not_found", "Klant niet gevonden", 404)
                    return self._json(http_s, pl)
                return self._json(200, h)
            if re.fullmatch(r"/api/customers/[^/]+/onboarding", path):
                cid = path.split("/")[3]
                if not db_fetchone("SELECT id FROM customers WHERE id=?", (cid,)):
                    http_s, pl = _api_error("not_found", "Klant niet gevonden", 404)
                    return self._json(http_s, pl)
                tenants = db_fetchall(
                    "SELECT id FROM tenants WHERE customer_id=? AND is_active=1", (cid,)
                )
                statuses = [get_tenant_onboarding_status(t["id"]) for t in tenants]
                return self._json(200, {"customer_id": cid, "tenants": statuses})
            if re.fullmatch(r"/api/customers/[^/]+/subscriptions", path):
                cid = path.split("/")[3]
                if not db_fetchone("SELECT id FROM customers WHERE id=?", (cid,)):
                    http_s, pl = _api_error("not_found", "Klant niet gevonden", 404)
                    return self._json(http_s, pl)
                tenants = db_fetchall(
                    "SELECT id FROM tenants WHERE customer_id=? AND is_active=1", (cid,)
                )
                all_subs = []
                for t in tenants:
                    all_subs.extend(list_subscriptions(t["id"]))
                return self._json(200, {"items": all_subs, "customer_id": cid})
            if re.fullmatch(r"/api/customers/[^/]+/access", path):
                cid = path.split("/")[3]
                if not db_fetchone("SELECT id FROM customers WHERE id=?", (cid,)):
                    http_s, pl = _api_error("not_found", "Klant niet gevonden", 404)
                    return self._json(http_s, pl)
                return self._json(200, {"items": list_user_customer_access(customer_id=cid)})
            # ── Portal Roles (GET) ─────────────────────────────────────────────
            if path == "/api/portal-roles":
                return self._json(200, {"items": list_portal_roles()})
            # ── Approvals (GET) ───────────────────────────────────────────────
            if path == "/api/approvals":
                status_f = qs.get("status", [None])[0]
                limit_f = int(qs.get("limit", ["100"])[0])
                return self._json(200, {"items": list_approvals(status=status_f, limit=limit_f)})
            # ── Integrations (GET) ────────────────────────────────────────────
            if re.fullmatch(r"/api/integrations/[^/]+", path):
                tid = path.split("/")[3]
                return self._json(200, {"items": list_integrations(tenant_id=tid), "tenant_id": tid})
            if path == "/api/integrations":
                return self._json(200, {"items": list_integrations()})
            # ── Audit API (GET) ────────────────────────────────────────────────
            if path == "/api/audit":
                logs = list_audit_logs(
                    tenant_id=qs.get("tenant_id", [None])[0],
                    user_email=qs.get("user_email", [None])[0],
                    action=qs.get("action", [None])[0],
                    date_from=qs.get("from", [None])[0],
                    date_to=qs.get("to", [None])[0],
                    limit=int(qs.get("limit", ["200"])[0]),
                )
                return self._json(200, {"items": logs, "count": len(logs), "_generated_at": now_iso()})
            if path == "/api/runs":
                tenant_id = qs.get("tenant_id", [None])[0]
                limit = int(qs.get("limit", ["100"])[0])
                return self._json(200, {"items": list_runs(tenant_id, limit)})
            if re.fullmatch(r"/api/runs/[^/]+", path):
                run = get_run(path.split("/")[3])
                if not run:
                    return self._json(404, {"error": "Run niet gevonden", "error_code": "not_found"})
                return self._json(200, run)
            if re.fullmatch(r"/api/runs/[^/]+/assessment-json", path):
                run_id = path.split("/")[3]
                try:
                    return self._json(200, _assessment_json_report_for_run(run_id))
                except ValueError as exc:
                    return self._json(404, {"ok": False, "error": str(exc)})
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
                    return self._json(404, {"error": "Run niet gevonden", "error_code": "not_found"})
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
            # ── Remediation routes (GET) ──
            if re.fullmatch(r"/api/remediate/[^/]+/catalog", path):
                tenant_id = path.split("/")[3]
                category = qs.get("category", [None])[0]
                return self._json(200, {"items": get_remediation_catalog(category)})
            if re.fullmatch(r"/api/remediate/[^/]+/history", path):
                tenant_id = path.split("/")[3]
                limit = int(qs.get("limit", ["100"])[0])
                return self._json(200, {"items": list_remediation_history(tenant_id, limit)})

            # ── User Management routes (GET) ──
            if re.fullmatch(r"/api/m365/[^/]+/users", path):
                tenant_id = path.split("/")[3]
                filter_q = qs.get("filter", [None])[0]
                strict_live = qs.get("strict_live", ["0"])[0] in {"1", "true", "yes"}
                live_error = None
                try:
                    result = _run_user_mgmt(tenant_id, "list-users", {"filter": filter_q})
                    r = result.get("result") or {}
                    if result.get("ok") and r.get("ok") is not False and r.get("users"):
                        r["users"] = [_normalize_user_license_payload(u) for u in (r.get("users") or []) if isinstance(u, dict)]
                        return self._json(200, r)
                    live_error = (r.get("error") if isinstance(r, dict) else None) or result.get("message")
                except Exception as exc:
                    live_error = str(exc)
                if strict_live:
                    return self._json(502, {"error": live_error or "Live gebruikersscan mislukt. Controleer tenantverbinding of app-autorisatie en probeer opnieuw."})
                # Fallback: serve from assessment snapshot
                snap = _latest_assessment_snapshot_for_tenant(tenant_id)
                users = _snapshot_as_users(tenant_id)
                if users:
                    return self._json(200, {
                        "ok": True,
                        "users": users,
                        "counts": snap.get("assessment_user_counts") or {},
                        "_source": "assessment_snapshot",
                    })
                return self._json(502, {"error": "Geen gebruikersdata beschikbaar. Voer een assessment uit of controleer de tenant auth-configuratie."})
            if re.fullmatch(r"/api/m365/[^/]+/users/[^/]+", path):
                parts = path.split("/")
                tenant_id = parts[3]; user_id = parts[5]
                try:
                    result = _run_user_mgmt(tenant_id, "get-user", {"user_id": user_id})
                    if result.get("ok"):
                        payload = result["result"] or {}
                        if isinstance(payload.get("user"), dict):
                            payload["user"] = _normalize_user_license_payload(payload["user"])
                        return self._json(200, payload)
                except Exception:
                    pass
                for user in _snapshot_as_users(tenant_id):
                    if not isinstance(user, dict):
                        continue
                    if user.get("id") == user_id or user.get("userPrincipalName") == user_id:
                        return self._json(200, {"ok": True, "user": user, "_source": "assessment_snapshot"})
                return self._json(404, {"error": "Gebruikersdetail niet beschikbaar"})
            if re.fullmatch(r"/api/m365/[^/]+/licenses", path):
                tenant_id = path.split("/")[3]
                try:
                    result = _run_user_mgmt(tenant_id, "list-licenses", {})
                    r = result.get("result") or {}
                    if result.get("ok") and r.get("ok") is not False and r.get("licenses"):
                        r["licenses"] = [_normalize_license_payload(lic) for lic in (r.get("licenses") or []) if isinstance(lic, dict)]
                        return self._json(200, r)
                except Exception:
                    pass
                # Fallback: serve from assessment snapshot
                licenses = _snapshot_as_licenses(tenant_id)
                if licenses:
                    return self._json(200, _attach_source_meta({"ok": True, "licenses": licenses}, "assessment_snapshot", tenant_id=tenant_id))
                return self._json(502, {"error": "Geen licentiedata beschikbaar. Voer een assessment uit."})
            if re.fullmatch(r"/api/m365/[^/]+/provisioning-history", path):
                tenant_id = path.split("/")[3]
                limit = int(qs.get("limit", ["100"])[0])
                return self._json(200, {"items": list_provisioning_history(tenant_id, limit)})

            # ── Baseline routes (GET) ──
            if path == "/api/baselines":
                return self._json(200, {"items": list_baselines()})
            if re.fullmatch(r"/api/baselines/[^/]+", path):
                bid = path.split("/")[3]
                row = get_baseline(bid)
                if not row:
                    return self._json(404, {"error": "Baseline niet gevonden", "error_code": "not_found"})
                return self._json(200, row)
            if re.fullmatch(r"/api/baselines/[^/]+/assignments", path):
                bid = path.split("/")[3]
                return self._json(200, {"items": list_assignments(baseline_id=bid)})
            if re.fullmatch(r"/api/baselines/[^/]+/history", path):
                bid = path.split("/")[3]
                limit = int(qs.get("limit", ["100"])[0])
                return self._json(200, {"items": list_baseline_history(baseline_id=bid, limit=limit)})
            if path == "/api/baselines/assignments/all":
                return self._json(200, {"items": list_assignments()})

            # ── Intune routes (GET) ──
            if re.fullmatch(r"/api/intune/[^/]+/devices", path):
                tid = path.split("/")[3]
                try:
                    data = _run_intune_ps(tid, "list-devices", {})
                    if data.get("ok") is not False and data.get("devices"):
                        return self._json(200, data)
                except Exception:
                    pass
                devices = _snapshot_as_intune_devices(tid)
                if devices:
                    return self._json(200, _attach_source_meta({"ok": True, "devices": devices}, "assessment_snapshot", tenant_id=tid))
                return self._json(200, _attach_source_meta({"ok": True, "devices": []}, "assessment_snapshot", tenant_id=tid))
            if re.fullmatch(r"/api/intune/[^/]+/devices/[^/]+", path):
                parts = path.split("/")
                tid, did = parts[3], parts[5]
                try:
                    data = _run_intune_ps(tid, "get-device", {"device_id": did})
                    if data.get("ok") is not False:
                        return self._json(200, data)
                except Exception:
                    pass
                devices = _snapshot_as_intune_devices(tid)
                device = next((d for d in devices if d.get("id") == did or d.get("deviceName") == did), None)
                if device:
                    return self._json(200, _attach_source_meta({"ok": True, "device": device}, "assessment_snapshot", tenant_id=tid))
                return self._json(404, {"error": "Apparaat niet gevonden"})
            if re.fullmatch(r"/api/intune/[^/]+/compliance", path):
                tid = path.split("/")[3]
                try:
                    data = _run_intune_ps(tid, "list-compliance", {})
                    if data.get("ok") is not False and data.get("policies"):
                        return self._json(200, data)
                except Exception:
                    pass
                policies = _snapshot_as_intune_compliance(tid)
                return self._json(200, _attach_source_meta({"ok": True, "policies": policies}, "assessment_snapshot", tenant_id=tid))
            if re.fullmatch(r"/api/intune/[^/]+/config", path):
                tid = path.split("/")[3]
                try:
                    data = _run_intune_ps(tid, "list-config", {})
                    if data.get("ok") is not False and data.get("profiles"):
                        return self._json(200, data)
                except Exception:
                    pass
                profiles = _snapshot_as_intune_config(tid)
                return self._json(200, _attach_source_meta({"ok": True, "profiles": profiles}, "assessment_snapshot", tenant_id=tid))
            if re.fullmatch(r"/api/intune/[^/]+/summary", path):
                tid = path.split("/")[3]
                try:
                    data = _run_intune_ps(tid, "get-compliance-summary", {})
                    if data.get("ok") is not False and ("score" in data or "total" in data):
                        return self._json(200, data)
                except Exception:
                    pass
                summary = _snapshot_as_intune_summary(tid)
                if summary:
                    return self._json(200, _attach_source_meta(summary, "assessment_snapshot", tenant_id=tid))
                return self._json(200, {"ok": False, "error": "Intune data niet beschikbaar"})
            if re.fullmatch(r"/api/intune/[^/]+/history", path):
                tid = path.split("/")[3]
                limit = int(qs.get("limit", ["50"])[0])
                return self._json(200, {"items": list_intune_history(tid, limit)})

            # ── Backup routes (GET) ──
            if re.fullmatch(r"/api/backup/[^/]+/summary", path):
                tid = path.split("/")[3]
                try:
                    data = _run_backup_ps(tid, "get-summary")
                    if data.get("ok"):
                        return self._json(200, data)
                except Exception:
                    pass
                sp = _snapshot_as_sharepoint_sites(tid)
                od = _snapshot_as_onedrive_backup(tid)
                return self._json(200, _attach_source_meta({
                    "ok": True,
                    "serviceStatus": "assessment_snapshot",
                    "sharePoint": {"policyCount": 1 if sp else 0, "resourceCount": len(sp)},
                    "oneDrive": {"policyCount": 1 if (od.get("policies") or []) else 0, "resourceCount": len((od.get("policies") or [{}])[0].get("drives") or []) if od.get("policies") else 0},
                    "exchange": {"policyCount": 0, "resourceCount": 0},
                }, "assessment_snapshot", tenant_id=tid))
            if re.fullmatch(r"/api/backup/[^/]+/status", path):
                tid = path.split("/")[3]
                data = _run_backup_ps(tid, "get-status")
                return self._json(200, data)
            if re.fullmatch(r"/api/backup/[^/]+/sharepoint", path):
                tid = path.split("/")[3]
                try:
                    data = _run_backup_ps(tid, "list-sharepoint")
                    if data.get("ok"):
                        return self._json(200, data)
                except Exception:
                    pass
                return self._json(200, _attach_source_meta(_snapshot_as_sharepoint_backup(tid), "assessment_snapshot", tenant_id=tid))
            if re.fullmatch(r"/api/backup/[^/]+/onedrive", path):
                tid = path.split("/")[3]
                try:
                    data = _run_backup_ps(tid, "list-onedrive")
                    if data.get("ok"):
                        return self._json(200, data)
                except Exception:
                    pass
                return self._json(200, _attach_source_meta(_snapshot_as_onedrive_backup(tid), "assessment_snapshot", tenant_id=tid))
            if re.fullmatch(r"/api/backup/[^/]+/exchange", path):
                tid = path.split("/")[3]
                data = _run_backup_ps(tid, "list-exchange")
                return self._json(200, data)
            if re.fullmatch(r"/api/backup/[^/]+/history", path):
                tid = path.split("/")[3]
                limit = int(qs.get("limit", ["50"])[0])
                return self._json(200, {"items": list_backup_history(tid, limit)})

            # ── CA routes (GET) ──
            if re.fullmatch(r"/api/ca/[^/]+/policies", path):
                tid = path.split("/")[3]
                try:
                    data = _run_ca_ps(tid, "list-policies", {})
                    if data.get("ok") and data.get("policies") is not None:
                        return self._json(200, data)
                except Exception:
                    pass
                # Fallback: snapshot CA policies
                policies = _snapshot_as_ca_policies(tid)
                return self._json(200, _attach_source_meta({"ok": True, "policies": policies, "count": len(policies)}, "assessment_snapshot", tenant_id=tid))
            if re.fullmatch(r"/api/ca/[^/]+/policies/[^/]+", path):
                parts = path.split("/")
                tid, pid = parts[3], parts[5]
                return self._json(200, _run_ca_ps(tid, "get-policy", {"policy_id": pid}))
            if re.fullmatch(r"/api/ca/[^/]+/named-locations", path):
                tid = path.split("/")[3]
                return self._json(200, _run_ca_ps(tid, "list-named-locations", {}))
            if re.fullmatch(r"/api/ca/[^/]+/history", path):
                tid = path.split("/")[3]
                limit = int(qs.get("limit", ["50"])[0])
                return self._json(200, {"items": list_ca_history(tid, limit)})

            # ── Compliance routes (GET) ──
            if re.fullmatch(r"/api/compliance/[^/]+/cis", path):
                tid = path.split("/")[3]
                payload = _snapshot_as_cis_data(tid)
                if isinstance(payload, dict):
                    return self._json(200, _attach_source_meta(payload, "assessment_snapshot", tenant_id=tid))
                return self._json(200, _attach_source_meta(
                    {"ok": True, "summary": {"pass": 0, "fail": 0, "warning": 0, "na": 0, "total": 0, "score": 0}, "items": [], "section": "compliance", "subsection": "cis"},
                    "assessment_snapshot", tenant_id=tid))

            if re.fullmatch(r"/api/compliance/[^/]+/zerotrust", path):
                tid = path.split("/")[3]
                folder = _zt_output_folder(tid)
                try:
                    data = _run_zerotrust_ps(tid, "get-status", folder)
                    if data.get("ok"):
                        if data.get("last_report"):
                            results = _run_zerotrust_ps(tid, "get-results", folder)
                            data["results"] = results if isinstance(results, dict) else None
                        return self._json(200, data)
                except Exception as e:
                    return self._json(200, {"ok": True, "module": {"installed": False}, "last_report": None, "error": str(e)})
                return self._json(200, {"ok": True, "module": {"installed": False}, "last_report": None})

            # ── Hybrid Identity routes (GET) ──
            if re.fullmatch(r"/api/hybrid/[^/]+/sync", path):
                tid = path.split("/")[3]
                payload = _snapshot_as_hybrid_sync(tid)
                if isinstance(payload, dict):
                    return self._json(200, _attach_source_meta(payload, "assessment_snapshot", tenant_id=tid))
                return self._json(200, _attach_source_meta(
                    {"ok": True, "summary": {"isHybrid": False, "syncEnabled": False, "authType": "Cloud Only", "totalUsers": 0}, "items": [], "section": "hybrid", "subsection": "sync"},
                    "assessment_snapshot", tenant_id=tid))

            # ── Domains routes (GET) ──
            if re.fullmatch(r"/api/domains/[^/]+/list", path):
                tid = path.split("/")[3]
                try:
                    data = _run_domains_ps(tid, "list-domains", {})
                    if data.get("ok") and data.get("domains") is not None:
                        return self._json(200, data)
                except Exception:
                    pass
                # Fallback: snapshot DomainDnsChecks as domain list
                domains = _snapshot_as_domains(tid)
                return self._json(200, {"ok": True, "domains": domains, "count": len(domains), "_source": "assessment_snapshot"})
            if re.fullmatch(r"/api/domains/[^/]+/analyse", path):
                tid = path.split("/")[3]
                domain = qs.get("domain", [None])[0]
                if not domain:
                    return self._json(400, {"error": "domain parameter vereist"})
                return self._json(200, _run_domains_ps(tid, "analyse-domain", {"domain": domain}))

            # ── Identity routes (GET) ──
            if re.fullmatch(r"/api/identity/[^/]+/mfa", path):
                tid = path.split("/")[3]
                try:
                    data = _run_identity_ps(tid, "list-mfa", {})
                    if data.get("ok") is not False:
                        return self._json(200, data)
                except Exception:
                    pass
                snap = _latest_assessment_snapshot_for_tenant(tid)
                payload = _assessment_json_payload(snap, "identity", "mfa")
                if isinstance(payload, dict):
                    summary = payload.get("summary") or {}
                    return self._json(200, _attach_source_meta({
                        "ok": True,
                        "items": payload.get("items") or [],
                        "count": len(payload.get("items") or []),
                        "enabledMemberUsers": int(summary.get("enabledMemberUsers") or 0),
                        "usersWithMfa": int(summary.get("usersWithMfa") or 0),
                        "usersWithoutMfa": int(summary.get("usersWithoutMfa") or 0),
                        "mfaCoveragePct": summary.get("mfaCoveragePct"),
                        "checkFailed": bool(summary.get("checkFailed")),
                        "notes": ((payload.get("meta") or {}).get("notes") or []),
                    }, "assessment_snapshot", tenant_id=tid))
                return self._json(200, {"ok": False, "items": [], "error": "MFA-data niet beschikbaar"})
            if re.fullmatch(r"/api/identity/[^/]+/guests", path):
                tid = path.split("/")[3]
                return self._json(200, _run_identity_ps(tid, "list-guests", {}))
            if re.fullmatch(r"/api/identity/[^/]+/admin-roles", path):
                tid = path.split("/")[3]
                try:
                    data = _run_identity_ps(tid, "list-admin-roles", {})
                    if data.get("ok") is not False:
                        return self._json(200, data)
                except Exception:
                    pass
                snap = _latest_assessment_snapshot_for_tenant(tid)
                payload = _assessment_json_payload(snap, "identity", "admin-roles")
                if isinstance(payload, dict):
                    items = []
                    for item in payload.get("items") or []:
                        if not isinstance(item, dict):
                            continue
                        items.append({
                            "displayName": _payload_value(item, "DisplayName", "displayName", default=""),
                            "userPrincipalName": _payload_value(item, "UserPrincipalName", "userPrincipalName", default=""),
                            "lastPasswordChange": _payload_value(item, "LastPasswordChange", "lastPasswordChange"),
                            "passwordAgeDays": _payload_value(item, "PasswordAgeDays", "passwordAgeDays"),
                            "status": _payload_value(item, "Status", "status"),
                        })
                    return self._json(200, _attach_source_meta({"ok": True, "items": items, "count": len(items)}, "assessment_snapshot", tenant_id=tid))
                return self._json(200, {"ok": False, "items": [], "error": "Rollen-data niet beschikbaar"})
            if re.fullmatch(r"/api/identity/[^/]+/security-defaults", path):
                tid = path.split("/")[3]
                return self._json(200, _run_identity_ps(tid, "get-security-defaults", {}))
            if re.fullmatch(r"/api/identity/[^/]+/legacy-auth", path):
                tid = path.split("/")[3]
                return self._json(200, _run_identity_ps(tid, "list-legacy-auth", {}))

            # ── Findings & health routes (GET) ──
            if re.fullmatch(r"/api/findings/[^/]+/health", path):
                tid = path.split("/")[3]
                return self._json(200, _get_tenant_health_score(tid))

            if re.fullmatch(r"/api/findings/[^/]+/trend", path):
                tid = path.split("/")[3]
                days = int(qs.get("days", ["30"])[0])
                trend = _get_findings_trend(tid, days)
                return self._json(200, {"ok": True, "tenant_id": tid, "days": days, "trend": trend})

            if re.fullmatch(r"/api/findings/[^/]+/list", path):
                tid = path.split("/")[3]
                domain_filter = qs.get("domain", [None])[0]
                status_filter = qs.get("status", [None])[0]
                limit = min(int(qs.get("limit", ["200"])[0]), 1000)
                where, args = ["f.tenant_id=?"], [tid]
                if domain_filter:
                    where.append("f.domain=?"); args.append(domain_filter)
                if status_filter:
                    where.append("f.status=?"); args.append(status_filter)
                conn = get_conn()
                try:
                    rows = conn.execute(
                        f"""
                        SELECT f.id, f.domain, f.control, f.title, f.status, f.finding,
                               f.impact, f.recommendation, f.service, f.metric_value, f.scanned_at
                        FROM scan_findings f
                        INNER JOIN (
                            SELECT domain, control, MAX(scanned_at) AS max_at
                            FROM scan_findings WHERE tenant_id=?
                            GROUP BY domain, control
                        ) latest ON f.domain=latest.domain AND f.control=latest.control AND f.scanned_at=latest.max_at
                        WHERE {' AND '.join(where)}
                        ORDER BY f.domain, f.status DESC, f.control
                        LIMIT ?
                        """,
                        [tid] + args + [limit],
                    ).fetchall()
                    findings = [dict(r) for r in rows]
                finally:
                    conn.close()
                return self._json(200, {"ok": True, "tenant_id": tid, "findings": findings, "count": len(findings)})

            if re.fullmatch(r"/api/findings/overview", path):
                conn = get_conn()
                try:
                    tenants_rows = conn.execute(
                        "SELECT id, customer_name, tenant_name FROM tenants WHERE is_active=1 ORDER BY customer_name"
                    ).fetchall()
                    overview = []
                    for t in tenants_rows:
                        score_data = _get_tenant_health_score(t["id"])
                        overview.append({
                            "tenant_id": t["id"],
                            "customer_name": t["customer_name"],
                            "tenant_name": t["tenant_name"],
                            "score": score_data.get("score"),
                            "total": score_data.get("total", 0),
                            "ok_count": score_data.get("ok_count", 0),
                            "warning_count": score_data.get("warning_count", 0),
                            "critical_count": score_data.get("critical_count", 0),
                        })
                finally:
                    conn.close()
                return self._json(200, {"ok": True, "tenants": overview})

            # ── MSP Aggregaat (cross-tenant) ──────────────────────────────────
            if path == "/api/msp/aggregate":
                rows = db_fetchall(
                    "SELECT id FROM tenants WHERE is_active=1"
                )
                total = len(rows)
                critical_tenants, no_assess, scores = 0, 0, []
                for r in rows:
                    tid_agg = r["id"]
                    run = _latest_completed_run_for_tenant(tid_agg)
                    if not run:
                        no_assess += 1
                        continue
                    if (run.get("critical_count") or 0) > 0:
                        critical_tenants += 1
                    sc = run.get("score_overall")
                    if sc is not None:
                        scores.append(float(sc))
                avg_score = round(sum(scores) / len(scores), 1) if scores else None
                # Tenants zonder assessment in afgelopen 30 dagen
                from datetime import datetime, timedelta, timezone
                threshold = (datetime.now(timezone.utc) - timedelta(days=30)).isoformat()
                stale_rows = db_fetchall(
                    "SELECT id FROM tenants WHERE is_active=1 AND id NOT IN "
                    "(SELECT tenant_id FROM assessment_runs WHERE status='completed' AND completed_at>=?)",
                    (threshold,),
                )
                stale_count = len(stale_rows)
                return self._json(200, {
                    "total_tenants": total,
                    "tenants_with_critical": critical_tenants,
                    "tenants_no_assessment": no_assess,
                    "tenants_stale_assessment": stale_count,
                    "avg_score": avg_score,
                    "assessed_count": len(scores),
                })
            # ── App Registration routes (GET) ──
            if re.fullmatch(r"/api/apps/[^/]+/registrations", path):
                tid = path.split("/")[3]
                try:
                    data = _run_appregs_ps(tid, "list-appregs", {})
                    if data.get("ok") is not False and ("items" in data or "registrations" in data):
                        return self._json(200, data)
                except Exception:
                    pass
                snap = _latest_assessment_snapshot_for_tenant(tid)
                payload = _assessment_json_payload(snap, "apps", "registrations")
                if isinstance(payload, dict):
                    items = []
                    for item in payload.get("items") or []:
                        if not isinstance(item, dict):
                            continue
                        perms = _payload_value(item, "Permissions", "permissions", default=None)
                        items.append({
                            "displayName": _payload_value(item, "DisplayName", "displayName", default=""),
                            "appId": _payload_value(item, "AppId", "appId", default=""),
                            "objectId": _payload_value(item, "ObjectId", "objectId", default=""),
                            "createdAt": _payload_value(item, "CreatedDateTime", "createdAt"),
                            "secretCount": int(_payload_value(item, "SecretCount", "secretCount", default=0) or 0),
                            "secretExpiration": _payload_value(item, "SecretExpiration", "secretExpiration"),
                            "secretExpirationStatus": _payload_value(item, "SecretExpirationStatus", "secretExpirationStatus"),
                            "certificateCount": int(_payload_value(item, "CertificateCount", "certificateCount", default=0) or 0),
                            "certificateExpiration": _payload_value(item, "CertificateExpiration", "certificateExpiration"),
                            "certificateExpirationStatus": _payload_value(item, "CertificateExpirationStatus", "certificateExpirationStatus"),
                            "permissionCount": int(_payload_value(item, "PermissionCount", "permissionCount", default=0) or 0),
                            "hasEnterpriseApp": bool(_payload_value(item, "HasEnterpriseApp", "hasEnterpriseApp", default=False)),
                            "permissions": list(perms) if isinstance(perms, list) else [],
                        })
                    return self._json(200, _attach_source_meta({"ok": True, "items": items, "count": len(items)}, "assessment_snapshot", tenant_id=tid))
                return self._json(200, {"ok": False, "items": [], "error": "App Registraties niet beschikbaar"})
            if re.fullmatch(r"/api/apps/[^/]+/registrations/[^/]+", path):
                parts = path.split("/")
                tid, app_id = parts[3], parts[5]
                try:
                    data = _run_appregs_ps(tid, "get-appreg", {"app_id": app_id})
                    if data.get("ok") is not False:
                        return self._json(200, data)
                except Exception:
                    pass
                # Assessment fallback: find by appId
                snap = _latest_assessment_snapshot_for_tenant(tid)
                payload = _assessment_json_payload(snap, "apps", "registrations")
                if isinstance(payload, dict):
                    for item in payload.get("items") or []:
                        if not isinstance(item, dict):
                            continue
                        if (_payload_value(item, "AppId", "appId", default="") or "").lower() == app_id.lower():
                            perms = _payload_value(item, "Permissions", "permissions", default=None)
                            return self._json(200, _attach_source_meta({
                                "ok": True,
                                "displayName": _payload_value(item, "DisplayName", "displayName", default=""),
                                "appId": _payload_value(item, "AppId", "appId", default=""),
                                "signInAudience": None,
                                "createdAt": _payload_value(item, "CreatedDateTime", "createdAt"),
                                "hasEnterpriseApp": bool(_payload_value(item, "HasEnterpriseApp", "hasEnterpriseApp", default=False)),
                                "secrets": ([{"hint": "•••", "statusLabel": _payload_value(item, "SecretExpirationStatus", "secretExpirationStatus")}]
                                            if int(_payload_value(item, "SecretCount", "secretCount", default=0) or 0) > 0 else []),
                                "certs": ([{"type": "Certificate", "statusLabel": _payload_value(item, "CertificateExpirationStatus", "certificateExpirationStatus")}]
                                          if int(_payload_value(item, "CertificateCount", "certificateCount", default=0) or 0) > 0 else []),
                                "redirectUris": [],
                                "identifierUris": [],
                                "requiredResourceAccess": [],
                                "permissions": list(perms) if isinstance(perms, list) else [],
                            }, "assessment_snapshot", tenant_id=tid))
                return self._json(404, {"ok": False, "error": "App Registratie niet gevonden"})

            # ── Collaboration routes (GET) ──
            if re.fullmatch(r"/api/collaboration/[^/]+/sharepoint/sites", path):
                tid = path.split("/")[3]
                try:
                    data = _run_collab_ps(tid, "list-sharepoint", {})
                    if data.get("ok") and data.get("sites") is not None:
                        enriched = dict(data)
                        enriched.update(_build_sharepoint_capacity_summary(tid, enriched.get("sites") or []))
                        return self._json(200, enriched)
                except Exception:
                    pass
                sites = _snapshot_as_sharepoint_sites(tid)
                payload = {"ok": True, "sites": sites, "count": len(sites)}
                payload.update(_build_sharepoint_capacity_summary(tid, sites))
                return self._json(200, _attach_source_meta(payload, "assessment_snapshot", tenant_id=tid))
            if re.fullmatch(r"/api/collaboration/[^/]+/sharepoint/settings", path):
                tid = path.split("/")[3]
                try:
                    data = _run_collab_ps(tid, "get-sharepoint-settings", {})
                    if data.get("ok") is not False:
                        return self._json(200, data)
                except Exception:
                    pass
                snap_settings = _snapshot_as_sharepoint_settings(tid)
                if snap_settings:
                    return self._json(200, _attach_source_meta(snap_settings, "assessment_snapshot", tenant_id=tid))
                return self._json(200, {"ok": False, "message": "SharePoint-instellingen niet beschikbaar"})
            if re.fullmatch(r"/api/collaboration/[^/]+/teams", path):
                tid = path.split("/")[3]
                try:
                    data = _run_collab_ps(tid, "list-teams", {})
                    if data.get("ok") and "teams" in data:
                        return self._json(200, data)
                except Exception:
                    pass
                teams = _snapshot_as_teams(tid)
                return self._json(200, _attach_source_meta({"ok": True, "teams": teams, "count": len(teams)}, "assessment_snapshot", tenant_id=tid))
            if re.fullmatch(r"/api/collaboration/[^/]+/teams/[^/]+", path):
                parts = path.split("/")
                tid, team_id = parts[3], parts[5]
                return self._json(200, _run_collab_ps(tid, "get-team", {"team_id": team_id}))
            if re.fullmatch(r"/api/collaboration/[^/]+/groups", path):
                tid = path.split("/")[3]
                try:
                    data = _run_collab_ps(tid, "list-groups", {})
                    if data.get("ok") and "groups" in data:
                        return self._json(200, data)
                except Exception:
                    pass
                return self._json(200, {"ok": True, "groups": [], "count": 0, "stats": {}})

            # ── Alerts routes (GET) ──
            if re.fullmatch(r"/api/alerts/[^/]+/audit-logs", path):
                tid = path.split("/")[3]
                limit = int(qs.get("limit", ["100"])[0])
                try:
                    data = _run_alerts_ps(tid, "list-audit-logs", {"limit": limit})
                    if data.get("ok") and "items" in data:
                        return self._json(200, data)
                except Exception:
                    pass
                snap = _latest_assessment_snapshot_for_tenant(tid)
                payload = _assessment_json_payload(snap, "alerts", "audit-logs")
                if isinstance(payload, dict):
                    items = []
                    for item in payload.get("items") or []:
                        if not isinstance(item, dict):
                            continue
                        items.append({
                            "title": _payload_value(item, "Title", "title", default=""),
                            "severity": _payload_value(item, "Severity", "severity", default=""),
                            "category": _payload_value(item, "Category", "category"),
                            "status": _payload_value(item, "Status", "status"),
                            "createdDateTime": _payload_value(item, "CreatedDateTime", "createdDateTime"),
                        })
                    return self._json(200, _attach_source_meta({"ok": True, "items": items, "count": len(items)}, "assessment_snapshot", tenant_id=tid))
                return self._json(200, _attach_source_meta({"ok": True, "items": [], "message": "Auditlog niet beschikbaar — voer een live sessie uit voor realtime data"}, "assessment_snapshot", tenant_id=tid))
            if re.fullmatch(r"/api/alerts/[^/]+/secure-score", path):
                tid = path.split("/")[3]
                try:
                    data = _run_alerts_ps(tid, "get-secure-score", {})
                    if data.get("ok") is not False and ("score" in data or "currentScore" in data):
                        return self._json(200, data)
                except Exception:
                    pass
                snap = _latest_assessment_snapshot_for_tenant(tid)
                payload = _assessment_json_payload(snap, "alerts", "secure-score")
                if isinstance(payload, dict):
                    summary = payload.get("summary") or {}
                    return self._json(200, _attach_source_meta({
                        "ok": True,
                        "score": round(float(summary.get("percentage") or 0)),
                        "currentScore": float(summary.get("currentScore") or 0),
                        "maxScore": float(summary.get("maxScore") or 100),
                        "recommendations": payload.get("items") or [],
                        "createdAt": payload.get("generated_at") or snap.get("assessment_generated_at"),
                    }, "assessment_snapshot", tenant_id=tid))
                # Fallback: snapshot SecureScorePct
                metrics = _snapshot_raw_metrics(tid)
                score = metrics.get("SecureScorePct")
                if score is not None:
                    return self._json(200, _attach_source_meta({
                        "ok": True, "score": round(float(score)),
                        "currentScore": round(float(score)), "maxScore": 100,
                        "createdAt": snap.get("assessment_generated_at"),
                    }, "assessment_snapshot", tenant_id=tid))
                return self._json(200, {"ok": False, "message": "Security score niet beschikbaar"})
            if re.fullmatch(r"/api/alerts/[^/]+/sign-ins", path):
                tid = path.split("/")[3]
                limit = int(qs.get("limit", ["50"])[0])
                try:
                    data = _run_alerts_ps(tid, "list-sign-ins", {"limit": limit})
                    if data.get("ok") and "items" in data:
                        return self._json(200, data)
                except Exception:
                    pass
                return self._json(200, _attach_source_meta({"ok": True, "items": [], "message": "Aanmeldingen niet beschikbaar — voer een live sessie uit voor realtime data"}, "assessment_snapshot", tenant_id=tid))
            if re.fullmatch(r"/api/alerts/[^/]+/config", path):
                tid = path.split("/")[3]
                return self._json(200, {"ok": True, "config": get_alert_config(tid)})

            # ── Exchange routes (GET) ──
            if re.fullmatch(r"/api/exchange/[^/]+/mailboxes", path):
                tid = path.split("/")[3]
                try:
                    data = _run_exchange_ps(tid, "list-mailboxes", {})
                    if data.get("ok") is not False and "mailboxes" in data:
                        return self._json(200, data)
                except Exception:
                    pass
                # Fallback: snapshot UserMailboxes als basis mailbox-lijst
                mailboxes = _snapshot_as_mailboxes(tid)
                if mailboxes:
                    return self._json(200, _attach_source_meta({"ok": True, "mailboxes": mailboxes}, "assessment_snapshot", tenant_id=tid))
                return self._json(200, {"ok": False, "mailboxes": [], "error": "Exchange data niet beschikbaar"})
            if re.fullmatch(r"/api/exchange/[^/]+/mailboxes/[^/]+", path):
                parts = path.split("/")
                tid, uid = parts[3], parts[5]
                try:
                    data = _run_exchange_ps(tid, "get-mailbox", {"user_id": uid})
                    if data.get("ok"):
                        return self._json(200, data)
                except Exception:
                    pass
                # Fallback: basic detail from snapshot
                detail = _snapshot_as_mailbox_detail(tid, uid)
                if detail:
                    return self._json(200, detail)
                return self._json(200, {"ok": False, "error": "Mailbox detail niet beschikbaar"})
            if re.fullmatch(r"/api/exchange/[^/]+/forwarding", path):
                tid = path.split("/")[3]
                return self._json(200, _run_exchange_ps(tid, "list-forwarding", {}))
            if re.fullmatch(r"/api/exchange/[^/]+/mailbox-rules", path):
                tid = path.split("/")[3]
                return self._json(200, _run_exchange_ps(tid, "list-mailbox-rules", {}))

            # ── Identiteit & Toegang routes (GET) ──────────────────────────────
            if re.fullmatch(r"/api/identity/[^/]+/mfa", path):
                tid = path.split("/")[3]
                try:
                    data = _run_identity_ps(tid, "list-mfa", {})
                    if data.get("ok") is not False:
                        return self._json(200, data)
                except Exception:
                    pass
                return self._json(200, {"ok": False, "items": [], "error": "MFA-data niet beschikbaar"})
            if re.fullmatch(r"/api/identity/[^/]+/guests", path):
                tid = path.split("/")[3]
                try:
                    data = _run_identity_ps(tid, "list-guests", {})
                    if data.get("ok") is not False:
                        return self._json(200, data)
                except Exception:
                    pass
                return self._json(200, {"ok": False, "items": [], "error": "Gast-data niet beschikbaar"})
            if re.fullmatch(r"/api/identity/[^/]+/admin-roles", path):
                tid = path.split("/")[3]
                try:
                    data = _run_identity_ps(tid, "list-admin-roles", {})
                    if data.get("ok") is not False:
                        return self._json(200, data)
                except Exception:
                    pass
                return self._json(200, {"ok": False, "items": [], "error": "Rollen-data niet beschikbaar"})
            if re.fullmatch(r"/api/identity/[^/]+/security-defaults", path):
                tid = path.split("/")[3]
                try:
                    data = _run_identity_ps(tid, "get-security-defaults", {})
                    if data.get("ok") is not False:
                        return self._json(200, data)
                except Exception:
                    pass
                return self._json(200, {"ok": False, "error": "Security Defaults niet beschikbaar"})
            if re.fullmatch(r"/api/identity/[^/]+/legacy-auth", path):
                tid = path.split("/")[3]
                try:
                    data = _run_identity_ps(tid, "list-legacy-auth", {})
                    if data.get("ok") is not False:
                        return self._json(200, data)
                except Exception:
                    pass
                return self._json(200, {"ok": False, "items": [], "error": "Legacy-auth data niet beschikbaar"})

            # ── App Registraties routes (GET) ───────────────────────────────────
            if re.fullmatch(r"/api/apps/[^/]+/registrations", path):
                tid = path.split("/")[3]
                try:
                    data = _run_appregs_ps(tid, "list-appregs", {})
                    if data.get("ok") is not False and "items" in data:
                        return self._json(200, data)
                except Exception:
                    pass
                return self._json(200, {"ok": False, "items": [], "error": "App Registraties niet beschikbaar"})
            if re.fullmatch(r"/api/apps/[^/]+/registrations/[^/]+", path):
                parts = path.split("/")
                tid, app_id = parts[3], parts[5]
                try:
                    data = _run_appregs_ps(tid, "get-appreg", {"app_id": app_id})
                    if data.get("ok") is not False:
                        return self._json(200, data)
                except Exception:
                    pass
                return self._json(404, {"ok": False, "error": "App Registratie niet gevonden"})

            # ── Samenwerking: SharePoint & Teams routes (GET) ───────────────────
            if re.fullmatch(r"/api/collaboration/[^/]+/sharepoint/sites", path):
                tid = path.split("/")[3]
                try:
                    data = _run_collab_ps(tid, "list-sharepoint", {})
                    if data.get("ok") is not False and "sites" in data:
                        enriched = dict(data)
                        enriched.update(_build_sharepoint_capacity_summary(tid, enriched.get("sites") or []))
                        return self._json(200, enriched)
                except Exception:
                    pass
                sites = _snapshot_as_sharepoint_sites(tid)
                if sites:
                    payload = {"ok": True, "sites": sites}
                    payload.update(_build_sharepoint_capacity_summary(tid, sites))
                    return self._json(200, _attach_source_meta(payload, "assessment_snapshot", tenant_id=tid))
                return self._json(200, {"ok": False, "sites": [], "error": "SharePoint-data niet beschikbaar"})
            if re.fullmatch(r"/api/collaboration/[^/]+/sharepoint/settings", path):
                tid = path.split("/")[3]
                try:
                    data = _run_collab_ps(tid, "get-sharepoint-settings", {})
                    if data.get("ok") is not False:
                        return self._json(200, data)
                except Exception:
                    pass
                settings = _snapshot_as_sharepoint_settings(tid)
                if settings:
                    return self._json(200, _attach_source_meta(settings, "assessment_snapshot", tenant_id=tid))
                return self._json(200, {"ok": False, "error": "SharePoint-instellingen niet beschikbaar"})
            if re.fullmatch(r"/api/collaboration/[^/]+/teams", path):
                tid = path.split("/")[3]
                try:
                    data = _run_collab_ps(tid, "list-teams", {})
                    if data.get("ok") is not False and "teams" in data:
                        return self._json(200, data)
                except Exception:
                    pass
                teams = _snapshot_as_teams(tid)
                if teams:
                    return self._json(200, _attach_source_meta({"ok": True, "teams": teams}, "assessment_snapshot", tenant_id=tid))
                return self._json(200, {"ok": False, "teams": [], "error": "Teams-data niet beschikbaar"})
            if re.fullmatch(r"/api/collaboration/[^/]+/teams/[^/]+", path):
                parts = path.split("/")
                tid, team_id = parts[3], parts[5]
                try:
                    data = _run_collab_ps(tid, "get-team", {"team_id": team_id})
                    if data.get("ok") is not False:
                        return self._json(200, data)
                except Exception:
                    pass
                return self._json(404, {"ok": False, "error": "Team detail niet beschikbaar"})
            if re.fullmatch(r"/api/collaboration/[^/]+/groups", path):
                tid = path.split("/")[3]
                try:
                    data = _run_collab_ps(tid, "list-groups", {})
                    if data.get("ok") and "groups" in data:
                        return self._json(200, data)
                except Exception:
                    pass
                return self._json(200, {"ok": True, "groups": [], "count": 0, "stats": {}})

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
            if re.fullmatch(r"/api/kb/[^/]+/assets/\d+/findings", path):
                tid = _kb_tid(path)
                asset_id = int(path.split("/")[5])
                return self._json(200, {"items": list_actions_for_asset(tid, asset_id)})
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
            if re.fullmatch(r"/api/kb/[^/]+/appregs", path):
                tid = _kb_tid(path)
                snap = _latest_assessment_snapshot_for_tenant(tid) or {}
                payload = _assessment_json_payload(snap, "apps", "registrations")
                items = []
                if isinstance(payload, dict):
                    for item in payload.get("items") or []:
                        if not isinstance(item, dict):
                            continue
                        perms = _payload_value(item, "Permissions", "permissions", default=None)
                        items.append({
                            "displayName": _payload_value(item, "DisplayName", "displayName", default=""),
                            "appId": _payload_value(item, "AppId", "appId", default=""),
                            "createdAt": _payload_value(item, "CreatedDateTime", "createdAt"),
                            "secretCount": int(_payload_value(item, "SecretCount", "secretCount", default=0) or 0),
                            "secretExpiration": _payload_value(item, "SecretExpiration", "secretExpiration"),
                            "secretExpirationStatus": _payload_value(item, "SecretExpirationStatus", "secretExpirationStatus"),
                            "certificateCount": int(_payload_value(item, "CertificateCount", "certificateCount", default=0) or 0),
                            "certificateExpiration": _payload_value(item, "CertificateExpiration", "certificateExpiration"),
                            "certificateExpirationStatus": _payload_value(item, "CertificateExpirationStatus", "certificateExpirationStatus"),
                            "permissionCount": int(_payload_value(item, "PermissionCount", "permissionCount", default=0) or 0),
                            "hasEnterpriseApp": bool(_payload_value(item, "HasEnterpriseApp", "hasEnterpriseApp", default=False)),
                            "permissions": list(perms) if isinstance(perms, list) else [],
                        })
                return self._json(200, {"ok": True, "items": items, "generated_at": snap.get("assessment_generated_at")})
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
                db_execute("UPDATE users SET last_login_at=? WHERE id=?", (now_iso(), user["id"]))
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
            if re.fullmatch(r"/api/tenants/[^/]+/auth-config", path):
                if _sess.get("role") != "admin":
                    return self._json(403, {"error": "Onvoldoende rechten."})
                tenant_id = path.split("/")[3]
                return self._json(200, save_tenant_auth_profile(tenant_id, self._read_json()))
            if re.fullmatch(r"/api/tenants/[^/]+/delete", path):
                tenant_id = path.split("/")[3]
                payload = self._read_json()
                mode = payload.get("mode") or parse_qs(parsed.query).get("mode", ["soft"])[0]
                return self._json(200, delete_tenant(tenant_id, mode))
            # ── Customer API (POST) ────────────────────────────────────────────
            if path == "/api/customers":
                if _sess.get("role") != "admin":
                    return self._json(403, {"error": "Onvoldoende rechten.", "error_code": "forbidden"})
                return self._json(201, create_customer(self._read_json()))
            if re.fullmatch(r"/api/customers/[^/]+/services", path):
                if _sess.get("role") != "admin":
                    return self._json(403, {"error": "Onvoldoende rechten.", "error_code": "forbidden"})
                cid = path.split("/")[3]
                if not db_fetchone("SELECT id FROM customers WHERE id=?", (cid,)):
                    http_s, pl = _api_error("not_found", "Klant niet gevonden", 404)
                    return self._json(http_s, pl)
                body = self._read_json()
                service_key = (body.get("service_key") or "").strip()
                if not service_key:
                    http_s, pl = _api_error("validation_error", "service_key is verplicht", 400)
                    return self._json(http_s, pl)
                sid = str(uuid.uuid4())
                db_execute(
                    "INSERT OR REPLACE INTO customer_services (id, customer_id, service_key, is_enabled, onboarded_at, notes) "
                    "VALUES (?, ?, ?, ?, ?, ?)",
                    (sid, cid, service_key,
                     1 if body.get("is_enabled", True) else 0,
                     body.get("onboarded_at") or now_iso(),
                     (body.get("notes") or "").strip() or None),
                )
                return self._json(201, db_fetchone("SELECT * FROM customer_services WHERE id=?", (sid,)) or {})
            # ── Approvals (POST) ──────────────────────────────────────────────
            if path == "/api/approvals":
                if _sess.get("role") != "admin":
                    return self._json(403, {"error": "Onvoldoende rechten.", "error_code": "forbidden"})
                body = self._read_json()
                action_log_id = (body.get("action_log_id") or "").strip()
                requested_by = _sess.get("email", "")
                reason = body.get("reason") or None
                if not action_log_id:
                    http_s, pl = _api_error("validation_error", "action_log_id is verplicht", 400)
                    return self._json(http_s, pl)
                return self._json(201, create_approval(action_log_id, requested_by, reason))
            if re.fullmatch(r"/api/approvals/[^/]+/approve", path):
                if _sess.get("role") != "admin":
                    return self._json(403, {"error": "Onvoldoende rechten.", "error_code": "forbidden"})
                appr_id = path.split("/")[3]
                body = self._read_json()
                if not db_fetchone("SELECT id FROM approvals WHERE id=?", (appr_id,)):
                    return self._json(404, {"error": "Goedkeuring niet gevonden", "error_code": "not_found"})
                return self._json(200, decide_approval(appr_id, "approved", _sess.get("email", ""), body.get("reason") or None))
            if re.fullmatch(r"/api/approvals/[^/]+/reject", path):
                if _sess.get("role") != "admin":
                    return self._json(403, {"error": "Onvoldoende rechten.", "error_code": "forbidden"})
                appr_id = path.split("/")[3]
                body = self._read_json()
                if not db_fetchone("SELECT id FROM approvals WHERE id=?", (appr_id,)):
                    return self._json(404, {"error": "Goedkeuring niet gevonden", "error_code": "not_found"})
                return self._json(200, decide_approval(appr_id, "rejected", _sess.get("email", ""), body.get("reason") or None))
            # ── Integrations (POST) ───────────────────────────────────────────
            if re.fullmatch(r"/api/integrations/[^/]+/[^/]+", path):
                if _sess.get("role") != "admin":
                    return self._json(403, {"error": "Onvoldoende rechten.", "error_code": "forbidden"})
                parts = path.split("/")
                tid = parts[3]
                itype = parts[4]
                if not db_fetchone("SELECT id FROM tenants WHERE id=?", (tid,)):
                    return self._json(404, {"error": "Tenant niet gevonden", "error_code": "not_found"})
                body = self._read_json()
                return self._json(201, upsert_integration(tid, itype, body))
            # ── Tenant subscriptions (POST) ───────────────────────────────────
            if re.fullmatch(r"/api/tenants/[^/]+/subscriptions", path):
                if _sess.get("role") != "admin":
                    return self._json(403, {"error": "Onvoldoende rechten.", "error_code": "forbidden"})
                tenant_id = path.split("/")[3]
                if not db_fetchone("SELECT id FROM tenants WHERE id=?", (tenant_id,)):
                    return self._json(404, {"error": "Tenant niet gevonden", "error_code": "not_found"})
                return self._json(201, upsert_subscription(tenant_id, self._read_json()))
            # ── Customer access (POST) ────────────────────────────────────────
            if re.fullmatch(r"/api/customers/[^/]+/access/[^/]+", path):
                if _sess.get("role") != "admin":
                    return self._json(403, {"error": "Onvoldoende rechten.", "error_code": "forbidden"})
                parts = path.split("/")
                cid = parts[3]
                uid = parts[5]
                body = self._read_json()
                role_key = (body.get("role_key") or "read_only").strip()
                if not db_fetchone("SELECT id FROM customers WHERE id=?", (cid,)):
                    http_s, pl = _api_error("not_found", "Klant niet gevonden", 404)
                    return self._json(http_s, pl)
                if not db_fetchone("SELECT id FROM users WHERE id=?", (uid,)):
                    http_s, pl = _api_error("not_found", "Gebruiker niet gevonden", 404)
                    return self._json(http_s, pl)
                return self._json(201, grant_customer_access(cid, uid, role_key, _sess.get("email", "")))
            # ── Azure/Alert/Cost Snapshots (POST) ────────────────────────────
            if re.fullmatch(r"/api/tenants/[^/]+/azure-snapshots/[^/]+/[^/]+", path):
                parts = path.split("/")
                tid = parts[3]; sec = parts[5]; sub = parts[6]
                return self._json(201, upsert_azure_snapshot(tid, sec, sub, self._read_json()))
            if re.fullmatch(r"/api/tenants/[^/]+/alert-snapshots/[^/]+", path):
                parts = path.split("/")
                tid = parts[3]; atype = parts[5]
                return self._json(201, upsert_alert_snapshot(tid, atype, self._read_json()))
            if re.fullmatch(r"/api/tenants/[^/]+/cost-snapshots", path):
                tid = path.split("/")[3]
                return self._json(201, upsert_cost_snapshot(tid, self._read_json()))
            # ── Job Queue (POST) ──────────────────────────────────────────────
            if path == "/api/jobs":
                if _sess.get("role") != "admin":
                    return self._json(403, {"error": "Onvoldoende rechten.", "error_code": "forbidden"})
                body = self._read_json()
                job_type = (body.get("job_type") or "").strip()
                if not job_type:
                    http_s, pl = _api_error("validation_error", "job_type is verplicht", 400)
                    return self._json(http_s, pl)
                return self._json(201, enqueue_job(
                    job_type,
                    tenant_id=body.get("tenant_id") or None,
                    payload=body.get("payload") or {},
                    priority=int(body.get("priority") or 5),
                    scheduled_at=body.get("scheduled_at") or None,
                ))
            if re.fullmatch(r"/api/jobs/[^/]+/cancel", path):
                if _sess.get("role") != "admin":
                    return self._json(403, {"error": "Onvoldoende rechten.", "error_code": "forbidden"})
                job_id = path.split("/")[3]
                return self._json(200, cancel_job(job_id))
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
                if not _check_rate_limit(self.client_address[0], max_attempts=5, window_secs=300):
                    return self._json(429, {"error": "Te veel assessment-aanvragen. Wacht enkele minuten.", "error_code": "rate_limited"})
                return self._json(201, create_run(self._read_json()))
            if path == "/api/scheduled-runs":
                body = self._read_json()
                tid = (body.get("tenant_id") or "").strip()
                sched = (body.get("scheduled_at") or "").strip()
                phases = body.get("phases") or []
                if not tid or not sched:
                    return self._json(400, {"error": "tenant_id en scheduled_at zijn verplicht", "error_code": "validation_error"})
                job = enqueue_job(
                    "assessment", tenant_id=tid,
                    payload={"phases": phases, "note": body.get("note", "")},
                    scheduled_at=sched,
                )
                return self._json(201, {"ok": True, "job": job})
            # ── Zero Trust Assessment run ──
            if re.fullmatch(r"/api/compliance/[^/]+/zerotrust/run", path):
                tid = path.split("/")[3]
                folder = _zt_output_folder(tid)
                def _zt_bg():
                    try:
                        _run_zerotrust_ps(tid, "run", folder)
                    except Exception:
                        pass
                threading.Thread(target=_zt_bg, daemon=True).start()
                return self._json(202, {"ok": True, "message": "Zero Trust Assessment gestart als achtergrondtaak. Dit kan uren duren.", "tenant_id": tid})
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
                          "assessment_ui_v1"):
                    if k in payload:
                        cfg[k] = payload[k]
                save_config(cfg)
                # Geef config terug zonder geheimen
                safe = {k: v for k, v in cfg.items() if k not in ("auth_client_secret",)}
                return self._json(200, safe)
            # ── Remediation routes (POST) ──
            if re.fullmatch(r"/api/remediate/[^/]+/execute", path):
                if not _check_rate_limit(self.client_address[0], max_attempts=15, window_secs=60):
                    return self._json(429, {"error": "Te veel herstelacties tegelijk. Wacht even.", "error_code": "rate_limited"})
                if _sess.get("role") != "admin":
                    return self._json(403, {"error": "Onvoldoende rechten."})
                tenant_id = path.split("/")[3]
                payload   = self._read_json()
                rem_id    = (payload.get("remediation_id") or "").strip()
                params    = payload.get("params") or {}
                dry_run   = bool(payload.get("dry_run", False))
                if not rem_id:
                    return self._json(400, {"error": "remediation_id is verplicht"})
                if not isinstance(params, dict):
                    params = {}
                result = execute_remediation(
                    tenant_id, rem_id, params, dry_run,
                    executed_by=_sess.get("email", "admin"),
                )
                return self._json(200, result)

            # ── User Management routes (POST) ──
            if re.fullmatch(r"/api/m365/[^/]+/users", path):
                tenant_id = path.split("/")[3]
                payload   = self._read_json()
                dry_run   = bool(payload.pop("dry_run", False))
                result = _run_user_mgmt(
                    tenant_id, "create-user", payload, dry_run,
                    executed_by=_sess.get("email", "admin"),
                )
                if not result["ok"]:
                    return self._json(502, {"error": result.get("error", "Fout bij aanmaken gebruiker")})
                return self._json(200, result["result"])
            if re.fullmatch(r"/api/m365/[^/]+/users/[^/]+/offboard", path):
                parts     = path.split("/")
                tenant_id = parts[3]; user_id = parts[5]
                payload   = self._read_json()
                dry_run   = bool(payload.pop("dry_run", False))
                payload["user_id"] = user_id
                result = _run_user_mgmt(
                    tenant_id, "offboard-user", payload, dry_run,
                    executed_by=_sess.get("email", "admin"),
                )
                if not result["ok"]:
                    return self._json(502, {"error": result.get("error", "Fout bij offboarding")})
                return self._json(200, result["result"])

            # ── Baseline routes (POST) ──
            if path == "/api/baselines":
                payload = self._read_json()
                config  = payload.get("config") or {}
                row = create_baseline(
                    name=payload.get("name", ""),
                    description=payload.get("description", ""),
                    config=config,
                    source_tenant_id=payload.get("source_tenant_id"),
                    source_tenant_name=payload.get("source_tenant_name"),
                    created_by=_sess.get("email", "admin"),
                )
                return self._json(201, row)
            # Export (Gold Tenant → baseline)
            if re.fullmatch(r"/api/baselines/export/[^/]+", path):
                tenant_id = path.split("/")[4]
                payload   = self._read_json()
                result = _run_baseline_ps(tenant_id, "export-baseline", {})
                if not result["ok"]:
                    return self._json(502, {"error": result.get("error", "Export mislukt")})
                exported_config = result["result"].get("baseline", {})
                tenant_row = db_fetchone("SELECT customer_name FROM tenants WHERE id=?", (tenant_id,))
                tenant_name = tenant_row["customer_name"] if tenant_row else tenant_id
                row = create_baseline(
                    name=payload.get("name") or f"Baseline {tenant_name}",
                    description=payload.get("description") or f"Geëxporteerd van {tenant_name}",
                    config=exported_config,
                    source_tenant_id=tenant_id,
                    source_tenant_name=tenant_name,
                    created_by=_sess.get("email", "admin"),
                )
                return self._json(201, row)
            # Assign baseline aan tenant
            if re.fullmatch(r"/api/baselines/[^/]+/assign", path):
                bid     = path.split("/")[3]
                payload = self._read_json()
                tid     = payload.get("tenant_id", "")
                if not tid:
                    return self._json(400, {"error": "tenant_id is verplicht"})
                row = assign_baseline(bid, tid, _sess.get("email", "admin"))
                return self._json(201, row)
            # Compliance check
            if re.fullmatch(r"/api/baselines/[^/]+/check/[^/]+", path):
                parts = path.split("/")
                bid = parts[3]; tid = parts[5]
                result = check_baseline_compliance(bid, tid, _sess.get("email", "admin"))
                return self._json(200, result)
            # Apply baseline
            if re.fullmatch(r"/api/baselines/[^/]+/apply/[^/]+", path):
                parts   = path.split("/")
                bid = parts[3]; tid = parts[5]
                payload = self._read_json()
                dry_run = bool(payload.get("dry_run", False))
                result = apply_baseline_to_tenant(bid, tid, dry_run, _sess.get("email", "admin"))
                return self._json(200, result)

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
            # ── Intune routes (POST) ──
            if re.fullmatch(r"/api/intune/[^/]+/deploy-config", path):
                tid     = path.split("/")[3]
                payload = self._read_json()
                dry_run = payload.pop("dry_run", False)
                result  = _run_intune_ps(tid, "deploy-config", payload, dry_run,
                                         executed_by=_sess.get("email", "admin"))
                if not result.get("ok"):
                    return self._json(502, {"error": result.get("error", "Fout bij toewijzen profiel")})
                return self._json(200, result)

            # ── CA routes (POST) ──
            if re.fullmatch(r"/api/ca/[^/]+/policies/[^/]+/toggle", path):
                parts = path.split("/")
                tid, pid = parts[3], parts[5]
                payload = self._read_json()
                action = "enable-policy" if payload.get("action") == "enable" else "disable-policy"
                result = _run_ca_ps(tid, action, {"policy_id": pid},
                                    executed_by=_sess.get("email", "admin"))
                if not result.get("ok"):
                    return self._json(502, {"error": result.get("error", "Fout bij toggle")})
                return self._json(200, result)

            # ── Alerts routes (POST) ──
            if re.fullmatch(r"/api/alerts/[^/]+/config", path):
                tid = path.split("/")[3]
                payload = self._read_json()
                upsert_alert_config(
                    tid,
                    payload.get("webhook_url", ""),
                    payload.get("webhook_type", "teams"),
                    payload.get("email_addr", ""),
                )
                return self._json(200, {"ok": True})
            if re.fullmatch(r"/api/alerts/[^/]+/test-webhook", path):
                payload = self._read_json()
                webhook_url = payload.get("webhook_url", "")
                webhook_type = payload.get("webhook_type", "teams")
                if not webhook_url:
                    return self._json(400, {"error": "webhook_url vereist"})
                result = send_test_webhook(webhook_url, webhook_type)
                return self._json(200 if result.get("ok") else 502, result)

            # ── Upload-report (PowerShell scripts → HTML rapport opslaan) ──
            if path == "/api/upload-report":
                data = self._read_json()
                filename = os.path.basename(data.get("filename") or "M365-Complete-Baseline-latest.html")
                content = data.get("content") or ""
                if not filename.lower().endswith(".html"):
                    return self._json(400, {"error": "Only .html files are allowed"})
                if not content:
                    return self._json(400, {"error": "No content provided"})
                DEFAULT_REPORTS_DIR.mkdir(parents=True, exist_ok=True)
                (DEFAULT_REPORTS_DIR / filename).write_text(content, encoding="utf-8")
                return self._json(200, {"path": f"/reports/{filename}", "filename": filename})
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
            # ── Customer API (DELETE) ──────────────────────────────────────────
            if re.fullmatch(r"/api/customers/[^/]+", path):
                return self._json(200, delete_customer(path.split("/")[3]))
            if re.fullmatch(r"/api/customers/[^/]+/access/[^/]+", path):
                if _sess.get("role") != "admin":
                    return self._json(403, {"error": "Onvoldoende rechten.", "error_code": "forbidden"})
                parts = path.split("/")
                cid = parts[3]
                uid = parts[5]
                result = revoke_customer_access(cid, uid)
                if not result.get("ok"):
                    return self._json(404, {"error": "Toegang niet gevonden", "error_code": "not_found"})
                return self._json(200, result)
            # ── Gebruikersbeheer ──
            if re.fullmatch(r"/api/users/[^/]+", path):
                uid = path.split("/")[3]
                return self._json(200, delete_user_account(uid, _sess.get("email", "")))
            if re.fullmatch(r"/api/runs/[^/]+", path):
                run_id = path.split("/")[3]
                return self._json(200, delete_run(run_id))
            # ── Baseline routes (DELETE) ──
            if re.fullmatch(r"/api/baselines/[^/]+", path):
                bid = path.split("/")[3]
                return self._json(200, delete_baseline(bid))
            if re.fullmatch(r"/api/baselines/[^/]+/assign/[^/]+", path):
                parts = path.split("/")
                bid = parts[3]; tid = parts[5]
                return self._json(200, unassign_baseline(bid, tid))
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
            # ── Baseline routes (PATCH) ──
            if re.fullmatch(r"/api/baselines/[^/]+", path):
                bid = path.split("/")[3]
                return self._json(200, update_baseline(bid, self._read_json()))
            # ── Customer API (PATCH) ───────────────────────────────────────────
            if re.fullmatch(r"/api/customers/[^/]+", path):
                return self._json(200, update_customer(path.split("/")[3], self._read_json()))
            # ── Integrations (PATCH) ──────────────────────────────────────────
            if re.fullmatch(r"/api/integrations/[^/]+", path):
                integ_id = path.split("/")[3]
                row = get_integration(integ_id)
                if not row:
                    return self._json(404, {"error": "Integratie niet gevonden", "error_code": "not_found"})
                body = self._read_json()
                return self._json(200, upsert_integration(
                    row["tenant_id"], row["integration_type"], body
                ))
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
    JOB_DISPATCHER.start()
    server = ThreadingHTTPServer((host, port), Handler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")
    finally:
        JOB_DISPATCHER.stop()
        server.server_close()


if __name__ == "__main__":
    run(
        host=os.environ.get("M365_LOCAL_WEBAPP_HOST", "127.0.0.1"),
        port=int(os.environ.get("M365_LOCAL_WEBAPP_PORT", "8787")),
    )
