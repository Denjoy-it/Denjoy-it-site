const API_BASE = '';

/**
 * Centrale API-route definities.
 * Gebruik API.tenants.list() ipv losse strings — makkelijker te refactoren.
 */
const API = {
  auth: {
    verify:    () => '/api/auth/verify',
    login:     () => '/api/auth/login',
    logout:    () => '/api/auth/logout',
    csrfToken: () => '/api/auth/csrf-token',
  },
  config:    () => '/api/config',
  tenants: {
    list:     () => '/api/tenants',
    get:      (id) => `/api/tenants/${id}`,
    overview: (id) => `/api/tenants/${id}/overview`,
    runs:     (id) => `/api/tenants/${id}/runs`,
    diff:     (id) => `/api/tenants/${id}/runs/diff`,
    actions:  (id) => `/api/tenants/${id}/actions`,
  },
  runs: {
    list:     (tid) => tid ? `/api/runs?tenant_id=${tid}` : '/api/runs',
    get:      (id) => `/api/runs/${id}`,
    logs:     (id) => `/api/runs/${id}/logs`,
    create:   () => '/api/runs',
    delete:   (id) => `/api/runs/${id}/delete`,
    archive:  (id) => `/api/reports/${id}/archive`,
    restore:  (id) => `/api/reports/${id}/restore`,
  },
  reports: {
    list:     () => '/api/reports/list',
    stats:    (tid) => tid ? `/api/reports/stats?tenant_id=${tid}` : '/api/reports/stats',
    retention: () => '/api/reports/retention/apply',
  },
  actions: {
    list:   (tid) => `/api/tenants/${tid}/actions`,
    create: () => '/api/actions',
    update: (id) => `/api/actions/${id}`,
  },
  remediate: {
    catalog:  (tid) => `/api/remediate/${tid}/catalog`,
    history:  (tid) => `/api/remediate/${tid}/history`,
    execute:  (tid) => `/api/remediate/${tid}/execute`,
  },
  capabilities: {
    tenant:      (tid) => `/api/capabilities/${tid}`,
    subsection:  (tid, section, subsection) => `/api/capabilities/${tid}/${section}/${subsection}`,
  },
  m365: {
    users:               (tid)      => `/api/m365/${tid}/users`,
    user:                (tid, uid) => `/api/m365/${tid}/users/${uid}`,
    offboard:            (tid, uid) => `/api/m365/${tid}/users/${uid}/offboard`,
    licenses:            (tid)      => `/api/m365/${tid}/licenses`,
    provisioningHistory: (tid)      => `/api/m365/${tid}/provisioning-history`,
  },
  baselines: {
    list:        ()           => `/api/baselines`,
    get:         (bid)        => `/api/baselines/${bid}`,
    create:      ()           => `/api/baselines`,
    update:      (bid)        => `/api/baselines/${bid}`,
    delete:      (bid)        => `/api/baselines/${bid}`,
    export:      (tid)        => `/api/baselines/export/${tid}`,
    assign:      (bid)        => `/api/baselines/${bid}/assign`,
    unassign:    (bid, tid)   => `/api/baselines/${bid}/assign/${tid}`,
    assignments: (bid)        => `/api/baselines/${bid}/assignments`,
    allAssign:   ()           => `/api/baselines/assignments/all`,
    check:       (bid, tid)   => `/api/baselines/${bid}/check/${tid}`,
    apply:       (bid, tid)   => `/api/baselines/${bid}/apply/${tid}`,
    history:     (bid)        => `/api/baselines/${bid}/history`,
  },
  backup: {
    summary:    (tid) => `/api/backup/${tid}/summary`,
    status:     (tid) => `/api/backup/${tid}/status`,
    sharepoint: (tid) => `/api/backup/${tid}/sharepoint`,
    onedrive:   (tid) => `/api/backup/${tid}/onedrive`,
    exchange:   (tid) => `/api/backup/${tid}/exchange`,
    history:    (tid) => `/api/backup/${tid}/history`,
  },
  ca: {
    policies:      (tid)      => `/api/ca/${tid}/policies`,
    policy:        (tid, pid) => `/api/ca/${tid}/policies/${pid}`,
    policyToggle:  (tid, pid) => `/api/ca/${tid}/policies/${pid}/toggle`,
    namedLocations:(tid)      => `/api/ca/${tid}/named-locations`,
    history:       (tid)      => `/api/ca/${tid}/history`,
  },
  domains: {
    list:    (tid)        => `/api/domains/${tid}/list`,
    analyse: (tid, domain)=> `/api/domains/${tid}/analyse?domain=${encodeURIComponent(domain)}`,
  },
  alerts: {
    auditLogs:   (tid) => `/api/alerts/${tid}/audit-logs`,
    secureScore: (tid) => `/api/alerts/${tid}/secure-score`,
    signIns:     (tid) => `/api/alerts/${tid}/sign-ins`,
    config:      (tid) => `/api/alerts/${tid}/config`,
    testWebhook: (tid) => `/api/alerts/${tid}/test-webhook`,
  },
  exchange: {
    mailboxes:    (tid)      => `/api/exchange/${tid}/mailboxes`,
    mailbox:      (tid, uid) => `/api/exchange/${tid}/mailboxes/${uid}`,
    forwarding:   (tid)      => `/api/exchange/${tid}/forwarding`,
    rules:        (tid)      => `/api/exchange/${tid}/mailbox-rules`,
  },
  identity: {
    mfa:              (tid) => `/api/identity/${tid}/mfa`,
    guests:           (tid) => `/api/identity/${tid}/guests`,
    adminRoles:       (tid) => `/api/identity/${tid}/admin-roles`,
    securityDefaults: (tid) => `/api/identity/${tid}/security-defaults`,
    legacyAuth:       (tid) => `/api/identity/${tid}/legacy-auth`,
  },
  apps: {
    registrations: (tid) => `/api/apps/${tid}/registrations`,
    registration:  (tid, appId) => `/api/apps/${tid}/registrations/${appId}`,
  },
  collaboration: {
    sharepointSites:    (tid) => `/api/collaboration/${tid}/sharepoint/sites`,
    sharepointSettings: (tid) => `/api/collaboration/${tid}/sharepoint/settings`,
    teams:              (tid) => `/api/collaboration/${tid}/teams`,
    team:               (tid, teamId) => `/api/collaboration/${tid}/teams/${teamId}`,
  },
  kb: {
    assets:     (tid) => `/api/kb/${tid}/assets`,
    asset:      (tid, id) => `/api/kb/${tid}/assets/${id}`,
    assetTypes: (tid) => `/api/kb/${tid}/asset-types`,
    vlans:      (tid) => `/api/kb/${tid}/vlans`,
    vlan:       (tid, id) => `/api/kb/${tid}/vlans/${id}`,
    pages:      (tid) => `/api/kb/${tid}/pages`,
    page:       (tid, id) => `/api/kb/${tid}/pages/${id}`,
    contacts:   (tid) => `/api/kb/${tid}/contacts`,
    contact:    (tid, id) => `/api/kb/${tid}/contacts/${id}`,
    passwords:  (tid) => `/api/kb/${tid}/passwords`,
    password:   (tid, id) => `/api/kb/${tid}/passwords/${id}`,
    software:   (tid) => `/api/kb/${tid}/software`,
    softwareItem: (tid, id) => `/api/kb/${tid}/software/${id}`,
    domains:    (tid) => `/api/kb/${tid}/domains`,
    domain:     (tid, id) => `/api/kb/${tid}/domains/${id}`,
    m365:       (tid) => `/api/kb/${tid}/m365`,
    changelog:  (tid) => `/api/kb/${tid}/changelog`,
    changelogItem: (tid, id) => `/api/kb/${tid}/changelog/${id}`,
    meta:       (tid) => `/api/kb/${tid}/meta`,
  },
};

let currentTenantId = null;
let localConfig = null;
let allTenants = [];
let _liveModuleContext = null;

function applyTheme(theme) {
  const effective = theme === 'dark' ? 'dark' : 'light';
  document.documentElement.setAttribute('data-theme', effective);
  try { localStorage.setItem('m365LocalTheme', effective); } catch (_) {}
  const lightBtn = document.getElementById('theme-light-btn');
  const darkBtn = document.getElementById('theme-dark-btn');
  const cycleBtn = document.getElementById('themeCycleButton');
  if (lightBtn) lightBtn.classList.toggle('active', effective === 'light');
  if (darkBtn) darkBtn.classList.toggle('active', effective === 'dark');
  if (cycleBtn) {
    cycleBtn.textContent = effective === 'dark' ? '☀️' : '🌙';
    cycleBtn.title = effective === 'dark' ? 'Zet licht thema aan' : 'Zet donker thema aan';
  }
  document.dispatchEvent(new CustomEvent('m365-theme-changed', { detail: { theme: effective } }));
}

/**
 * Toon een toast notificatie.
 * @param {string} message  - Tekst om te tonen
 * @param {'info'|'success'|'warning'|'error'} type - Type toast (bepaalt kleur)
 * @param {number} duration - Milliseconden zichtbaar (default 4000, 0 = permanent)
 */
function showToast(message, type = 'info', duration = 4000) {
  const container = document.getElementById('toast-container');
  if (!container) return;

  const icons = { info: 'ℹ', success: '✓', warning: '⚠', error: '✕' };
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.setAttribute('role', 'alert');
  toast.innerHTML =
    `<span class="toast-icon" aria-hidden="true">${icons[type] || icons.info}</span>` +
    `<span class="toast-body">${message}</span>` +
    `<button type="button" class="toast-close" aria-label="Sluiten">×</button>`;

  const close = () => {
    toast.classList.add('toast-hiding');
    toast.addEventListener('animationend', () => toast.remove(), { once: true });
  };
  toast.querySelector('.toast-close').addEventListener('click', close);
  container.appendChild(toast);

  if (duration > 0) setTimeout(close, duration);
}
window.showToast = showToast;

function getSessionToken() {
  return localStorage.getItem('denjoy_auth_token') || localStorage.getItem('denjoy_token') || '';
}

async function fetchCapabilityStatus(tenantId, section, subsection, { forceRefresh = false } = {}) {
  if (!tenantId || !section || !subsection) return null;
  const path = API.capabilities.subsection(tenantId, section, subsection);
  if (forceRefresh && window.cacheClear) window.cacheClear(path);
  if (!forceRefresh && window.cacheGet) {
    const hit = window.cacheGet(path);
    if (hit !== null) return hit.capability || hit;
  }
  const token = getSessionToken();
  const res = await fetch(path, {
    credentials: 'include',
    headers: {
      'Content-Type': 'application/json',
      ...(token ? { Authorization: `Bearer ${token}` } : {}),
    },
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    throw new Error(data.error || `HTTP ${res.status}`);
  }
  if (window.cacheSet) window.cacheSet(path, data, 2 * 60 * 1000);
  return data.capability || data;
}

function describeCapabilityStatus(capability) {
  if (!capability) {
    return {
      label: 'Capability onbekend',
      detail: 'Er is nog geen capability-profiel geladen voor dit subhoofdstuk.',
      className: 'is-neutral',
    };
  }
  const status = String(capability.status || '');
  if (status === 'ready') {
    return { label: capability.status_label || 'Live beschikbaar', detail: capability.status_reason || 'Live ophalen is beschikbaar.', className: 'is-live' };
  }
  if (status === 'validation_required') {
    return { label: capability.status_label || 'Validatie vereist', detail: capability.status_reason || 'Controleer toegangsmodel en rechten.', className: 'is-warn' };
  }
  if (status === 'config_required') {
    return { label: capability.status_label || 'Configuratie vereist', detail: capability.status_reason || 'App-configuratie ontbreekt.', className: 'is-stale' };
  }
  if (status === 'snapshot_only') {
    return { label: capability.status_label || 'Assessment-only', detail: capability.status_reason || 'Dit onderdeel gebruikt alleen snapshotdata.', className: 'is-neutral' };
  }
  return { label: capability.status_label || 'Niet beschikbaar', detail: capability.status_reason || 'Live ophalen is nu niet beschikbaar.', className: 'is-stale' };
}

window.denjoyFetchCapabilityStatus = fetchCapabilityStatus;
window.denjoyDescribeCapabilityStatus = describeCapabilityStatus;
window.denjoySetLiveModuleContext = function setLiveModuleContext(context) {
  _liveModuleContext = context || null;
  if (_currentSection && _liveModuleContext && _liveModuleContext.section === _currentSection) {
    renderContextRail(_currentSection);
  }
};
window.denjoyGetLiveModuleContext = function getLiveModuleContext() {
  return _liveModuleContext;
};

function initThemeControls() {
  const lightBtn = document.getElementById('theme-light-btn');
  const darkBtn = document.getElementById('theme-dark-btn');
  const cycleBtn = document.getElementById('themeCycleButton');
  if (lightBtn) lightBtn.addEventListener('click', () => applyTheme('light'));
  if (darkBtn) darkBtn.addEventListener('click', () => applyTheme('dark'));
  if (cycleBtn) {
    cycleBtn.addEventListener('click', () => {
      const current = document.documentElement.getAttribute('data-theme') || 'light';
      applyTheme(current === 'dark' ? 'light' : 'dark');
    });
  }
  let saved = null;
  try { saved = localStorage.getItem('m365LocalTheme'); } catch (_) {}
  if (!saved) {
    // Geen voorkeur opgeslagen — volg OS dark mode instelling
    saved = (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) ? 'dark' : 'light';
  }
  applyTheme(saved);
  // Luister naar OS dark mode wijzigingen (alleen als gebruiker nog geen handmatige keuze heeft)
  if (window.matchMedia) {
    window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', (e) => {
      let stored = null;
      try { stored = localStorage.getItem('m365LocalTheme'); } catch (_) {}
      if (!stored) applyTheme(e.matches ? 'dark' : 'light');
    });
  }
}

// ── Loading bar (slanke top-bar, gedeeld over alle API calls) ────────────────
let _loadingCount = 0;

function _showLoadingBar() {
  _loadingCount++;
  const bar = document.getElementById('topLoadingBar');
  if (bar) { bar.classList.add('loading'); bar.classList.remove('done'); }
}

function _hideLoadingBar() {
  _loadingCount = Math.max(0, _loadingCount - 1);
  if (_loadingCount === 0) {
    const bar = document.getElementById('topLoadingBar');
    if (bar) {
      bar.classList.add('done');
      setTimeout(() => bar.classList.remove('loading', 'done'), 400);
    }
  }
}

async function apiFetch(path, options = {}) {
  _showLoadingBar();
  try {
    const token = localStorage.getItem('denjoy_token');
    const authHeader = token ? { 'Authorization': `Bearer ${token}` } : {};
    const res = await fetch(`${API_BASE}${path}`, {
      credentials: 'include',
      headers: { 'Content-Type': 'application/json', ...authHeader, ...(options.headers || {}) },
      ...options,
    });
    let data = null;
    try { data = await res.json(); } catch (_) { data = null; }
    if (res.status === 401) {
      // Sessie verlopen of niet ingelogd — stuur terug naar login
      localStorage.removeItem('denjoy_token');
      window.location.href = '/site/login.html';
      return null;
    }
    if (!res.ok) {
      throw new Error((data && data.error) || `HTTP ${res.status}`);
    }
    return data;
  } finally {
    _hideLoadingBar();
  }
}

// ── Client-side TTL-cache (sessionStorage) ───────────────────────────────────
// Gebruik: apiFetchCached(url, {}, 5 * 60 * 1000) voor 5 minuten cache
// Cache is tenant-aware: de sleutel bevat de volledige URL inclusief tenant-id.
// Bij navigatie naar een andere tenant wordt de cache niet automatisch gewist —
// de URL-sleutel zorgt voor isolatie.

const _CACHE_PREFIX = 'djc:';

function cacheSet(key, data, ttlMs) {
  try {
    sessionStorage.setItem(_CACHE_PREFIX + key, JSON.stringify({
      data,
      expires: Date.now() + ttlMs,
    }));
  } catch (_) { /* sessionStorage vol of niet beschikbaar */ }
}

function cacheGet(key) {
  try {
    const raw = sessionStorage.getItem(_CACHE_PREFIX + key);
    if (!raw) return null;
    const entry = JSON.parse(raw);
    if (Date.now() > entry.expires) {
      sessionStorage.removeItem(_CACHE_PREFIX + key);
      return null;
    }
    return entry.data;
  } catch (_) { return null; }
}

function cacheClear(keyPrefix) {
  try {
    const toRemove = [];
    for (let i = 0; i < sessionStorage.length; i++) {
      const k = sessionStorage.key(i);
      if (k && k.startsWith(_CACHE_PREFIX + (keyPrefix || ''))) toRemove.push(k);
    }
    toRemove.forEach(k => sessionStorage.removeItem(k));
  } catch (_) {}
}

// TTL-constanten (milliseconden)
const CACHE_TTL = {
  policies:   5 * 60 * 1000,   // CA Policies, App Registrations — traag te laden
  domains:    5 * 60 * 1000,   // DNS-checks veranderen nauwelijks
  teams:      3 * 60 * 1000,   // Teams/SharePoint
  mailboxes:  2 * 60 * 1000,   // Mailboxen kunnen vaker wijzigen
  short:      1 * 60 * 1000,   // Audit logs, sign-ins — relatief vers houden
};

async function apiFetchCached(path, options = {}, ttlMs = CACHE_TTL.short) {
  const cached = cacheGet(path);
  if (cached !== null) return cached;
  const data = await apiFetch(path, options);
  if (data !== null) cacheSet(path, data, ttlMs);
  return data;
}

window.cacheSet = cacheSet;
window.cacheGet = cacheGet;
window.cacheClear = cacheClear;
window.CACHE_TTL = CACHE_TTL;
window.apiFetchCached = apiFetchCached;

// ── Skeleton screen helpers ───────────────────────────────────────────────────
// skeletonTable(rows, cols)  → HTML voor gebruik in <tbody>
// skeletonCards(n)           → HTML voor lijst/kaart weergave
// skeletonLines(n)           → HTML voor eenvoudige regel-loader

function skeletonTable(rows = 5, cols = 5) {
  const widths = ['sk-w-60','sk-w-40','sk-w-30','sk-w-50','sk-w-30','sk-w-40','sk-w-20'];
  const row = `<tr class="sk-table-row">${
    Array.from({ length: cols }, (_, i) =>
      `<td><span class="sk-shimmer sk-line ${widths[i % widths.length]}">&nbsp;</span></td>`
    ).join('')
  }</tr>`;
  return Array(rows).fill(row).join('');
}

function skeletonCards(n = 4) {
  return Array.from({ length: n }, (_, i) => `
    <div class="sk-card-block">
      <span class="sk-shimmer sk-line sk-line-lg ${i % 2 === 0 ? 'sk-w-50' : 'sk-w-40'}">&nbsp;</span>
      <span class="sk-shimmer sk-line sk-line-sm sk-w-80">&nbsp;</span>
      <span class="sk-shimmer sk-line sk-line-sm sk-w-60">&nbsp;</span>
    </div>`).join('');
}

function skeletonLines(n = 3) {
  const widths = ['sk-w-80','sk-w-60','sk-w-70','sk-w-50','sk-w-40'];
  return `<div style="padding:.75rem 0">` +
    Array.from({ length: n }, (_, i) =>
      `<div><span class="sk-shimmer sk-line ${widths[i % widths.length]}">&nbsp;</span></div>`
    ).join('') + `</div>`;
}

window.skeletonTable = skeletonTable;
window.skeletonCards = skeletonCards;
window.skeletonLines = skeletonLines;

// ── Sub-nav configuratie per sectie ──────────────────────────────────────────
const SUBNAV_CONFIG = {
  overview: [],
  assessment: [
    { label: 'Assessment starten', section: 'assessment' },
    { label: 'Rapporten',          section: 'results' },
  ],
  results: [
    { label: 'Rapport',      resultsPanel: 'viewer' },
    { label: 'Vergelijking', resultsPanel: 'diff' },
    { label: 'Beheer',       resultsPanel: 'management' },
    { label: 'Acties',       resultsPanel: 'actions' },
  ],
  herstel: [
    { label: 'Catalogus',    remTab: 'catalogus' },
    { label: 'Geschiedenis', remTab: 'geschiedenis' },
  ],
  gebruikers: [
    { label: 'Gebruikers',   gbTab: 'gebruikers' },
    { label: 'Licenties',    gbTab: 'licenties' },
    { label: 'Geschiedenis', gbTab: 'geschiedenis' },
  ],
  teams: [
    { label: 'Teams',   liveTab: 'teams' },
    { label: 'Groepen', liveTab: 'groepen' },
  ],
  sharepoint: [
    { label: 'Sites',        liveTab: 'sharepoint-sites' },
    { label: 'Instellingen', liveTab: 'sharepoint-settings' },
    { label: 'Backup',       liveTab: 'sharepoint-backup' },
  ],
  identity: [
    { label: 'MFA',               liveTab: 'mfa' },
    { label: 'Guests',            liveTab: 'guests' },
    { label: 'Admin Roles',       liveTab: 'admin-roles' },
    { label: 'Security Defaults', liveTab: 'security-defaults' },
    { label: 'Legacy Auth',       liveTab: 'legacy-auth' },
  ],
  apps: [
    { label: 'Registrations', liveTab: 'registrations' },
  ],
  baseline: [
    { label: 'Baselines',   baselineTab: 'baselines' },
    { label: 'Gold Tenant', baselineTab: 'gold' },
    { label: 'Koppelingen', baselineTab: 'assignments' },
    { label: 'Geschiedenis', baselineTab: 'history' },
  ],
  intune: [
    { label: 'Overzicht',    liveTab: 'overzicht' },
    { label: 'Apparaten',    liveTab: 'apparaten' },
    { label: 'Compliance',   liveTab: 'compliance' },
    { label: 'Configuratie', liveTab: 'configuratie' },
    { label: 'Geschiedenis', liveTab: 'geschiedenis' },
  ],
  backup: [
    { label: 'Overzicht',    liveTab: 'overzicht' },
    { label: 'OneDrive',     liveTab: 'onedrive' },
    { label: 'Exchange',     liveTab: 'exchange' },
    { label: 'Geschiedenis', liveTab: 'geschiedenis' },
  ],
  ca: [
    { label: 'Policies',        caTab: 'policies' },
    { label: 'Named Locations', caTab: 'locations' },
    { label: 'Geschiedenis',    caTab: 'geschiedenis' },
  ],
  domains: [
    { label: 'Domeinen', liveTab: 'domains-list' },
    { label: 'Analyse',  liveTab: 'domains-analyse' },
  ],
  alerts: [
    { label: 'Audit Log',    liveTab: 'auditlog' },
    { label: 'Secure Score', liveTab: 'securescr' },
    { label: 'Aanmeldingen', liveTab: 'signins' },
    { label: 'Notificaties', liveTab: 'config' },
  ],
  exchange: [
    { label: 'Mailboxen',    liveTab: 'mailboxen' },
    { label: 'Forwarding',   liveTab: 'forwarding' },
    { label: 'Inbox Regels', liveTab: 'regels' },
  ],
  compliance: [
    { label: 'CIS Benchmark',  liveTab: 'cis' },
    { label: 'Zero Trust',     liveTab: 'zerotrust' },
  ],
  hybrid: [
    { label: 'AD Connect', liveTab: 'sync' },
  ],
  kb: [
    { label: 'Overzicht',      kbTab: 'overview'   },
    { label: 'Apparaten',      kbTab: 'assets',    countId: 'nbCountAssets' },
    { label: 'VLANs',          kbTab: 'vlans',     countId: 'nbCountVlans' },
    { label: 'Documenten',     kbTab: 'pages',     countId: 'nbCountPages' },
    { label: 'Contacten',      kbTab: 'contacts',  countId: 'nbCountContacts' },
    { label: 'Passwords',      kbTab: 'passwords', countId: 'nbCountPasswords' },
    { label: 'Software',       kbTab: 'software',  countId: 'nbCountSoftware' },
    { label: 'Domeinen',       kbTab: 'domains',   countId: 'nbCountDomains' },
    { label: 'Microsoft 365',  kbTab: 'm365' },
    { label: 'Wijzigingslog',  kbTab: 'changelog', countId: 'nbCountChangelog' },
  ],
  settings: [
    { label: 'Tenants',      settingsTab: 'tenant' },
    { label: 'Configuratie', settingsTab: 'general' },
    { label: 'Integraties',  settingsTab: 'integrations' },
  ],
};

const SECTION_META = {
  overview: {
    eyebrow: 'Dashboard',
    title: 'Tenant Command Center',
    meta: 'Centraal overzicht van posture, recente runs en operationele status per tenant.',
  },
  assessment: {
    eyebrow: 'Assessment',
    title: 'Assessment Orchestrator',
    meta: 'Start scans, kies fases en werk vanuit een vaste tenantcontext zoals in een MSP-workspace.',
  },
  results: {
    eyebrow: 'Rapportage',
    title: 'Rapporten & Trends',
    meta: 'Werk vanuit een rapportworkspace met viewer, vergelijking, beheer en follow-up acties.',
  },
  herstel: {
    eyebrow: 'Remediation',
    title: 'Herstel & Acties',
    meta: 'Voer herstelacties uit, bekijk risico\'s en houd historie centraal in beeld.',
  },
  gebruikers: {
    eyebrow: 'Identiteit',
    title: 'Gebruikers & Licenties',
    meta: 'Beheer identities, licenties en provisioning-historie vanuit een vaste beheerstructuur.',
  },
  teams: {
    eyebrow: 'Samenwerking',
    title: 'Teams Workspace',
    meta: 'Werk tenant-specifiek met Teams, leden, owners en zichtbaarheid zonder SharePoint-koppelingen in dezelfde werkruimte.',
  },
  sharepoint: {
    eyebrow: 'Samenwerking',
    title: 'SharePoint Workspace',
    meta: 'Werk tenant-specifiek met SharePoint-sites, tenantinstellingen en SharePoint-backup vanuit één losse werkruimte.',
  },
  identity: {
    eyebrow: 'Security',
    title: 'Identity & Access',
    meta: 'Werk tenant-specifiek met MFA, gasten, adminrollen en authenticatie-inzichten.',
  },
  apps: {
    eyebrow: 'Security',
    title: 'App Registrations',
    meta: 'Live overzicht van app-registraties, secrets, certificaten en gerelateerde risico-indicatoren.',
  },
  baseline: {
    eyebrow: 'Compliance',
    title: 'Baseline Engine',
    meta: 'Werk met gold tenants, assignments en historie om tenantconfiguraties consistent te houden.',
  },
  intune: {
    eyebrow: 'Devices',
    title: 'Intune Operations',
    meta: 'Apparaten, compliance en configuratie in een beheeropzet met duidelijke vervolgstappen.',
  },
  backup: {
    eyebrow: 'Protection',
    title: 'Backup Monitoring',
    meta: 'Toon status, dekking en historie voor SharePoint, OneDrive en Exchange.',
  },
  ca: {
    eyebrow: 'Security',
    title: 'Conditional Access',
    meta: 'Beheer policies en named locations met directe context voor risico en impact.',
  },
  domains: {
    eyebrow: 'Security',
    title: 'Domein Analyse',
    meta: 'Analyseer e-mailbeveiliging en DNS-records met concrete aanbevelingen aan de zijkant.',
  },
  alerts: {
    eyebrow: 'Monitoring',
    title: 'Alerts & Signalering',
    meta: 'Combineer audit events, secure score en notificaties in een centrale signaleringslaag.',
  },
  exchange: {
    eyebrow: 'Messaging',
    title: 'Exchange Review',
    meta: 'Mailboxen, forwarding en regels met focus op verdachte patronen en vervolgacties.',
  },
  kb: {
    eyebrow: 'Documentatie',
    title: 'Kennisbank Workspace',
    meta: 'Navigeer tenantdocumentatie, inventaris en wijzigingen vanuit een vaste knowledge-rail.',
  },
  settings: {
    eyebrow: 'Administratie',
    title: 'Platform Beheer',
    meta: 'Tenantbeheer, lokale configuratie en integraties in een eigen admin-workspace.',
  },
  tenantoverzicht: {
    eyebrow: 'Administrator',
    title: 'Tenant Overzicht',
    meta: 'Health scores, bevindingen en laatste scandatum van alle beheerde tenants.',
  },
  bevindingen: {
    eyebrow: 'Security',
    title: 'Bevindingen & Health',
    meta: 'Gestructureerde security-bevindingen per domein op basis van live scandata.',
  },
  compliance: {
    eyebrow: 'Compliance',
    title: 'Compliance & Benchmarks',
    meta: 'CIS Benchmark-score en afwijkingen t.o.v. het geharde referentiemodel per controledomein.',
  },
  hybrid: {
    eyebrow: 'Infrastructuur',
    title: 'Hybrid & AD Connect',
    meta: 'Synchronisatiestatus, objectfouten en configuratiecheck voor Azure AD Connect en hybride identiteiten.',
  },
};
window.SECTION_META = SECTION_META;

// Mapping van sectie → nav-groep label (voor breadcrumb in subnav)
const NAV_GROUP_MAP = {
  gebruikers: 'Gebruikers & Identiteit',
  identity:   'Gebruikers & Identiteit',
  apps:       'Gebruikers & Identiteit',
  ca:         'Gebruikers & Identiteit',
  hybrid:     'Gebruikers & Identiteit',
  alerts:     'Security & Compliance',
  compliance: 'Security & Compliance',
  bevindingen:'Security & Compliance',
  teams:      'Samenwerking & Email',
  sharepoint: 'Samenwerking & Email',
  exchange:   'Samenwerking & Email',
  domains:    'Samenwerking & Email',
  backup:     'Samenwerking & Email',
  intune:     'Devices & Intune',
  assessment: 'Assessment & Opvolging',
  results:    'Assessment & Opvolging',
  herstel:    'Assessment & Opvolging',
  baseline:   'Assessment & Opvolging',
  kb:         'Kennisbank',
  tenantoverzicht: 'MSP Admin',
  settings:   'MSP Admin',
  klantenbeheer: 'MSP Admin',
};

const UI_PREFS_KEY = 'portal_ui_prefs_v1';
const QUICK_ACTIONS = {
  overview: [
    { id: 'refreshWorkspace', label: 'Ververs workspace', kind: 'secondary' },
    { id: 'goResults', label: 'Rapporten', kind: 'ghost' },
    { id: 'goKb', label: 'Kennisbank', kind: 'ghost' },
  ],
  assessment: [
    { id: 'refreshWorkspace', label: 'Ververs workspace', kind: 'secondary' },
    { id: 'goAssessment', label: 'Start assessment', kind: 'primary' },
    { id: 'goResults', label: 'Open rapporten', kind: 'ghost' },
  ],
  results: [
    { id: 'refreshWorkspace', label: 'Ververs workspace', kind: 'secondary' },
    { id: 'resultsViewer', label: 'Viewer', kind: 'primary' },
    { id: 'resultsActions', label: 'Acties', kind: 'ghost' },
  ],
  herstel: [
    { id: 'refreshWorkspace', label: 'Ververs workspace', kind: 'secondary' },
    { id: 'goRemCatalog', label: 'Catalogus', kind: 'primary' },
    { id: 'goRemHistory', label: 'Geschiedenis', kind: 'ghost' },
  ],
  gebruikers: [
    { id: 'refreshWorkspace', label: 'Ververs workspace', kind: 'secondary' },
    { id: 'scanUsersLive', label: 'Live scan gebruikers', kind: 'ghost' },
    { id: 'goUsers', label: 'Gebruikers', kind: 'primary' },
    { id: 'goLicenses', label: 'Licenties', kind: 'ghost' },
  ],
  teams: [
    { id: 'refreshWorkspace', label: 'Ververs workspace', kind: 'secondary' },
    { id: 'goTeams', label: 'Teams', kind: 'primary' },
  ],
  sharepoint: [
    { id: 'refreshWorkspace', label: 'Ververs workspace', kind: 'secondary' },
    { id: 'goSharePointSites', label: 'Sites', kind: 'primary' },
    { id: 'goSharePointBackup', label: 'Backup', kind: 'ghost' },
  ],
  identity: [
    { id: 'refreshWorkspace', label: 'Ververs workspace', kind: 'secondary' },
    { id: 'goIdentityMfa', label: 'MFA', kind: 'primary' },
    { id: 'goIdentityGuests', label: 'Guests', kind: 'ghost' },
  ],
  apps: [
    { id: 'refreshWorkspace', label: 'Ververs workspace', kind: 'secondary' },
    { id: 'goAppsRegistrations', label: 'Registrations', kind: 'primary' },
  ],
  baseline: [
    { id: 'refreshWorkspace', label: 'Ververs workspace', kind: 'secondary' },
    { id: 'goBaselines', label: 'Baselines', kind: 'primary' },
    { id: 'goAssignments', label: 'Koppelingen', kind: 'ghost' },
  ],
  intune: [
    { id: 'refreshWorkspace', label: 'Ververs workspace', kind: 'secondary' },
    { id: 'goDevices', label: 'Apparaten', kind: 'primary' },
    { id: 'goCompliance', label: 'Compliance', kind: 'ghost' },
  ],
  backup: [
    { id: 'refreshWorkspace', label: 'Ververs workspace', kind: 'secondary' },
    { id: 'goBackupOverview', label: 'Overzicht', kind: 'primary' },
    { id: 'goBackupHistory', label: 'Geschiedenis', kind: 'ghost' },
  ],
  ca: [
    { id: 'refreshWorkspace', label: 'Ververs workspace', kind: 'secondary' },
    { id: 'goCaPolicies', label: 'Policies', kind: 'primary' },
    { id: 'goCaLocations', label: 'Locations', kind: 'ghost' },
  ],
  domains: [
    { id: 'refreshWorkspace', label: 'Ververs workspace', kind: 'secondary' },
    { id: 'goDomains', label: 'Analyse openen', kind: 'primary' },
  ],
  alerts: [
    { id: 'refreshWorkspace', label: 'Ververs workspace', kind: 'secondary' },
    { id: 'goAlertsAudit', label: 'Audit Log', kind: 'primary' },
    { id: 'goAlertsScore', label: 'Secure Score', kind: 'ghost' },
  ],
  exchange: [
    { id: 'refreshWorkspace', label: 'Ververs workspace', kind: 'secondary' },
    { id: 'goExchangeMail', label: 'Mailboxen', kind: 'primary' },
    { id: 'goExchangeRules', label: 'Regels', kind: 'ghost' },
  ],
  kb: [
    { id: 'refreshWorkspace', label: 'Ververs workspace', kind: 'secondary' },
    { id: 'goKbAssets', label: 'Apparaten', kind: 'primary' },
    { id: 'goKbChanges', label: 'Wijzigingslog', kind: 'ghost' },
  ],
  settings: [
    { id: 'refreshWorkspace', label: 'Ververs workspace', kind: 'secondary' },
    { id: 'goSettingsTenant', label: 'Tenants', kind: 'primary' },
    { id: 'goSettingsIntegrations', label: 'Integraties', kind: 'ghost' },
  ],
};

const NAV_GROUP_SECTIONS = {
  gbid: [
    'gebruikers',
    { section: 'identity', subItems: ['mfa', 'guests', 'admin-roles'] },
    'ca',
    'hybrid',
  ],
  security: [
    'alerts',
    'apps',
    'bevindingen',
    'compliance',
    { section: 'identity', subItems: ['legacy-auth'] },
  ],
  collab: ['exchange', 'teams', 'sharepoint', 'domains', 'backup'],
  devices: [
    { section: 'intune', subItems: ['overzicht', 'apparaten', 'compliance', 'configuratie', 'geschiedenis'] },
  ],
  followup: ['assessment', 'results', 'herstel', 'baseline'],
  kb: ['kb'],
  admin: ['settings', 'tenantoverzicht', 'klantenbeheer', 'goedkeuringen', 'kosten', 'jobmonitor'],
};

let _currentSection = 'overview';
let _currentSubItem = null;
let _contextRailOpen = true;
let _sidebarCompact = false;

function getUiPrefs() {
  try {
    return JSON.parse(localStorage.getItem(UI_PREFS_KEY) || '{}');
  } catch (_) {
    return {};
  }
}

function saveUiPrefs(patch) {
  const next = { ...getUiPrefs(), ...patch };
  try { localStorage.setItem(UI_PREFS_KEY, JSON.stringify(next)); } catch (_) {}
}

function getCurrentTenantLabel() {
  const tenant = allTenants.find((item) => item.id === currentTenantId);
  return tenant ? (tenant.customer_name || tenant.tenant_name || 'Actieve tenant') : 'Geen tenant geselecteerd';
}

function parseMetricValue(text) {
  const match = String(text || '').match(/-?\d+/);
  return match ? Number(match[0]) : null;
}

function formatCompactBytes(bytes) {
  const value = Number(bytes);
  if (!Number.isFinite(value) || value <= 0) return '—';
  if (value >= 1024 ** 4) return `${(value / (1024 ** 4)).toFixed(1)} TB`;
  if (value >= 1024 ** 3) return `${(value / (1024 ** 3)).toFixed(1)} GB`;
  if (value >= 1024 ** 2) return `${(value / (1024 ** 2)).toFixed(1)} MB`;
  return `${Math.round(value / 1024)} KB`;
}

function setTextContent(id, value) {
  const el = document.getElementById(id);
  if (el) el.textContent = value;
}

function parseSourceDate(value) {
  if (!value) return null;
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? null : date;
}

function formatSourceStamp(value) {
  if (!value) return 'datum onbekend';
  const date = parseSourceDate(value);
  if (!date) return `laatst: ${String(value)}`;
  try {
    return `laatst: ${date.toLocaleString('nl-NL', { dateStyle: 'short', timeStyle: 'short' })}`;
  } catch (_) {
    return `laatst: ${value}`;
  }
}

function formatSourceAge(value) {
  const date = parseSourceDate(value);
  if (!date) return formatSourceStamp(value);
  const ageMs = Date.now() - date.getTime();
  const ageMin = Math.max(0, Math.round(ageMs / 60000));
  if (ageMin < 1) return 'zojuist';
  if (ageMin < 60) return `${ageMin} min oud`;
  const ageHours = Math.round(ageMin / 60);
  if (ageHours < 24) return `${ageHours} uur oud`;
  const ageDays = Math.round(ageHours / 24);
  return `${ageDays} dag${ageDays === 1 ? '' : 'en'} oud`;
}

function describeSourceMeta(data) {
  const isAssessment = data && data._source === 'assessment_snapshot';
  const generatedAt = data?._generated_at || data?.generated_at || data?.assessment_generated_at || data?.createdAt || null;
  const stale = isAssessment ? (typeof data?._stale === 'boolean' ? data._stale : ((Date.now() - (parseSourceDate(generatedAt)?.getTime() || Date.now())) > 30 * 60 * 1000)) : false;
  return {
    label: isAssessment ? 'Assessment' : 'Live',
    detail: isAssessment ? formatSourceAge(generatedAt) : 'actueel',
    className: isAssessment ? (stale ? 'is-stale' : 'is-assessment') : 'is-live',
  };
}

window.denjoyDescribeSourceMeta = describeSourceMeta;

function setOverviewServiceCard(prefix, primary = '—', secondary = '—', meta = 'Live data niet beschikbaar', data = null) {
  setTextContent(`overview${prefix}Primary`, primary);
  setTextContent(`overview${prefix}Secondary`, secondary);
  setTextContent(`overview${prefix}Meta`, meta);
  const source = document.getElementById(`overview${prefix}Source`);
  if (source) {
    if (!data) {
      source.textContent = 'Bron onbekend';
      source.className = 'ov-service-source';
      return;
    }
    const info = describeSourceMeta(data);
    source.textContent = `${info.label} · ${info.detail}`;
    source.className = `ov-service-source ${info.className}`;
  }
}

function resetOverviewServiceCards() {
  setOverviewServiceCard('Exchange', '—', 'Mailboxen', 'Open module', null);
  setOverviewServiceCard('Teams', '—', 'Teams', 'Open module', null);
  setOverviewServiceCard('SharePoint', '—', 'Sites', 'Open module', null);
  setOverviewServiceCard('OneDrive', '—', 'Accounts', 'Open module', null);
  setOverviewServiceCard('Licenses', '—', 'Toegewezen', 'Open module', null);
}

function bindOverviewActions() {
  document.querySelectorAll('[data-overview-nav]').forEach((btn) => {
    if (btn.dataset.ovBound === '1') return;
    btn.dataset.ovBound = '1';
    btn.addEventListener('click', () => {
      const section = btn.dataset.overviewNav;
      if (!section) return;
      const opts = {};
      if (btn.dataset.liveTab) opts.liveTab = btn.dataset.liveTab;
      if (btn.dataset.gbTab) opts.gbTab = btn.dataset.gbTab;
      if (btn.dataset.caTab) opts.caTab = btn.dataset.caTab;
      showSection(section, opts);
    });
  });
  const refreshBtn = document.getElementById('overviewRefreshButton');
  if (refreshBtn && refreshBtn.dataset.ovBound !== '1') {
    refreshBtn.dataset.ovBound = '1';
    refreshBtn.addEventListener('click', () => refreshTenantData());
  }
}

async function loadOverviewServiceCards() {
  resetOverviewServiceCards();
  if (!currentTenantId) return;

  const requests = await Promise.allSettled([
    apiFetch(API.exchange.mailboxes(currentTenantId)),
    apiFetch(API.collaboration.teams(currentTenantId)),
    apiFetch(API.collaboration.sharepointSites(currentTenantId)),
    apiFetch(API.backup.summary(currentTenantId)),
    apiFetch(API.m365.licenses(currentTenantId)),
  ]);

  const [exchangeRes, teamsRes, sharePointRes, backupRes, licenseRes] = requests;

  if (exchangeRes.status === 'fulfilled' && exchangeRes.value?.ok) {
    const payload = exchangeRes.value;
    const total = Number(payload.count || payload.mailboxes?.length || 0);
    const active = (payload.mailboxes || []).filter((item) => item && item.accountEnabled !== false).length;
    setOverviewServiceCard('Exchange', String(total), 'Mailboxen', active ? `${active} actief` : 'Mailboxoverzicht', payload);
  } else {
    setOverviewServiceCard('Exchange', '—', 'Mailboxen', 'Live data niet beschikbaar', null);
  }

  if (teamsRes.status === 'fulfilled' && teamsRes.value?.ok) {
    const payload = teamsRes.value;
    const total = Number(payload.count || payload.teams?.length || 0);
    const publicCount = Number(payload.publicCount || 0);
    setOverviewServiceCard('Teams', String(total), 'Teams', publicCount ? `${publicCount} publiek` : 'Teams-overzicht', payload);
  } else {
    setOverviewServiceCard('Teams', '—', 'Teams', 'Live data niet beschikbaar', null);
  }

  if (sharePointRes.status === 'fulfilled' && sharePointRes.value?.ok) {
    const payload = sharePointRes.value;
    const total = Number(payload.count || payload.sites?.length || 0);
    const storage = (payload.sites || []).reduce((sum, item) => sum + (Number(item?.storageUsed) || 0), 0);
    setOverviewServiceCard('SharePoint', String(total), 'Sites', storage > 0 ? `${formatCompactBytes(storage)} opslag` : 'Sites-overzicht', payload);
  } else {
    setOverviewServiceCard('SharePoint', '—', 'Sites', 'Live data niet beschikbaar', null);
  }

  if (backupRes.status === 'fulfilled' && backupRes.value?.ok) {
    const payload = backupRes.value;
    const count = Number(payload.oneDrive?.resourceCount || 0);
    const policies = Number(payload.oneDrive?.policyCount || 0);
    setOverviewServiceCard('OneDrive', String(count), 'Accounts', policies ? `${policies} policy${policies === 1 ? '' : '\'s'}` : 'Backup-overzicht', payload);
  } else {
    setOverviewServiceCard('OneDrive', '—', 'Accounts', 'Live data niet beschikbaar', null);
  }

  if (licenseRes.status === 'fulfilled' && licenseRes.value?.ok) {
    const licenses = licenseRes.value.licenses || [];
    const assigned = licenses.reduce((sum, item) => sum + (Number(item?.consumed) || 0), 0);
    const available = licenses.reduce((sum, item) => sum + (Number(item?.available) || 0), 0);
    setOverviewServiceCard('Licenses', String(assigned), 'Toegewezen', `${available} beschikbaar`, licenseRes.value);
  } else {
    setOverviewServiceCard('Licenses', '—', 'Toegewezen', 'Live data niet beschikbaar', null);
  }
}

function updateWorkspaceHeader(sectionName) {
  const meta = SECTION_META[sectionName] || SECTION_META.overview;
  const eyebrowEl = document.getElementById('workspaceEyebrow');
  const titleEl = document.getElementById('workspaceTitle');
  const metaEl = document.getElementById('workspaceMeta');
  if (eyebrowEl) eyebrowEl.textContent = meta.eyebrow;
  if (titleEl) titleEl.textContent = meta.title;
  if (metaEl) metaEl.textContent = `${meta.meta} Tenant: ${getCurrentTenantLabel()}.`;
  renderWorkspaceActions(sectionName);
}

function setSidebarCompact(isCompact) {
  _sidebarCompact = !!isCompact;
  document.body.classList.toggle('portal-sidebar-compact', _sidebarCompact);
  const btn = document.getElementById('sidebarDensityButton');
  if (btn) {
    btn.classList.toggle('active', _sidebarCompact);
    btn.title = _sidebarCompact ? 'Normale sidebar' : 'Compacte sidebar';
  }
  saveUiPrefs({ sidebarCompact: _sidebarCompact });
}

function getQuickActionHandlers() {
  return {
    refreshWorkspace: () => refreshTenantData(),
    scanUsersLive: () => {
      showSection('gebruikers', { gbTab: 'gebruikers' });
      if (typeof scanGebruikersLive === 'function') scanGebruikersLive();
    },
    goResults: () => showSection('results', { resultsPanel: 'viewer' }),
    goKb: () => showSection('kb', { kbTab: 'overview' }),
    goAssessment: () => showSection('assessment'),
    resultsViewer: () => showSection('results', { resultsPanel: 'viewer' }),
    resultsActions: () => showSection('results', { resultsPanel: 'actions' }),
    goRemCatalog: () => showSection('herstel', { remTab: 'catalogus' }),
    goRemHistory: () => showSection('herstel', { remTab: 'geschiedenis' }),
    goUsers: () => showSection('gebruikers', { gbTab: 'gebruikers' }),
    goLicenses: () => showSection('gebruikers', { gbTab: 'licenties' }),
    goTeams: () => showSection('teams', { liveTab: 'teams' }),
    goSharePointSites: () => showSection('sharepoint', { liveTab: 'sharepoint-sites' }),
    goSharePointBackup: () => showSection('sharepoint', { liveTab: 'sharepoint-backup' }),
    goIdentityMfa: () => showSection('identity', { liveTab: 'mfa' }),
    goIdentityGuests: () => showSection('identity', { liveTab: 'guests' }),
    goAppsRegistrations: () => showSection('apps', { liveTab: 'registrations' }),
    goBaselines: () => showSection('baseline', { baselineTab: 'baselines' }),
    goAssignments: () => showSection('baseline', { baselineTab: 'assignments' }),
    goDevices: () => showSection('intune', { liveTab: 'apparaten' }),
    goCompliance: () => showSection('intune', { liveTab: 'compliance' }),
    goBackupOverview: () => showSection('backup', { liveTab: 'overzicht' }),
    goBackupHistory: () => showSection('backup', { liveTab: 'geschiedenis' }),
    goCaPolicies: () => showSection('ca', { caTab: 'policies' }),
    goCaLocations: () => showSection('ca', { caTab: 'locations' }),
    goDomains: () => showSection('domains', { liveTab: 'domains-list' }),
    goAlertsAudit: () => showSection('alerts', { liveTab: 'auditlog' }),
    goAlertsScore: () => showSection('alerts', { liveTab: 'securescr' }),
    goExchangeMail: () => showSection('exchange', { liveTab: 'mailboxen' }),
    goExchangeRules: () => showSection('exchange', { liveTab: 'regels' }),
    goKbAssets: () => showSection('kb', { kbTab: 'assets' }),
    goKbChanges: () => showSection('kb', { kbTab: 'changelog' }),
    goSettingsTenant: () => showSection('settings', { settingsTab: 'tenant' }),
    goSettingsIntegrations: () => showSection('settings', { settingsTab: 'integrations' }),
  };
}

function renderWorkspaceActions(sectionName = _currentSection) {
  const root = document.getElementById('workspaceQuickActions');
  if (!root) return;
  const actions = QUICK_ACTIONS[sectionName] || QUICK_ACTIONS.overview;
  root.innerHTML = actions.map((item) => `
    <button type="button" class="workspace-action-btn workspace-action-btn--${escapeHtml(item.kind || 'ghost')}" data-workspace-action="${escapeHtml(item.id)}">
      ${escapeHtml(item.label)}
    </button>
  `).join('');
  const handlers = getQuickActionHandlers();
  root.querySelectorAll('[data-workspace-action]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const handler = handlers[btn.dataset.workspaceAction];
      if (typeof handler === 'function') handler();
    });
  });
}

function getContextEntries(sectionName) {
  const tenantLabel = getCurrentTenantLabel();
  if (!currentTenantId) {
    return [{
      tone: 'info',
      title: 'Selecteer eerst een tenant',
      body: 'Kies links een tenant om aanbevelingen, opmerkingen en urgente signalen te laden.',
    }];
  }

  if (sectionName === 'overview') {
    const score = parseMetricValue(document.getElementById('secureScore')?.textContent);
    const mfa = parseMetricValue(document.getElementById('mfaStatus')?.textContent);
    const policies = parseMetricValue(document.getElementById('caPolicies')?.textContent);
    const assessments = document.querySelectorAll('#recentAssessmentsList .assessment-item').length;
    const entries = [{
      tone: 'info',
      title: tenantLabel,
      body: 'Gebruik dit paneel als snelle operator-samenvatting met aandachtspunten voordat je een module induikt.',
    }];
    if (score != null && score < 65) {
      entries.push({ tone: 'urgent', title: 'Secure Score vraagt aandacht', body: `De huidige secure score staat rond ${score}%. Prioriteer rapportvergelijking en openstaande acties.`, badge: score });
    }
    if (mfa != null && mfa < 90) {
      entries.push({ tone: 'warn', title: 'MFA-dekking is niet volledig', body: `De huidige MFA-coverage is ${mfa}%. Controleer gebruikersbeheer en Conditional Access.`, badge: `${mfa}%` });
    }
    if (policies != null && policies < 3) {
      entries.push({ tone: 'warn', title: 'Beperkte CA-set', body: `Er lijken maar ${policies} CA-policies actief. Controleer minimaal basisblokkades en admin-beveiliging.`, badge: policies });
    }
    entries.push({ tone: 'info', title: 'Recente assessments', body: `${assessments || 0} recente run(s) beschikbaar voor deze tenant.`, badge: assessments || 0 });
    return entries;
  }

  if (sectionName === 'results') {
    const critical = parseMetricValue(document.getElementById('kpiCritical')?.textContent);
    const warning = parseMetricValue(document.getElementById('kpiWarning')?.textContent);
    const reportRuns = parseMetricValue(document.getElementById('metaReportRuns')?.textContent);
    const entries = [{
      tone: 'info',
      title: 'Rapportworkspace',
      body: 'Werk vanuit viewer, vergelijking en acties om bevindingen direct om te zetten naar opvolging.',
    }];
    if (critical != null && critical > 0) {
      entries.push({ tone: 'urgent', title: 'Kritieke findings gedetecteerd', body: `${critical} kritieke finding(s) vragen directe opvolging of escalatie.`, badge: critical });
    }
    if (warning != null && warning > 0) {
      entries.push({ tone: 'warn', title: 'Waarschuwingen beschikbaar', body: `${warning} waarschuwing(en) kunnen worden omgezet in acties of baseline-wijzigingen.`, badge: warning });
    }
    entries.push({ tone: 'info', title: 'Rapportgeschiedenis', body: `${reportRuns || 0} rapport-run(s) beschikbaar voor vergelijking en retentiebeheer.`, badge: reportRuns || 0 });
    return entries;
  }

  if (sectionName === 'kb') {
    const countIds = ['nbCountAssets', 'nbCountPages', 'nbCountContacts', 'nbCountSoftware', 'nbCountDomains', 'nbCountChangelog'];
    const totalKnown = countIds.reduce((sum, id) => sum + (parseMetricValue(document.getElementById(id)?.textContent) || 0), 0);
    return [
      { tone: 'info', title: 'Knowledge posture', body: `Ongeveer ${totalKnown} gedocumenteerde items zichtbaar. Gebruik dit om lacunes in documentatie snel te herkennen.`, badge: totalKnown },
      { tone: 'warn', title: 'Let op actualiteit', body: 'Controleer vooral passwords, changelog en documenten als hier veel operationele wijzigingen lopen.' },
    ];
  }

  const introText = document.querySelector(`#${sectionName}Section .section-intro`)?.textContent?.trim();
  const subnavLabels = (SUBNAV_CONFIG[sectionName] || []).map((item) => item.label).join(', ');
  const entries = [{
    tone: 'info',
    title: SECTION_META[sectionName]?.title || 'Werkruimte',
    body: introText || SECTION_META[sectionName]?.meta || 'Moduleoverzicht voor de geselecteerde tenant.',
  }];
  if (subnavLabels) {
    entries.push({ tone: 'info', title: 'Snelle routes', body: `Beschikbare onderdelen: ${subnavLabels}.` });
  }
  if (['herstel', 'alerts', 'exchange', 'ca'].includes(sectionName)) {
    entries.push({ tone: 'warn', title: 'Controleer impact', body: 'Wijzigingen in deze module kunnen direct effect hebben op bereikbaarheid, authenticatie of compliance.' });
  }
  const liveContext = window.denjoyGetLiveModuleContext?.();
  if (liveContext && liveContext.section === sectionName) {
    const capability = liveContext.capability || null;
    const describe = window.denjoyDescribeCapabilityStatus;
    const subnavItems = SUBNAV_CONFIG[sectionName] || [];
    const activeSubnav = subnavItems.find((item) => item.liveTab === liveContext.tab);
    const activeLabel = activeSubnav?.label || liveContext.tab || 'Subhoofdstuk';
    if (typeof describe === 'function' && capability) {
      const info = describe(capability);
      entries.push({
        tone: info.className === 'is-live' ? 'info' : (info.className === 'is-warn' ? 'warn' : 'info'),
        title: `${activeLabel} status`,
        body: info.detail || 'Capabilitystatus beschikbaar voor dit subhoofdstuk.',
      });
      const engine = capability.engine || '—';
      const roles = (capability.extra_roles || []).length ? capability.extra_roles.join(', ') : 'Geen extra rollen';
      const consent = (capability.extra_consent || []).length ? capability.extra_consent.join(', ') : 'Geen extra consent';
      entries.push({
        tone: 'info',
        title: `${activeLabel} toegang`,
        body: `Engine: ${engine}. Rollen: ${roles}. Consent: ${consent}.`,
      });
    }
  }
  return entries;
}

function setContextRailOpen(isOpen) {
  _contextRailOpen = !!isOpen;
  const rail = document.getElementById('portalContextRail');
  const toggle = document.getElementById('portalContextToggle');
  const contentArea = document.querySelector('.content-area');
  if (rail) rail.classList.toggle('open', _contextRailOpen);
  if (contentArea) contentArea.classList.toggle('with-context-rail', _contextRailOpen);
  if (toggle) {
    toggle.classList.toggle('is-open', _contextRailOpen);
    toggle.setAttribute('aria-expanded', String(_contextRailOpen));
    toggle.textContent = _contextRailOpen ? '✕ Inzichten' : 'Inzichten';
  }
  saveUiPrefs({ contextRailOpen: _contextRailOpen });
}

function renderContextRail(sectionName = _currentSection) {
  const titleEl = document.getElementById('portalContextTitle');
  const contentEl = document.getElementById('portalContextContent');
  if (!titleEl || !contentEl) return;
  // Reset detail-modus als we van sectie wisselen
  document.getElementById('portalContextRail')?.classList.remove('portal-context-rail--detail');
  document.querySelector('.content-area')?.classList.remove('detail-rail-open');
  titleEl.textContent = SECTION_META[sectionName]?.title || 'Inzichten';
  const entries = getContextEntries(sectionName);
  contentEl.innerHTML = entries.map((item) => `
    <article class="portal-context-card portal-context-card--${escapeHtml(item.tone || 'info')}">
      <div class="portal-context-card-top">
        <span class="portal-context-pill">${escapeHtml((item.tone || 'info').toUpperCase())}</span>
        ${item.badge != null ? `<span class="portal-context-count">${escapeHtml(item.badge)}</span>` : ''}
      </div>
      <h4>${escapeHtml(item.title || 'Notitie')}</h4>
      <p>${escapeHtml(item.body || '')}</p>
    </article>
  `).join('');
}

// ── Detail side-rail ─────────────────────────────────────────────────────────
// Opent het Inzichten-paneel met item-detail in plaats van een modale popup.
// Gebruik window.openSideRailDetail(kicker, title) om te openen,
// daarna window.updateSideRailDetail(title, html) voor async geladen inhoud.

function openSideRailDetail(kicker, title) {
  const titleEl = document.getElementById('portalContextTitle');
  const contentEl = document.getElementById('portalContextContent');
  const rail = document.getElementById('portalContextRail');
  const contentArea = document.querySelector('.content-area');
  if (!contentEl) return;

  // Sla huidige context op voor de terug-knop
  const savedTitle = titleEl ? titleEl.textContent : '';
  const savedHtml  = contentEl.innerHTML;

  if (titleEl) titleEl.textContent = title || 'Details';
  if (rail) rail.classList.add('portal-context-rail--detail');
  if (contentArea) contentArea.classList.add('detail-rail-open');

  contentEl.innerHTML = `
    <div class="dp-header">
      <span class="dp-kicker">${escapeHtml(kicker || 'Detail')}</span>
      <h4 class="dp-title">${escapeHtml(title || 'Details')}</h4>
      <button type="button" class="dp-back" id="dpBackBtn">← Terug naar inzichten</button>
    </div>
    <div class="dp-body" id="dpBody">
      <p class="live-module-empty">Laden…</p>
    </div>
  `;

  document.getElementById('dpBackBtn')?.addEventListener('click', () => {
    if (titleEl) titleEl.textContent = savedTitle;
    contentEl.innerHTML = savedHtml;
    if (rail) rail.classList.remove('portal-context-rail--detail');
    if (contentArea) contentArea.classList.remove('detail-rail-open');
  });

  if (!_contextRailOpen) setContextRailOpen(true);
}
window.openSideRailDetail = openSideRailDetail;

function updateSideRailDetail(title, bodyHtml) {
  const titleEl = document.getElementById('portalContextTitle');
  if (titleEl && title) titleEl.textContent = title;
  const body = document.getElementById('dpBody');
  if (body) body.innerHTML = bodyHtml || '<p class="live-module-empty">Geen data beschikbaar.</p>';
}
window.updateSideRailDetail = updateSideRailDetail;

function setNavSignal(target, text, tone = 'info') {
  const host = document.querySelector(target);
  if (!host) return;
  let badge = host.querySelector('.portal-nav-signal');
  if (!text && text !== 0) {
    if (badge) badge.remove();
    return;
  }
  if (!badge) {
    badge = document.createElement('span');
    badge.className = 'portal-nav-signal';
    host.appendChild(badge);
  }
  badge.className = `portal-nav-signal portal-nav-signal--${tone}`;
  badge.textContent = String(text);
}

function getSectionOptionsFromDataset(dataset) {
  const opts = {};
  if (dataset.resultsPanel) opts.resultsPanel = dataset.resultsPanel;
  if (dataset.kbTab) opts.kbTab = dataset.kbTab;
  if (dataset.settingsTab) opts.settingsTab = dataset.settingsTab;
  if (dataset.remTab) opts.remTab = dataset.remTab;
  if (dataset.gbTab) opts.gbTab = dataset.gbTab;
  if (dataset.baselineTab) opts.baselineTab = dataset.baselineTab;
  if (dataset.itTab) opts.itTab = dataset.itTab;
  if (dataset.bkTab) opts.bkTab = dataset.bkTab;
  if (dataset.caTab) opts.caTab = dataset.caTab;
  if (dataset.alTab) opts.alTab = dataset.alTab;
  if (dataset.exTab) opts.exTab = dataset.exTab;
  if (dataset.liveTab) opts.liveTab = dataset.liveTab;
  return opts;
}

function renderNavSignals() {
  const score = parseMetricValue(document.getElementById('secureScore')?.textContent);
  const critical = parseMetricValue(document.getElementById('kpiCritical')?.textContent);
  const reportRuns = parseMetricValue(document.getElementById('metaReportRuns')?.textContent);
  const kbCountIds = ['nbCountAssets', 'nbCountPages', 'nbCountContacts', 'nbCountSoftware', 'nbCountDomains', 'nbCountChangelog'];
  const kbTotal = kbCountIds.reduce((sum, id) => sum + (parseMetricValue(document.getElementById(id)?.textContent) || 0), 0);
  const userCount = parseMetricValue(document.getElementById('userCount')?.textContent);

  setNavSignal('[data-nav-group="followup"] > .portal-nav-link', critical > 0 ? critical : reportRuns || '', critical > 0 ? 'urgent' : 'info');
  setNavSignal('[data-nav-group="kb"] > .portal-nav-link', kbTotal || '', 'info');
  setNavSignal('[data-nav-group="gebruikers"] > .portal-nav-link', userCount || '', 'info');
  setNavSignal('[data-nav-group="identity"] > .portal-nav-link', '', 'info');
  setNavSignal('[data-nav-group="apps"] > .portal-nav-link', '', 'info');
  setNavSignal('[data-nav-group="ca"] > .portal-nav-link', '', 'info');
  setNavSignal('[data-nav-group="alerts"] > .portal-nav-link', score != null ? `${score}%` : '', score != null && score < 65 ? 'warn' : 'info');
  setNavSignal('[data-nav-group="domains"] > .portal-nav-link', '', 'info');
  setNavSignal('[data-nav-group="teams"] > .portal-nav-link', '', 'info');
  setNavSignal('[data-nav-group="sharepoint"] > .portal-nav-link', '', 'info');
  setNavSignal('[data-nav-group="exchange"] > .portal-nav-link', '', 'info');
  setNavSignal('[data-nav-group="intune"] > .portal-nav-link', '', 'info');
  setNavSignal('[data-nav-group="backup"] > .portal-nav-link', '', 'info');
}

function getSubnavItemMeta(item) {
  const pairs = [
    ['kbTab', 'kb'],
    ['settingsTab', 'settings'],
    ['resultsPanel', 'results'],
    ['remTab', 'rem'],
    ['gbTab', 'gebruikers'],
    ['baselineTab', 'baseline'],
    ['itTab', 'intune'],
    ['bkTab', 'backup'],
    ['caTab', 'ca'],
    ['alTab', 'alerts'],
    ['exTab', 'exchange'],
    ['liveTab', 'live'],
    ['section', 'section'],
  ];
  for (const [prop, type] of pairs) {
    if (item[prop]) return { key: item[prop], type };
  }
  return { key: '', type: 'section' };
}

function activateSectionSubtab(sectionName, tabKey) {
  const switchers = {
    herstel: window.switchRemediationTab,
    gebruikers: window.switchGebruikersTab,
    baseline: window.switchBaselineTab,
    intune: window.switchIntuneTab,
    backup: window.switchBackupTab,
    ca: window.switchCaTab,
    alerts: window.switchAlertsTab,
    exchange: window.switchExchangeTab,
  };
  const switcher = switchers[sectionName];
  if (typeof switcher === 'function' && tabKey) switcher(tabKey);
  setActiveSubnavItem(tabKey || null);
}

function updateSubnav(sectionName, activeItem) {
  const subnav = document.getElementById('portalSubnav');
  if (!subnav) return;
  const items = SUBNAV_CONFIG[sectionName] || [];
  if (!items.length) {
    subnav.style.display = 'none';
    return;
  }
  subnav.style.display = '';
  const navGroup = NAV_GROUP_MAP[sectionName] || null;
  const sectionTitle = SECTION_META[sectionName]?.title || null;
  const breadcrumb = (navGroup && sectionTitle)
    ? `<span class="subnav-breadcrumb"><span class="subnav-bc-group">${escapeHtml(navGroup)}</span><span class="subnav-bc-sep">›</span><span class="subnav-bc-section">${escapeHtml(sectionTitle)}</span></span>`
    : '';
  subnav.innerHTML = breadcrumb + items.map((item) => {
    const count = item.countId ? (document.getElementById(item.countId)?.textContent || '') : '';
    const showCount = count && count !== '—' && count !== '';
    const countBadge = showCount ? `<span class="subnav-count">${escapeHtml(count)}</span>` : '';
    const { key, type } = getSubnavItemMeta(item);
    const isActive = activeItem ? key === activeItem : false;
    return `<button type="button" class="subnav-item${isActive ? ' active' : ''}"
      data-subnav-key="${escapeHtml(key)}"
      data-subnav-type="${escapeHtml(type)}"
    >${escapeHtml(item.label)}${countBadge}</button>`;
  }).join('');

  subnav.querySelectorAll('.subnav-item').forEach((btn) => {
    btn.addEventListener('click', () => {
      const key = btn.dataset.subnavKey;
      const type = btn.dataset.subnavType;
      if (type === 'kb') {
        if (typeof kbSwitchTab === 'function') kbSwitchTab(key);
        setActiveSubnavItem(key);
      } else if (type === 'settings') {
        switchSettingsTab(key);
        setActiveSubnavItem(key);
      } else if (type === 'results') {
        showResultsPanel(key);
      } else if (['rem', 'gebruikers', 'baseline', 'intune', 'backup', 'ca', 'alerts', 'exchange'].includes(type)) {
        activateSectionSubtab(_currentSection, key);
      } else if (type === 'live') {
        if (typeof switchLiveModuleTab === 'function') switchLiveModuleTab(_currentSection, key);
        setActiveSubnavItem(key);
      } else if (type === 'section') {
        // item kan ook een resultsPanel target hebben (bijv. subnav → rapporten tab)
        const item = (SUBNAV_CONFIG[_currentSection] || []).find(
          (i) => getSubnavItemMeta(i).key === key
        );
        if (item && item.resultsPanel) {
          showSection('results', { resultsPanel: item.resultsPanel });
        } else {
          showSection(key);
        }
      }
    });
  });
}
window.updateSubnav = updateSubnav;

function setActiveSubnavItem(key) {
  _currentSubItem = key;
  document.querySelectorAll('.subnav-item').forEach((btn) => {
    btn.classList.toggle('active', btn.dataset.subnavKey === key);
  });
}
window.setActiveSubnavItem = setActiveSubnavItem;

// Herlaad tellers in subnav (na KB load)
function refreshSubnavCounts() {
  if (_currentSection !== 'kb') return;
  const subnav = document.getElementById('portalSubnav');
  if (!subnav) return;
  subnav.querySelectorAll('.subnav-item').forEach((btn) => {
    const item = SUBNAV_CONFIG.kb.find((i) => (i.kbTab || '') === btn.dataset.subnavKey);
    if (!item || !item.countId) return;
    const count = document.getElementById(item.countId)?.textContent || '';
    const showCount = count && count !== '—' && count !== '';
    let badge = btn.querySelector('.subnav-count');
    if (showCount) {
      if (!badge) {
        badge = document.createElement('span');
        badge.className = 'subnav-count';
        btn.appendChild(badge);
      }
      badge.textContent = count;
    } else if (badge) {
      badge.remove();
    }
  });
  renderNavSignals();
}
window.refreshSubnavCounts = refreshSubnavCounts;

function showResultsPanel(panelName) {
  // Update panes (.nb-pane[data-results-panel])
  document.querySelectorAll('.nb-pane[data-results-panel]').forEach((el) => {
    el.classList.toggle('active', el.dataset.resultsPanel === panelName);
  });
  // Sync tabbar tabs
  document.querySelectorAll('#resultsTabbar [data-results-panel]').forEach((el) => {
    el.classList.toggle('active', el.dataset.resultsPanel === panelName);
  });
  _currentSubItem = panelName;
  setActiveNav('results');
  setActiveSubnavItem(panelName);
  // Laad panel-specifieke data lazy
  if (panelName === 'diff')         loadRunDiffPanel();
  if (panelName === 'management')   loadReportsManagementPanel();
  if (panelName === 'actions')      loadActionsPanel();
}
window.showResultsPanel = showResultsPanel;

function setDropdownOpen(dropdown, isOpen) {
  if (!dropdown) return;
  dropdown.classList.toggle('open', !!isOpen);
  dropdown.querySelector('.nav-dropdown-toggle')?.setAttribute('aria-expanded', String(!!isOpen));
}

function getNavItemSubKey(item) {
  return item.dataset.resultsPanel
    || item.dataset.kbTab
    || item.dataset.settingsTab
    || item.dataset.remTab
    || item.dataset.gbTab
    || item.dataset.baselineTab
    || item.dataset.itTab
    || item.dataset.bkTab
    || item.dataset.caTab
    || item.dataset.alTab
    || item.dataset.exTab
    || item.dataset.liveTab
    || '';
}

function isNavItemActive(item, navSection) {
  if (!item?.dataset?.section || item.dataset.section !== navSection) return false;
  if (item.classList.contains('nav-dropdown-toggle')) return true;
  const subKey = getNavItemSubKey(item);
  if (!subKey) return true;
  return subKey === _currentSubItem;
}

function setActiveNav(sectionName) {
  const navSection = sectionName === 'history' ? 'results' : sectionName;
  document.querySelectorAll('.portal-nav-link[data-section], .nav-dropdown-link[data-section]').forEach((item) => {
    item.classList.toggle('active', isNavItemActive(item, navSection));
  });
  document.querySelectorAll('.nav-dropdown').forEach((dropdown) => {
    const groupSections = NAV_GROUP_SECTIONS[dropdown.dataset.navGroup || ''] || [];
    const hasActiveChild = groupSections.some((entry) => {
      if (typeof entry === 'string') return entry === navSection;
      if (!entry || entry.section !== navSection) return false;
      if (!Array.isArray(entry.subItems) || !entry.subItems.length) return true;
      return entry.subItems.includes(_currentSubItem);
    });
    dropdown.classList.toggle('has-active', hasActiveChild);
    if (hasActiveChild) setDropdownOpen(dropdown, true);
  });
  // Close groups that don't own the active section, but only when a group was found
  // (prevents collapsing everything when on overview or hub sections)
  const activeGroup = document.querySelector('.nav-dropdown.has-active');
  if (activeGroup) {
    document.querySelectorAll('.nav-dropdown:not(.has-active)').forEach((dropdown) => {
      setDropdownOpen(dropdown, false);
    });
  }
}

function showSection(sectionName, opts = {}) {
  // history is een alias voor results (gecombineerde Rapporten pagina)
  if (sectionName === 'history') sectionName = 'results';
  _currentSection = sectionName;
  _currentSubItem = opts.resultsPanel
    || opts.kbTab
    || opts.settingsTab
    || opts.remTab
    || opts.gbTab
    || opts.liveTab
    || opts.baselineTab
    || opts.itTab
    || opts.bkTab
    || opts.caTab
    || opts.alTab
    || opts.exTab
    || (sectionName === 'assessment' ? 'assessment' : null);
  window._currentSection = sectionName;
  updateWorkspaceHeader(sectionName);
  renderContextRail(sectionName);

  document.querySelectorAll('.content-section').forEach((s) => s.classList.remove('active'));
  const el = document.getElementById(`${sectionName}Section`);
  if (el) el.classList.add('active');
  setActiveNav(sectionName);

  if (sectionName === 'assessment' && typeof loadAssessmentExperience === 'function') {
    loadAssessmentExperience();
  }
  if (sectionName === 'results') {
    updateSubnav('results', opts.resultsPanel || 'viewer');
    _currentSubItem = opts.resultsPanel || 'viewer';
    loadResultsSection().then(() => {
      showResultsPanel(opts.resultsPanel || 'viewer');
    });
    return;
  }
  if (sectionName === 'settings') {
    populateSettings();
    const activeTab = opts.settingsTab || 'tenant';
    switchSettingsTab(activeTab);
    updateSubnav('settings', activeTab);
    _currentSubItem = activeTab;
    return;
  }
  if (sectionName === 'kb') {
    const activeTab = opts.kbTab || 'overview';
    if (typeof kbSwitchTab === 'function') kbSwitchTab(activeTab);
    updateSubnav('kb', activeTab);
    _currentSubItem = activeTab;
    // Tellers na kleine delay (KB laadt async)
    setTimeout(refreshSubnavCounts, 600);
    return;
  }
  if (sectionName === 'herstel') {
    const activeTab = opts.remTab || 'catalogus';
    updateSubnav('herstel', activeTab);
    _currentSubItem = activeTab;
    if (typeof loadHerstellSection === 'function') loadHerstellSection();
    activateSectionSubtab('herstel', activeTab);
    return;
  }
  if (sectionName === 'gebruikers') {
    const activeTab = opts.gbTab || 'gebruikers';
    updateSubnav('gebruikers', activeTab);
    _currentSubItem = activeTab;
    if (typeof loadGebruikersSection === 'function') loadGebruikersSection();
    activateSectionSubtab('gebruikers', activeTab);
    return;
  }
  if (sectionName === 'baseline') {
    const activeTab = opts.baselineTab || 'baselines';
    updateSubnav('baseline', activeTab);
    _currentSubItem = activeTab;
    if (typeof loadBaselineSection === 'function') loadBaselineSection();
    activateSectionSubtab('baseline', activeTab);
    return;
  }
  if (sectionName === 'teams' || sectionName === 'sharepoint' || sectionName === 'identity' || sectionName === 'apps' || sectionName === 'domains' || sectionName === 'exchange' || sectionName === 'intune' || sectionName === 'backup' || sectionName === 'alerts' || sectionName === 'compliance' || sectionName === 'hybrid') {
    const defaultTabs = {
      teams: 'teams',
      sharepoint: 'sharepoint-sites',
      identity: 'mfa',
      apps: 'registrations',
      domains: 'domains-list',
      exchange: 'mailboxen',
      intune: 'overzicht',
      backup: 'overzicht',
      alerts: 'auditlog',
      compliance: 'cis',
      hybrid: 'sync',
    };
    const activeTab = opts.liveTab || defaultTabs[sectionName];
    updateSubnav(sectionName, activeTab);
    _currentSubItem = activeTab;
    if (typeof loadLiveModuleSection === 'function') loadLiveModuleSection(sectionName, activeTab);
    return;
  }
  if (sectionName === 'ca') {
    const activeTab = opts.caTab || 'policies';
    updateSubnav('ca', activeTab);
    _currentSubItem = activeTab;
    if (typeof loadCaSection === 'function') loadCaSection();
    activateSectionSubtab('ca', activeTab);
    return;
  }
  if (sectionName === 'bevindingen') {
    updateSubnav('bevindingen', null);
    _currentSubItem = null;
    if (typeof loadBevindingenSection === 'function') loadBevindingenSection();
    return;
  }
  if (sectionName === 'tenantoverzicht') {
    // Dubbele toegangscontrole: alleen admin mag deze sectie laden
    if (_currentUserRole !== 'admin') {
      showToast('Onvoldoende rechten om het tenant overzicht te bekijken.', 'error');
      showSection('overview');
      return;
    }
    updateSubnav('tenantoverzicht', null);
    _currentSubItem = null;
    loadTenantHealthDashboard();
    return;
  }
  if (sectionName === 'klantenbeheer') {
    if (_currentUserRole !== 'admin') {
      showToast('Onvoldoende rechten.', 'error');
      showSection('overview');
      return;
    }
    updateSubnav('klantenbeheer', null);
    _currentSubItem = null;
    loadKlantenbeheer();
    return;
  }
  if (sectionName === 'goedkeuringen') {
    if (_currentUserRole !== 'admin') {
      showToast('Onvoldoende rechten.', 'error');
      showSection('overview');
      return;
    }
    updateSubnav('goedkeuringen', null);
    _currentSubItem = null;
    loadGoedkeuringen();
    return;
  }
  if (sectionName === 'kosten') {
    if (_currentUserRole !== 'admin') {
      showToast('Onvoldoende rechten.', 'error');
      showSection('overview');
      return;
    }
    updateSubnav('kosten', null);
    _currentSubItem = null;
    loadKostenSection();
    return;
  }
  if (sectionName === 'jobmonitor') {
    if (_currentUserRole !== 'admin') {
      showToast('Onvoldoende rechten.', 'error');
      showSection('overview');
      return;
    }
    updateSubnav('jobmonitor', null);
    _currentSubItem = null;
    loadJobMonitor();
    return;
  }
  if (sectionName === 'gbidHub') {
    updateSubnav('gbidHub', null);
    _currentSubItem = null;
    loadHubSection('gbid');
    return;
  }
  if (sectionName === 'securityHub') {
    updateSubnav('securityHub', null);
    _currentSubItem = null;
    loadHubSection('security');
    return;
  }
  if (sectionName === 'collabHub') {
    updateSubnav('collabHub', null);
    _currentSubItem = null;
    loadHubSection('collab');
    return;
  }
  if (sectionName === 'devicesHub') {
    updateSubnav('devicesHub', null);
    _currentSubItem = null;
    loadHubSection('devices');
    return;
  }
  if (sectionName === 'assessmentHub') {
    updateSubnav('assessmentHub', null);
    _currentSubItem = null;
    loadHubSection('assessment');
    return;
  }
  // Assessment sub-nav
  updateSubnav(sectionName, sectionName === 'assessment' ? 'assessment' : null);
  _currentSubItem = sectionName === 'assessment' ? 'assessment' : null;
}
window.showSection = showSection;

function switchSettingsTab(tabName) {
  const tabs = document.querySelectorAll('.settings-tab');
  const panes = document.querySelectorAll('.settings-pane');
  tabs.forEach((btn) => btn.classList.toggle('active', btn.dataset.tab === tabName));
  panes.forEach((pane) => pane.classList.toggle('active', pane.dataset.tab === tabName));
  // Sync subnav actief item
  if (_currentSection === 'settings') setActiveSubnavItem(tabName);
  // Laad tabspecifieke data
  if (tabName === 'roles') loadRolesTab();
  if (tabName === 'tenant') loadIntegratieStatusGrid();
}
window.switchSettingsTab = switchSettingsTab;

function formatDate(value) {
  if (!value) return '-';
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return value;
  return d.toLocaleString();
}

function statusBadge(status) {
  const color = {
    queued: '#667085',
    running: '#0ea5e9',
    completed: '#16a34a',
    failed: '#dc2626',
    partial: '#f59e0b',
  }[status] || '#667085';
  return `<span style="display:inline-block;padding:2px 8px;border-radius:999px;background:${color};color:#fff;font-size:12px;font-weight:600;">${status || '-'}</span>`;
}

function formatPhaseList(phases) {
  if (!Array.isArray(phases) || !phases.length) return 'Alle fases';
  return phases.map((p) => p.replace('phase', 'F')).join(', ');
}

function escapeHtml(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * Koppelt event-listeners aan knoppen met data-action attributen.
 * Gebruik data-action="actionName", data-id="...", data-extra="..." op buttons.
 * Voorkomt inline onclick handlers (XSS-vector).
 */
function bindActions(root) {
  root.querySelectorAll('[data-action]').forEach((btn) => {
    if (btn._actionBound) return;
    btn._actionBound = true;
    btn.addEventListener('click', (e) => {
      e.preventDefault();
      const action = btn.dataset.action;
      const id = btn.dataset.id || '';
      const extra = btn.dataset.extra || '';
      switch (action) {
        case 'selectTenant': selectTenantFromManagement(id); break;
        case 'selectTenantPill': selectTenantFromPill(id); break;
        case 'deleteTenant': deleteTenantFromManagement(id); break;
        case 'viewRun': viewRunDetails(id); break;
        case 'openUrl': if (id) window.open(id, '_blank'); break;
        case 'archiveRun': archiveReportRun(id); break;
        case 'restoreRun': restoreReportRun(id); break;
        case 'deleteRun': deleteRunPermanently(id); break;
        case 'stopRun': stopRunById(id); break;
        case 'setStatus': setActionStatus(id, extra); break;
        default: console.warn('Onbekende actie:', action);
      }
    });
  });
}

function toQuery(params) {
  const q = new URLSearchParams();
  Object.entries(params || {}).forEach(([k, v]) => {
    if (v != null && v !== '') q.set(k, String(v));
  });
  const s = q.toString();
  return s ? `?${s}` : '';
}

function deltaText(value, reverseGood = false) {
  const n = Number(value || 0);
  const sign = n > 0 ? '+' : '';
  const className = n === 0 ? 'diff-neutral' : ((n > 0) !== reverseGood ? 'diff-bad' : 'diff-good');
  return `<span class="${className}">${sign}${n}</span>`;
}

function actionStatusBadge(status) {
  const map = {
    open: '#334155',
    in_progress: '#2563eb',
    done: '#16a34a',
    accepted: '#f59e0b',
  };
  const color = map[status] || '#64748b';
  return `<span style="display:inline-block;padding:2px 8px;border-radius:999px;background:${color};color:#fff;font-size:11px;font-weight:700;">${escapeHtml(status || '-')}</span>`;
}

function severityBadge(severity) {
  const map = {
    critical: '#dc2626',
    warning: '#f59e0b',
    info: '#2563eb',
  };
  const color = map[severity] || '#64748b';
  return `<span style="display:inline-block;padding:2px 8px;border-radius:999px;background:${color};color:#fff;font-size:11px;font-weight:700;">${escapeHtml(severity || '-')}</span>`;
}

async function loadTenants() {
  const data = await apiFetch('/api/tenants');
  const tenants = data.items || [];
  allTenants = tenants;
  const select = document.getElementById('tenantSelect');
  select.innerHTML = '';

  tenants.forEach((tenant) => {
    const opt = document.createElement('option');
    opt.value = tenant.id;
    const status = tenant.status ? ` [${tenant.status}]` : '';
    opt.textContent = `${tenant.customer_name} / ${tenant.tenant_name}${status}`;
    select.appendChild(opt);
  });

  const stored = localStorage.getItem('local_m365_current_tenant');
  if (stored && tenants.some((t) => t.id === stored)) {
    currentTenantId = stored;
  } else if (tenants[0]) {
    currentTenantId = tenants[0].id;
  } else {
    currentTenantId = null;
  }

  if (currentTenantId) {
    select.value = currentTenantId;
    localStorage.setItem('local_m365_current_tenant', currentTenantId);
  }
  window.currentTenantId = currentTenantId;

  updateTenantPill(tenants, currentTenantId);
  updateHeroVisibility();
  updateWorkspaceHeader(_currentSection);
  renderContextRail(_currentSection);
  renderNavSignals();

  if (document.getElementById('tenantManagementTableBody')) {
    renderTenantManagementTable(tenants);
  }

  return tenants;
}

function renderTenantManagementTable(tenants) {
  const tbody = document.getElementById('tenantManagementTableBody');
  if (!tbody) return;
  const items = Array.isArray(tenants) ? tenants : [];
  if (!items.length) {
    tbody.innerHTML = '<tr><td colspan="5" class="empty-state">Geen actieve tenants.</td></tr>';
    return;
  }
  tbody.innerHTML = items.map((t) => `
    <tr>
      <td>${escapeHtml(t.customer_name || '-')}</td>
      <td>${escapeHtml(t.tenant_name || '-')}</td>
      <td>${escapeHtml(t.tenant_guid || '-')}</td>
      <td>${escapeHtml(t.status || '-')}</td>
      <td>
        <div class="results-row-actions">
          <button class="btn btn-secondary btn-sm" data-action="selectTenant" data-id="${escapeHtml(t.id)}">Selecteer</button>
          <button class="btn btn-warning btn-sm" data-action="deleteTenant" data-id="${escapeHtml(t.id)}">Verwijder</button>
        </div>
      </td>
    </tr>
  `).join('');
  bindActions(tbody);
}
window.renderTenantManagementTable = renderTenantManagementTable;

async function selectTenantFromManagement(tenantId) {
  currentTenantId = tenantId;
  window.currentTenantId = currentTenantId;
  localStorage.setItem('local_m365_current_tenant', tenantId);
  const select = document.getElementById('tenantSelect');
  if (select) select.value = tenantId;
  updateTenantPill(allTenants, tenantId);
  updateHeroVisibility();
  updateWorkspaceHeader(_currentSection);
  await populateSettings();
  await refreshTenantData();
  renderNavSignals();
}
window.selectTenantFromManagement = selectTenantFromManagement;

async function deleteTenantFromManagement(tenantId) {
  const tenants = await apiFetch('/api/tenants');
  const target = (tenants.items || []).find((t) => t.id === tenantId);
  const label = target ? `${target.customer_name} / ${target.tenant_name}` : tenantId;
  const confirmed = window.confirm(`Tenant verwijderen?\n\n${label}\n\nOK = soft delete (inactief).\nCancel = annuleren.`);
  if (!confirmed) return;
  const hard = window.confirm('Ook alle geschiedenis/rapporten permanent verwijderen?\nOK = hard delete\nCancel = soft delete');
  const mode = hard ? 'hard' : 'soft';
  try {
    await apiFetch(`/api/tenants/${tenantId}?mode=${mode}`, { method: 'DELETE' });
  } catch (e) {
    await apiFetch(`/api/tenants/${tenantId}/delete?mode=${mode}`, {
      method: 'POST',
      body: JSON.stringify({ mode }),
    });
  }

  if (currentTenantId === tenantId) {
    currentTenantId = null;
    localStorage.removeItem('local_m365_current_tenant');
  }
  await loadTenants();
  await populateSettings();
  await refreshTenantData();
  alert(`Tenant verwijderd (${mode}).`);
}
window.deleteTenantFromManagement = deleteTenantFromManagement;

async function loadOverview() {
  bindOverviewActions();
  const statTargets = {
    userCount: ['userCount', 'heroUserCount'],
    mfaStatus: ['mfaStatus', 'heroMfaStatus'],
    secureScore: ['secureScore', 'heroSecureScore'],
    caPolicies: ['caPolicies', 'heroCaPolicies'],
  };
  Object.values(statTargets).flat().forEach((id) => {
    const el = document.getElementById(id);
    if (el) el.textContent = '-';
  });
  const list = document.getElementById('recentAssessmentsList');
  if (!currentTenantId) {
    list.innerHTML = '<p class="empty-state">Geen tenant geselecteerd</p>';
    resetOverviewServiceCards();
    renderContextRail('overview');
    renderNavSignals();
    return;
  }

  await loadOverviewServiceCards();

  const stats = await apiFetch(`/api/tenants/${currentTenantId}/overview`);
  if (!stats.hasData) {
    list.innerHTML = '<p class="empty-state">Nog geen assessments uitgevoerd</p>';
    renderContextRail('overview');
    renderNavSignals();
    return;
  }

  const values = {
    userCount: stats.totalUsers || '-',
    mfaStatus: stats.mfaCoverage != null ? `${Math.round(stats.mfaCoverage)}%` : '-',
    secureScore: stats.secureScorePercentage != null ? `${Math.round(stats.secureScorePercentage)}%` : (stats.scoreOverall ?? '-'),
    caPolicies: stats.caPolicies ?? '-',
  };
  Object.entries(statTargets).forEach(([key, targetIds]) => {
    targetIds.forEach((id) => {
      const el = document.getElementById(id);
      if (el) el.textContent = values[key];
    });
  });

  const runs = await apiFetch(`/api/tenants/${currentTenantId}/runs`);
  const items = (runs.items || []).slice(0, 5);
  if (!items.length) {
    list.innerHTML = '<p class="empty-state">Nog geen assessments uitgevoerd</p>';
    return;
  }
  const isAdmin = typeof _currentUserRole !== 'undefined' && _currentUserRole === 'admin';
  list.innerHTML = items.map((r) => `
    <div class="assessment-item">
      <div><strong>${formatDate(r.completed_at || r.started_at)}</strong> - ${statusBadge(r.status)}</div>
      <div style="margin-top:6px;font-size:0.9em;color:#666;">${(r.phases || []).join(', ') || 'alle fases'}</div>
      <div style="margin-top:8px;display:flex;gap:8px;flex-wrap:wrap;">
        <button class="btn btn-secondary btn-sm" data-action="viewRun" data-id="${escapeHtml(r.id)}">Details</button>
        <button class="btn btn-secondary btn-sm" data-action="showSection" data-id="assessment">Assessment</button>
        ${isAdmin && r.status === 'running' ? `<button class="btn btn-warning btn-sm" data-action="stopRun" data-id="${escapeHtml(r.id)}">&#9646; Stop</button>` : ''}
        ${isAdmin ? `<button class="btn btn-danger btn-sm" data-action="deleteRun" data-id="${escapeHtml(r.id)}">&#128465; Verwijder</button>` : ''}
      </div>
    </div>`).join('');
  bindActions(list);
  renderContextRail('overview');
  renderNavSignals();
}

async function loadHistorySection() {
  const tbody = document.getElementById('historyTableBody');
  if (!tbody) return;
  if (!currentTenantId) {
    tbody.innerHTML = '<tr><td colspan="5" class="empty-state">Geen tenant geselecteerd</td></tr>';
    return;
  }
  const data = await apiFetch(`/api/tenants/${currentTenantId}/runs`);
  const runs = data.items || [];
  if (!runs.length) {
    tbody.innerHTML = '<tr><td colspan="5" class="empty-state">Geen geschiedenis beschikbaar</td></tr>';
    return;
  }
  tbody.innerHTML = runs.map((r) => `
    <tr>
      <td>${formatDate(r.completed_at || r.started_at)}</td>
      <td>${r.tenant_name || '-'}</td>
      <td>${(r.phases || []).join(', ') || '-'}</td>
      <td>${statusBadge(r.status)}</td>
      <td>
        <button class="btn btn-secondary btn-sm" data-action="viewRun" data-id="${escapeHtml(r.id)}">Details</button>
        ${r.report_path ? `<button class="btn btn-secondary btn-sm" data-action="openUrl" data-id="${escapeHtml(r.report_path)}">Rapport</button>` : ''}
        ${r.is_archived ? `<button class="btn btn-secondary btn-sm" data-action="restoreRun" data-id="${escapeHtml(r.id)}">Herstel</button>` : `<button class="btn btn-secondary btn-sm" data-action="archiveRun" data-id="${escapeHtml(r.id)}">Archiveer</button>`}
        <button class="btn btn-warning btn-sm" data-action="deleteRun" data-id="${escapeHtml(r.id)}">Verwijder</button>
      </td>
    </tr>`).join('');
  bindActions(tbody);
}
window.loadHistorySection = loadHistorySection;

async function loadResultsSection() {
  const emptyEl   = document.getElementById('resultsViewerEmpty');
  const contentEl = document.getElementById('resultsViewerContent');

  function showEmpty(msg) {
    if (emptyEl)   { emptyEl.style.display = ''; emptyEl.textContent = msg; }
    if (contentEl) contentEl.style.display = 'none';
  }

  if (!currentTenantId) { showEmpty('Geen tenant geselecteerd.'); renderContextRail('results'); renderNavSignals(); return; }

  const runs = await apiFetch(`/api/tenants/${currentTenantId}/runs`);
  const allRuns = runs.items || [];
  const latest  = allRuns.find((r) => !!(r.snapshot_path || r.report_path));

  if (!latest) { showEmpty('Start een assessment om resultaten te zien.'); renderContextRail('results'); renderNavSignals(); return; }

  // Toon content
  if (emptyEl)   emptyEl.style.display = 'none';
  if (contentEl) contentEl.style.display = '';

  const reportRuns    = allRuns.filter((r) => !!(r.snapshot_path || r.report_path)).slice(0, 8);
  const completedRuns = allRuns.filter((r) => r.status === 'completed').length;
  const failedRuns    = allRuns.filter((r) => r.status === 'failed').length;

  // ── KPI ──
  const setText = (id, v) => { const el = document.getElementById(id); if (el) el.textContent = v; };
  const setHtml = (id, v) => { const el = document.getElementById(id); if (el) el.innerHTML   = v; };

  setText('kpiScore',    latest.score_overall ?? '—');
  setText('kpiCritical', latest.critical_count ?? 0);
  setText('kpiWarning',  latest.warning_count  ?? 0);
  setText('kpiInfo',     latest.info_count     ?? 0);

  // ── Meta ──
  setText('metaRunId',       latest.id);
  setHtml('metaStatus',      statusBadge(latest.status));
  setText('metaStarted',     formatDate(latest.started_at));
  setText('metaCompleted',   formatDate(latest.completed_at));
  setText('metaRunMode',     latest.run_mode || '—');
  setText('metaPhases',      formatPhaseList(latest.phases));
  setText('metaTenantName',  latest.tenant_name  || '—');
  setText('metaCustomer',    latest.customer_name || '—');
  setText('metaTotalRuns',   allRuns.length);
  setText('metaCompletedRuns', completedRuns);
  setText('metaFailedRuns',  failedRuns);
  setText('metaReportRuns',  reportRuns.length);

  // ── Runs count badges ──
  setText('resultsRunCount',  reportRuns.length);
  setText('resultsRunsCount', `${reportRuns.length} item(s)`);
  setText('resultsTabCount',  reportRuns.length);

  // ── Runs tabel ──
  // ── Rapportviewer ──
  if (typeof initResultsViewer === 'function') {
    initResultsViewer('resultsViewerContainer', latest.report_path || '', {
      tenantId: currentTenantId,
      latestRun: latest,
      reportRuns,
      summary: {
        totalRuns: allRuns.length,
        completedRuns,
        failedRuns,
        reportRuns: reportRuns.length,
      },
      latestReportUrl: latest.report_path || '',
      latestCsvUrl: latest.report_path ? latest.report_path.replace(/\/[^/]+\.html$/, '') + '/_Assessment-Summary.csv' : '',
      formatDate,
      formatPhaseList,
      statusBadge,
      escapeHtml,
    });
  }
  renderContextRail('results');
  renderNavSignals();
}
window.loadResultsSection = loadResultsSection;

async function loadRunDiffPanel() {
  const el = document.getElementById('runDiffContainer');
  if (!el || !currentTenantId) return;
  try {
    const diff = await apiFetch(`/api/tenants/${currentTenantId}/runs/diff`);
    if (!diff.hasDiff) {
      el.innerHTML = '<div class="empty-state">Nog onvoldoende runs voor vergelijking (minimaal 2 rapport-runs nodig).</div>';
      return;
    }
    el.innerHTML = `
      <div class="diff-grid">
        <div class="diff-card"><span>Trend</span><strong>${escapeHtml(diff.trend)}</strong></div>
        <div class="diff-card"><span>Score Δ</span><strong>${deltaText(diff.delta.score_overall, false)}</strong></div>
        <div class="diff-card"><span>Critical Δ</span><strong>${deltaText(diff.delta.critical_count, true)}</strong></div>
        <div class="diff-card"><span>Warning Δ</span><strong>${deltaText(diff.delta.warning_count, true)}</strong></div>
        <div class="diff-card"><span>Info Δ</span><strong>${deltaText(diff.delta.info_count, true)}</strong></div>
      </div>
      <div class="diff-footnote">
        Van ${formatDate(diff.from.completed_at)} naar ${formatDate(diff.to.completed_at)}
      </div>
    `;
  } catch (e) {
    el.innerHTML = `<div class="empty-state">Vergelijking mislukt: ${escapeHtml(e.message)}</div>`;
  }
}
window.loadRunDiffPanel = loadRunDiffPanel;

function getReportsFilters() {
  return {
    tenant_id: currentTenantId || '',
    status: document.getElementById('reportsFilterStatus')?.value || '',
    archived: document.getElementById('reportsFilterArchived')?.value || 'exclude',
    from: document.getElementById('reportsFilterFrom')?.value || '',
    to: document.getElementById('reportsFilterTo')?.value || '',
    q: document.getElementById('reportsFilterSearch')?.value || '',
    limit: 300,
  };
}

async function loadReportsManagementPanel() {
  const tbody = document.getElementById('reportsManagementTableBody');
  if (!tbody || !currentTenantId) return;
  try {
    const filters = getReportsFilters();
    const data = await apiFetch(`/api/reports${toQuery(filters)}`);
    const items = data.items || [];
    if (!items.length) {
      tbody.innerHTML = '<tr><td colspan="8" class="empty-state">Geen rapporten gevonden met huidige filters.</td></tr>';
      return;
    }
    tbody.innerHTML = items.map((r) => `
      <tr>
        <td>${formatDate(r.completed_at || r.started_at)}</td>
        <td>${escapeHtml(r.tenant_name || '-')}</td>
        <td>${statusBadge(r.status)}</td>
        <td>${r.is_archived ? '<span class="diff-neutral">Gearchiveerd</span>' : '<span class="diff-good">Actief</span>'}</td>
        <td>${escapeHtml(r.score_overall ?? '-')}</td>
        <td>${escapeHtml(r.critical_count ?? 0)} / ${escapeHtml(r.warning_count ?? 0)} / ${escapeHtml(r.info_count ?? 0)}</td>
        <td title="${escapeHtml(r.report_filename || '')}">${escapeHtml(r.report_filename || '-')}</td>
        <td>
          <div class="results-row-actions">
            <button class="btn btn-secondary btn-sm" data-action="viewRun" data-id="${escapeHtml(r.id)}">Details</button>
            <button class="btn btn-secondary btn-sm" data-action="openUrl" data-id="${escapeHtml(r.report_path)}">Open</button>
            ${r.is_archived
              ? `<button class="btn btn-secondary btn-sm" data-action="restoreRun" data-id="${escapeHtml(r.id)}">Herstel</button>`
              : `<button class="btn btn-secondary btn-sm" data-action="archiveRun" data-id="${escapeHtml(r.id)}">Archiveer</button>`}
            <button class="btn btn-warning btn-sm" data-action="deleteRun" data-id="${escapeHtml(r.id)}">Verwijder</button>
          </div>
        </td>
      </tr>
    `).join('');
    bindActions(tbody);
  } catch (e) {
    tbody.innerHTML = `<tr><td colspan="8" class="empty-state">Rapportbeheer laden mislukt: ${escapeHtml(e.message)}</td></tr>`;
  }
}
window.loadReportsManagementPanel = loadReportsManagementPanel;

function clearReportsFilters() {
  const ids = ['reportsFilterSearch', 'reportsFilterStatus', 'reportsFilterArchived', 'reportsFilterFrom', 'reportsFilterTo'];
  ids.forEach((id) => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
  const archivedEl = document.getElementById('reportsFilterArchived');
  if (archivedEl) archivedEl.value = 'exclude';
  loadReportsManagementPanel();
}
window.clearReportsFilters = clearReportsFilters;

function exportReportsCsv() {
  if (!currentTenantId) return;
  const url = `/api/reports/export.csv${toQuery(getReportsFilters())}`;
  window.open(url, '_blank');
}
window.exportReportsCsv = exportReportsCsv;

async function archiveReportRun(runId) {
  const reason = window.prompt('Reden voor archiveren (optioneel):', 'Handmatig gearchiveerd');
  await apiFetch(`/api/reports/${runId}/archive`, {
    method: 'POST',
    body: JSON.stringify({ reason: reason || '' }),
  });
  await Promise.allSettled([loadReportsManagementPanel(), loadResultsSection()]);
}
window.archiveReportRun = archiveReportRun;

async function restoreReportRun(runId) {
  await apiFetch(`/api/reports/${runId}/restore`, { method: 'POST', body: '{}' });
  await Promise.allSettled([loadReportsManagementPanel(), loadResultsSection()]);
}
window.restoreReportRun = restoreReportRun;

async function deleteRunPermanently(runId) {
  const confirmed = window.confirm(`Run ${runId} permanent verwijderen?\n\nDit verwijdert ook rapportbestanden en run-logs.`);
  if (!confirmed) return;
  try {
    await apiFetch(`/api/runs/${runId}`, { method: 'DELETE' });
  } catch (e) {
    await apiFetch(`/api/runs/${runId}/delete`, { method: 'POST', body: '{}' });
  }
  await Promise.allSettled([loadOverview(), loadResultsSection()]);
  alert('Run permanent verwijderd.');
}
window.deleteRunPermanently = deleteRunPermanently;

async function stopRunById(runId) {
  const confirmed = window.confirm(`Assessment run stoppen?\n\nDe run wordt gestopt en gemarkeerd als geannuleerd.`);
  if (!confirmed) return;
  try {
    await apiFetch(`/api/runs/${runId}/stop`, { method: 'POST', body: '{}' });
    if (typeof showToast === 'function') showToast('Run gestopt.', 'success');
  } catch (e) {
    if (typeof showToast === 'function') showToast('Kon run niet stoppen: ' + (e.message || e), 'error');
  }
  await loadOverview();
}
window.stopRunById = stopRunById;

async function applyRetentionPolicy() {
  if (!currentTenantId) {
    alert('Selecteer eerst een tenant.');
    return;
  }
  const keepLatest = parseInt(document.getElementById('retentionKeepLatestInput')?.value || '10', 10);
  const keepDays = parseInt(document.getElementById('retentionKeepDaysInput')?.value || '90', 10);
  const scope = document.getElementById('retentionScopeSelect')?.value || 'tenant';

  const result = await apiFetch('/api/reports/retention/apply', {
    method: 'POST',
    body: JSON.stringify({
      tenant_id: scope === 'all' ? null : currentTenantId,
      keep_latest: Number.isFinite(keepLatest) ? keepLatest : 10,
      keep_days: Number.isFinite(keepDays) ? keepDays : 90,
    }),
  });
  alert(`Retentie toegepast. Gescand: ${result.scanned}, gearchiveerd: ${result.archived}.`);
  await loadReportsManagementPanel();
}
window.applyRetentionPolicy = applyRetentionPolicy;

async function loadActionsPanel() {
  const tbody = document.getElementById('actionsTableBody');
  if (!tbody || !currentTenantId) return;
  const status = document.getElementById('actionsStatusFilter')?.value || 'all';
  try {
    const data = await apiFetch(`/api/tenants/${currentTenantId}/actions${toQuery({ status })}`);
    const items = data.items || [];
    if (!items.length) {
      tbody.innerHTML = '<tr><td colspan="7" class="empty-state">Nog geen acties voor deze tenant.</td></tr>';
      return;
    }
    tbody.innerHTML = items.map((a) => `
      <tr>
        <td><strong>${escapeHtml(a.finding_key)}</strong><br><span>${escapeHtml(a.title)}</span></td>
        <td>${severityBadge(a.severity)}</td>
        <td>${escapeHtml(a.owner || '-')}</td>
        <td>${actionStatusBadge(a.status)}</td>
        <td>${escapeHtml(a.due_date || '-')}</td>
        <td>${escapeHtml(a.run_id ? a.run_id.slice(0, 8) : '-')}</td>
        <td>
          <div class="results-row-actions">
            <button class="btn btn-secondary btn-sm" data-action="setStatus" data-id="${escapeHtml(a.id)}" data-extra="open">Open</button>
            <button class="btn btn-secondary btn-sm" data-action="setStatus" data-id="${escapeHtml(a.id)}" data-extra="in_progress">In progress</button>
            <button class="btn btn-secondary btn-sm" data-action="setStatus" data-id="${escapeHtml(a.id)}" data-extra="done">Done</button>
          </div>
        </td>
      </tr>
    `).join('');
    bindActions(tbody);
  } catch (e) {
    tbody.innerHTML = `<tr><td colspan="7" class="empty-state">Acties laden mislukt: ${escapeHtml(e.message)}</td></tr>`;
  }
}
window.loadActionsPanel = loadActionsPanel;

async function createFindingAction() {
  if (!currentTenantId) {
    alert('Selecteer eerst een tenant.');
    return;
  }
  const payload = {
    tenant_id: currentTenantId,
    title: document.getElementById('actionTitleInput')?.value || '',
    finding_key: document.getElementById('actionKeyInput')?.value || '',
    severity: document.getElementById('actionSeverityInput')?.value || 'warning',
    owner: document.getElementById('actionOwnerInput')?.value || '',
    due_date: document.getElementById('actionDueDateInput')?.value || '',
    status: 'open',
  };
  if (!payload.title.trim()) {
    alert('Vul een titel in.');
    return;
  }
  await apiFetch('/api/actions', { method: 'POST', body: JSON.stringify(payload) });
  ['actionTitleInput', 'actionKeyInput', 'actionOwnerInput', 'actionDueDateInput'].forEach((id) => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
  await loadActionsPanel();
}
window.createFindingAction = createFindingAction;

async function setActionStatus(actionId, status) {
  await apiFetch(`/api/actions/${actionId}`, { method: 'PATCH', body: JSON.stringify({ status }) });
  await loadActionsPanel();
}
window.setActionStatus = setActionStatus;

async function populateSettings() {
  try {
    localConfig = await apiFetch('/api/config');
    const modeEl   = document.getElementById('runModeSelect');
    const scriptEl = document.getElementById('scriptPathInput');
    const tenantAuthEl = document.getElementById('authTenantIdInput');
    const clientIdAuthEl = document.getElementById('authClientIdInput');
    const certThumbEl = document.getElementById('authCertThumbInput');
    if (modeEl)       modeEl.value       = localConfig.default_run_mode       || 'demo';
    if (scriptEl)     scriptEl.value     = localConfig.script_path             || '';
    if (tenantAuthEl) tenantAuthEl.value = localConfig.auth_tenant_id          || '';
    if (clientIdAuthEl) clientIdAuthEl.value = localConfig.auth_client_id      || '';
    if (certThumbEl)  certThumbEl.value  = localConfig.auth_cert_thumbprint    || '';
    // Client secret wordt NIET teruggeladen om veiligheidsredenen
  } catch (e) {
    console.warn('Config laden mislukt', e);
  }

  // Keep legacy settings harmless in local mode
  const clientId = document.getElementById('clientIdInput');
  const tenantId = document.getElementById('tenantIdInput');
  if (clientId) clientId.value = 'Lokale modus';
  if (tenantId) tenantId.value = currentTenantId || '-';

  try {
    const integrationCfg = JSON.parse(localStorage.getItem('m365LocalIntegrations') || '{}');
    const webhookUrlEl = document.getElementById('integrationWebhookUrlInput');
    const webhookEnabledEl = document.getElementById('integrationWebhookEnabledInput');
    if (webhookUrlEl) webhookUrlEl.value = integrationCfg.webhook_url || '';
    if (webhookEnabledEl) webhookEnabledEl.value = integrationCfg.webhook_enabled || 'off';
  } catch (_) {}

  try {
    const tenantsData = await apiFetch('/api/tenants');
    renderTenantManagementTable(tenantsData.items || []);
  } catch (_) {}

  if (!currentTenantId) return;
  try {
    const tenant = await apiFetch(`/api/tenants/${currentTenantId}`);
    const statusEl = document.getElementById('tenantStatusInput');
    const riskEl = document.getElementById('tenantRiskProfileInput');
    const ownerPrimaryEl = document.getElementById('tenantOwnerPrimaryInput');
    const ownerBackupEl = document.getElementById('tenantOwnerBackupInput');
    const tagsEl = document.getElementById('tenantTagsInput');
    if (statusEl) statusEl.value = tenant.status || 'active';
    if (riskEl) riskEl.value = tenant.risk_profile || 'standard';
    if (ownerPrimaryEl) ownerPrimaryEl.value = tenant.owner_primary || '';
    if (ownerBackupEl) ownerBackupEl.value = tenant.owner_backup || '';
    if (tagsEl) tagsEl.value = tenant.tags_csv || '';
  } catch (e) {
    console.warn('Tenant governance laden mislukt', e);
  }
}
window.populateSettings = populateSettings;

async function saveLocalConfig() {
  const clientSecret = document.getElementById('authClientSecretInput')?.value || '';
  const payload = {
    default_run_mode:       document.getElementById('runModeSelect')?.value       || 'demo',
    script_path:            document.getElementById('scriptPathInput')?.value      || '',
    auth_tenant_id:         document.getElementById('authTenantIdInput')?.value    || '',
    auth_client_id:         document.getElementById('authClientIdInput')?.value    || '',
    auth_cert_thumbprint:   document.getElementById('authCertThumbInput')?.value   || '',
    // Secret alleen opslaan als ingevuld, anders bestaande waarde bewaren
    ...(clientSecret ? { auth_client_secret: clientSecret } : {}),
  };
  localConfig = await apiFetch('/api/config', { method: 'POST', body: JSON.stringify(payload) });
  // Wis het secret veld na opslaan
  const secretEl = document.getElementById('authClientSecretInput');
  if (secretEl) secretEl.value = '';
  alert('Config opgeslagen.');
}

async function createTenantFromForm() {
  const payload = {
    customer_name: document.getElementById('newCustomerNameInput')?.value || '',
    tenant_name: document.getElementById('newTenantNameInput')?.value || '',
    tenant_guid: document.getElementById('newTenantGuidInput')?.value || '',
    status: document.getElementById('tenantStatusInput')?.value || 'active',
    risk_profile: document.getElementById('tenantRiskProfileInput')?.value || 'standard',
    owner_primary: document.getElementById('tenantOwnerPrimaryInput')?.value || '',
    owner_backup: document.getElementById('tenantOwnerBackupInput')?.value || '',
    tags_csv: document.getElementById('tenantTagsInput')?.value || '',
  };
  if (!payload.customer_name && !payload.tenant_name) {
    alert('Vul minimaal klantnaam of tenant naam in.');
    return;
  }
  await apiFetch('/api/tenants', { method: 'POST', body: JSON.stringify(payload) });
  await loadTenants();
  await refreshTenantData();
  ['newCustomerNameInput', 'newTenantNameInput', 'newTenantGuidInput'].forEach((id) => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
  alert('Tenant aangemaakt.');
}

async function saveTenantGovernance() {
  if (!currentTenantId) {
    alert('Geen tenant geselecteerd.');
    return;
  }
  const payload = {
    status: document.getElementById('tenantStatusInput')?.value || 'active',
    risk_profile: document.getElementById('tenantRiskProfileInput')?.value || 'standard',
    owner_primary: document.getElementById('tenantOwnerPrimaryInput')?.value || '',
    owner_backup: document.getElementById('tenantOwnerBackupInput')?.value || '',
    tags_csv: document.getElementById('tenantTagsInput')?.value || '',
  };
  await apiFetch(`/api/tenants/${currentTenantId}`, { method: 'PATCH', body: JSON.stringify(payload) });
  await loadTenants();
  await refreshTenantData();
  alert('Tenant governance opgeslagen.');
}

function saveIntegrationSettings() {
  const payload = {
    webhook_url: document.getElementById('integrationWebhookUrlInput')?.value || '',
    webhook_enabled: document.getElementById('integrationWebhookEnabledInput')?.value || 'off',
  };
  localStorage.setItem('m365LocalIntegrations', JSON.stringify(payload));
  alert('Integratie-instellingen opgeslagen.');
}

async function applySettingsRetentionPolicy() {
  if (!currentTenantId) {
    alert('Selecteer eerst een tenant.');
    return;
  }
  const keepLatest = parseInt(document.getElementById('settingsRetentionKeepLatestInput')?.value || '10', 10);
  const keepDays = parseInt(document.getElementById('settingsRetentionKeepDaysInput')?.value || '90', 10);
  const scope = document.getElementById('settingsRetentionScopeSelect')?.value || 'tenant';
  const result = await apiFetch('/api/reports/retention/apply', {
    method: 'POST',
    body: JSON.stringify({
      tenant_id: scope === 'all' ? null : currentTenantId,
      keep_latest: Number.isFinite(keepLatest) ? keepLatest : 10,
      keep_days: Number.isFinite(keepDays) ? keepDays : 90,
    }),
  });
  alert(`Retentie toegepast. Gescand: ${result.scanned}, gearchiveerd: ${result.archived}.`);
  await refreshTenantData();
}

async function refreshTenantData() {
  await Promise.allSettled([loadOverview(), loadResultsSection()]);
  if (document.getElementById('assessmentSection')?.classList.contains('active') && typeof loadAssessmentExperience === 'function') {
    await loadAssessmentExperience({ force: true });
  }
  if (['teams', 'sharepoint', 'identity', 'apps', 'domains', 'exchange', 'intune', 'backup', 'alerts'].includes(_currentSection) && typeof loadLiveModuleSection === 'function') {
    await loadLiveModuleSection(_currentSection, _currentSubItem || null);
  }
  updateWorkspaceHeader(_currentSection);
  renderContextRail(_currentSection);
  renderNavSignals();
}
window.refreshTenantData = refreshTenantData;

async function viewRunDetails(runId) {
  const [run, logs] = await Promise.all([
    apiFetch(`/api/runs/${runId}`),
    apiFetch(`/api/runs/${runId}/logs`),
  ]);
  const logText = (logs.lines || []).slice(-25).join('\n');
  let message = `Run: ${run.id}\nStatus: ${run.status}\nGestart: ${formatDate(run.started_at)}\nVoltooid: ${formatDate(run.completed_at)}\nMode: ${run.run_mode}\nFases: ${(run.phases || []).join(', ')}\n`;
  if (run.report_path) message += `Rapport: ${window.location.origin}${run.report_path}\n`;
  if (run.error_message) message += `Fout: ${run.error_message}\n`;
  message += `\nLaatste logregels:\n${logText || '(geen logs)'}`;
  alert(message);
}
window.viewRunDetails = viewRunDetails;

function setupNavigation() {
  document.querySelectorAll('.portal-nav-link[data-section]:not(.nav-dropdown-toggle), .nav-dropdown-link[data-section]').forEach((item) => {
    item.addEventListener('click', (e) => {
      e.preventDefault();
      const opts = getSectionOptionsFromDataset(item.dataset);
      showSection(item.dataset.section, opts);
    });
  });
  document.querySelectorAll('.nav-dropdown-toggle').forEach((toggle) => {
    toggle.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      const dropdown = toggle.closest('.nav-dropdown');
      const shouldOpen = !dropdown.classList.contains('open') || !dropdown.classList.contains('has-active');
      document.querySelectorAll('.nav-dropdown.open').forEach((item) => {
        if (item !== dropdown) {
          setDropdownOpen(item, false);
        }
      });
      if (toggle.dataset.section) {
        showSection(toggle.dataset.section, getSectionOptionsFromDataset(toggle.dataset));
      }
      setDropdownOpen(dropdown, shouldOpen);
    });
  });
  document.querySelectorAll('.nav-dropdown').forEach((dropdown) => {
    dropdown.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        setDropdownOpen(dropdown, false);
        dropdown.querySelector('.nav-dropdown-toggle')?.focus();
      }
    });
  });
}

/* ── Header helpers ── */

function getInitials(name) {
  if (!name || name === 'Lokaal') return 'LK';
  const parts = name.trim().split(/\s+/);
  return (parts.length >= 2 ? parts[0][0] + parts[parts.length - 1][0] : name.substring(0, 2)).toUpperCase();
}

function updateTenantPill(tenants, selectedId) {
  const nameEl = document.getElementById('tenantPillName');
  const dropdown = document.getElementById('tenantPillDropdown');
  if (!nameEl || !dropdown) return;

  const selected = tenants.find((t) => t.id === selectedId);
  nameEl.textContent = selected ? (selected.customer_name || selected.tenant_name) : 'Geen tenant';

  dropdown.innerHTML = tenants.length
    ? tenants.map((t) => `
        <button type="button" class="tenant-dd-item${t.id === selectedId ? ' active' : ''}"
                data-action="selectTenantPill" data-id="${escapeHtml(t.id)}">
          ${escapeHtml(t.customer_name || t.tenant_name)}
        </button>`).join('')
    : '<div class="tenant-dd-empty">Geen tenants</div>';

  bindActions(dropdown);
}

function updateHeroVisibility() {
  const hero = document.getElementById('portalHero');
  if (hero) hero.style.display = currentTenantId ? 'none' : '';
}

async function selectTenantFromPill(tenantId) {
  const dropdown = document.getElementById('tenantPillDropdown');
  if (dropdown) dropdown.style.display = 'none';
  const pill = document.getElementById('tenantPill');
  if (pill) pill.classList.remove('open');

  currentTenantId = tenantId;
  window.currentTenantId = currentTenantId;
  localStorage.setItem('local_m365_current_tenant', tenantId);
  const select = document.getElementById('tenantSelect');
  if (select) select.value = tenantId;

  updateTenantPill(allTenants, tenantId);
  updateHeroVisibility();
  updateWorkspaceHeader(_currentSection);
  await populateSettings();
  await refreshTenantData();
}

function setupHeaderActions() {
  const refreshBtn = document.getElementById('logoutButton');
  if (refreshBtn) {
    refreshBtn.addEventListener('click', refreshTenantData);
  }

  const tenantSelect = document.getElementById('tenantSelect');
  if (tenantSelect) {
    tenantSelect.addEventListener('change', async (e) => {
      currentTenantId = e.target.value || null;
      window.currentTenantId = currentTenantId;
      if (currentTenantId) localStorage.setItem('local_m365_current_tenant', currentTenantId);
      updateTenantPill(allTenants, currentTenantId);
      updateHeroVisibility();
      await populateSettings();
      await refreshTenantData();
    });
  }

  // Tenant pill toggle
  const pill = document.getElementById('tenantPill');
  const pillDropdown = document.getElementById('tenantPillDropdown');
  if (pill && pillDropdown) {
    pill.addEventListener('click', (e) => {
      e.stopPropagation();
      const isOpen = pillDropdown.style.display !== 'none';
      pillDropdown.style.display = isOpen ? 'none' : 'block';
      pill.classList.toggle('open', !isOpen);
      if (!isOpen) closeUserDropdown();
    });
    pill.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); pill.click(); }
    });
  }

  // Avatar dropdown toggle
  const avatarBtn = document.getElementById('userAvatarBtn');
  const userDropdown = document.getElementById('portalUserDropdown');
  if (avatarBtn && userDropdown) {
    avatarBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      const isOpen = userDropdown.style.display !== 'none';
      userDropdown.style.display = isOpen ? 'none' : 'block';
      if (!isOpen) {
        if (pillDropdown) pillDropdown.style.display = 'none';
        if (pill) pill.classList.remove('open');
      }
    });
  }

  // Signout button — afmelden (lokale setup: reload)
  const signoutBtn = document.getElementById('signoutBtn');
  if (signoutBtn) {
    signoutBtn.addEventListener('click', async () => {
      localStorage.removeItem('local_m365_current_tenant');
      try { await apiFetch('/api/auth/logout', { method: 'POST' }); } catch (_) {}
      window.location.href = '/site/login.html';
    });
  }

  // Sluit dropdowns bij klik buiten
  document.addEventListener('click', closeAllDropdowns);

  const contextToggle = document.getElementById('portalContextToggle');
  const contextClose = document.getElementById('portalContextClose');
  if (contextToggle) {
    contextToggle.addEventListener('click', () => setContextRailOpen(!_contextRailOpen));
  }
  if (contextClose) {
    contextClose.addEventListener('click', () => setContextRailOpen(false));
  }
}

function closeUserDropdown() {
  const d = document.getElementById('portalUserDropdown');
  if (d) d.style.display = 'none';
}

function closeAllDropdowns() {
  const pd = document.getElementById('tenantPillDropdown');
  const ud = document.getElementById('portalUserDropdown');
  const pill = document.getElementById('tenantPill');
  if (pd) pd.style.display = 'none';
  if (ud) ud.style.display = 'none';
  if (pill) pill.classList.remove('open');
  document.querySelectorAll('.nav-dropdown.open').forEach((dropdown) => {
    if (!dropdown.classList.contains('has-active')) setDropdownOpen(dropdown, false);
  });
}

function setupSettingsActions() {
  const saveBtn = document.getElementById('saveLocalConfigButton');
  if (saveBtn) saveBtn.addEventListener('click', saveLocalConfig);
  const createBtn = document.getElementById('createTenantButton');
  if (createBtn) createBtn.addEventListener('click', createTenantFromForm);
  const governanceBtn = document.getElementById('saveTenantGovernanceButton');
  if (governanceBtn) governanceBtn.addEventListener('click', saveTenantGovernance);
  const saveIntegrationBtn = document.getElementById('saveIntegrationSettingsButton');
  if (saveIntegrationBtn) saveIntegrationBtn.addEventListener('click', saveIntegrationSettings);
  const applyRetentionBtn = document.getElementById('applySettingsRetentionButton');
  if (applyRetentionBtn) applyRetentionBtn.addEventListener('click', applySettingsRetentionPolicy);
  const refreshTenantMgmtBtn = document.getElementById('refreshTenantManagementButton');
  if (refreshTenantMgmtBtn) refreshTenantMgmtBtn.addEventListener('click', async () => {
    const tenantsData = await apiFetch('/api/tenants');
    renderTenantManagementTable(tenantsData.items || []);
  });

  const regenBtn = document.getElementById('regenerateAppButton');
  if (regenBtn) {
    regenBtn.textContent = 'Herlaad dashboard';
    regenBtn.addEventListener('click', () => window.location.reload());
  }
}

// ── Tenant Health Dashboard ────────────────────────────────────────────────

/**
 * Berekent een health-klasse op basis van de score.
 * @param {number|null} score
 * @returns {'good'|'warning'|'critical'|'unknown'}
 */
function _thHealthClass(score) {
  if (score == null) return 'unknown';
  if (score >= 85)   return 'good';
  if (score >= 60)   return 'warning';
  return 'critical';
}

/**
 * Formatteert een ISO timestamp als "X dagen geleden" of een datum.
 * @param {string|null} iso
 * @returns {string}
 */
function _thRelativeDate(iso) {
  if (!iso) return 'onbekend';
  try {
    const diff = Date.now() - new Date(iso).getTime();
    const days = Math.floor(diff / 86400000);
    if (days === 0) return 'vandaag';
    if (days === 1) return 'gisteren';
    if (days < 30)  return `${days} dagen geleden`;
    const months = Math.floor(days / 30);
    return `${months} maand${months > 1 ? 'en' : ''} geleden`;
  } catch (_) { return iso.slice(0, 10); }
}

/**
 * Bouwt een tenant health kaart als DOM-element.
 * @param {object} tenant  - tenant object incl. latest_run
 * @returns {HTMLElement}
 */
function _thBuildCard(tenant) {
  const run = tenant.latest_run;
  const score = run?.score_overall ?? null;
  const health = _thHealthClass(score);
  const hasScan = !!run && run.status === 'completed';

  const badgeClass = `th-score-badge--${health}`;
  const badgeLabel = score != null ? score : '—';

  const card = document.createElement('div');
  card.className = 'th-card';
  card.dataset.health = health;
  card.dataset.tenantId = tenant.id;

  // Risicoprofiel badge
  const riskColors = { low: '#16a34a', standard: '#b45309', high: '#ea580c', critical: '#dc2626' };
  const riskColor  = riskColors[tenant.risk_profile] || riskColors.standard;
  const riskLabel  = { low: 'Laag', standard: 'Standaard', high: 'Hoog', critical: 'Kritiek' }[tenant.risk_profile] || 'Standaard';

  // Signals opbouwen
  let signalsHtml = '';
  if (hasScan) {
    const crit = run.critical_count || 0;
    const warn = run.warning_count  || 0;
    const info = run.info_count     || 0;

    if (crit > 0) {
      signalsHtml += `<div class="th-signal"><span class="th-signal-dot th-signal-dot--crit"></span>${crit} kritieke bevinding${crit !== 1 ? 'en' : ''}</div>`;
    }
    if (warn > 0) {
      signalsHtml += `<div class="th-signal"><span class="th-signal-dot th-signal-dot--warn"></span>${warn} waarschuwing${warn !== 1 ? 'en' : ''}</div>`;
    }
    if (crit === 0 && warn === 0) {
      signalsHtml += `<div class="th-signal"><span class="th-signal-dot th-signal-dot--ok"></span>Geen kritieke bevindingen</div>`;
    }
    if (info > 0) {
      signalsHtml += `<div class="th-signal"><span class="th-signal-dot th-signal-dot--muted"></span>${info} informatief item${info !== 1 ? 's' : ''}</div>`;
    }
  } else {
    signalsHtml = `<div class="th-no-scan">Nog geen voltooide scan beschikbaar</div>`;
  }

  const scanDate = hasScan
    ? `Laatste scan: ${_thRelativeDate(run.completed_at || run.started_at)}`
    : 'Nog niet gescand';

  const statusColors = { active: '#16a34a', onboarding: '#b45309', paused: '#6b7280', offboarded: '#dc2626' };
  const statusColor  = statusColors[tenant.status] || '#6b7280';

  card.innerHTML = `
    <div class="th-card-head">
      <div>
        <p class="th-card-name">${escapeHtml(tenant.customer_name || tenant.tenant_name || 'Onbekend')}</p>
        <p class="th-card-sub">${escapeHtml(tenant.tenant_name || '')}</p>
        <span style="display:inline-flex;align-items:center;gap:.3rem;font-size:.7rem;font-weight:600;color:${riskColor};margin-top:.3rem;">
          <span style="width:6px;height:6px;border-radius:50%;background:${riskColor};display:inline-block;"></span>
          Risico: ${riskLabel}
        </span>
      </div>
      <div class="th-score-badge ${badgeClass}" title="Overallscore laatste scan">
        ${badgeLabel}
        <small>score</small>
      </div>
    </div>
    <div class="th-signals">${signalsHtml}</div>
    <div class="th-card-meta">
      <span style="width:7px;height:7px;border-radius:50%;background:${statusColor};display:inline-block;flex-shrink:0;"></span>
      ${scanDate}
    </div>
    <div class="th-card-footer">
      <button type="button" class="th-btn th-btn--primary" data-action="open" data-tenant-id="${escapeHtml(tenant.id)}">
        Open →
      </button>
      <button type="button" class="th-btn th-btn--secondary" data-action="scan" data-tenant-id="${escapeHtml(tenant.id)}" title="Assessment starten voor deze tenant">
        ▷ Scan
      </button>
    </div>
  `;

  // Knoppen koppelen
  card.querySelector('[data-action="open"]').addEventListener('click', async () => {
    await selectTenantFromManagement(tenant.id);
    showSection('overview');
  });
  card.querySelector('[data-action="scan"]').addEventListener('click', async () => {
    await selectTenantFromManagement(tenant.id);
    showSection('assessment');
  });

  return card;
}

/**
 * Rendert skeletons tijdens het laden.
 * @param {HTMLElement} grid
 * @param {number} count
 */
function _thRenderSkeletons(grid, count = 6) {
  grid.innerHTML = '';
  for (let i = 0; i < count; i++) {
    const sk = document.createElement('div');
    sk.className = 'th-card';
    sk.dataset.health = 'unknown';
    sk.innerHTML = `
      <div class="th-card-head">
        <div style="flex:1;">
          <div class="th-skeleton" style="height:14px;width:65%;margin-bottom:8px;"></div>
          <div class="th-skeleton" style="height:10px;width:45%;"></div>
        </div>
        <div class="th-skeleton" style="width:52px;height:52px;border-radius:10px;flex-shrink:0;"></div>
      </div>
      <div class="th-signals">
        <div class="th-skeleton" style="height:10px;width:80%;margin-bottom:6px;"></div>
        <div class="th-skeleton" style="height:10px;width:60%;"></div>
      </div>
      <div class="th-skeleton" style="height:10px;width:55%;margin-bottom:.9rem;"></div>
      <div class="th-card-footer" style="border-top:none;padding-top:0;">
        <div class="th-skeleton" style="height:32px;flex:1;border-radius:8px;"></div>
        <div class="th-skeleton" style="height:32px;width:60px;border-radius:8px;"></div>
      </div>`;
    grid.appendChild(sk);
  }
}

/**
 * Laadt en rendert het Tenant Health Dashboard.
 * Roept GET /api/tenants aan (bestaat al, geeft latest_run mee).
 * Veilig: elke fout toont een melding zonder de rest te breken.
 */
async function loadTenantHealthDashboard() {
  // Dubbele check — mocht de functie toch aangeroepen worden zonder admin-rol
  if (_currentUserRole !== 'admin') return;

  const grid   = document.getElementById('thGrid');
  const pills  = document.getElementById('thSummaryPills');
  if (!grid) return;

  _thRenderSkeletons(grid);

  let tenants = [];
  try {
    const data = await apiFetch('/api/tenants');
    tenants = (data && data.items) ? data.items : [];
  } catch (e) {
    grid.innerHTML = `<div class="empty-state" style="grid-column:1/-1;">Tenants laden mislukt: ${escapeHtml(e.message)}</div>`;
    return;
  }

  if (tenants.length === 0) {
    grid.innerHTML = `<div class="empty-state" style="grid-column:1/-1;">Nog geen tenants aangemaakt. Voeg een tenant toe via Admin → Tenants.</div>`;
    if (pills) pills.innerHTML = '';
    return;
  }

  // Samenvattingspillen berekenen
  const total   = tenants.length;
  const scanned = tenants.filter((t) => t.latest_run?.status === 'completed').length;
  const crits   = tenants.filter((t) => (t.latest_run?.critical_count || 0) > 0).length;
  const ok      = tenants.filter((t) => {
    const r = t.latest_run;
    return r?.status === 'completed' && (r.critical_count || 0) === 0 && (r.warning_count || 0) === 0;
  }).length;

  if (pills) {
    pills.innerHTML = `
      <span class="th-pill th-pill--total">◈ ${total} tenant${total !== 1 ? 's' : ''}</span>
      ${scanned > 0 ? `<span class="th-pill th-pill--ok">✓ ${ok} schoon</span>` : ''}
      ${crits  > 0 ? `<span class="th-pill th-pill--crit">✕ ${crits} kritiek</span>` : ''}
      ${(scanned < total) ? `<span class="th-pill th-pill--warn">⚑ ${total - scanned} niet gescand</span>` : ''}
    `;
  }

  // Sorteren: kritiek bovenaan, dan waarschuwing, dan goed, dan onbekend
  const order = { critical: 0, warning: 1, good: 2, unknown: 3 };
  tenants.sort((a, b) => {
    return (order[_thHealthClass(a.latest_run?.score_overall)] ?? 3)
         - (order[_thHealthClass(b.latest_run?.score_overall)] ?? 3);
  });

  // Kaarten renderen
  grid.innerHTML = '';
  tenants.forEach((tenant) => grid.appendChild(_thBuildCard(tenant)));

  // Laad MSP aggregate KPI's
  _loadMspAggregate();
}
window.loadTenantHealthDashboard = loadTenantHealthDashboard;

async function _loadMspAggregate() {
  const bar = document.getElementById('mspAggregateBar');
  if (!bar) return;
  bar.innerHTML = '<div style="color:var(--text-muted);font-size:0.8rem;padding:0.5rem">Statistieken laden…</div>';
  try {
    const d = await apiFetch('/api/msp/aggregate');
    const kpis = [
      { label: 'Tenants',         value: d.total_tenants ?? '—',           tone: '' },
      { label: 'Gem. score',      value: d.avg_score != null ? d.avg_score + '%' : '—', tone: d.avg_score != null && d.avg_score < 60 ? 'warn' : 'ok' },
      { label: 'Kritiek',         value: d.tenants_with_critical ?? '—',    tone: d.tenants_with_critical > 0 ? 'bad' : 'ok' },
      { label: 'Geen assessment', value: d.tenants_no_assessment ?? '—',    tone: d.tenants_no_assessment > 0 ? 'warn' : 'ok' },
      { label: 'Verouderd (>30d)',value: d.tenants_stale_assessment ?? '—', tone: d.tenants_stale_assessment > 0 ? 'warn' : 'ok' },
    ];
    bar.innerHTML = kpis.map(k => {
      const col = k.tone === 'bad' ? '#dc2626' : k.tone === 'warn' ? '#d97706' : k.tone === 'ok' ? '#16a34a' : 'var(--text)';
      return `<div style="background:var(--card,#fff);border:1px solid var(--border,#e2e8f0);border-radius:10px;padding:0.75rem 1rem;">
        <div style="font-size:0.7rem;font-weight:600;text-transform:uppercase;letter-spacing:0.06em;color:var(--text-muted);margin-bottom:0.25rem">${escapeHtml(k.label)}</div>
        <div style="font-size:1.4rem;font-weight:800;color:${col};line-height:1">${escapeHtml(String(k.value))}</div>
      </div>`;
    }).join('');
  } catch (_) {
    bar.innerHTML = '';
  }
}

// Huidige gebruikersrol — standaard 'klant' totdat /api/auth/verify bevestigt
let _currentUserRole = 'klant';

// Setter zodat het inline auth-script in dashboard.html de rol ook kan doorzetten
window._setDashboardRole = function(role) {
  _currentUserRole = role;
  _applyRoleVisibility();
};

// Getter zodat andere modules (bijv. live-modules.js) de huidige rol kunnen opvragen
window._getDashboardRole = function() { return _currentUserRole; };

/**
 * Haalt de huidige sessierol op via /api/auth/verify.
 * Bij elke fout valt het veilig terug op 'klant' — admin functionaliteit
 * blijft dan verborgen totdat de rol expliciet bevestigd is.
 */
async function _loadCurrentRole() {
  try {
    const res = await apiFetch('/api/auth/verify');
    if (res && res.ok && res.role) {
      _currentUserRole = res.role;
      // Naam en initialen bijwerken vanuit sessie
      const name = res.display_name || res.email || 'Lokaal';
      const userNameEl = document.getElementById('userName');
      if (userNameEl) userNameEl.textContent = name;
      const initialsEl = document.getElementById('userInitials');
      if (initialsEl) initialsEl.textContent = getInitials(name);
      const avatarBtn = document.getElementById('userAvatarBtn');
      if (avatarBtn) avatarBtn.title = name;
    }
  } catch (_) {
    // Stil falen — rol blijft 'klant' (veiligste standaard)
  }
}

/**
 * Toont admin-only UI elementen als de huidige rol 'admin' is.
 * Verwijdert de 'admin-only-nav' class zodat het element zichtbaar wordt.
 * De CSS !important op die class wint altijd van inline styles —
 * daarom class-remove i.p.v. style.removeProperty.
 */
function _applyRoleVisibility() {
  if (_currentUserRole === 'admin') {
    document.querySelectorAll('.admin-only-nav').forEach((el) => {
      el.classList.remove('admin-only-nav');
    });
  }
}

async function bootstrap() {
  try {
    // Zet alvast een veilige standaardnaam
    const fallbackName = 'Lokaal';
    const userNameEl = document.getElementById('userName');
    if (userNameEl) userNameEl.textContent = fallbackName;
    const initialsEl = document.getElementById('userInitials');
    if (initialsEl) initialsEl.textContent = getInitials(fallbackName);

    initThemeControls();
    setupNavigation();
    setupHeaderActions();
    setupSettingsActions();
    switchSettingsTab('tenant');
    updateSubnav('overview');
    const prefs = getUiPrefs();
    setSidebarCompact(!!prefs.sidebarCompact);
    updateWorkspaceHeader('overview');
    setContextRailOpen(prefs.contextRailOpen === true);

    // Rol ophalen vóór tenants laden zodat admin UI correct zichtbaar is
    await _loadCurrentRole();
    _applyRoleVisibility();

    await loadTenants();
    await populateSettings();
    await refreshTenantData();
    renderContextRail('overview');
    renderNavSignals();
  } catch (e) {
    console.error(e);
    alert(`Dashboard initialisatie mislukt: ${e.message}`);
  }
}

document.addEventListener('DOMContentLoaded', () => {
  bootstrap();

  // ── Hamburger menu toggle (mobiel) ─────────────────────────────────────────
  const hamburger = document.getElementById('mobileMenuToggle');
  if (hamburger) {
    const navBar = hamburger.closest('.portal-nav-bar');
    hamburger.addEventListener('click', () => {
      const isOpen = navBar.classList.toggle('nav-mobile-open');
      hamburger.setAttribute('aria-expanded', String(isOpen));
      hamburger.setAttribute('aria-label', isOpen ? 'Menu sluiten' : 'Menu openen');
    });
    // Sluit menu bij klik buiten de nav
    document.addEventListener('click', (e) => {
      if (navBar && !navBar.contains(e.target)) {
        navBar.classList.remove('nav-mobile-open');
        hamburger.setAttribute('aria-expanded', 'false');
        hamburger.setAttribute('aria-label', 'Menu openen');
      }
    });
    // Sluit menu na navigatie (click op een nav-link)
    document.querySelectorAll('.portal-nav-links .portal-nav-link, .portal-nav-links .nav-dropdown-link').forEach((link) => {
      link.addEventListener('click', () => {
        navBar.classList.remove('nav-mobile-open');
        hamburger.setAttribute('aria-expanded', 'false');
      });
    });
  }

  // Vernieuwen-knop tenant health dashboard
  const thRefreshBtn = document.getElementById('thRefreshBtn');
  if (thRefreshBtn) {
    thRefreshBtn.addEventListener('click', () => loadTenantHealthDashboard());
  }

  // Wire up Rapporten sidebar tabbar clicks
  document.querySelectorAll('#resultsTabbar [data-results-panel]').forEach((tab) => {
    tab.addEventListener('click', (e) => {
      e.preventDefault();
      showResultsPanel(tab.dataset.resultsPanel);
    });
  });

  // Klantenbeheer knoppen
  const kbhRefreshBtn = document.getElementById('kbhRefreshBtn');
  if (kbhRefreshBtn) kbhRefreshBtn.addEventListener('click', () => loadKlantenbeheer());

  const kbhAddBtn = document.getElementById('kbhAddBtn');
  if (kbhAddBtn) kbhAddBtn.addEventListener('click', () => _showKlantForm(null));

  const kbhDetailClose = document.getElementById('kbhDetailClose');
  if (kbhDetailClose) kbhDetailClose.addEventListener('click', () => {
    const p = document.getElementById('kbhDetailPanel');
    if (p) p.style.display = 'none';
  });

  const kbhSearch = document.getElementById('kbhSearch');
  if (kbhSearch) kbhSearch.addEventListener('input', () => _filterKlantTable(kbhSearch.value));

  // Goedkeuringen filter knoppen
  document.querySelectorAll('.gdk-filter-btn').forEach((btn) => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.gdk-filter-btn').forEach((b) => b.classList.remove('active'));
      btn.classList.add('active');
      loadGoedkeuringen(btn.dataset.status || null);
    });
  });

  const gdkRefreshBtn = document.getElementById('gdkRefreshBtn');
  if (gdkRefreshBtn) gdkRefreshBtn.addEventListener('click', () => {
    const active = document.querySelector('.gdk-filter-btn.active');
    loadGoedkeuringen(active ? (active.dataset.status || null) : 'pending');
  });

  // Kosten knoppen
  const kostenRefreshBtn = document.getElementById('kostenRefreshBtn');
  if (kostenRefreshBtn) kostenRefreshBtn.addEventListener('click', () => loadKostenSection());
  const kostenTenantSel = document.getElementById('kostenTenantSel');
  if (kostenTenantSel) kostenTenantSel.addEventListener('change', () => loadKostenSection());

  // Job monitor knoppen
  document.querySelectorAll('.jm-filter').forEach((btn) => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.jm-filter').forEach((b) => b.classList.remove('active'));
      btn.classList.add('active');
      loadJobMonitor(btn.dataset.status || null);
    });
  });
  const jmRefreshBtn = document.getElementById('jmRefreshBtn');
  if (jmRefreshBtn) jmRefreshBtn.addEventListener('click', () => {
    const active = document.querySelector('.jm-filter.active');
    loadJobMonitor(active ? (active.dataset.status || null) : null);
  });
  const jmEnqueueBtn = document.getElementById('jmEnqueueBtn');
  if (jmEnqueueBtn) jmEnqueueBtn.addEventListener('click', () => _jmEnqueue());

  // Hub tile clicks
  document.addEventListener('click', (e) => {
    const tile = e.target.closest('.hub-tile[data-section]');
    if (!tile) return;
    e.preventDefault();
    const ds = tile.dataset;
    const opts = {};
    if (ds.gbTab) opts.gbTab = ds.gbTab;
    if (ds.liveTab) opts.liveTab = ds.liveTab;
    if (ds.caTab) opts.caTab = ds.caTab;
    if (ds.kbTab) opts.kbTab = ds.kbTab;
    if (ds.resultsPanel) opts.resultsPanel = ds.resultsPanel;
    if (ds.settingsTab) opts.settingsTab = ds.settingsTab;
    showSection(ds.section, opts);
  });

  // Rollen tab knoppen
  const rolesRefreshBtn = document.getElementById('rolesRefreshBtn');
  if (rolesRefreshBtn) rolesRefreshBtn.addEventListener('click', () => loadRolesTab());
  const rolesAddUserBtn = document.getElementById('rolesAddUserBtn');
  if (rolesAddUserBtn) rolesAddUserBtn.addEventListener('click', async () => {
    const email = window.prompt('E-mailadres nieuwe gebruiker:');
    if (!email || !email.trim()) return;
    const name = window.prompt('Naam:', '') || email;
    const role = window.prompt('Rol (admin / klant):', 'klant') || 'klant';
    const pw = window.prompt('Tijdelijk wachtwoord:');
    if (!pw) return;
    try {
      const resp = await apiFetch('/api/users', {
        method: 'POST',
        body: JSON.stringify({ email: email.trim(), display_name: name.trim(), role: role.trim(), password: pw }),
      });
      if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
      showToast('Gebruiker aangemaakt.', 'success');
      loadRolesTab();
    } catch (e) {
      showToast(`Fout: ${e}`, 'error');
    }
  });
});

// ── Klantenbeheer ─────────────────────────────────────────────────────────────
let _klantenCache = [];

async function loadKlantenbeheer() {
  const tbody = document.getElementById('kbhTableBody');
  const summary = document.getElementById('kbhSummary');
  if (tbody) tbody.innerHTML = '<tr><td colspan="6" style="text-align:center;padding:2rem;color:var(--text-muted,#6b7280);">Laden…</td></tr>';
  try {
    const data = await apiFetch('/api/customers');
    _klantenCache = (data && data.items) || [];
    _renderKlantTable(_klantenCache);
    if (summary) {
      const active = _klantenCache.filter((c) => c.status === 'active').length;
      summary.innerHTML = `
        <span style="padding:.25rem .7rem;border-radius:999px;background:#dcfce7;color:#166534;font-size:.8rem;font-weight:600;">${active} actief</span>
        <span style="padding:.25rem .7rem;border-radius:999px;background:#f3f4f6;color:#374151;font-size:.8rem;">${_klantenCache.length} totaal</span>
      `;
    }
  } catch (e) {
    if (tbody) tbody.innerHTML = `<tr><td colspan="6" style="text-align:center;color:#dc2626;padding:2rem;">Fout bij laden klanten: ${escapeHtml(String(e))}</td></tr>`;
  }
}

function _renderKlantTable(items) {
  const tbody = document.getElementById('kbhTableBody');
  if (!tbody) return;
  if (!items.length) {
    tbody.innerHTML = '<tr><td colspan="6" style="text-align:center;color:var(--text-muted,#6b7280);padding:2rem;">Geen klanten gevonden.</td></tr>';
    return;
  }
  tbody.innerHTML = items.map((c) => `
    <tr data-klant-id="${escapeHtml(c.id)}">
      <td><strong>${escapeHtml(c.name || '-')}</strong></td>
      <td><span style="padding:.2rem .6rem;border-radius:999px;font-size:.78rem;font-weight:600;background:${c.status === 'active' ? '#dcfce7' : '#fee2e2'};color:${c.status === 'active' ? '#166534' : '#991b1b'};">${escapeHtml(c.status || '-')}</span></td>
      <td>${c.tenant_count ?? '-'}</td>
      <td>${escapeHtml(c.primary_contact_name || '-')}</td>
      <td>${formatDate(c.created_at)}</td>
      <td style="white-space:nowrap;">
        <button type="button" class="btn btn-secondary" style="font-size:.75rem;padding:.25rem .6rem;" onclick="_showKlantDetail('${escapeHtml(c.id)}')">Details</button>
        <button type="button" class="btn btn-secondary" style="font-size:.75rem;padding:.25rem .6rem;" onclick="_showKlantForm('${escapeHtml(c.id)}')">Bewerken</button>
      </td>
    </tr>
  `).join('');
}

function _filterKlantTable(query) {
  const q = (query || '').toLowerCase();
  const filtered = q ? _klantenCache.filter((c) => (c.name || '').toLowerCase().includes(q)) : _klantenCache;
  _renderKlantTable(filtered);
}

async function _showKlantDetail(customerId) {
  const panel = document.getElementById('kbhDetailPanel');
  const nameEl = document.getElementById('kbhDetailName');
  const body = document.getElementById('kbhDetailBody');
  if (!panel || !body) return;
  panel.style.display = 'block';
  body.innerHTML = '<p style="color:var(--text-muted,#6b7280);">Laden…</p>';
  try {
    const [c, onb] = await Promise.all([
      apiFetch(`/api/customers/${customerId}`),
      apiFetch(`/api/customers/${customerId}/onboarding`),
    ]);
    if (nameEl) nameEl.textContent = c.name || 'Klant';
    const tenants = (c.tenants || []);
    const onbTenants = (onb.tenants || []);
    body.innerHTML = `
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:1rem;margin-bottom:1rem;">
        <div><strong>Status:</strong> ${escapeHtml(c.status || '-')}</div>
        <div><strong>Contactpersoon:</strong> ${escapeHtml(c.primary_contact_name || '-')}</div>
        <div><strong>E-mail:</strong> ${escapeHtml(c.primary_contact_email || '-')}</div>
        <div><strong>Tenants:</strong> ${tenants.length}</div>
      </div>
      ${c.notes ? `<p style="color:var(--text-muted,#6b7280);font-size:.875rem;">${escapeHtml(c.notes)}</p>` : ''}
      <h4 style="margin:.75rem 0 .5rem;">Onboarding voortgang</h4>
      ${onbTenants.map((t) => `
        <div style="margin-bottom:.75rem;padding:.75rem;border:1px solid var(--border-color,#e5e7eb);border-radius:6px;">
          <div style="display:flex;justify-content:space-between;margin-bottom:.4rem;">
            <strong style="font-size:.875rem;">${escapeHtml(t.tenant_name || t.tenant_id)}</strong>
            <span style="font-size:.8rem;color:var(--text-muted,#6b7280);">${t.completion_pct ?? 0}%</span>
          </div>
          <div style="height:6px;background:#f3f4f6;border-radius:3px;overflow:hidden;">
            <div style="height:100%;width:${t.completion_pct ?? 0}%;background:#2563eb;border-radius:3px;transition:width .3s;"></div>
          </div>
          <div style="display:flex;flex-wrap:wrap;gap:.35rem;margin-top:.5rem;">
            ${(t.steps || []).map((s) => `
              <span style="font-size:.72rem;padding:.1rem .5rem;border-radius:999px;background:${s.done ? '#dcfce7' : '#fee2e2'};color:${s.done ? '#166534' : '#991b1b'};">${s.done ? '✓' : '✗'} ${escapeHtml(s.label)}</span>
            `).join('')}
          </div>
        </div>
      `).join('') || '<p style="color:var(--text-muted,#6b7280);font-size:.875rem;">Geen tenants gekoppeld.</p>'}
    `;
  } catch (e) {
    body.innerHTML = `<p style="color:#dc2626;">Fout bij laden: ${escapeHtml(String(e))}</p>`;
  }
}

async function _showKlantForm(customerId) {
  const isNew = !customerId;
  let existing = {};
  if (!isNew) {
    try {
      existing = await apiFetch(`/api/customers/${customerId}`) || {};
    } catch (_) {}
  }
  const name = window.prompt(isNew ? 'Naam nieuwe klant:' : `Naam (huidig: ${existing.name || ''}):`);
  if (name === null) return;
  const contact = window.prompt('Contactpersoon (naam):', existing.primary_contact_name || '');
  if (contact === null) return;
  const email = window.prompt('Contactpersoon e-mail:', existing.primary_contact_email || '');
  if (email === null) return;
  try {
    const payload = { name: name.trim(), primary_contact_name: contact.trim(), primary_contact_email: email.trim() };
    if (!isNew) payload.status = existing.status || 'active';
    const resp = await apiFetch(
      isNew ? '/api/customers' : `/api/customers/${customerId}`,
      { method: isNew ? 'POST' : 'PATCH', body: JSON.stringify(payload) }
    );
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    showToast(isNew ? 'Klant aangemaakt.' : 'Klant bijgewerkt.', 'success');
    loadKlantenbeheer();
  } catch (e) {
    showToast(`Fout: ${e}`, 'error');
  }
}
window.loadKlantenbeheer = loadKlantenbeheer;
window._showKlantDetail = _showKlantDetail;
window._showKlantForm = _showKlantForm;

// ── Goedkeuringen ─────────────────────────────────────────────────────────────
async function loadGoedkeuringen(statusFilter = 'pending') {
  const tbody = document.getElementById('gdkTableBody');
  if (tbody) tbody.innerHTML = '<tr><td colspan="7" style="text-align:center;padding:2rem;color:var(--text-muted,#6b7280);">Laden…</td></tr>';
  try {
    const params = new URLSearchParams({ limit: '200' });
    if (statusFilter) params.set('status', statusFilter);
    const data = await apiFetch(`/api/approvals?${params}`);
    const items = (data && data.items) || [];
    if (!tbody) return;
    if (!items.length) {
      tbody.innerHTML = '<tr><td colspan="7" style="text-align:center;color:var(--text-muted,#6b7280);padding:2rem;">Geen goedkeuringen gevonden.</td></tr>';
      return;
    }
    const statusColor = { pending: '#fef3c7:#92400e', approved: '#dcfce7:#166534', rejected: '#fee2e2:#991b1b' };
    tbody.innerHTML = items.map((a) => {
      const [bg, fg] = (statusColor[a.approval_status] || '#f3f4f6:#374151').split(':');
      return `<tr>
        <td style="font-size:.82rem;max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${escapeHtml(a.action_log_id || '-')}</td>
        <td><span style="padding:.2rem .6rem;border-radius:999px;font-size:.78rem;font-weight:600;background:${bg};color:${fg};">${escapeHtml(a.approval_status || '-')}</span></td>
        <td>${escapeHtml(a.requested_by || '-')}</td>
        <td>${formatDate(a.requested_at)}</td>
        <td style="max-width:180px;overflow:hidden;text-overflow:ellipsis;">${escapeHtml(a.reason || '-')}</td>
        <td>${escapeHtml(a.approved_by || '-')}</td>
        <td style="white-space:nowrap;">
          ${a.approval_status === 'pending' ? `
            <button type="button" class="btn btn-primary" style="font-size:.75rem;padding:.25rem .6rem;" onclick="_gdkDecide('${escapeHtml(a.id)}','approve')">Goedkeuren</button>
            <button type="button" class="btn btn-secondary" style="font-size:.75rem;padding:.25rem .6rem;color:#dc2626;" onclick="_gdkDecide('${escapeHtml(a.id)}','reject')">Afwijzen</button>
          ` : ''}
        </td>
      </tr>`;
    }).join('');
  } catch (e) {
    if (tbody) tbody.innerHTML = `<tr><td colspan="7" style="text-align:center;color:#dc2626;padding:2rem;">Fout: ${escapeHtml(String(e))}</td></tr>`;
  }
}

async function _gdkDecide(approvalId, decision) {
  const reason = window.prompt(decision === 'approve' ? 'Reden goedkeuring (optioneel):' : 'Reden afwijzing (optioneel):', '');
  if (reason === null) return;
  try {
    const resp = await apiFetch(`/api/approvals/${approvalId}/${decision === 'approve' ? 'approve' : 'reject'}`, {
      method: 'POST',
      body: JSON.stringify({ reason: reason.trim() || undefined }),
    });
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    showToast(decision === 'approve' ? 'Goedgekeurd.' : 'Afgewezen.', 'success');
    const active = document.querySelector('.gdk-filter-btn.active');
    loadGoedkeuringen(active ? (active.dataset.status || null) : 'pending');
  } catch (e) {
    showToast(`Fout: ${e}`, 'error');
  }
}
window.loadGoedkeuringen = loadGoedkeuringen;
window._gdkDecide = _gdkDecide;

// ── Kosten Overzicht ──────────────────────────────────────────────────────────
async function loadKostenSection() {
  const sel = document.getElementById('kostenTenantSel');
  const tbody = document.getElementById('kostenTableBody');
  const summary = document.getElementById('kostenSummary');

  // Vul tenant selector de eerste keer
  if (sel && sel.options.length <= 1) {
    try {
      const r = await apiFetch('/api/tenants');
      const tenants = (r && r.items) || [];
      tenants.forEach((t) => {
        const opt = document.createElement('option');
        opt.value = t.id;
        opt.textContent = `${t.customer_name || t.tenant_name}`;
        sel.appendChild(opt);
      });
    } catch (_) {}
  }

  const tenantId = sel ? sel.value : '';
  if (!tenantId) {
    if (tbody) tbody.innerHTML = '<tr><td colspan="6" style="text-align:center;color:var(--text-muted,#6b7280);padding:2rem;">Selecteer een tenant.</td></tr>';
    return;
  }
  if (tbody) tbody.innerHTML = '<tr><td colspan="6" style="text-align:center;padding:2rem;color:var(--text-muted,#6b7280);">Laden…</td></tr>';
  try {
    const data = await apiFetch(`/api/tenants/${tenantId}/cost-snapshots`);
    const items = (data && data.items) || [];
    if (!tbody) return;
    if (!items.length) {
      tbody.innerHTML = '<tr><td colspan="6" style="text-align:center;color:var(--text-muted,#6b7280);padding:2rem;">Geen kostendata beschikbaar voor deze tenant.</td></tr>';
      if (summary) summary.innerHTML = '';
      return;
    }
    let totalCost = 0;
    tbody.innerHTML = items.map((s) => {
      let sumObj = {};
      try { sumObj = JSON.parse(s.summary_json || '{}'); } catch (_) {}
      const cost = parseFloat(sumObj.total_cost || sumObj.totalCost || 0);
      totalCost += cost;
      const currency = sumObj.currency || 'EUR';
      return `<tr>
        <td>${escapeHtml(s.tenant_id || '-')}</td>
        <td style="font-size:.8rem;color:var(--text-muted,#6b7280);">${escapeHtml(s.subscription_id || 'Alle')}</td>
        <td>${escapeHtml(s.period_start || '-')} – ${escapeHtml(s.period_end || '-')}</td>
        <td style="font-weight:600;">${cost > 0 ? `€ ${cost.toFixed(2)}` : '-'}</td>
        <td>${escapeHtml(currency)}</td>
        <td>${formatDate(s.generated_at)}</td>
      </tr>`;
    }).join('');
    if (summary) {
      summary.innerHTML = `
        <span style="padding:.25rem .7rem;border-radius:999px;background:#dbeafe;color:#1e40af;font-size:.8rem;font-weight:600;">Totaal: € ${totalCost.toFixed(2)}</span>
        <span style="padding:.25rem .7rem;border-radius:999px;background:#f3f4f6;color:#374151;font-size:.8rem;">${items.length} perioden</span>
      `;
    }
  } catch (e) {
    if (tbody) tbody.innerHTML = `<tr><td colspan="6" style="text-align:center;color:#dc2626;padding:2rem;">Fout: ${escapeHtml(String(e))}</td></tr>`;
  }
}
window.loadKostenSection = loadKostenSection;

// ── Job Monitor ───────────────────────────────────────────────────────────────
const _jmStatusColors = {
  pending:   ['#fef3c7', '#92400e'],
  running:   ['#dbeafe', '#1e40af'],
  completed: ['#dcfce7', '#166534'],
  failed:    ['#fee2e2', '#991b1b'],
  cancelled: ['#f3f4f6', '#6b7280'],
};

async function loadJobMonitor(statusFilter = null) {
  const tbody = document.getElementById('jmTableBody');
  const summary = document.getElementById('jmSummary');
  if (tbody) tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:2rem;color:var(--text-muted,#6b7280);">Laden…</td></tr>';
  try {
    const params = new URLSearchParams({ limit: '200' });
    if (statusFilter) params.set('status', statusFilter);
    const data = await apiFetch(`/api/jobs?${params}`);
    const items = (data && data.items) || [];

    // summary pills
    if (summary) {
      const counts = {};
      items.forEach((j) => { counts[j.status] = (counts[j.status] || 0) + 1; });
      summary.innerHTML = Object.entries(counts).map(([st, cnt]) => {
        const [bg, fg] = _jmStatusColors[st] || ['#f3f4f6', '#374151'];
        return `<span style="padding:.25rem .7rem;border-radius:999px;background:${bg};color:${fg};font-size:.8rem;font-weight:600;">${cnt} ${st}</span>`;
      }).join('');
    }

    if (!tbody) return;
    if (!items.length) {
      tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;color:var(--text-muted,#6b7280);padding:2rem;">Geen jobs gevonden.</td></tr>';
      return;
    }
    tbody.innerHTML = items.map((j) => {
      const [bg, fg] = _jmStatusColors[j.status] || ['#f3f4f6', '#374151'];
      const canCancel = j.status === 'pending' || j.status === 'failed';
      return `<tr>
        <td style="font-family:monospace;font-size:.8rem;">${escapeHtml(j.job_type)}</td>
        <td style="font-size:.78rem;color:var(--text-muted,#6b7280);">${escapeHtml(j.tenant_id || '-')}</td>
        <td><span style="padding:.2rem .6rem;border-radius:999px;font-size:.78rem;font-weight:600;background:${bg};color:${fg};">${j.status}</span></td>
        <td style="text-align:center;">${j.priority}</td>
        <td style="text-align:center;">${j.attempt_count}/${3}</td>
        <td style="font-size:.78rem;">${formatDate(j.scheduled_at)}</td>
        <td style="font-size:.78rem;">${j.completed_at ? formatDate(j.completed_at) : '-'}</td>
        <td style="font-size:.75rem;color:#dc2626;max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${escapeHtml(j.error_message || '')}">${escapeHtml(j.error_message || '-')}</td>
        <td>${canCancel ? `<button type="button" class="btn btn-secondary" style="font-size:.75rem;padding:.2rem .5rem;color:#dc2626;" onclick="_jmCancel('${escapeHtml(j.id)}')">Annuleren</button>` : ''}</td>
      </tr>`;
    }).join('');
  } catch (e) {
    if (tbody) tbody.innerHTML = `<tr><td colspan="9" style="text-align:center;color:#dc2626;padding:2rem;">Fout: ${escapeHtml(String(e))}</td></tr>`;
  }
}

async function _jmCancel(jobId) {
  if (!confirm('Job annuleren?')) return;
  try {
    const resp = await apiFetch(`/api/jobs/${jobId}/cancel`, { method: 'POST', body: '{}' });
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    showToast('Job geannuleerd.', 'success');
    const active = document.querySelector('.jm-filter.active');
    loadJobMonitor(active ? (active.dataset.status || null) : null);
  } catch (e) {
    showToast(`Fout: ${e}`, 'error');
  }
}

async function _jmEnqueue() {
  const jobType = window.prompt('Job type (bijv. assessment_run, snapshot_import):', 'assessment_run');
  if (!jobType) return;
  const tenantId = window.prompt('Tenant ID (leeg = geen):', '') || null;
  try {
    const resp = await apiFetch('/api/jobs', {
      method: 'POST',
      body: JSON.stringify({ job_type: jobType.trim(), tenant_id: tenantId }),
    });
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    showToast('Job aangemaakt.', 'success');
    loadJobMonitor();
  } catch (e) {
    showToast(`Fout: ${e}`, 'error');
  }
}
window.loadJobMonitor = loadJobMonitor;
window._jmCancel = _jmCancel;

// ── Integratiestatus Grid (settings > tenant tab) ──────────────────────────────
async function loadIntegratieStatusGrid() {
  const grid = document.getElementById('integratieStatusGrid');
  if (!grid) return;
  // Gebruik actieve tenant als beschikbaar
  const tenantId = currentTenantId || (allTenants && allTenants[0] && allTenants[0].id);
  if (!tenantId) {
    grid.innerHTML = '<p style="color:var(--text-muted,#6b7280);font-size:.875rem;grid-column:1/-1;">Selecteer een tenant om integratiestatus te laden.</p>';
    return;
  }
  grid.innerHTML = '<p style="color:var(--text-muted,#6b7280);font-size:.875rem;grid-column:1/-1;">Laden…</p>';
  try {
    const data = await apiFetch(`/api/tenants/${tenantId}/integrations`);
    const items = (data && data.items) || [];
    if (!items.length) {
      grid.innerHTML = '<p style="color:var(--text-muted,#6b7280);font-size:.875rem;grid-column:1/-1;">Geen integraties geconfigureerd voor deze tenant.</p>';
      return;
    }
    const statusIcon = { active: '✓', unknown: '?', error: '✗', inactive: '○' };
    const statusColor = { active: '#16a34a', unknown: '#d97706', error: '#dc2626', inactive: '#6b7280' };
    grid.innerHTML = items.map((integ) => {
      const st = integ.status || 'unknown';
      const color = statusColor[st] || '#6b7280';
      const icon = statusIcon[st] || '?';
      return `<div style="border:1px solid var(--border-color,#e5e7eb);border-radius:8px;padding:.875rem 1rem;background:var(--card-bg,#fff);">
        <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:.5rem;">
          <strong style="font-size:.9rem;">${escapeHtml(integ.integration_type || '-')}</strong>
          <span style="font-size:.8rem;font-weight:600;color:${color};">${icon} ${escapeHtml(st)}</span>
        </div>
        ${integ.gdap_status ? `<div style="font-size:.78rem;color:var(--text-muted,#6b7280);">GDAP: <strong>${escapeHtml(integ.gdap_status)}</strong></div>` : ''}
        ${integ.app_registration_status ? `<div style="font-size:.78rem;color:var(--text-muted,#6b7280);">App Reg: <strong>${escapeHtml(integ.app_registration_status)}</strong></div>` : ''}
        ${integ.lighthouse_status ? `<div style="font-size:.78rem;color:var(--text-muted,#6b7280);">Lighthouse: <strong>${escapeHtml(integ.lighthouse_status)}</strong></div>` : ''}
        ${integ.last_validated_at ? `<div style="font-size:.75rem;color:var(--text-muted,#6b7280);margin-top:.35rem;">Gevalideerd: ${formatDate(integ.last_validated_at)}</div>` : ''}
      </div>`;
    }).join('');
  } catch (e) {
    grid.innerHTML = `<p style="color:#dc2626;font-size:.875rem;grid-column:1/-1;">Fout bij laden integraties: ${escapeHtml(String(e))}</p>`;
  }
}
window.loadIntegratieStatusGrid = loadIntegratieStatusGrid;

// ── Gebruikers & Rollen Tab ───────────────────────────────────────────────────
let _rolesData = { users: [], roles: [] };

async function loadRolesTab() {
  const tbody = document.getElementById('rolesUsersTableBody');
  const pillsEl = document.getElementById('rolesRolePills');
  if (tbody) tbody.innerHTML = '<tr><td colspan="7" style="text-align:center;color:var(--text-muted,#6b7280);padding:1.5rem;">Laden…</td></tr>';
  try {
    const [usersResp, rolesResp] = await Promise.all([
      apiFetch('/api/users'),
      apiFetch('/api/portal-roles'),
    ]);
    _rolesData.users = (usersResp && usersResp.items) || [];
    _rolesData.roles = (rolesResp && rolesResp.items) || [];

    // Rol-pillen tonen
    if (pillsEl) {
      pillsEl.innerHTML = _rolesData.roles.map((r) =>
        `<span style="padding:.25rem .75rem;border-radius:999px;background:#eff6ff;color:#1d4ed8;font-size:.78rem;font-weight:600;" title="${escapeHtml(r.description || '')}">${escapeHtml(r.label || r.role_key)}</span>`
      ).join('');
    }

    _renderRolesTable(_rolesData.users);
  } catch (e) {
    if (tbody) tbody.innerHTML = `<tr><td colspan="7" style="text-align:center;color:#dc2626;padding:1.5rem;">Fout: ${escapeHtml(String(e))}</td></tr>`;
  }
}

function _renderRolesTable(users) {
  const tbody = document.getElementById('rolesUsersTableBody');
  if (!tbody) return;
  if (!users.length) {
    tbody.innerHTML = '<tr><td colspan="7" style="text-align:center;color:var(--text-muted,#6b7280);padding:1.5rem;">Geen gebruikers gevonden.</td></tr>';
    return;
  }
  tbody.innerHTML = users.map((u) => `
    <tr>
      <td>${escapeHtml(u.display_name || '-')}</td>
      <td style="font-size:.82rem;">${escapeHtml(u.email || '-')}</td>
      <td>
        <select onchange="_updateUserRole('${escapeHtml(u.id)}', this.value)" style="padding:.2rem .5rem;border:1px solid var(--border-color,#d1d5db);border-radius:4px;font-size:.8rem;background:var(--input-bg,#fff);">
          <option value="admin" ${u.role === 'admin' ? 'selected' : ''}>admin</option>
          <option value="klant" ${u.role === 'klant' ? 'selected' : ''}>klant</option>
        </select>
      </td>
      <td style="font-size:.78rem;color:var(--text-muted,#6b7280);">${escapeHtml(u.linked_tenant_id || '-')}</td>
      <td>
        <span style="padding:.2rem .5rem;border-radius:999px;font-size:.75rem;background:${u.is_active ? '#dcfce7' : '#fee2e2'};color:${u.is_active ? '#166534' : '#991b1b'};">${u.is_active ? 'Actief' : 'Inactief'}</span>
      </td>
      <td style="font-size:.78rem;">${formatDate(u.last_login_at || u.created_at)}</td>
      <td>
        <button type="button" class="btn btn-secondary" style="font-size:.75rem;padding:.2rem .5rem;" onclick="_toggleUserActive('${escapeHtml(u.id)}', ${u.is_active})">
          ${u.is_active ? 'Deactiveren' : 'Activeren'}
        </button>
      </td>
    </tr>
  `).join('');
}

async function _updateUserRole(userId, newRole) {
  try {
    const resp = await apiFetch(`/api/users/${userId}`, {
      method: 'PATCH',
      body: JSON.stringify({ role: newRole }),
    });
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    showToast('Rol bijgewerkt.', 'success');
  } catch (e) {
    showToast(`Fout bij bijwerken rol: ${e}`, 'error');
    loadRolesTab(); // reset
  }
}

async function _toggleUserActive(userId, currentlyActive) {
  const action = currentlyActive ? 'deactiveren' : 'activeren';
  if (!confirm(`Gebruiker ${action}?`)) return;
  try {
    const resp = await apiFetch(`/api/users/${userId}`, {
      method: 'PATCH',
      body: JSON.stringify({ is_active: currentlyActive ? 0 : 1 }),
    });
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    showToast(`Gebruiker ${action === 'deactiveren' ? 'gedeactiveerd' : 'geactiveerd'}.`, 'success');
    loadRolesTab();
  } catch (e) {
    showToast(`Fout: ${e}`, 'error');
  }
}
window.loadRolesTab = loadRolesTab;
window._updateUserRole = _updateUserRole;
window._toggleUserActive = _toggleUserActive;

// ── Hub secties lader ─────────────────────────────────────────────────────────
async function loadHubSection(hubId) {
  const tid = currentTenantId;

  if (hubId === 'assessment') {
    // Laad laatste run stats
    const metaEl = document.getElementById('assessmentHubMeta');
    const lastEl = document.getElementById('hubStatAssessmentLast');
    const scoreEl = document.getElementById('hubStatAssessmentScore');
    const actionsEl = document.getElementById('hubStatAssessmentActions');
    if (tid && metaEl) {
      try {
        const r = await apiFetch(`/api/tenants/${tid}/runs?limit=1`);
        const last = (r && r.items || [])[0];
        if (last) {
          if (lastEl) lastEl.textContent = `Laatste: ${formatDate(last.completed_at || last.created_at)}`;
          if (scoreEl) scoreEl.textContent = last.score_overall != null ? `Score: ${last.score_overall}%` : '—';
          if (actionsEl) actionsEl.textContent = last.critical_count != null ? `${last.critical_count} kritiek` : '—';
          metaEl.innerHTML = `
            <span style="padding:.25rem .7rem;border-radius:999px;background:#eff6ff;color:#1d4ed8;font-size:.8rem;font-weight:600;">Laatste run: ${formatDate(last.completed_at || last.created_at)}</span>
            ${last.score_overall != null ? `<span style="padding:.25rem .7rem;border-radius:999px;background:${last.score_overall >= 70 ? '#dcfce7' : '#fee2e2'};color:${last.score_overall >= 70 ? '#166534' : '#991b1b'};font-size:.8rem;font-weight:600;">Score: ${last.score_overall}%</span>` : ''}
            ${last.critical_count ? `<span style="padding:.25rem .7rem;border-radius:999px;background:#fee2e2;color:#991b1b;font-size:.8rem;font-weight:600;">${last.critical_count} kritiek</span>` : ''}
          `;
        }
      } catch (_) {}
    }
    return;
  }

  if (!tid) return;

  // Generieke snapshot stats ophalen voor hub tiles
  const fetchStat = async (section, subsection, statId, formatter) => {
    const el = document.getElementById(statId);
    if (!el) return;
    try {
      const data = await apiFetch(`/api/tenants/${tid}/snapshots/${section}/${subsection}`);
      if (!data) { el.textContent = '—'; return; }
      el.textContent = formatter(data);
    } catch (_) { el.textContent = '—'; }
  };

  if (hubId === 'gbid') {
    fetchStat('gebruikers', 'users', 'hubStatUsers', (d) => {
      const cnt = d?.data?.TotalUsers ?? d?.TotalUsers ?? '—';
      return cnt !== '—' ? `${cnt} gebruikers` : '—';
    });
    fetchStat('gebruikers', 'licenses', 'hubStatLicenses', (d) => {
      const cnt = (d?.data?.Licenses ?? d?.Licenses ?? []).length;
      return cnt ? `${cnt} licenties` : '—';
    });
    fetchStat('identity', 'mfa', 'hubStatMfa', (d) => {
      const pct = d?.data?.MfaRegisteredPct ?? d?.MfaRegisteredPct;
      return pct != null ? `${pct}% gedekt` : '—';
    });
  }

  if (hubId === 'security') {
    fetchStat('alerts', 'secure-score', 'hubStatSecureScore', (d) => {
      const score = d?.data?.CurrentScore ?? d?.CurrentScore;
      const max = d?.data?.MaxScore ?? d?.MaxScore;
      return score != null ? `${score}${max ? '/' + max : ''} pts` : '—';
    });
    fetchStat('alerts', 'audit-logs', 'hubStatAudit', (d) => {
      const cnt = (d?.data?.Events ?? d?.Events ?? []).length;
      return cnt ? `${cnt} events` : '—';
    });
    fetchStat('apps', 'registrations', 'hubStatApps', (d) => {
      const cnt = (d?.data?.Apps ?? d?.Apps ?? []).length;
      return cnt ? `${cnt} apps` : '—';
    });
  }

  if (hubId === 'collab') {
    fetchStat('exchange', 'mailboxes', 'hubStatExchange', (d) => {
      const cnt = (d?.data?.Mailboxes ?? d?.Mailboxes ?? []).length;
      return cnt ? `${cnt} mailboxen` : '—';
    });
    fetchStat('teams', 'teams', 'hubStatTeams', (d) => {
      const cnt = (d?.data?.Teams ?? d?.Teams ?? []).length;
      return cnt ? `${cnt} teams` : '—';
    });
    fetchStat('sharepoint', 'sharepoint-sites', 'hubStatSharePoint', (d) => {
      const cnt = (d?.data?.Sites ?? d?.Sites ?? []).length;
      return cnt ? `${cnt} sites` : '—';
    });
    fetchStat('domains', 'domains', 'hubStatDomains', (d) => {
      const cnt = (d?.data?.Domains ?? d?.Domains ?? []).length;
      return cnt ? `${cnt} domeinen` : '—';
    });
  }

  if (hubId === 'devices') {
    fetchStat('intune', 'summary', 'hubStatIntuneOvz', (d) => {
      const total = d?.data?.TotalDevices ?? d?.TotalDevices;
      return total != null ? `${total} devices` : '—';
    });
    fetchStat('intune', 'devices', 'hubStatDevices', (d) => {
      const cnt = (d?.data?.Devices ?? d?.Devices ?? []).length;
      return cnt ? `${cnt} apparaten` : '—';
    });
    fetchStat('intune', 'compliance', 'hubStatCompliance', (d) => {
      const ok = (d?.data?.CompliantDevices ?? d?.CompliantDevices);
      return ok != null ? `${ok} compliant` : '—';
    });
  }
}
window.loadHubSection = loadHubSection;
