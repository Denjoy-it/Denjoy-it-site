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
  let saved = 'light';
  try { saved = localStorage.getItem('m365LocalTheme') || 'light'; } catch (_) {}
  applyTheme(saved);
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
      window.location.href = '/login.html';
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

// ── Sub-nav configuratie per sectie ──────────────────────────────────────────
const SUBNAV_CONFIG = {
  overview: [],
  assessment: [
    { label: 'Assessment starten', section: 'assessment' },
    { label: 'Rapporten',          section: 'results' },
  ],
  entrafalcon: [
    { label: 'Scan starten', section: 'entrafalcon' },
    { label: 'Rapporten',    section: 'results', resultsPanel: 'entrafalcon' },
  ],
  results: [
    { label: 'Rapport',      resultsPanel: 'viewer' },
    { label: 'Vergelijking', resultsPanel: 'diff' },
    { label: 'Beheer',       resultsPanel: 'management' },
    { label: 'Acties',       resultsPanel: 'actions' },
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

let _currentSection = 'overview';
let _currentSubItem = null;

function updateSubnav(sectionName, activeItem) {
  const subnav = document.getElementById('portalSubnav');
  if (!subnav) return;
  const items = SUBNAV_CONFIG[sectionName] || [];
  if (!items.length) {
    subnav.style.display = 'none';
    return;
  }
  subnav.style.display = '';
  subnav.innerHTML = items.map((item) => {
    const count = item.countId ? (document.getElementById(item.countId)?.textContent || '') : '';
    const showCount = count && count !== '—' && count !== '';
    const countBadge = showCount ? `<span class="subnav-count">${escapeHtml(count)}</span>` : '';
    const key = item.kbTab || item.settingsTab || item.resultsPanel || item.section || '';
    const type = item.kbTab ? 'kb' : item.settingsTab ? 'settings' : item.resultsPanel ? 'results' : 'section';
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
      } else if (type === 'section') {
        // item kan ook een resultsPanel target hebben (bijv. entrafalcon subnav → rapporten tab)
        const item = (SUBNAV_CONFIG[_currentSection] || []).find(
          (i) => (i.kbTab || i.settingsTab || i.resultsPanel || i.section || '') === key
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
  setActiveSubnavItem(panelName);
  // Laad panel-specifieke data lazy
  if (panelName === 'diff')         loadRunDiffPanel();
  if (panelName === 'management')   loadReportsManagementPanel();
  if (panelName === 'actions')      loadActionsPanel();
  if (panelName === 'entrafalcon' && typeof loadEntraFalconResultsPane === 'function') loadEntraFalconResultsPane();
}
window.showResultsPanel = showResultsPanel;

function setActiveNav(sectionName) {
  const navSection = sectionName === 'history' ? 'results' : sectionName;
  document.querySelectorAll('.portal-nav-link[data-section]').forEach((item) => {
    item.classList.toggle('active', item.dataset.section === navSection);
  });
}

function showSection(sectionName, opts = {}) {
  // history is een alias voor results (gecombineerde Rapporten pagina)
  if (sectionName === 'history') sectionName = 'results';
  _currentSection = sectionName;

  document.querySelectorAll('.content-section').forEach((s) => s.classList.remove('active'));
  const el = document.getElementById(`${sectionName}Section`);
  if (el) el.classList.add('active');
  setActiveNav(sectionName);

  if (sectionName === 'assessment' && typeof loadAssessmentExperience === 'function') {
    loadAssessmentExperience();
  }
  if (sectionName === 'entrafalcon' && typeof loadEntraFalconSection === 'function') {
    loadEntraFalconSection();
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

  updateTenantPill(tenants, currentTenantId);
  updateHeroVisibility();

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
  localStorage.setItem('local_m365_current_tenant', tenantId);
  const select = document.getElementById('tenantSelect');
  if (select) select.value = tenantId;
  updateTenantPill(allTenants, tenantId);
  updateHeroVisibility();
  await populateSettings();
  await refreshTenantData();
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
    return;
  }

  const stats = await apiFetch(`/api/tenants/${currentTenantId}/overview`);
  if (!stats.hasData) {
    list.innerHTML = '<p class="empty-state">Nog geen assessments uitgevoerd</p>';
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
  list.innerHTML = items.map((r) => `
    <div class="assessment-item">
      <div><strong>${formatDate(r.completed_at || r.started_at)}</strong> - ${statusBadge(r.status)}</div>
      <div style="margin-top:6px;font-size:0.9em;color:#666;">${(r.phases || []).join(', ') || 'alle fases'}</div>
      <div style="margin-top:8px;display:flex;gap:8px;flex-wrap:wrap;">
        <button class="btn btn-secondary btn-sm" data-action="viewRun" data-id="${escapeHtml(r.id)}">Details</button>
        ${r.report_path ? `<button class="btn btn-secondary btn-sm" data-action="openUrl" data-id="${escapeHtml(r.report_path)}">Rapport</button>` : ''}
      </div>
    </div>`).join('');
  bindActions(list);
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
  const downloadBar = document.getElementById('resultsDownloadBar');

  function showEmpty(msg) {
    if (emptyEl)   { emptyEl.style.display = ''; emptyEl.textContent = msg; }
    if (contentEl) contentEl.style.display = 'none';
    if (downloadBar) downloadBar.style.display = 'none';
  }

  if (!currentTenantId) { showEmpty('Geen tenant geselecteerd.'); return; }

  const runs = await apiFetch(`/api/tenants/${currentTenantId}/runs`);
  const allRuns = runs.items || [];
  const latest  = allRuns.find((r) => !!r.report_path);

  if (!latest) { showEmpty('Start een assessment om resultaten te zien.'); return; }

  // Toon content
  if (emptyEl)   emptyEl.style.display = 'none';
  if (contentEl) contentEl.style.display = '';

  const reportRuns    = allRuns.filter((r) => !!r.report_path).slice(0, 8);
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
  const tbody = document.getElementById('resultsRunsTableBody');
  if (tbody) {
    tbody.innerHTML = reportRuns.map((r) => `
      <tr>
        <td>${formatDate(r.completed_at || r.started_at)}</td>
        <td>${statusBadge(r.status)}</td>
        <td>${r.run_mode || '—'}</td>
        <td>${formatPhaseList(r.phases)}</td>
        <td>${r.score_overall ?? '—'}</td>
        <td>${r.critical_count ?? 0} / ${r.warning_count ?? 0} / ${r.info_count ?? 0}</td>
        <td>
          <div class="results-row-actions">
            <button class="btn btn-secondary btn-sm" data-action="viewRun" data-id="${escapeHtml(r.id)}">Details</button>
            <button class="btn btn-secondary btn-sm" data-action="openUrl" data-id="${escapeHtml(r.report_path)}">Rapport</button>
          </div>
        </td>
      </tr>
    `).join('') || '<tr><td colspan="7" class="empty-state">Geen rapport-runs.</td></tr>';
    bindActions(tbody.closest('section') || document.getElementById('resultsSection'));
  }

  // ── Download bar ──
  if (downloadBar && latest.report_path) {
    downloadBar.style.display = 'flex';
    const btnReport = document.getElementById('btnOpenReport');
    const btnCsv    = document.getElementById('btnDownloadCsv');
    if (btnReport) btnReport.href = latest.report_path;
    if (btnCsv) {
      const csvDir = latest.report_path.replace(/\/[^/]+\.html$/, '');
      btnCsv.href  = csvDir + '/_Assessment-Summary.csv';
      btnCsv.title = 'Download _Assessment-Summary.csv — overige CSV-bestanden staan in dezelfde map';
    }
  }

  // ── Rapportviewer ──
  if (typeof initResultsViewer === 'function') {
    initResultsViewer('resultsViewerContainer', latest.report_path);
  }
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
  if (document.getElementById('entrafalconSection')?.classList.contains('active') && typeof loadEntraFalconSection === 'function') {
    loadEntraFalconSection();
  }
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
  document.querySelectorAll('.portal-nav-link[data-section]').forEach((item) => {
    item.addEventListener('click', (e) => {
      e.preventDefault();
      showSection(item.dataset.section);
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
  localStorage.setItem('local_m365_current_tenant', tenantId);
  const select = document.getElementById('tenantSelect');
  if (select) select.value = tenantId;

  updateTenantPill(allTenants, tenantId);
  updateHeroVisibility();
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

async function bootstrap() {
  try {
    const displayName = 'Lokaal';
    const userNameEl = document.getElementById('userName');
    if (userNameEl) userNameEl.textContent = displayName;
    const initialsEl = document.getElementById('userInitials');
    if (initialsEl) initialsEl.textContent = getInitials(displayName);
    const avatarBtn = document.getElementById('userAvatarBtn');
    if (avatarBtn) avatarBtn.title = displayName;
    initThemeControls();
    setupNavigation();
    setupHeaderActions();
    setupSettingsActions();
    switchSettingsTab('tenant');
    updateSubnav('overview'); // verborgen op startpagina
    await loadTenants();
    await populateSettings();
    await refreshTenantData();
  } catch (e) {
    console.error(e);
    alert(`Dashboard initialisatie mislukt: ${e.message}`);
  }
}

document.addEventListener('DOMContentLoaded', () => {
  bootstrap();
  // Wire up Rapporten sidebar tabbar clicks
  document.querySelectorAll('#resultsTabbar [data-results-panel]').forEach((tab) => {
    tab.addEventListener('click', (e) => {
      e.preventDefault();
      showResultsPanel(tab.dataset.resultsPanel);
    });
  });
});
