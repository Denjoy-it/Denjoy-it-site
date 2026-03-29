/**
 * Denjoy IT Platform — Intune & Device Management module (Fase 4)
 * Tabs: Overzicht | Apparaten | Compliance | Configuratie | Geschiedenis
 */
(function () {
  'use strict';

  // ── State ────────────────────────────────────────────────────────────────

  let _devices      = [];
  let _compliance   = [];
  let _config       = [];
  let _summaryData  = null;
  let _filterOs     = 'all';
  let _filterState  = 'all';
  let _searchQ      = '';
  let _loading      = false;
  let _tabsBound    = false;

  // ── API helper ───────────────────────────────────────────────────────────

  function itApiFetch(url, opts = {}) {
    const token   = localStorage.getItem('denjoy_auth_token') || localStorage.getItem('denjoy_token') || '';
    const headers = Object.assign({ 'Content-Type': 'application/json' }, opts.headers || {});
    if (token) headers['Authorization'] = `Bearer ${token}`;
    return fetch(url, Object.assign({}, opts, { headers })).then((r) => {
      if (!r.ok) return r.json().then((e) => Promise.reject(e.error || r.statusText));
      return r.json();
    });
  }

  function getTenantId() {
    if (typeof window.currentTenantId !== 'undefined') return window.currentTenantId;
    const sel = document.getElementById('tenantSelect');
    return sel ? sel.value : null;
  }

  // ── Helpers ──────────────────────────────────────────────────────────────

  function esc(s) {
    if (!s) return '';
    return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  }

  function fmtDate(iso) {
    if (!iso) return '—';
    try { return new Date(iso).toLocaleDateString('nl-NL'); } catch { return iso; }
  }

  function fmtDateTime(iso) {
    if (!iso) return '—';
    try { return new Date(iso).toLocaleString('nl-NL'); } catch { return iso; }
  }

  function osIcon(os) {
    const icons = { Windows: '🖥', iOS: '📱', Android: '📱', macOS: '💻', linux: '🐧' };
    return icons[os] || '💻';
  }

  function complianceBadge(state) {
    const map = {
      compliant:    { cls: 'it-compliance-compliant',     label: 'Compliant' },
      noncompliant: { cls: 'it-compliance-noncompliant',  label: 'Non-compliant' },
      inGracePeriod:{ cls: 'it-compliance-inGracePeriod', label: 'Grace period' },
      unknown:      { cls: 'it-compliance-unknown',       label: 'Onbekend' },
    };
    const d = map[state] || map.unknown;
    return `<span class="it-compliance ${d.cls}"><span class="it-dot"></span>${d.label}</span>`;
  }

  function renderWorkspaceSource(data) {
    const wrap = document.getElementById('itWorkspaceSource');
    const describe = window.denjoyDescribeSourceMeta;
    if (!wrap || typeof describe !== 'function' || !data) return;
    const info = describe(data);
    wrap.innerHTML = `
      <div class="live-module-source">
        <span class="live-module-source-pill ${esc(info.className || '')}">${esc(info.label)}</span>
        <span>${esc(info.detail)}</span>
      </div>`;
  }

  function renderIntuneOverview() {
    const wrap = document.getElementById('itServiceOverview');
    if (!wrap) return;
    const score = _summaryData ? `${Number(_summaryData.score || 0)}%` : '—';
    const total = _summaryData ? Number(_summaryData.total || _devices.length || 0) : (_devices.length ? _devices.length : '—');
    const compliant = _summaryData ? Number(_summaryData.compliantCount || 0) : '—';
    const policies = _compliance.length ? _compliance.length : '—';
    wrap.innerHTML = `
      <div class="workspace-service-overview">
        <article class="workspace-service-card"><span class="workspace-service-label">Apparaten</span><strong class="workspace-service-value">${total}</strong><span class="workspace-service-meta">managed</span></article>
        <article class="workspace-service-card"><span class="workspace-service-label">Compliant</span><strong class="workspace-service-value">${compliant}</strong><span class="workspace-service-meta">apparaten</span></article>
        <article class="workspace-service-card"><span class="workspace-service-label">Policies</span><strong class="workspace-service-value">${policies}</strong><span class="workspace-service-meta">compliance</span></article>
        <article class="workspace-service-card"><span class="workspace-service-label">Score</span><strong class="workspace-service-value">${score}</strong><span class="workspace-service-meta">tenant posture</span></article>
      </div>`;
  }

  // ── Tab switching ─────────────────────────────────────────────────────────

  function switchItTab(tab) {
    document.querySelectorAll('.it-tab').forEach((btn) => {
      btn.classList.toggle('active', btn.dataset.itTab === tab);
    });
    document.querySelectorAll('.it-tab-panel').forEach((panel) => {
      panel.style.display = panel.dataset.itPanel === tab ? '' : 'none';
    });
    if (tab === 'overzicht')   loadSummary();
    if (tab === 'apparaten')   { if (!_devices.length) loadDevices(); else renderDevicesTable(); }
    if (tab === 'compliance')  { if (!_compliance.length) loadCompliancePolicies(); }
    if (tab === 'configuratie'){ if (!_config.length) loadConfigProfiles(); }
    if (tab === 'geschiedenis') loadIntuneHistory();
  }

  function bindItTabs() {
    if (_tabsBound) return;
    _tabsBound = true;
    document.querySelectorAll('.it-tab[data-it-tab]').forEach((btn) => {
      btn.addEventListener('click', () => switchItTab(btn.dataset.itTab));
    });

    const search = document.getElementById('itSearchInput');
    if (search) search.addEventListener('input', () => { _searchQ = search.value.trim(); renderDevicesTable(); });

    document.querySelectorAll('.it-filter-tab[data-filter-os]').forEach((btn) => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('.it-filter-tab[data-filter-os]').forEach((b) => b.classList.remove('active'));
        btn.classList.add('active');
        _filterOs = btn.dataset.filterOs;
        renderDevicesTable();
      });
    });
    document.querySelectorAll('.it-filter-tab[data-filter-state]').forEach((btn) => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('.it-filter-tab[data-filter-state]').forEach((b) => b.classList.remove('active'));
        btn.classList.add('active');
        _filterState = btn.dataset.filterState;
        renderDevicesTable();
      });
    });

    const btnRefresh = document.getElementById('itBtnRefresh');
    if (btnRefresh) btnRefresh.addEventListener('click', () => { _devices = []; loadDevices(); });

    const btnRefreshCompliance = document.getElementById('itBtnRefreshCompliance');
    if (btnRefreshCompliance) btnRefreshCompliance.addEventListener('click', () => { _compliance = []; loadCompliancePolicies(); });

    const btnRefreshConfig = document.getElementById('itBtnRefreshConfig');
    if (btnRefreshConfig) btnRefreshConfig.addEventListener('click', () => { _config = []; loadConfigProfiles(); });
  }

  // ── Overzicht / Summary ──────────────────────────────────────────────────

  function loadSummary() {
    const tid = getTenantId();
    if (!tid) { renderSummaryEmpty('Selecteer een tenant.'); return; }

    const wrap = document.getElementById('itSummaryWrap');
    if (wrap) wrap.innerHTML = '<p class="it-empty">Overzicht laden...</p>';

    itApiFetch(`/api/intune/${tid}/summary`)
      .then((data) => renderSummary(data))
      .catch((err) => { if (wrap) wrap.innerHTML = `<p class="it-empty">Fout: ${esc(String(err))}</p>`; });
  }

  function renderSummaryEmpty(msg) {
    const wrap = document.getElementById('itSummaryWrap');
    if (wrap) wrap.innerHTML = `<p class="it-empty">${esc(msg)}</p>`;
  }

  function renderSummary(data) {
    const wrap = document.getElementById('itSummaryWrap');
    if (!wrap) return;
    _summaryData = data;
    renderWorkspaceSource(data);
    renderIntuneOverview();

    const score    = data.score || 0;
    const total    = data.total || 0;
    const comp     = data.compliantCount || 0;
    const nonComp  = total - comp;
    const byOs     = data.byOs || {};

    // Score ring SVG
    const r = 36; const circ = 2 * Math.PI * r;
    const offset = circ - (score / 100) * circ;
    const ringCls = score >= 80 ? 'good' : score >= 50 ? 'warn' : 'bad';

    // OS breakdown
    const osRows = Object.entries(byOs).map(([os, s]) => {
      const pct = s.total > 0 ? Math.round((s.compliant / s.total) * 100) : 0;
      const barCls = pct >= 80 ? 'green' : pct < 50 ? 'red' : '';
      return `<div class="it-os-row">
        <span class="it-os-label">${esc(os)} <small style="color:var(--text-muted)">(${s.total})</small></span>
        <div class="it-os-bar-wrap"><div class="it-os-bar ${barCls}" style="width:${pct}%"></div></div>
        <span class="it-os-count">${pct}%</span>
      </div>`;
    }).join('');

    wrap.innerHTML = `
      <div class="it-summary-grid">
        <div class="it-summary-card">
          <div class="it-summary-label">Apparaten</div>
          <div class="it-summary-value">${total}</div>
        </div>
        <div class="it-summary-card">
          <div class="it-summary-label">Compliant</div>
          <div class="it-summary-value compliant">${comp}</div>
        </div>
        <div class="it-summary-card">
          <div class="it-summary-label">Non-compliant</div>
          <div class="it-summary-value noncompliant">${nonComp}</div>
        </div>
        <div class="it-summary-card">
          <div class="it-summary-label">Compliance score</div>
          <div class="it-summary-value score">${score}%</div>
        </div>
      </div>

      <div style="display:flex;gap:1.5rem;flex-wrap:wrap;align-items:flex-start">
        <div class="it-score-wrap">
          <svg class="it-score-ring" width="90" height="90" viewBox="0 0 90 90">
            <circle class="it-score-ring-bg" cx="45" cy="45" r="${r}" stroke-width="8"/>
            <circle class="it-score-ring-fg ${ringCls}" cx="45" cy="45" r="${r}" stroke-width="8"
              stroke-dasharray="${circ}" stroke-dashoffset="${offset}"
              transform="rotate(-90 45 45)"/>
          </svg>
          <div class="it-score-info">
            <div class="it-score-pct">${score}%</div>
            <div class="it-score-label">Compliance score</div>
            <div style="font-size:.78rem;color:var(--text-muted);margin-top:.25rem">${comp} van ${total} compliant</div>
          </div>
        </div>

        ${osRows ? `<div style="flex:1;min-width:200px">
          <div style="font-size:.78rem;font-weight:600;text-transform:uppercase;letter-spacing:.05em;color:var(--text-muted);margin-bottom:.75rem">Per besturingssysteem</div>
          <div class="it-os-list">${osRows}</div>
        </div>` : ''}
      </div>`;
  }

  // ── Apparaten ─────────────────────────────────────────────────────────────

  function loadDevices() {
    const tid = getTenantId();
    if (!tid) { showDevicesEmpty('Selecteer een tenant.'); return; }
    if (_loading) return;
    _loading = true;
    showDevicesLoading();

    itApiFetch(`/api/intune/${tid}/devices`)
      .then((data) => {
        _devices = data.devices || [];
        renderWorkspaceSource(data);
        renderIntuneOverview();
        renderDevicesTable();
        const counter = document.getElementById('itDeviceCount');
        if (counter) counter.textContent = `${_devices.length} apparaat${_devices.length !== 1 ? 'en' : ''}`;
      })
      .catch((err) => showDevicesEmpty(`Fout: ${esc(String(err))}`))
      .finally(() => { _loading = false; });
  }

  function showDevicesLoading() {
    const tbody = document.getElementById('itDeviceTableBody');
    if (tbody) tbody.innerHTML = `<tr><td colspan="6" class="it-loading">Apparaten laden...</td></tr>`;
  }

  function showDevicesEmpty(msg) {
    const tbody = document.getElementById('itDeviceTableBody');
    if (tbody) tbody.innerHTML = `<tr><td colspan="6" class="it-table-empty">${esc(msg)}</td></tr>`;
  }

  function renderDevicesTable() {
    const tbody = document.getElementById('itDeviceTableBody');
    if (!tbody) return;

    let filtered = _devices;
    if (_filterOs !== 'all')    filtered = filtered.filter((d) => (d.operatingSystem || '').toLowerCase() === _filterOs.toLowerCase());
    if (_filterState !== 'all') filtered = filtered.filter((d) => (d.complianceState || 'unknown') === _filterState);
    if (_searchQ) {
      const q = _searchQ.toLowerCase();
      filtered = filtered.filter((d) =>
        (d.deviceName || '').toLowerCase().includes(q) ||
        (d.userPrincipalName || '').toLowerCase().includes(q) ||
        (d.userDisplayName || '').toLowerCase().includes(q) ||
        (d.model || '').toLowerCase().includes(q)
      );
    }

    if (!filtered.length) {
      tbody.innerHTML = `<tr><td colspan="6" class="it-table-empty">${
        _searchQ || _filterOs !== 'all' || _filterState !== 'all'
          ? 'Geen apparaten gevonden voor deze filter.'
          : 'Geen apparaten geladen. Klik op Verversen.'
      }</td></tr>`;
      return;
    }

    tbody.innerHTML = filtered.map((d) => `
      <tr data-devid="${esc(d.id)}">
        <td>
          <div class="it-device-cell">
            <div class="it-device-icon">${osIcon(d.operatingSystem)}</div>
            <div>
              <div class="it-device-name">${esc(d.deviceName || '—')}</div>
              <div class="it-device-model">${esc(d.manufacturer || '')} ${esc(d.model || '')}</div>
            </div>
          </div>
        </td>
        <td>${esc(d.operatingSystem || '—')} ${esc(d.osVersion || '')}</td>
        <td>${esc(d.userDisplayName || d.userPrincipalName || '—')}</td>
        <td>${complianceBadge(d.complianceState || 'unknown')}</td>
        <td>${fmtDate(d.lastSyncDateTime)}</td>
        <td>
          <div class="it-row-actions">
            <button class="it-btn it-btn-ghost it-btn-detail" data-devid="${esc(d.id)}" data-devname="${esc(d.deviceName)}">Detail</button>
          </div>
        </td>
      </tr>`).join('');

    tbody.querySelectorAll('tr[data-devid]').forEach((row) => {
      row.addEventListener('click', (e) => {
        if (e.target.closest('.it-row-actions')) return;
        openDeviceDetail(row.dataset.devid, row.querySelector('.it-device-name')?.textContent || '');
      });
    });
    tbody.querySelectorAll('.it-btn-detail').forEach((btn) => {
      btn.addEventListener('click', (e) => { e.stopPropagation(); openDeviceDetail(btn.dataset.devid, btn.dataset.devname); });
    });
  }

  // ── Device detail modal ──────────────────────────────────────────────────

  function openDeviceDetail(deviceId, fallbackName) {
    const tid = getTenantId();
    if (!tid) return;

    // Open het Inzichten-paneel
    if (typeof window.openSideRailDetail === 'function') {
      window.openSideRailDetail('Apparaat', fallbackName || deviceId);
    }

    itApiFetch(`/api/intune/${tid}/devices/${deviceId}`)
      .then((data) => {
        const d = data.device || {};
        const storage = d.totalStorageSpaceInBytes
          ? `${Math.round(d.freeStorageSpaceInBytes / 1024 / 1024 / 1024)} GB vrij / ${Math.round(d.totalStorageSpaceInBytes / 1024 / 1024 / 1024)} GB totaal`
          : '—';
        const ram = d.physicalMemoryInBytes
          ? `${Math.round(d.physicalMemoryInBytes / 1024 / 1024 / 1024)} GB`
          : '—';
        const policyRows = (d.compliancePolicies || []).map((p) =>
          `<tr><td>${esc(p.displayName)}</td><td>${complianceBadge(p.state)}</td></tr>`).join('');
        const configRows = (d.configProfiles || []).slice(0, 10).map((c) =>
          `<tr><td>${esc(c.displayName)}</td><td>${complianceBadge(c.state)}</td></tr>`).join('');
        const bodyHtml = `
          <div class="it-detail-grid">
            <div class="it-detail-item"><div class="it-detail-key">Naam</div><div class="it-detail-val">${esc(d.deviceName || '—')}</div></div>
            <div class="it-detail-item"><div class="it-detail-key">Status</div><div class="it-detail-val">${complianceBadge(d.complianceState || 'unknown')}</div></div>
            <div class="it-detail-item"><div class="it-detail-key">OS</div><div class="it-detail-val">${esc(d.operatingSystem || '—')} ${esc(d.osVersion || '')}</div></div>
            <div class="it-detail-item"><div class="it-detail-key">Fabrikant / Model</div><div class="it-detail-val">${esc(d.manufacturer || '—')} ${esc(d.model || '')}</div></div>
            <div class="it-detail-item"><div class="it-detail-key">Serienummer</div><div class="it-detail-val">${esc(d.serialNumber || '—')}</div></div>
            <div class="it-detail-item"><div class="it-detail-key">Eigenaar</div><div class="it-detail-val">${esc(d.ownerType || '—')}</div></div>
            <div class="it-detail-item"><div class="it-detail-key">Gebruiker</div><div class="it-detail-val">${esc(d.userDisplayName || d.userPrincipalName || '—')}</div></div>
            <div class="it-detail-item"><div class="it-detail-key">Ingeschreven</div><div class="it-detail-val">${fmtDate(d.enrolledDateTime)}</div></div>
            <div class="it-detail-item"><div class="it-detail-key">Laatste sync</div><div class="it-detail-val">${fmtDateTime(d.lastSyncDateTime)}</div></div>
            <div class="it-detail-item"><div class="it-detail-key">Versleuteld</div><div class="it-detail-val">${d.isEncrypted ? '✓ Ja' : '✗ Nee'}</div></div>
            <div class="it-detail-item"><div class="it-detail-key">Opslag</div><div class="it-detail-val">${esc(storage)}</div></div>
            <div class="it-detail-item"><div class="it-detail-key">RAM</div><div class="it-detail-val">${esc(ram)}</div></div>
            <div class="it-detail-item"><div class="it-detail-key">Entra Device ID</div><div class="it-detail-val" style="font-size:.72rem;word-break:break-all">${esc(d.azureADDeviceId || '—')}</div></div>
            <div class="it-detail-item"><div class="it-detail-key">Azure AD geregistreerd</div><div class="it-detail-val">${d.azureADRegistered ? '✓ Ja' : '✗ Nee'}</div></div>
          </div>
          ${policyRows ? `<div class="it-detail-section-title" style="margin-top:1rem">Compliance policies</div><div class="it-table-wrap" style="margin-top:.5rem"><table class="it-table"><thead><tr><th>Policy</th><th>Status</th></tr></thead><tbody>${policyRows}</tbody></table></div>` : ''}
          ${configRows ? `<div class="it-detail-section-title" style="margin-top:1rem">Configuratieprofielen (top 10)</div><div class="it-table-wrap" style="margin-top:.5rem"><table class="it-table"><thead><tr><th>Profiel</th><th>Status</th></tr></thead><tbody>${configRows}</tbody></table></div>` : ''}`;
        if (typeof window.updateSideRailDetail === 'function') {
          window.updateSideRailDetail(d.deviceName || fallbackName || 'Apparaat', bodyHtml);
        }
      })
      .catch((err) => {
        if (typeof window.updateSideRailDetail === 'function') {
          window.updateSideRailDetail('Fout', `<p class="it-empty">Fout: ${esc(String(err))}</p>`);
        }
      });
  }

  // ── Compliance policies ──────────────────────────────────────────────────

  function loadCompliancePolicies() {
    const tid = getTenantId();
    if (!tid) { renderComplianceEmpty('Selecteer een tenant.'); return; }
    const grid = document.getElementById('itComplianceGrid');
    if (grid) grid.innerHTML = '<p class="it-empty">Policies laden...</p>';

    itApiFetch(`/api/intune/${tid}/compliance`)
      .then((data) => {
        _compliance = data.policies || [];
        renderWorkspaceSource(data);
        renderIntuneOverview();
        renderComplianceGrid();
      })
      .catch((err) => renderComplianceEmpty(`Fout: ${esc(String(err))}`));
  }

  function renderComplianceEmpty(msg) {
    const grid = document.getElementById('itComplianceGrid');
    if (grid) grid.innerHTML = `<p class="it-empty">${esc(msg)}</p>`;
  }

  function renderComplianceGrid() {
    const grid = document.getElementById('itComplianceGrid');
    if (!grid) return;
    if (!_compliance.length) { renderComplianceEmpty('Geen compliance policies gevonden.'); return; }

    const counter = document.getElementById('itComplianceCount');
    if (counter) counter.textContent = `${_compliance.length} polic${_compliance.length !== 1 ? 'ies' : 'y'}`;

    grid.innerHTML = _compliance.map((p) => {
      const total    = p.totalDevices || 0;
      const ok       = p.compliantCount || 0;
      const err      = p.nonCompliantCount || 0;
      const pctOk    = total > 0 ? Math.round((ok / total) * 100) : 0;
      const pctErr   = total > 0 ? Math.round((err / total) * 100) : 0;
      const platform = (p.platform || '').replace(/^[a-z]/, (c) => c.toUpperCase()) || 'Onbekend';
      return `<div class="it-policy-card">
        <div class="it-policy-name">${esc(p.displayName)}</div>
        <div class="it-policy-platform">${esc(platform)}</div>
        <div class="it-policy-stats">
          <span class="it-policy-stat-ok">${ok}</span><span class="it-policy-stat-lbl"> compliant</span>
          &nbsp;&nbsp;
          <span class="it-policy-stat-err">${err}</span><span class="it-policy-stat-lbl"> non-compliant</span>
          &nbsp;&nbsp;
          <span class="it-policy-stat-lbl">${total} totaal</span>
        </div>
        <div class="it-policy-bar-wrap">
          <span class="it-policy-bar-ok" style="width:${pctOk}%"></span>
          <span class="it-policy-bar-err" style="width:${pctErr}%"></span>
        </div>
      </div>`;
    }).join('');
  }

  // ── Config profiles ──────────────────────────────────────────────────────

  function loadConfigProfiles() {
    const tid = getTenantId();
    if (!tid) { renderConfigEmpty('Selecteer een tenant.'); return; }
    const grid = document.getElementById('itConfigGrid');
    if (grid) grid.innerHTML = '<p class="it-empty">Profielen laden...</p>';

    itApiFetch(`/api/intune/${tid}/config`)
      .then((data) => {
        _config = data.profiles || [];
        renderWorkspaceSource(data);
        renderIntuneOverview();
        renderConfigGrid();
      })
      .catch((err) => renderConfigEmpty(`Fout: ${esc(String(err))}`));
  }

  function renderConfigEmpty(msg) {
    const grid = document.getElementById('itConfigGrid');
    if (grid) grid.innerHTML = `<p class="it-empty">${esc(msg)}</p>`;
  }

  function renderConfigGrid() {
    const grid = document.getElementById('itConfigGrid');
    if (!grid) return;
    if (!_config.length) { renderConfigEmpty('Geen configuratieprofielen gevonden.'); return; }

    const counter = document.getElementById('itConfigCount');
    if (counter) counter.textContent = `${_config.length} profiel${_config.length !== 1 ? 'en' : ''}`;

    grid.innerHTML = _config.map((c) => {
      const typeChip   = c.type === 'catalog'
        ? '<span class="it-chip catalog">Settings Catalog</span>'
        : '<span class="it-chip">Legacy template</span>';
      const assignChip = c.isAssigned
        ? '<span class="it-chip assigned">Toegewezen</span>'
        : '<span class="it-chip">Niet toegewezen</span>';
      return `<div class="it-config-card">
        <div class="it-config-name">${esc(c.displayName || c.name || '—')}</div>
        <div class="it-config-meta">
          ${typeChip}
          ${c.platforms && c.platforms !== 'legacy' ? `<span class="it-chip">${esc(c.platforms)}</span>` : ''}
          ${c.settingCount ? `<span class="it-chip">${c.settingCount} instellingen</span>` : ''}
          ${assignChip}
        </div>
        ${c.description ? `<div style="font-size:.78rem;color:var(--text-muted);margin-bottom:.5rem">${esc(c.description)}</div>` : ''}
        <div class="it-config-action">
          <button class="it-btn it-btn-ghost it-btn-deploy" data-pid="${esc(c.id)}" data-pname="${esc(c.displayName || c.name || '')}">
            Toewijzen
          </button>
        </div>
      </div>`;
    }).join('');

    grid.querySelectorAll('.it-btn-deploy').forEach((btn) => {
      btn.addEventListener('click', () => openDeployModal(btn.dataset.pid, btn.dataset.pname));
    });
  }

  // ── Deploy modal ─────────────────────────────────────────────────────────

  function openDeployModal(policyId, policyName) {
    const tid = getTenantId();
    if (!tid) return;

    const overlay = document.createElement('div');
    overlay.className = 'it-modal-overlay';
    overlay.addEventListener('click', (e) => { if (e.target === overlay) overlay.remove(); });

    const modal = document.createElement('div');
    modal.className = 'it-modal';
    modal.innerHTML = `
      <div class="it-modal-header">
        <div>
          <div class="it-modal-title">Profiel toewijzen</div>
          <div class="it-modal-subtitle">${esc(policyName)}</div>
        </div>
        <button class="it-modal-close">✕</button>
      </div>
      <div class="it-modal-body">
        <div id="itDeployResult"></div>
        <div class="it-form-group">
          <label class="it-form-label">Groep ID <span style="color:#ef4444">*</span></label>
          <input class="it-form-input" id="itGroupIdInput" placeholder="Azure AD Groep GUID" autocomplete="off">
          <div style="font-size:.75rem;color:var(--text-muted);margin-top:.25rem">
            Vind de Group ID via Entra ID → Groepen → Eigenschappen
          </div>
        </div>
        <div>
          <label style="font-size:.82rem;color:var(--text-secondary)">
            <input type="checkbox" id="itDryRunDeploy" style="accent-color:var(--accent)">
            Dry-run (preview zonder toewijzing)
          </label>
        </div>
      </div>
      <div class="it-modal-footer">
        <button class="it-btn it-btn-secondary it-modal-cancel">Annuleren</button>
        <button class="it-btn it-btn-primary it-btn-confirm">Toewijzen</button>
      </div>`;

    overlay.appendChild(modal);
    document.body.appendChild(overlay);
    modal.querySelector('.it-modal-close').onclick = () => overlay.remove();
    modal.querySelector('.it-modal-cancel').onclick = () => overlay.remove();

    modal.querySelector('.it-btn-confirm').onclick = () => {
      const groupId = modal.querySelector('#itGroupIdInput')?.value.trim();
      if (!groupId) { modal.querySelector('#itGroupIdInput').style.borderColor = '#ef4444'; return; }
      const dryRun = modal.querySelector('#itDryRunDeploy')?.checked ?? false;
      const resultDiv = modal.querySelector('#itDeployResult');
      if (resultDiv) resultDiv.innerHTML = '<p style="font-size:.82rem;color:var(--text-muted)">Bezig...</p>';
      modal.querySelector('.it-btn-confirm').disabled = true;

      itApiFetch(`/api/intune/${tid}/deploy-config`, {
        method: 'POST',
        body: JSON.stringify({ policy_id: policyId, group_id: groupId, dry_run: dryRun }),
      })
        .then((data) => {
          const cls = dryRun ? 'it-result-dryrun' : 'it-result-ok';
          const icon = dryRun ? 'ℹ️' : '✅';
          if (resultDiv) resultDiv.innerHTML = `<div class="it-result ${cls}"><div class="it-result-icon">${icon}</div><div class="it-result-msg">${esc(data.message || 'Toewijzing geslaagd')}</div></div>`;
          modal.querySelector('.it-modal-cancel').textContent = 'Sluiten';
          modal.querySelector('.it-btn-confirm').style.display = 'none';
        })
        .catch((err) => {
          if (resultDiv) resultDiv.innerHTML = `<div class="it-result it-result-error"><div class="it-result-icon">❌</div><div class="it-result-msg">Fout: ${esc(String(err))}</div></div>`;
          modal.querySelector('.it-btn-confirm').disabled = false;
        });
    };
  }

  // ── Geschiedenis ─────────────────────────────────────────────────────────

  function loadIntuneHistory() {
    const tid = getTenantId();
    if (!tid) { renderHistoryEmpty('Selecteer een tenant.'); return; }
    const tbody = document.getElementById('itHistoryBody');
    if (tbody) tbody.innerHTML = `<tr><td colspan="5" class="it-loading">Laden...</td></tr>`;

    itApiFetch(`/api/intune/${tid}/history`)
      .then((data) => {
        const items = data.items || [];
        if (!items.length) { renderHistoryEmpty('Nog geen Intune acties gelogd.'); return; }
        if (!tbody) return;
        tbody.innerHTML = items.map((h) => {
          const statusBadge = h.status === 'success'
            ? `<span class="it-compliance it-compliance-compliant"><span class="it-dot"></span>Geslaagd</span>`
            : h.status === 'dry_run'
              ? `<span class="it-compliance it-compliance-unknown"><span class="it-dot"></span>Dry-run</span>`
              : `<span class="it-compliance it-compliance-noncompliant"><span class="it-dot"></span>Mislukt</span>`;
          const actionLabels = {
            'list-devices': 'Apparaten laden',
            'get-device': 'Device detail',
            'list-compliance': 'Compliance laden',
            'list-config': 'Profielen laden',
            'deploy-config': 'Profiel toewijzen',
            'get-compliance-summary': 'Overzicht laden',
          };
          return `<tr>
            <td>${esc(fmtDateTime(h.executed_at))}</td>
            <td>${esc(actionLabels[h.action] || h.action)}</td>
            <td>${statusBadge}</td>
            <td>${esc(h.executed_by || '—')}</td>
            <td style="max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(h.error_message || '—')}</td>
          </tr>`;
        }).join('');
      })
      .catch((err) => renderHistoryEmpty(`Fout: ${esc(String(err))}`));
  }

  function renderHistoryEmpty(msg) {
    const tbody = document.getElementById('itHistoryBody');
    if (tbody) tbody.innerHTML = `<tr><td colspan="5" class="it-table-empty">${esc(msg)}</td></tr>`;
  }

  // ── Publieke interface ───────────────────────────────────────────────────

  window.loadIntuneSection = function () {
    _tabsBound = false;
    _devices   = [];
    _compliance = [];
    _config    = [];
    _summaryData = null;
    bindItTabs();
    loadSummary();
  };
  window.switchIntuneTab = switchItTab;

})();
