/**
 * Denjoy IT Platform — Fase 5: Backup Module
 * IIFE module — publiek interface: window.loadBackupSection
 */
(function () {
  'use strict';

  // ── State ──────────────────────────────────────────────────────────────────
  let _summary    = null;
  let _sharepoint = null;
  let _onedrive   = null;
  let _exchange   = null;
  let _tabsBound  = false;
  let _loading    = false;

  // ── Helpers ────────────────────────────────────────────────────────────────

  function getTenantId() {
    const sel = document.getElementById('tenantSelect');
    return sel ? sel.value : '';
  }

  function escHtml(s) {
    if (s == null) return '';
    return String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function fmtDate(iso) {
    if (!iso) return '—';
    try { return new Date(iso).toLocaleString('nl-NL'); } catch (_) { return iso; }
  }

  function fmtBytes(bytes) {
    if (!bytes || bytes === 0) return '—';
    const gb = bytes / (1024 * 1024 * 1024);
    if (gb >= 1) return gb.toFixed(2) + ' GB';
    const mb = bytes / (1024 * 1024);
    if (mb >= 1) return mb.toFixed(1) + ' MB';
    return Math.round(bytes / 1024) + ' KB';
  }

  function bkApiFetch(url, opts = {}) {
    const token = localStorage.getItem('denjoy_auth_token') || localStorage.getItem('denjoy_token') || '';
    const headers = { 'Content-Type': 'application/json' };
    if (token) headers['Authorization'] = 'Bearer ' + token;
    return fetch(url, { credentials: 'include', headers, ...opts })
      .then((r) => r.json());
  }

  function statusBadge(status) {
    const cls = status ? status.toLowerCase().replace(/\s+/g, '') : 'unknown';
    const label = status || 'unknown';
    return `<span class="bk-badge bk-badge-${escHtml(cls)}">${escHtml(label)}</span>`;
  }

  function loading(msg) {
    return `<div class="bk-loading"><div class="bk-spinner"></div><span>${escHtml(msg)}</span></div>`;
  }

  function renderWorkspaceSource(data) {
    const wrap = document.getElementById('bkWorkspaceSource');
    const describe = window.denjoyDescribeSourceMeta;
    if (!wrap || typeof describe !== 'function' || !data) return;
    const info = describe(data);
    wrap.innerHTML = `
      <div class="live-module-source">
        <span class="live-module-source-pill ${escHtml(info.className || '')}">${escHtml(info.label)}</span>
        <span>${escHtml(info.detail)}</span>
      </div>`;
  }

  function renderBackupOverview() {
    const wrap = document.getElementById('bkServiceOverview');
    if (!wrap) return;
    const spResources = _summary ? Number(_summary?.sharePoint?.resourceCount || 0) : (_sharepoint?.policies ? (_sharepoint.policies.reduce((sum, p) => sum + Number(p.siteCount || (p.sites || []).length || 0), 0)) : '—');
    const odResources = _summary ? Number(_summary?.oneDrive?.resourceCount || 0) : (_onedrive?.policies ? (_onedrive.policies.reduce((sum, p) => sum + Number(p.driveCount || (p.drives || []).length || 0), 0)) : '—');
    const exResources = _summary ? Number(_summary?.exchange?.resourceCount || 0) : (_exchange?.policies ? (_exchange.policies.reduce((sum, p) => sum + Number(p.mailboxCount || (p.mailboxes || []).length || 0), 0)) : '—');
    const serviceStatus = _summary ? String(_summary.serviceStatus || 'unknown') : '—';
    wrap.innerHTML = `
      <div class="workspace-service-overview">
        <article class="workspace-service-card"><span class="workspace-service-label">Service</span><strong class="workspace-service-value">${escHtml(serviceStatus)}</strong><span class="workspace-service-meta">backup status</span></article>
        <article class="workspace-service-card"><span class="workspace-service-label">SharePoint</span><strong class="workspace-service-value">${spResources}</strong><span class="workspace-service-meta">beschermde sites</span></article>
        <article class="workspace-service-card"><span class="workspace-service-label">OneDrive</span><strong class="workspace-service-value">${odResources}</strong><span class="workspace-service-meta">beschermde drives</span></article>
        <article class="workspace-service-card"><span class="workspace-service-label">Exchange</span><strong class="workspace-service-value">${exResources}</strong><span class="workspace-service-meta">beschermde mailboxen</span></article>
      </div>`;
  }

  // ── Tab switching ──────────────────────────────────────────────────────────

  function switchBkTab(tab) {
    document.querySelectorAll('#backupSection .bk-tab').forEach((btn) => {
      btn.classList.toggle('active', btn.dataset.bkTab === tab);
    });
    document.querySelectorAll('#backupSection .bk-tab-panel').forEach((panel) => {
      panel.style.display = panel.dataset.bkPanel === tab ? '' : 'none';
    });

    const tid = getTenantId();
    if (!tid) return;

    if (tab === 'overzicht' && !_summary)      loadSummary();
    if (tab === 'sharepoint' && !_sharepoint)  loadSharePoint();
    if (tab === 'onedrive' && !_onedrive)      loadOneDrive();
    if (tab === 'exchange' && !_exchange)      loadExchange();
    if (tab === 'geschiedenis')               loadBackupHistory();
  }

  function bindBkTabs() {
    if (_tabsBound) return;
    _tabsBound = true;

    document.querySelectorAll('#backupSection .bk-tab[data-bk-tab]').forEach((btn) => {
      btn.addEventListener('click', () => switchBkTab(btn.dataset.bkTab));
    });

    const bkBtnRefresh = document.getElementById('bkBtnRefreshSummary');
    if (bkBtnRefresh) bkBtnRefresh.addEventListener('click', () => { _summary = null; loadSummary(); });

    const bkBtnRefreshSP = document.getElementById('bkBtnRefreshSP');
    if (bkBtnRefreshSP) bkBtnRefreshSP.addEventListener('click', () => { _sharepoint = null; loadSharePoint(); });

    const bkBtnRefreshOD = document.getElementById('bkBtnRefreshOD');
    if (bkBtnRefreshOD) bkBtnRefreshOD.addEventListener('click', () => { _onedrive = null; loadOneDrive(); });

    const bkBtnRefreshEX = document.getElementById('bkBtnRefreshEX');
    if (bkBtnRefreshEX) bkBtnRefreshEX.addEventListener('click', () => { _exchange = null; loadExchange(); });
  }

  // ── Overzicht ──────────────────────────────────────────────────────────────

  function loadSummary() {
    const tid = getTenantId();
    if (!tid) return;
    const wrap = document.getElementById('bkSummaryWrap');
    if (!wrap) return;
    wrap.innerHTML = loading('Backup samenvatting laden…');

    bkApiFetch(`/api/backup/${tid}/summary`)
      .then((data) => {
        _summary = data;
        renderSummary(data);
      })
      .catch((err) => {
        wrap.innerHTML = `<p class="bk-empty">Fout bij laden: ${escHtml(err.message)}</p>`;
      });
  }

  function renderSummary(data) {
    const wrap = document.getElementById('bkSummaryWrap');
    if (!wrap) return;
    renderWorkspaceSource(data);
    renderBackupOverview();

    if (!data.ok) {
      wrap.innerHTML = `<p class="bk-empty">${escHtml(data.error || 'Onbekende fout')}</p>`;
      return;
    }

    const svcStatus = data.serviceStatus || 'unknown';
    const svcIcon = svcStatus === 'active' || svcStatus === 'enabled' ? '🟢' : svcStatus === 'notEnabled' ? '⚪' : '🟡';

    const storageStr = fmtBytes(data.storageUsed);

    const spCount  = data.sharePoint ? data.sharePoint.resourceCount : 0;
    const odCount  = data.oneDrive   ? data.oneDrive.resourceCount   : 0;
    const exCount  = data.exchange   ? data.exchange.resourceCount   : 0;
    const spPol    = data.sharePoint ? data.sharePoint.policyCount   : 0;
    const odPol    = data.oneDrive   ? data.oneDrive.policyCount     : 0;
    const exPol    = data.exchange   ? data.exchange.policyCount     : 0;

    wrap.innerHTML = `
      <div class="bk-status-card">
        <div class="bk-status-icon">${svcIcon}</div>
        <div>
          <div class="bk-status-label">Backup service status</div>
          <div class="bk-status-value">${statusBadge(svcStatus)}</div>
          ${data.lastModified ? `<div class="bk-status-sub">Laatst gewijzigd: ${fmtDate(data.lastModified)}</div>` : ''}
          ${storageStr !== '—' ? `<div class="bk-status-sub">Opslag gebruik: ${escHtml(storageStr)}</div>` : ''}
        </div>
      </div>

      ${svcStatus === 'notEnabled' ? `
        <div class="bk-not-enabled">
          <strong>M365 Backup niet ingeschakeld</strong>
          De Microsoft 365 Backup Storage service is nog niet geactiveerd voor deze tenant.
          Activeer dit via het <strong>Microsoft 365 Admin Center</strong> (Apps &gt; Microsoft 365 Backup)
          en zorg dat de app-registratie de permissie <code>BackupRestore-Configuration.Read.All</code> heeft.
        </div>
      ` : `
      <div class="bk-summary-grid">
        <div class="bk-summary-card">
          <div class="bk-summary-card-header">
            <span class="bk-summary-card-icon">🗂️</span>
            <span class="bk-summary-card-title">SharePoint</span>
          </div>
          <div class="bk-summary-stat"><span>Policies</span><strong>${spPol}</strong></div>
          <div class="bk-summary-stat"><span>Beschermde sites</span><strong>${spCount}</strong></div>
        </div>
        <div class="bk-summary-card">
          <div class="bk-summary-card-header">
            <span class="bk-summary-card-icon">☁️</span>
            <span class="bk-summary-card-title">OneDrive</span>
          </div>
          <div class="bk-summary-stat"><span>Policies</span><strong>${odPol}</strong></div>
          <div class="bk-summary-stat"><span>Beschermde drives</span><strong>${odCount}</strong></div>
        </div>
        <div class="bk-summary-card">
          <div class="bk-summary-card-header">
            <span class="bk-summary-card-icon">📧</span>
            <span class="bk-summary-card-title">Exchange</span>
          </div>
          <div class="bk-summary-stat"><span>Policies</span><strong>${exPol}</strong></div>
          <div class="bk-summary-stat"><span>Beschermde mailboxen</span><strong>${exCount}</strong></div>
        </div>
      </div>
      `}
    `;
  }

  // ── SharePoint ─────────────────────────────────────────────────────────────

  function loadSharePoint() {
    const tid = getTenantId();
    if (!tid) return;
    const wrap = document.getElementById('bkSPList');
    if (!wrap) return;
    const info = document.getElementById('bkSPCount');
    if (info) info.textContent = '— policies';
    wrap.innerHTML = loading('SharePoint backup policies laden…');

    bkApiFetch(`/api/backup/${tid}/sharepoint`)
      .then((data) => {
        _sharepoint = data;
        renderWorkspaceSource(data);
        renderBackupOverview();
        renderPolicyList(wrap, data, 'sites', 'site', ['siteName', 'siteUrl', 'status'], ['Site', 'URL', 'Status']);
        if (info) info.textContent = `${(data.policies || []).length} policies`;
      })
      .catch((err) => {
        wrap.innerHTML = `<p class="bk-empty">Fout: ${escHtml(err.message)}</p>`;
      });
  }

  // ── OneDrive ───────────────────────────────────────────────────────────────

  function loadOneDrive() {
    const tid = getTenantId();
    if (!tid) return;
    const wrap = document.getElementById('bkODList');
    if (!wrap) return;
    const info = document.getElementById('bkODCount');
    if (info) info.textContent = '— policies';
    wrap.innerHTML = loading('OneDrive backup policies laden…');

    bkApiFetch(`/api/backup/${tid}/onedrive`)
      .then((data) => {
        _onedrive = data;
        renderWorkspaceSource(data);
        renderBackupOverview();
        renderPolicyList(wrap, data, 'drives', 'drive', ['ownerName', 'status'], ['Eigenaar', 'Status']);
        if (info) info.textContent = `${(data.policies || []).length} policies`;
      })
      .catch((err) => {
        wrap.innerHTML = `<p class="bk-empty">Fout: ${escHtml(err.message)}</p>`;
      });
  }

  // ── Exchange ───────────────────────────────────────────────────────────────

  function loadExchange() {
    const tid = getTenantId();
    if (!tid) return;
    const wrap = document.getElementById('bkEXList');
    if (!wrap) return;
    const info = document.getElementById('bkEXCount');
    if (info) info.textContent = '— policies';
    wrap.innerHTML = loading('Exchange backup policies laden…');

    bkApiFetch(`/api/backup/${tid}/exchange`)
      .then((data) => {
        _exchange = data;
        renderWorkspaceSource(data);
        renderBackupOverview();
        renderPolicyList(wrap, data, 'mailboxes', 'mailbox', ['displayName', 'emailAddress', 'status'], ['Naam', 'E-mail', 'Status']);
        if (info) info.textContent = `${(data.policies || []).length} policies`;
      })
      .catch((err) => {
        wrap.innerHTML = `<p class="bk-empty">Fout: ${escHtml(err.message)}</p>`;
      });
  }

  // ── Gemeenschappelijke policy list renderer ────────────────────────────────

  function renderPolicyList(wrap, data, resourceKey, resourceLabel, fields, headers) {
    if (!data.ok) {
      wrap.innerHTML = `<p class="bk-empty">${escHtml(data.error || 'Onbekende fout')}</p>`;
      return;
    }

    const policies = data.policies || [];

    if (policies.length === 0) {
      const note = data.note || `Geen backup policies geconfigureerd voor dit service.`;
      wrap.innerHTML = `
        <div class="bk-not-enabled">
          <strong>Geen policies gevonden</strong>
          ${escHtml(note)}
        </div>`;
      return;
    }

    const cards = policies.map((pol) => {
      const resources = pol[resourceKey] || [];
      const retDays   = pol.retentionPeriodInDays ? `${pol.retentionPeriodInDays} dagen` : '—';
      const created   = fmtDate(pol.createdAt);
      const resourceCount = pol[`${resourceLabel}Count`] || resources.length;

      const rows = resources.map((r) => {
        const cells = fields.map((f) => {
          const v = r[f] || '—';
          if (f === 'status') return `<td>${statusBadge(v)}</td>`;
          if (f === 'siteUrl' || f === 'emailAddress') {
            return v !== '—' ? `<td><span title="${escHtml(v)}" style="max-width:220px;display:inline-block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${escHtml(v)}</span></td>` : '<td>—</td>';
          }
          return `<td>${escHtml(v)}</td>`;
        }).join('');
        return `<tr>${cells}</tr>`;
      }).join('');

      const thead = headers.map((h) => `<th>${escHtml(h)}</th>`).join('');

      return `
        <div class="bk-policy-card">
          <div class="bk-policy-header">
            <span class="bk-policy-name">${escHtml(pol.displayName || pol.id)}</span>
            <span class="bk-policy-meta">${resourceCount} ${resourceLabel}s &nbsp;·&nbsp; Retentie: ${escHtml(retDays)}</span>
            ${statusBadge(pol.status)}
            <span class="bk-policy-chevron">▶</span>
          </div>
          <div class="bk-policy-body">
            <table class="bk-resource-table">
              <thead><tr>${thead}</tr></thead>
              <tbody>${rows || `<tr><td colspan="${headers.length}" class="bk-history-table-empty">Geen ${resourceLabel}s in policy</td></tr>`}</tbody>
            </table>
          </div>
        </div>`;
    }).join('');

    wrap.innerHTML = `<div class="bk-policy-list">${cards}</div>`;

    // Accordion toggle
    wrap.querySelectorAll('.bk-policy-header').forEach((hdr) => {
      hdr.addEventListener('click', () => {
        hdr.closest('.bk-policy-card').classList.toggle('expanded');
      });
    });
  }

  // ── Geschiedenis ───────────────────────────────────────────────────────────

  function loadBackupHistory() {
    const tid = getTenantId();
    if (!tid) return;
    const tbody = document.getElementById('bkHistoryBody');
    if (!tbody) return;
    tbody.innerHTML = `<tr><td colspan="5" class="bk-history-table-empty">${loading('Geschiedenis laden…')}</td></tr>`;

    bkApiFetch(`/api/backup/${tid}/history`)
      .then((data) => {
        const items = data.items || [];
        if (!items.length) {
          tbody.innerHTML = '<tr><td colspan="5" class="bk-history-table-empty">Nog geen backup acties gelogd.</td></tr>';
          return;
        }
        tbody.innerHTML = items.map((row) => `
          <tr>
            <td>${fmtDate(row.executed_at)}</td>
            <td>${escHtml(row.action)}</td>
            <td>${statusBadge(row.status)}</td>
            <td>${escHtml(row.executed_by || '—')}</td>
            <td>${escHtml(row.error_message || '—')}</td>
          </tr>`).join('');
      })
      .catch((err) => {
        tbody.innerHTML = `<tr><td colspan="5" class="bk-history-table-empty">Fout: ${escHtml(err.message)}</td></tr>`;
      });
  }

  // ── Publieke ingang ────────────────────────────────────────────────────────

  window.loadBackupSection = function () {
    const tid = getTenantId();

    // Reset state bij tenant wissel
    const prevTid = window._bkLastTenantId;
    if (prevTid !== tid) {
      _summary = _sharepoint = _onedrive = _exchange = null;
      _tabsBound = false;
      window._bkLastTenantId = tid;
    }

    bindBkTabs();

    // Zorg dat het actieve tabblad zichtbaar is
    const activeTab = document.querySelector('#backupSection .bk-tab.active');
    const activePanel = activeTab ? activeTab.dataset.bkTab : 'overzicht';
    switchBkTab(activePanel);
  };
  window.switchBackupTab = switchBkTab;

})();
