/**
 * Denjoy IT Platform — Fase 6: Conditional Access
 * IIFE module — window.loadCaSection
 */
(function () {
  'use strict';

  let _policies = null;
  let _locations = null;
  let _tabsBound = false;
  let _currentCapability = null;

  function getTid() { const s = document.getElementById('tenantSelect'); return s ? s.value : ''; }
  function esc(s) { if (s == null) return ''; return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
  function fmtDate(iso) { if (!iso) return '—'; try { return new Date(iso).toLocaleString('nl-NL'); } catch(_) { return iso; } }

  function apiFetch(url, opts = {}) {
    const token = localStorage.getItem('denjoy_auth_token') || localStorage.getItem('denjoy_token') || '';
    const headers = { 'Content-Type': 'application/json' };
    if (token) headers['Authorization'] = 'Bearer ' + token;
    return fetch(url, { credentials: 'include', headers, ...opts }).then(r => r.json());
  }

  function apiFetchCached(url, opts, ttlMs) {
    const get = window.cacheGet; const set = window.cacheSet;
    const ttl = ttlMs || (window.CACHE_TTL ? window.CACHE_TTL.policies : 300000);
    if (get) { const hit = get(url); if (hit !== null) return Promise.resolve(hit); }
    return apiFetch(url, opts).then(data => { if (data !== null && set) set(url, data, ttl); return data; });
  }

  function loading(msg, type = 'cards') {
    if (type === 'table' && window.skeletonTable) return `<table class="ca-table"><tbody>${window.skeletonTable(5, 7)}</tbody></table>`;
    if (type === 'cards' && window.skeletonCards) return window.skeletonCards(4);
    return `<div class="ca-loading"><div class="ca-spinner"></div><span>${esc(msg)}</span></div>`;
  }

  function stateBadge(state) {
    const labels = { enabled: 'Ingeschakeld', disabled: 'Uitgeschakeld', enabledForReportingButNotEnforced: 'Rapportage-modus' };
    const label = labels[state] || state || '—';
    return `<span class="ca-state ca-state-${esc(state)}"><span class="ca-state-dot"></span>${esc(label)}</span>`;
  }

  function renderWorkspaceSource(data) {
    const wrap = document.getElementById('caWorkspaceSource');
    const describe = window.denjoyDescribeSourceMeta;
    if (!wrap || typeof describe !== 'function' || !data) return;
    const info = describe(data);
    wrap.innerHTML = `
      <div class="live-module-source">
        <span class="live-module-source-pill ${esc(info.className || '')}">${esc(info.label)}</span>
        <span>${esc(info.detail)}</span>
      </div>`;
  }

  async function renderCapabilityBanner(tab, forceRefresh = false) {
    const tid = getTid();
    const wrap = document.getElementById('caCapabilityBanner');
    if (!wrap || !tid || typeof window.denjoyFetchCapabilityStatus !== 'function' || typeof window.denjoyDescribeCapabilityStatus !== 'function') {
      if (wrap) wrap.innerHTML = '';
      return null;
    }
    const map = { policies: 'policies', locations: 'named-locations', geschiedenis: 'history' };
    try {
      const capability = await window.denjoyFetchCapabilityStatus(tid, 'ca', map[tab] || 'policies', { forceRefresh });
      _currentCapability = capability;
      const info = window.denjoyDescribeCapabilityStatus(capability);
      const roles = (capability.extra_roles || []).slice(0, 3).join(', ');
      wrap.innerHTML = `
        <div class="live-module-source">
          <span class="live-module-source-pill ${esc(info.className || '')}">${esc(info.label)}</span>
          <span>${esc(info.detail)}</span>
        </div>
        <div class="gb-capability-meta">${esc(roles || 'Geen extra rollen gespecificeerd')}</div>`;
      const blocked = capability.status === 'config_required' || capability.status === 'not_implemented' || !capability.supports_live;
      const refreshButtons = [document.getElementById('caBtnRefresh'), document.getElementById('caBtnRefreshLoc')];
      refreshButtons.forEach((btn) => { if (btn) btn.disabled = blocked; });
      return capability;
    } catch (_) {
      wrap.innerHTML = '';
      return null;
    }
  }

  function renderCaOverview() {
    const wrap = document.getElementById('caServiceOverview');
    if (!wrap) return;
    const policies = _policies?.policies || [];
    const enabled = policies.filter((p) => p.state === 'enabled').length;
    const reportOnly = policies.filter((p) => p.state === 'enabledForReportingButNotEnforced').length;
    const locations = _locations ? (_locations.locations || []).length : '—';
    wrap.innerHTML = `
      <div class="workspace-service-overview">
        <article class="workspace-service-card"><span class="workspace-service-label">Policies</span><strong class="workspace-service-value">${_policies ? policies.length : '—'}</strong><span class="workspace-service-meta">conditional access</span></article>
        <article class="workspace-service-card"><span class="workspace-service-label">Actief</span><strong class="workspace-service-value">${_policies ? enabled : '—'}</strong><span class="workspace-service-meta">ingeschakeld</span></article>
        <article class="workspace-service-card"><span class="workspace-service-label">Rapportage</span><strong class="workspace-service-value">${_policies ? reportOnly : '—'}</strong><span class="workspace-service-meta">report-only</span></article>
        <article class="workspace-service-card"><span class="workspace-service-label">Locaties</span><strong class="workspace-service-value">${locations}</strong><span class="workspace-service-meta">named locations</span></article>
      </div>`;
  }

  // ── Tab switching ──
  function switchCaTab(tab) {
    document.querySelectorAll('#caSection .ca-tab').forEach(b => b.classList.toggle('active', b.dataset.caTab === tab));
    document.querySelectorAll('#caSection .ca-tab-panel').forEach(p => { p.style.display = p.dataset.caPanel === tab ? '' : 'none'; });
    renderCapabilityBanner(tab);
    if (tab === 'policies' && !_policies) loadPolicies();
    if (tab === 'locations' && !_locations) loadLocations();
    if (tab === 'geschiedenis') loadCaHistory();
  }

  function bindCaTabs() {
    if (_tabsBound) return;
    _tabsBound = true;
    document.querySelectorAll('#caSection .ca-tab[data-ca-tab]').forEach(b => {
      b.addEventListener('click', () => switchCaTab(b.dataset.caTab));
    });
    const r = document.getElementById('caBtnRefresh');
    if (r) r.addEventListener('click', () => { _policies = null; loadPolicies(); });
    const rl = document.getElementById('caBtnRefreshLoc');
    if (rl) rl.addEventListener('click', () => { _locations = null; loadLocations(); });
  }

  // ── Policies ──
  function loadPolicies() {
    const tid = getTid(); if (!tid) return;
    const wrap = document.getElementById('caPoliciesWrap');
    if (!wrap) return;
    wrap.innerHTML = loading('CA policies laden…', 'table');
    apiFetchCached(`/api/ca/${tid}/policies`, {}, window.CACHE_TTL ? window.CACHE_TTL.policies : 300000)
      .then(data => { _policies = data; renderPolicies(data); })
      .catch(err => { wrap.innerHTML = `<p class="ca-empty">Fout: ${esc(err.message)}</p>`; });
  }

  function renderPolicies(data) {
    const wrap = document.getElementById('caPoliciesWrap');
    const info = document.getElementById('caPolicyCount');
    const banner = document.getElementById('caSnapshotBanner');
    if (!wrap) return;
    renderWorkspaceSource(data);
    renderCaOverview();
    if (banner) {
      if (data._source === 'assessment_snapshot') {
        banner.style.display = '';
        banner.textContent = 'Gegevens uit laatste assessment. Live data vereist actieve verbinding.';
      } else {
        banner.style.display = 'none';
      }
    }
    if (!data.ok) { wrap.innerHTML = `<p class="ca-empty">${esc(data.error || 'Fout')}</p>`; return; }
    const policies = data.policies || [];
    if (info) info.textContent = `${policies.length} policies`;
    if (!policies.length) { wrap.innerHTML = '<p class="ca-empty">Geen Conditional Access policies gevonden.</p>'; return; }

    wrap.innerHTML = `
      <div class="ca-table-wrap">
        <table class="ca-table">
          <thead><tr>
            <th>Naam</th><th>Staat</th><th>Gebruikers</th><th>Apps</th><th>Grant</th><th>Gewijzigd</th><th>Acties</th>
          </tr></thead>
          <tbody>
            ${policies.map(p => `
              <tr>
                <td><strong>${esc(p.displayName)}</strong></td>
                <td>${stateBadge(p.state)}</td>
                <td>${esc(p.userScope || '—')}</td>
                <td>${esc(p.appScope || '—')}</td>
                <td>${esc(p.grantControl || '—')}</td>
                <td>${fmtDate(p.modifiedAt)}</td>
                <td>
                  <div style="display:flex;gap:.375rem;flex-wrap:wrap;">
                    <button class="ca-btn ca-btn-secondary" style="font-size:.75rem;padding:.2rem .6rem" data-action="detail" data-pid="${esc(p.id)}">Detail</button>
                    ${p.state === 'enabled'
                      ? `<button class="ca-btn ca-btn-danger" style="font-size:.75rem;padding:.2rem .6rem" data-action="disable" data-pid="${esc(p.id)}" data-name="${esc(p.displayName)}">Uitschakelen</button>`
                      : `<button class="ca-btn ca-btn-success" style="font-size:.75rem;padding:.2rem .6rem" data-action="enable" data-pid="${esc(p.id)}" data-name="${esc(p.displayName)}">Inschakelen</button>`
                    }
                  </div>
                </td>
              </tr>`).join('')}
          </tbody>
        </table>
      </div>`;

    wrap.querySelectorAll('[data-action]').forEach(btn => {
      btn.addEventListener('click', () => {
        const { action, pid, name } = btn.dataset;
        if (action === 'detail') openPolicyDetail(pid);
        if (action === 'enable')  togglePolicy(pid, 'enable',  name);
        if (action === 'disable') togglePolicy(pid, 'disable', name);
      });
    });
  }

  function _caToast(msg, type) {
    const wrap = document.getElementById('caPoliciesWrap') || document.body;
    const el = document.createElement('div');
    el.style.cssText = `margin:0.5rem 0;padding:0.65rem 1rem;border-radius:8px;font-size:0.83rem;
      background:${type === 'error' ? 'rgba(239,68,68,0.1)' : 'rgba(34,197,94,0.1)'};
      border:1px solid ${type === 'error' ? 'rgba(239,68,68,0.3)' : 'rgba(34,197,94,0.3)'};
      color:${type === 'error' ? '#dc2626' : '#16a34a'};`;
    el.textContent = msg;
    wrap.insertAdjacentElement('afterbegin', el);
    setTimeout(() => el.remove(), 4000);
  }

  function _caConfirm(msg, onConfirm) {
    const wrap = document.getElementById('caPoliciesWrap') || document.body;
    const existing = document.getElementById('caConfirmBanner');
    if (existing) existing.remove();
    const el = document.createElement('div');
    el.id = 'caConfirmBanner';
    el.style.cssText = 'margin:0.5rem 0;padding:0.75rem 1rem;background:rgba(249,115,22,0.08);border:1px solid rgba(249,115,22,0.3);border-radius:8px;display:flex;align-items:center;gap:0.75rem;font-size:0.83rem;';
    el.innerHTML = `<span style="flex:1">${msg}</span>
      <button id="caConfirmYes" style="padding:0.35rem 0.8rem;background:var(--orange,#f97316);color:#fff;border:none;border-radius:6px;font-size:0.8rem;cursor:pointer;">Bevestigen</button>
      <button id="caConfirmNo" style="padding:0.35rem 0.8rem;background:transparent;border:1px solid #ddd;border-radius:6px;font-size:0.8rem;cursor:pointer;">Annuleren</button>`;
    wrap.insertAdjacentElement('afterbegin', el);
    el.querySelector('#caConfirmYes').addEventListener('click', () => { el.remove(); onConfirm(); });
    el.querySelector('#caConfirmNo').addEventListener('click', () => el.remove());
  }

  function togglePolicy(pid, action, name) {
    const tid = getTid(); if (!tid) return;
    const verb = action === 'enable' ? 'inschakelen' : 'uitschakelen';
    _caConfirm(`Policy <strong>${name}</strong> ${verb}?`, () => {
      apiFetch(`/api/ca/${tid}/policies/${pid}/toggle`, {
        method: 'POST',
        body: JSON.stringify({ action })
      }).then(data => {
        if (!data.ok) { _caToast('Fout: ' + (data.error || 'Onbekend'), 'error'); return; }
        _policies = null;
        if (window.cacheClear) window.cacheClear(`/api/ca/${tid}/policies`);
        loadPolicies();
      }).catch(err => _caToast('Fout: ' + err.message, 'error'));
    });
  }

  function openPolicyDetail(pid) {
    const tid = getTid(); if (!tid) return;

    if (typeof window.openSideRailDetail === 'function') {
      window.openSideRailDetail('CA Policy', 'Laden…');
    }

    apiFetch(`/api/ca/${tid}/policies/${pid}`).then(data => {
      if (!data.ok || !data.policy) {
        if (typeof window.updateSideRailDetail === 'function') window.updateSideRailDetail('Fout', `<p class="ca-empty">${esc(data.error || 'Niet gevonden')}</p>`);
        return;
      }
      const p = data.policy;
      const bodyHtml = `
        <div class="ca-detail-grid">
          <div><div class="ca-detail-label">Staat</div><div class="ca-detail-value">${stateBadge(p.state)}</div></div>
          <div><div class="ca-detail-label">Sessiecontroles</div><div class="ca-detail-value">${esc(p.sessionCtrl)}</div></div>
          <div><div class="ca-detail-label">Gebruikers scope</div><div class="ca-detail-value">${esc(p.userScope || '—')}</div></div>
          <div><div class="ca-detail-label">Apps scope</div><div class="ca-detail-value">${esc(p.appScope || '—')}</div></div>
          <div><div class="ca-detail-label">Grant control</div><div class="ca-detail-value">${esc(p.grantControl || '—')}</div></div>
          <div><div class="ca-detail-label">Aangemaakt</div><div class="ca-detail-value">${fmtDate(p.createdAt)}</div></div>
          <div><div class="ca-detail-label">Gewijzigd</div><div class="ca-detail-value">${fmtDate(p.modifiedAt)}</div></div>
          <div><div class="ca-detail-label">Policy ID</div><div class="ca-detail-value" style="font-family:var(--mono);font-size:.72rem;word-break:break-all">${esc(p.id)}</div></div>
        </div>`;
      if (typeof window.updateSideRailDetail === 'function') window.updateSideRailDetail(p.displayName, bodyHtml);
    }).catch(err => {
      if (typeof window.updateSideRailDetail === 'function') window.updateSideRailDetail('Fout', `<p class="ca-empty">Fout: ${esc(err.message)}</p>`);
    });
  }

  // ── Named Locations ──
  function loadLocations() {
    const tid = getTid(); if (!tid) return;
    const wrap = document.getElementById('caLocationsWrap');
    if (!wrap) return;
    wrap.innerHTML = loading('Named locations laden…', 'cards');
    apiFetchCached(`/api/ca/${tid}/named-locations`, {}, window.CACHE_TTL ? window.CACHE_TTL.policies : 300000)
      .then(data => { _locations = data; renderLocations(data); })
      .catch(err => { wrap.innerHTML = `<p class="ca-empty">Fout: ${esc(err.message)}</p>`; });
  }

  function renderLocations(data) {
    const wrap = document.getElementById('caLocationsWrap');
    const info = document.getElementById('caLocCount');
    if (!wrap) return;
    renderWorkspaceSource(data);
    renderCaOverview();
    if (!data.ok) { wrap.innerHTML = `<p class="ca-empty">${esc(data.error || 'Fout')}</p>`; return; }
    const locs = data.locations || [];
    if (info) info.textContent = `${locs.length} locaties`;
    if (!locs.length) { wrap.innerHTML = '<p class="ca-empty">Geen named locations gevonden.</p>'; return; }
    wrap.innerHTML = `<div class="ca-loc-grid">${locs.map(l => `
      <div class="ca-loc-card">
        <div class="ca-loc-name">${esc(l.displayName)}</div>
        <div class="ca-loc-meta">
          <span>${esc(l.type === 'ipRange' ? '🌐 IP Range' : '🗺️ Landen')}</span>
          <span>${esc(l.detail)}</span>
          ${l.isTrusted ? '<span class="ca-loc-trusted">Vertrouwd</span>' : ''}
          <span style="color:var(--text-muted);font-size:.7rem">${fmtDate(l.createdAt)}</span>
        </div>
      </div>`).join('')}</div>`;
  }

  // ── Geschiedenis ──
  function loadCaHistory() {
    const tid = getTid(); if (!tid) return;
    const tbody = document.getElementById('caHistoryBody');
    if (!tbody) return;
    tbody.innerHTML = `<tr><td colspan="5" class="ca-history-table-empty">${loading('Laden…')}</td></tr>`;
    apiFetch(`/api/ca/${tid}/history`).then(data => {
      const items = data.items || [];
      if (!items.length) { tbody.innerHTML = '<tr><td colspan="5" class="ca-history-table-empty">Nog geen CA acties gelogd.</td></tr>'; return; }
      tbody.innerHTML = items.map(r => `<tr>
        <td>${fmtDate(r.executed_at)}</td>
        <td>${esc(r.action)}</td>
        <td>${esc(r.policy_id || '—')}</td>
        <td><span class="ca-state ca-state-${esc(r.status)}">${esc(r.status)}</span></td>
        <td>${esc(r.executed_by || '—')}</td>
      </tr>`).join('');
    }).catch(err => { tbody.innerHTML = `<tr><td colspan="5" class="ca-history-table-empty">Fout: ${esc(err.message)}</td></tr>`; });
  }

  // ── Publieke ingang ──
  window.loadCaSection = function () {
    const tid = getTid();
    if (window._caLastTid !== tid) { _policies = _locations = null; _tabsBound = false; window._caLastTid = tid; }
    bindCaTabs();
    const active = document.querySelector('#caSection .ca-tab.active');
    renderCapabilityBanner(active ? active.dataset.caTab : 'policies');
    switchCaTab(active ? active.dataset.caTab : 'policies');
  };
  window.switchCaTab = switchCaTab;
})();
