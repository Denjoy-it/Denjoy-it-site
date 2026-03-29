/**
 * Denjoy IT Platform — Fase 9: Exchange & Email module
 * IIFE module — window.loadExchangeSection
 */
(function () {
  'use strict';

  let _mailboxes = null;
  let _rules = null;
  let _forwarding = null;
  let _tabsBound = false;
  let _searchQ = '';

  function getTid() { const s = document.getElementById('tenantSelect'); return s ? s.value : ''; }
  function esc(s) { if (s == null) return ''; return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
  function fmtDate(iso) { if (!iso) return '—'; try { return new Date(iso).toLocaleString('nl-NL'); } catch(_) { return iso; } }
  function initials(name) { if (!name) return '?'; return name.split(' ').filter(Boolean).map(w => w[0]).slice(0,2).join('').toUpperCase(); }

  function apiFetch(url, opts = {}) {
    const token = localStorage.getItem('denjoy_auth_token') || localStorage.getItem('denjoy_token') || '';
    const headers = { 'Content-Type': 'application/json' };
    if (token) headers['Authorization'] = 'Bearer ' + token;
    return fetch(url, { credentials: 'include', headers, ...opts }).then(r => r.json());
  }

  function apiFetchCached(url, opts, ttlMs) {
    const get = window.cacheGet; const set = window.cacheSet;
    const ttl = ttlMs || (window.CACHE_TTL ? window.CACHE_TTL.mailboxes : 120000);
    if (get) { const hit = get(url); if (hit !== null) return Promise.resolve(hit); }
    return apiFetch(url, opts).then(data => { if (data !== null && set) set(url, data, ttl); return data; });
  }

  function loading(msg, type = 'lines') {
    if (type === 'table' && window.skeletonTable) return `<tr><td colspan="5">${window.skeletonCards(3)}</td></tr>`;
    if (type === 'lines' && window.skeletonLines) return window.skeletonLines(4);
    return `<div class="ex-loading"><div class="ex-spinner"></div><span>${esc(msg)}</span></div>`;
  }

  function renderWorkspaceSource(data) {
    const wrap = document.getElementById('exWorkspaceSource');
    const describe = window.denjoyDescribeSourceMeta;
    if (!wrap || typeof describe !== 'function' || !data) return;
    const info = describe(data);
    wrap.innerHTML = `
      <div class="live-module-source">
        <span class="live-module-source-pill ${esc(info.className || '')}">${esc(info.label)}</span>
        <span>${esc(info.detail)}</span>
      </div>`;
  }

  function renderExchangeOverview() {
    const wrap = document.getElementById('exServiceOverview');
    if (!wrap) return;
    const mailboxCount = _mailboxes ? (_mailboxes.mailboxes || []).length : '—';
    const forwardingCount = _forwarding ? (_forwarding.forwarding || []).length : '—';
    const rulesCount = _rules ? (_rules.rules || []).length : '—';
    const suspiciousCount = _rules ? Number(_rules.suspicious || 0) : '—';
    wrap.innerHTML = `
      <div class="workspace-service-overview">
        <article class="workspace-service-card"><span class="workspace-service-label">Mailboxen</span><strong class="workspace-service-value">${mailboxCount}</strong><span class="workspace-service-meta">tenantbreed</span></article>
        <article class="workspace-service-card"><span class="workspace-service-label">Forwarding</span><strong class="workspace-service-value">${forwardingCount}</strong><span class="workspace-service-meta">actief</span></article>
        <article class="workspace-service-card"><span class="workspace-service-label">Regels</span><strong class="workspace-service-value">${rulesCount}</strong><span class="workspace-service-meta">inboxregels</span></article>
        <article class="workspace-service-card"><span class="workspace-service-label">Verdacht</span><strong class="workspace-service-value">${suspiciousCount}</strong><span class="workspace-service-meta">controle vereist</span></article>
      </div>`;
  }

  // ── Tab switching ──
  function switchExTab(tab) {
    document.querySelectorAll('#exchangeSection .ex-tab').forEach(b => b.classList.toggle('active', b.dataset.exTab === tab));
    document.querySelectorAll('#exchangeSection .ex-tab-panel').forEach(p => { p.style.display = p.dataset.exPanel === tab ? '' : 'none'; });
    if (tab === 'mailboxen'   && !_mailboxes)  loadMailboxes();
    if (tab === 'forwarding'  && !_forwarding) loadForwarding();
    if (tab === 'regels'      && !_rules)      loadRules();
  }

  function bindExTabs() {
    if (_tabsBound) return;
    _tabsBound = true;
    document.querySelectorAll('#exchangeSection .ex-tab[data-ex-tab]').forEach(b => {
      b.addEventListener('click', () => switchExTab(b.dataset.exTab));
    });
    const r = document.getElementById('exBtnRefreshMbx');
    if (r) r.addEventListener('click', () => { _mailboxes = null; loadMailboxes(); });
    const rf = document.getElementById('exBtnRefreshFwd');
    if (rf) rf.addEventListener('click', () => { _forwarding = null; loadForwarding(); });
    const rr = document.getElementById('exBtnRefreshRules');
    if (rr) rr.addEventListener('click', () => { _rules = null; loadRules(); });
    const search = document.getElementById('exSearchInput');
    if (search) search.addEventListener('input', e => { _searchQ = e.target.value.toLowerCase(); renderMailboxes(_mailboxes); });
  }

  // ── Mailboxen ──
  function loadMailboxes() {
    const tid = getTid(); if (!tid) return;
    const wrap = document.getElementById('exMailboxTableBody');
    if (!wrap) return;
    wrap.innerHTML = `<tr><td colspan="5" class="ex-table-empty">${loading('Mailboxen laden…', 'table')}</td></tr>`;
    apiFetchCached(`/api/exchange/${tid}/mailboxes`, {}, window.CACHE_TTL ? window.CACHE_TTL.mailboxes : 120000)
      .then(data => { _mailboxes = data; renderMailboxes(data); })
      .catch(err => { wrap.innerHTML = `<tr><td colspan="5" class="ex-table-empty">Fout: ${esc(err.message)}</td></tr>`; });
  }

  function renderMailboxes(data) {
    const tbody = document.getElementById('exMailboxTableBody');
    const info = document.getElementById('exMbxCount');
    if (!tbody) return;
    renderWorkspaceSource(data);
    renderExchangeOverview();
    if (!data || !data.ok) { tbody.innerHTML = `<tr><td colspan="5" class="ex-table-empty">${esc(data?.error || 'Fout')}</td></tr>`; return; }
    let mbx = data.mailboxes || [];
    if (_searchQ) mbx = mbx.filter(m => (m.displayName + m.upn + (m.mail||'')).toLowerCase().includes(_searchQ));
    if (info) info.textContent = `${mbx.length} mailboxen`;
    if (!mbx.length) { tbody.innerHTML = '<tr><td colspan="5" class="ex-table-empty">Geen mailboxen gevonden.</td></tr>'; return; }
    tbody.innerHTML = mbx.map(m => {
      const statusCls = m.accountEnabled ? 'active' : 'disabled';
      const statusLabel = m.accountEnabled ? 'Actief' : 'Uitgeschakeld';
      const syncBadge = m.onPremSync ? '<span class="ex-badge ex-badge-sync">Sync</span> ' : '';
      const replyBadge = m.autoReplyEnabled ? '<span class="ex-badge ex-badge-warn">Auto-reply</span> ' : '';
      return `<tr data-uid="${esc(m.id)}">
        <td>
          <div class="ex-mailbox-cell">
            <div class="ex-mailbox-avatar">${esc(initials(m.displayName))}</div>
            <div>
              <div style="font-weight:600">${esc(m.displayName)}</div>
              <div style="font-size:.75rem;color:var(--text-muted)">${esc(m.upn)}</div>
            </div>
          </div>
        </td>
        <td>${esc(m.mail || '—')}</td>
        <td><span class="ex-badge ex-badge-${esc(statusCls)}">${esc(statusLabel)}</span> ${syncBadge}${replyBadge}</td>
        <td>${esc(m.timezone || '—')}</td>
        <td><button class="ex-btn" style="font-size:.75rem;padding:.2rem .55rem" data-detail="${esc(m.id)}">Detail</button></td>
      </tr>`;
    }).join('');
    tbody.querySelectorAll('[data-detail]').forEach(btn => {
      btn.addEventListener('click', () => openMailboxDetail(btn.dataset.detail));
    });
  }

  function _mailboxDetailHtml(m) {
    const isSnapshot = m._source === 'assessment_snapshot';
    const snapshotNote = isSnapshot
      ? '<div class="dp-snapshot-note">Gegevens uit laatste assessment. Live data vereist actieve verbinding.</div>'
      : '';
    const fwdHtml = m.forwarding && m.forwarding.enabled
      ? `<div class="ex-fwd-alert">⚠ Forwarding actief → ${esc(m.forwarding.address || '?')}</div>`
      : '<span style="color:var(--text-muted)">Geen forwarding</span>';
    const statusLabel = m.accountEnabled === false ? 'Uitgeschakeld' : m.accountEnabled === true ? 'Actief' : null;
    const extraRows = [
      m.recipientTypeDetails ? `<div class="ex-detail-item"><label>Type</label><span>${esc(m.recipientTypeDetails)}</span></div>` : '',
      statusLabel            ? `<div class="ex-detail-item"><label>Status</label><span>${esc(statusLabel)}</span></div>` : '',
      m.whenCreated          ? `<div class="ex-detail-item"><label>Aangemaakt</label><span>${fmtDate(m.whenCreated)}</span></div>` : '',
    ].join('');
    return `
      ${snapshotNote}
      <div class="ex-detail-grid">
        <div class="ex-detail-item"><label>E-mail</label><span>${esc(m.mail || '—')}</span></div>
        <div class="ex-detail-item"><label>UPN</label><span>${esc(m.upn || '—')}</span></div>
        <div class="ex-detail-item"><label>Afdeling</label><span>${esc(m.department || '—')}</span></div>
        <div class="ex-detail-item"><label>Functie</label><span>${esc(m.jobTitle || '—')}</span></div>
        <div class="ex-detail-item"><label>Tijdzone</label><span>${esc(m.timezone || '—')}</span></div>
        <div class="ex-detail-item"><label>Taal</label><span>${esc(m.language || '—')}</span></div>
        <div class="ex-detail-item"><label>Mobiel</label><span>${esc(m.mobile || '—')}</span></div>
        <div class="ex-detail-item"><label>Kantoor</label><span>${esc(m.office || '—')}</span></div>
        ${extraRows}
      </div>
      <div class="ex-detail-section">
        <div class="ex-detail-section-title">Forwarding</div>
        ${fwdHtml}
      </div>
      <div class="ex-detail-section">
        <div class="ex-detail-section-title">Auto-reply</div>
        <div style="font-size:.8125rem;color:var(--text-muted)">${esc(m.autoReply?.status || 'disabled')}</div>
      </div>
    `;
  }

  function _renderMailboxDetailModal(overlay, m) {
    // Verouderd pad — stuurt nu naar het Inzichten-paneel
    if (typeof window.updateSideRailDetail === 'function') {
      window.updateSideRailDetail(m.displayName || 'Mailbox', _mailboxDetailHtml(m));
    }
  }

  function _cachedMailbox(uid) {
    return (_mailboxes?.mailboxes || []).find(m => m.id === uid || m.primarySmtpAddress === uid || m.mail === uid);
  }

  function openMailboxDetail(uid) {
    const tid = getTid(); if (!tid) return;

    const cached = _cachedMailbox(uid);
    const fallbackName = cached?.displayName || uid;

    // Open het Inzichten-paneel direct
    if (typeof window.openSideRailDetail === 'function') {
      window.openSideRailDetail('Mailbox', fallbackName);
    }

    // Uit assessment snapshot? Direct renderen zonder round-trip
    if (cached && _mailboxes?._source === 'assessment_snapshot') {
      _renderMailboxDetailModal(null, {
        ok: true,
        displayName: cached.displayName,
        mail: cached.mail || cached.primarySmtpAddress || uid,
        upn: cached.upn || cached.primarySmtpAddress || uid,
        department: cached.department || null,
        jobTitle: cached.jobTitle || null,
        office: cached.office || null,
        mobile: cached.mobile || null,
        timezone: cached.timezone || null,
        language: cached.language || null,
        accountEnabled: cached.accountEnabled,
        recipientTypeDetails: cached.recipientTypeDetails || null,
        whenCreated: cached.whenCreated || null,
        autoReply: { status: cached.autoReplyEnabled ? 'enabled' : 'disabled' },
        forwarding: cached.forwarding || { enabled: false, address: null },
        _source: 'assessment_snapshot',
      });
      return;
    }

    const fallbackData = cached ? { ok: true, displayName: cached.displayName, mail: cached.mail || cached.primarySmtpAddress || uid, upn: cached.upn || uid, department: null, jobTitle: null, office: null, mobile: null, timezone: cached.timezone || null, language: null, accountEnabled: cached.accountEnabled, recipientTypeDetails: cached.recipientTypeDetails, whenCreated: cached.whenCreated, autoReply: { status: 'disabled' }, forwarding: { enabled: false }, _source: 'assessment_snapshot' } : null;

    apiFetch(`/api/exchange/${tid}/mailboxes/${uid}`).then(data => {
      if (!data.ok) {
        if (fallbackData) _renderMailboxDetailModal(null, fallbackData);
        else if (typeof window.updateSideRailDetail === 'function') window.updateSideRailDetail('Fout', `<p class="ex-empty">${esc(data.error || 'Fout')}</p>`);
        return;
      }
      _renderMailboxDetailModal(null, data);
    }).catch(err => {
      if (fallbackData) _renderMailboxDetailModal(null, fallbackData);
      else if (typeof window.updateSideRailDetail === 'function') window.updateSideRailDetail('Fout', `<p class="ex-empty">Fout: ${esc(err.message)}</p>`);
    });
  }

  // ── Forwarding ──
  function loadForwarding() {
    const tid = getTid(); if (!tid) return;
    const wrap = document.getElementById('exFwdWrap');
    if (!wrap) return;
    wrap.innerHTML = loading('Forwarding-instellingen laden…');
    apiFetch(`/api/exchange/${tid}/forwarding`).then(data => { _forwarding = data; renderForwarding(data); })
      .catch(err => { wrap.innerHTML = `<p class="ex-empty">Fout: ${esc(err.message)}</p>`; });
  }

  function renderForwarding(data) {
    const wrap = document.getElementById('exFwdWrap');
    const info = document.getElementById('exFwdCount');
    if (!wrap) return;
    renderWorkspaceSource(data);
    renderExchangeOverview();
    if (!data.ok) { wrap.innerHTML = `<p class="ex-empty">${esc(data.error || 'Fout')}</p>`; return; }
    const fwd = data.forwarding || [];
    if (info) info.textContent = `${fwd.length} actieve forwardings`;
    if (!fwd.length) { wrap.innerHTML = '<div class="ex-fwd-empty">✓ Geen actieve e-mail forwarding gevonden.</div>'; return; }
    wrap.innerHTML = `
      <div class="ex-fwd-banner">⚠ ${fwd.length} mailbox(en) met actieve forwarding — controleer of dit gewenst is.</div>
      <div class="ex-table-wrap">
        <table class="ex-table">
          <thead><tr><th>Gebruiker</th><th>UPN</th><th>Doorstuurt naar</th></tr></thead>
          <tbody>${fwd.map(f => `<tr>
            <td>${esc(f.displayName)}</td>
            <td>${esc(f.upn)}</td>
            <td><span class="ex-fwd-alert">⚠ ${esc(f.forwardTo)}</span></td>
          </tr>`).join('')}</tbody>
        </table>
      </div>`;
  }

  // ── Inbox regels ──
  function loadRules() {
    const tid = getTid(); if (!tid) return;
    const wrap = document.getElementById('exRulesWrap');
    if (!wrap) return;
    wrap.innerHTML = loading('Inbox regels analyseren (kan even duren)…');
    apiFetch(`/api/exchange/${tid}/mailbox-rules`).then(data => { _rules = data; renderRules(data); })
      .catch(err => { wrap.innerHTML = `<p class="ex-empty">Fout: ${esc(err.message)}</p>`; });
  }

  function renderRules(data) {
    const wrap = document.getElementById('exRulesWrap');
    const info = document.getElementById('exRulesCount');
    if (!wrap) return;
    renderWorkspaceSource(data);
    renderExchangeOverview();
    if (!data.ok) { wrap.innerHTML = `<p class="ex-empty">${esc(data.error || 'Fout')}</p>`; return; }
    const rules = data.rules || [];
    if (info) info.textContent = `${rules.length} regels (${data.suspicious || 0} verdacht)`;
    if (!rules.length) { wrap.innerHTML = `<p class="ex-empty">Geen inbox regels gevonden (${data.usersChecked || 0} mailboxen gecontroleerd).</p>`; return; }

    const suspicious = rules.filter(r => r.suspicious);
    const normal = rules.filter(r => !r.suspicious);

    wrap.innerHTML = `
      ${suspicious.length ? `
        <div class="ex-fwd-banner">⚠ ${suspicious.length} verdachte regel(s) gevonden — controleer direct.</div>
        <div class="ex-table-wrap" style="margin-bottom:1rem">
          <table class="ex-table">
            <thead><tr><th>Gebruiker</th><th>Regelsnaam</th><th>Melding</th><th>Actief</th><th>Doorstuurt naar</th></tr></thead>
            <tbody>${suspicious.map(r => `<tr>
              <td>${esc(r.userName)}<div style="font-size:.75rem;color:var(--text-muted)">${esc(r.userUpn)}</div></td>
              <td>${esc(r.ruleName)}</td>
              <td>${r.flags.map(f => `<span class="ex-rule-suspicious">⚠ ${esc(f)}</span>`).join(' ')}</td>
              <td>${r.enabled ? '✓' : '—'}</td>
              <td>${esc(r.forwardTo || '—')}</td>
            </tr>`).join('')}</tbody>
          </table>
        </div>` : ''}
      ${normal.length ? `
        <details>
          <summary style="cursor:pointer;font-size:.875rem;color:var(--text-muted);margin-bottom:.5rem">
            ${normal.length} normale regel(s) tonen
          </summary>
          <div class="ex-table-wrap">
            <table class="ex-table">
              <thead><tr><th>Gebruiker</th><th>Regelsnaam</th><th>Actief</th></tr></thead>
              <tbody>${normal.map(r => `<tr>
                <td>${esc(r.userName)}<div style="font-size:.75rem;color:var(--text-muted)">${esc(r.userUpn)}</div></td>
                <td>${esc(r.ruleName)}</td>
                <td>${r.enabled ? '✓' : '—'}</td>
              </tr>`).join('')}</tbody>
            </table>
          </div>
        </details>` : ''}`;
  }

  // ── Publieke ingang ──
  window.loadExchangeSection = function () {
    const tid = getTid();
    if (window._exLastTid !== tid) { _mailboxes = _rules = _forwarding = null; _tabsBound = false; _searchQ = ''; window._exLastTid = tid; }
    bindExTabs();
    const active = document.querySelector('#exchangeSection .ex-tab.active');
    switchExTab(active ? active.dataset.exTab : 'mailboxen');
  };
  window.switchExchangeTab = switchExTab;
})();
