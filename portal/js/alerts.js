/**
 * Denjoy IT Platform — Fase 8: Alerts & Notificaties
 * IIFE module — window.loadAlertsSection
 */
(function () {
  'use strict';

  let _tabsBound = false;
  let _secureScore = null;
  let _auditData = null;
  let _signInsData = null;

  function getTid() { const s = document.getElementById('tenantSelect'); return s ? s.value : ''; }
  function esc(s) { if (s == null) return ''; return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
  function fmtDate(iso) { if (!iso) return '—'; try { return new Date(iso).toLocaleString('nl-NL'); } catch(_) { return iso; } }

  function apiFetch(url, opts = {}) {
    const token = localStorage.getItem('denjoy_auth_token') || localStorage.getItem('denjoy_token') || '';
    const headers = { 'Content-Type': 'application/json' };
    if (token) headers['Authorization'] = 'Bearer ' + token;
    return fetch(url, { credentials: 'include', headers, ...opts }).then(r => r.json());
  }

  function loading(msg, type = 'lines') {
    if (type === 'lines' && window.skeletonLines) return window.skeletonLines(5);
    if (type === 'cards' && window.skeletonCards) return window.skeletonCards(3);
    return `<div class="al-loading"><div class="al-spinner"></div><span>${esc(msg)}</span></div>`;
  }

  function resultBadge(result) {
    const cls = (result || 'unknown').toLowerCase();
    return `<span class="al-result al-result-${esc(cls)}">${esc(result || '—')}</span>`;
  }

  function riskBadge(level) {
    const cls = (level || 'none').toLowerCase();
    return `<span class="al-risk al-risk-${esc(cls)}">${esc(level || 'none')}</span>`;
  }

  function renderWorkspaceSource(data) {
    const wrap = document.getElementById('alWorkspaceSource');
    const describe = window.denjoyDescribeSourceMeta;
    if (!wrap || typeof describe !== 'function' || !data) return;
    const info = describe(data);
    wrap.innerHTML = `
      <div class="live-module-source">
        <span class="live-module-source-pill ${esc(info.className || '')}">${esc(info.label)}</span>
        <span>${esc(info.detail)}</span>
      </div>`;
  }

  function renderAlertsOverview() {
    const wrap = document.getElementById('alServiceOverview');
    if (!wrap) return;
    const auditCount = _auditData ? (_auditData.items || []).length : '—';
    const score = _secureScore ? Number(_secureScore.currentScore ?? _secureScore.score ?? 0) : '—';
    const maxScore = _secureScore ? Number(_secureScore.maxScore || 100) : '—';
    const signIns = _signInsData ? (_signInsData.items || []).length : '—';
    const risky = _signInsData ? (_signInsData.items || []).filter((item) => String(item?.riskLevel || '').toLowerCase() !== 'none').length : '—';
    wrap.innerHTML = `
      <div class="workspace-service-overview">
        <article class="workspace-service-card"><span class="workspace-service-label">Audit log</span><strong class="workspace-service-value">${auditCount}</strong><span class="workspace-service-meta">events</span></article>
        <article class="workspace-service-card"><span class="workspace-service-label">Secure Score</span><strong class="workspace-service-value">${score}</strong><span class="workspace-service-meta">/ ${maxScore}</span></article>
        <article class="workspace-service-card"><span class="workspace-service-label">Aanmeldingen</span><strong class="workspace-service-value">${signIns}</strong><span class="workspace-service-meta">recent</span></article>
        <article class="workspace-service-card"><span class="workspace-service-label">Risico</span><strong class="workspace-service-value">${risky}</strong><span class="workspace-service-meta">risicovol</span></article>
      </div>`;
  }

  // ── Tab switching ──
  function switchAlTab(tab) {
    document.querySelectorAll('#alertsSection .al-tab').forEach(b => b.classList.toggle('active', b.dataset.alTab === tab));
    document.querySelectorAll('#alertsSection .al-tab-panel').forEach(p => { p.style.display = p.dataset.alPanel === tab ? '' : 'none'; });
    if (tab === 'auditlog')   loadAuditLog();
    if (tab === 'securescr')  loadSecureScore();
    if (tab === 'signins')    loadSignIns();
    if (tab === 'config')     loadNotifConfig();
  }

  function bindAlTabs() {
    if (_tabsBound) return;
    _tabsBound = true;
    document.querySelectorAll('#alertsSection .al-tab[data-al-tab]').forEach(b => {
      b.addEventListener('click', () => switchAlTab(b.dataset.alTab));
    });
    const r = document.getElementById('alBtnRefreshAudit');
    if (r) r.addEventListener('click', loadAuditLog);
    const rs = document.getElementById('alBtnRefreshScore');
    if (rs) rs.addEventListener('click', () => { _secureScore = null; loadSecureScore(); });
    const rsi = document.getElementById('alBtnRefreshSignIns');
    if (rsi) rsi.addEventListener('click', loadSignIns);
  }

  // ── Audit Log ──
  function loadAuditLog() {
    const tid = getTid(); if (!tid) return;
    const wrap = document.getElementById('alAuditWrap');
    if (!wrap) return;
    wrap.innerHTML = loading('Audit log laden…');
    apiFetch(`/api/alerts/${tid}/audit-logs`).then(data => renderAuditLog(data))
      .catch(err => { wrap.innerHTML = `<p class="al-empty">Fout: ${esc(err.message)}</p>`; });
  }

  function _showAlBanner(src) {
    const banner = document.getElementById('alSnapshotBanner');
    if (!banner) return;
    if (src === 'assessment_snapshot') {
      banner.style.display = '';
      banner.textContent = 'Gegevens uit laatste assessment. Live data vereist actieve verbinding.';
    } else {
      banner.style.display = 'none';
    }
  }

  function renderAuditLog(data) {
    const wrap = document.getElementById('alAuditWrap');
    const info = document.getElementById('alAuditCount');
    if (!wrap) return;
    _auditData = data;
    renderWorkspaceSource(data);
    renderAlertsOverview();
    _showAlBanner(data._source);
    if (!data.ok) { wrap.innerHTML = `<p class="al-empty">${esc(data.error || 'Fout')}</p>`; return; }
    const items = data.items || [];
    if (info) info.textContent = `${items.length} events`;
    if (!items.length) {
      const msg = data.message || 'Geen audit log events gevonden.';
      wrap.innerHTML = `<p class="al-empty">${esc(msg)}</p>`;
      return;
    }
    wrap.innerHTML = `
      <div class="al-table-wrap">
        <table class="al-table">
          <thead><tr><th>Tijdstip</th><th>Activiteit</th><th>Categorie</th><th>Resultaat</th><th>Geïnitieerd door</th><th>Doel</th></tr></thead>
          <tbody>${items.map(i => `<tr>
            <td style="white-space:nowrap">${fmtDate(i.activityDateTime)}</td>
            <td>${esc(i.activityDisplayName)}</td>
            <td>${esc(i.category || '—')}</td>
            <td>${resultBadge(i.result)}</td>
            <td>${esc(i.initiatedBy)}</td>
            <td>${esc(i.targetResources)}</td>
          </tr>`).join('')}</tbody>
        </table>
      </div>`;
  }

  // ── Secure Score ──
  function loadSecureScore() {
    const tid = getTid(); if (!tid) return;
    if (_secureScore) { renderSecureScore(_secureScore); return; }
    const wrap = document.getElementById('alScoreWrap');
    if (!wrap) return;
    wrap.innerHTML = loading('Secure Score laden…', 'cards');
    apiFetch(`/api/alerts/${tid}/secure-score`).then(data => { _secureScore = data; renderSecureScore(data); })
      .catch(err => { wrap.innerHTML = `<p class="al-empty">Fout: ${esc(err.message)}</p>`; });
  }

  function renderSecureScore(data) {
    const wrap = document.getElementById('alScoreWrap');
    if (!wrap) return;
    renderWorkspaceSource(data);
    renderAlertsOverview();
    _showAlBanner(data._source);
    if (!data.ok) { wrap.innerHTML = `<p class="al-empty">${esc(data.error || 'Fout')}</p>`; return; }
    if (!data.score && data.message) { wrap.innerHTML = `<p class="al-empty">${esc(data.message)}</p>`; return; }

    const pct = data.percentage || 0;
    const circ = 2 * Math.PI * 50;
    const offset = circ - (pct / 100) * circ;
    const scoreColor = pct >= 70 ? '#22c55e' : pct >= 40 ? '#f59e0b' : '#ef4444';

    const improvRows = (data.improvements || []).map(i => `
      <div class="al-improvement-row">
        <div class="al-improvement-name">${esc(i.control)}<div class="al-improvement-cat">${esc(i.category || '')}</div></div>
        <div class="al-improvement-pct">${i.current}%</div>
      </div>`).join('');

    wrap.innerHTML = `
      <div class="al-score-wrap">
        <div class="al-score-ring-wrap">
          <svg class="al-score-ring" width="120" height="120" viewBox="0 0 120 120">
            <circle class="al-score-track" cx="60" cy="60" r="50"/>
            <circle class="al-score-fill" cx="60" cy="60" r="50"
              stroke="${scoreColor}" stroke-dasharray="${circ.toFixed(1)}" stroke-dashoffset="${offset.toFixed(1)}"/>
            <text x="60" y="56" class="al-score-pct-text" text-anchor="middle" dominant-baseline="middle" transform="rotate(90,60,60)">${pct}%</text>
            <text x="60" y="72" class="al-score-sub-text" text-anchor="middle" dominant-baseline="middle" transform="rotate(90,60,60)">score</text>
          </svg>
        </div>
        <div class="al-score-info">
          <div><span class="al-score-big">${data.currentScore ?? '—'}</span> <span class="al-score-max">/ ${data.maxScore ?? '—'} punten</span></div>
          <div class="al-score-updated">Bijgewerkt: ${fmtDate(data.createdAt)}</div>
          ${improvRows ? `<div class="al-improvements"><div class="al-improvements-title">Verbeterpunten</div>${improvRows}</div>` : ''}
        </div>
      </div>`;
  }

  // ── Sign-ins ──
  function loadSignIns() {
    const tid = getTid(); if (!tid) return;
    const wrap = document.getElementById('alSignInsWrap');
    if (!wrap) return;
    wrap.innerHTML = loading('Aanmeldingen laden…');
    apiFetch(`/api/alerts/${tid}/sign-ins`).then(data => renderSignIns(data))
      .catch(err => { wrap.innerHTML = `<p class="al-empty">Fout: ${esc(err.message)}</p>`; });
  }

  function renderSignIns(data) {
    const wrap = document.getElementById('alSignInsWrap');
    if (!wrap) return;
    _signInsData = data;
    renderWorkspaceSource(data);
    renderAlertsOverview();
    _showAlBanner(data._source);
    if (!data.ok) { wrap.innerHTML = `<p class="al-empty">${esc(data.error || 'Fout')}</p>`; return; }
    const items = data.items || [];
    if (!items.length) {
      const msg = data.message || 'Geen aanmeldingen gevonden (P1/P2-licentie vereist voor volledige log).';
      wrap.innerHTML = `<p class="al-empty">${esc(msg)}</p>`;
      return;
    }
    wrap.innerHTML = `
      <div class="al-table-wrap">
        <table class="al-table">
          <thead><tr><th>Tijdstip</th><th>Gebruiker</th><th>App</th><th>IP / Locatie</th><th>Risico</th><th>Status</th></tr></thead>
          <tbody>${items.map(i => `<tr>
            <td style="white-space:nowrap">${fmtDate(i.createdDateTime)}</td>
            <td>${esc(i.userPrincipalName || '—')}</td>
            <td>${esc(i.appDisplayName || '—')}</td>
            <td>${esc(i.ipAddress || '—')}${i.location ? ` <span style="font-size:.75rem;color:var(--text-muted)">${esc(i.location)}</span>` : ''}</td>
            <td>${riskBadge(i.riskLevel)}</td>
            <td>${esc(i.statusDetail || (i.status === 0 ? 'Geslaagd' : `Code ${i.status}`))}</td>
          </tr>`).join('')}</tbody>
        </table>
      </div>`;
  }

  // ── Notificatie configuratie ──
  function loadNotifConfig() {
    const tid = getTid(); if (!tid) return;
    apiFetch(`/api/alerts/${tid}/config`).then(data => {
      const cfg = data.config || {};
      const wh = document.getElementById('alWebhookUrl');
      const wt = document.getElementById('alWebhookType');
      const em = document.getElementById('alEmailAddr');
      if (wh) wh.value = cfg.webhook_url || '';
      if (wt) wt.value = cfg.webhook_type || 'teams';
      if (em) em.value = cfg.email_addr || '';
    }).catch(() => {});
  }

  function saveNotifConfig() {
    const tid = getTid(); if (!tid) return;
    const wh = document.getElementById('alWebhookUrl');
    const wt = document.getElementById('alWebhookType');
    const em = document.getElementById('alEmailAddr');
    const res = document.getElementById('alConfigResult');
    apiFetch(`/api/alerts/${tid}/config`, {
      method: 'POST',
      body: JSON.stringify({
        webhook_url:  wh ? wh.value.trim() : '',
        webhook_type: wt ? wt.value : 'teams',
        email_addr:   em ? em.value.trim() : '',
      })
    }).then(data => {
      if (res) {
        res.className = 'al-test-result ' + (data.ok ? 'al-test-result-ok' : 'al-test-result-err');
        res.textContent = data.ok ? 'Configuratie opgeslagen.' : (data.error || 'Fout');
        res.style.display = 'block';
        setTimeout(() => { res.style.display = 'none'; }, 3000);
      }
    }).catch(err => {
      if (res) { res.className = 'al-test-result al-test-result-err'; res.textContent = err.message; res.style.display = 'block'; }
    });
  }

  function testWebhook() {
    const tid = getTid(); if (!tid) return;
    const wh = document.getElementById('alWebhookUrl');
    const wt = document.getElementById('alWebhookType');
    const res = document.getElementById('alConfigResult');
    if (!wh || !wh.value.trim()) { alert('Vul eerst een webhook URL in.'); return; }
    apiFetch(`/api/alerts/${tid}/test-webhook`, {
      method: 'POST',
      body: JSON.stringify({ webhook_url: wh.value.trim(), webhook_type: wt ? wt.value : 'teams' })
    }).then(data => {
      if (res) {
        res.className = 'al-test-result ' + (data.ok ? 'al-test-result-ok' : 'al-test-result-err');
        res.textContent = data.ok ? '✓ Test bericht verzonden.' : ('Fout: ' + (data.error || 'Onbekend'));
        res.style.display = 'block';
      }
    }).catch(err => {
      if (res) { res.className = 'al-test-result al-test-result-err'; res.textContent = err.message; res.style.display = 'block'; }
    });
  }

  // ── Publieke ingang ──
  window.loadAlertsSection = function () {
    const tid = getTid();
    if (window._alLastTid !== tid) { _secureScore = null; _auditData = null; _signInsData = null; _tabsBound = false; window._alLastTid = tid; }
    bindAlTabs();

    // Config knopen binden
    const savBtn = document.getElementById('alBtnSaveConfig');
    if (savBtn && !savBtn._bound) { savBtn._bound = true; savBtn.addEventListener('click', saveNotifConfig); }
    const tstBtn = document.getElementById('alBtnTestWebhook');
    if (tstBtn && !tstBtn._bound) { tstBtn._bound = true; tstBtn.addEventListener('click', testWebhook); }

    const active = document.querySelector('#alertsSection .al-tab.active');
    switchAlTab(active ? active.dataset.alTab : 'auditlog');
  };
  window.switchAlertsTab = switchAlTab;
})();
