/**
 * Denjoy IT Platform — Fase 7: Domains Analyser
 * IIFE module — window.loadDomainsSection
 */
(function () {
  'use strict';

  let _domains = null;
  let _analyses = {};

  function getTid() { const s = document.getElementById('tenantSelect'); return s ? s.value : ''; }
  function esc(s) { if (s == null) return ''; return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }

  function apiFetch(url, opts = {}) {
    const token = localStorage.getItem('denjoy_auth_token') || localStorage.getItem('denjoy_token') || '';
    const headers = { 'Content-Type': 'application/json' };
    if (token) headers['Authorization'] = 'Bearer ' + token;
    return fetch(url, { credentials: 'include', headers, ...opts }).then(r => r.json());
  }

  function apiFetchCached(url, opts, ttlMs) {
    const get = window.cacheGet; const set = window.cacheSet;
    const ttl = ttlMs || (window.CACHE_TTL ? window.CACHE_TTL.domains : 300000);
    if (get) { const hit = get(url); if (hit !== null) return Promise.resolve(hit); }
    return apiFetch(url, opts).then(data => { if (data !== null && set) set(url, data, ttl); return data; });
  }

  function loading(msg, type = 'cards') {
    if (type === 'cards' && window.skeletonCards) return window.skeletonCards(4);
    return `<div class="dm-loading"><div class="dm-spinner"></div><span>${esc(msg)}</span></div>`;
  }

  // Score ring SVG (r=22, circumference=138.2)
  function scoreRing(score, size = 52) {
    const r = (size / 2) - 6;
    const circ = 2 * Math.PI * r;
    const pct = Math.min(100, Math.max(0, score));
    const offset = circ - (pct / 100) * circ;
    const cls = score >= 80 ? 'ok' : score >= 50 ? 'warn' : 'poor';
    return `
      <svg class="dm-score-ring" width="${size}" height="${size}" viewBox="0 0 ${size} ${size}">
        <circle class="dm-score-track" cx="${size/2}" cy="${size/2}" r="${r}"/>
        <circle class="dm-score-fill dm-score-fill-${cls}" cx="${size/2}" cy="${size/2}" r="${r}"
          stroke-dasharray="${circ}" stroke-dashoffset="${offset}"/>
        <text x="${size/2}" y="${size/2}" class="dm-score-text" text-anchor="middle" dominant-baseline="central" transform="rotate(90,${size/2},${size/2})">${score}</text>
      </svg>`;
  }

  // ── Laad domeinen ──
  function loadDomains() {
    const tid = getTid(); if (!tid) return;
    const wrap = document.getElementById('dmDomainsWrap');
    if (!wrap) return;
    const info = document.getElementById('dmDomainCount');
    wrap.innerHTML = loading('Domeinen laden…');
    if (info) info.textContent = '— domeinen';

    apiFetchCached(`/api/domains/${tid}/list`, {}, window.CACHE_TTL ? window.CACHE_TTL.domains : 300000).then(data => {
      _domains = data;
      renderDomainList(data);
    }).catch(err => { wrap.innerHTML = `<p class="dm-empty">Fout: ${esc(err.message)}</p>`; });
  }

  function renderDomainList(data) {
    const wrap = document.getElementById('dmDomainsWrap');
    const info = document.getElementById('dmDomainCount');
    const banner = document.getElementById('dmSnapshotBanner');
    if (!wrap) return;
    if (banner) {
      if (data._source === 'assessment_snapshot') {
        banner.style.display = '';
        banner.textContent = 'Gegevens uit laatste assessment. Live data vereist actieve verbinding.';
      } else {
        banner.style.display = 'none';
      }
    }
    if (!data.ok) { wrap.innerHTML = `<p class="dm-empty">${esc(data.error || 'Fout')}</p>`; return; }
    const domains = (data.domains || []).filter(d => !d.isInitial); // verberg .onmicrosoft.com
    if (info) info.textContent = `${domains.length} domeinen`;
    if (!domains.length) { wrap.innerHTML = '<p class="dm-empty">Geen aangepaste domeinen gevonden.</p>'; return; }

    wrap.innerHTML = `
      <div class="dm-domain-list">
        ${domains.map(d => {
          const cached = _analyses[d.id];
          const ringHtml = cached ? scoreRing(cached.score) : `<div class="dm-score-label">—</div>`;
          return `
            <div class="dm-domain-row" data-domain="${esc(d.id)}">
              <div class="dm-domain-name">${esc(d.id)}</div>
              <div class="dm-domain-badges">
                ${d.isDefault ? '<span class="dm-badge dm-badge-default">Standaard</span>' : ''}
                ${d.isVerified ? '<span class="dm-badge dm-badge-verified">Geverifieerd</span>' : ''}
              </div>
              <div class="dm-score-wrap">${ringHtml}</div>
              <button class="dm-btn dm-btn-primary" style="font-size:.75rem;padding:.25rem .65rem" data-analyse="${esc(d.id)}">Analyseer</button>
            </div>
            <div class="dm-analyse-result" data-result-for="${esc(d.id)}" style="display:none;padding:0 .25rem .5rem;"></div>`;
        }).join('')}
      </div>`;

    wrap.querySelectorAll('[data-analyse]').forEach(btn => {
      btn.addEventListener('click', () => analyseDomain(btn.dataset.analyse));
    });
    wrap.querySelectorAll('.dm-domain-row').forEach(row => {
      row.addEventListener('click', e => {
        if (e.target.closest('button')) return;
        const domain = row.dataset.domain;
        if (_analyses[domain]) showAnalysis(domain);
      });
    });

    // Toon al gecachede analyses
    Object.keys(_analyses).forEach(d => showAnalysis(d));
  }

  function analyseDomain(domain) {
    const tid = getTid(); if (!tid) return;
    const resultDiv = document.querySelector(`[data-result-for="${CSS.escape(domain)}"]`);
    if (resultDiv) { resultDiv.style.display = ''; resultDiv.innerHTML = loading(`${domain} analyseren via DNS…`); }

    apiFetch(`/api/domains/${tid}/analyse?domain=${encodeURIComponent(domain)}`).then(data => {
      _analyses[domain] = data;
      showAnalysis(domain);
      renderDomainList(_domains); // update score ring
    }).catch(err => {
      if (resultDiv) resultDiv.innerHTML = `<p class="dm-empty">Fout: ${esc(err.message)}</p>`;
    });
  }

  function showAnalysis(domain) {
    const data = _analyses[domain]; if (!data) return;
    const resultDiv = document.querySelector(`[data-result-for="${CSS.escape(domain)}"]`);
    if (!resultDiv) return;
    resultDiv.style.display = '';

    if (!data.ok) { resultDiv.innerHTML = `<p class="dm-empty">${esc(data.error || 'Fout')}</p>`; return; }

    const labelCls = { 'Uitstekend': 'excellent', 'Goed': 'good', 'Matig': 'moderate', 'Zwak': 'weak' }[data.label] || 'weak';
    const checks = data.checks || [];

    const checkCards = checks.map(c => {
      const pct = c.maxScore > 0 ? Math.round((c.score / c.maxScore) * 100) : 0;
      const iconMap = { SPF: '📋', DMARC: '🛡️', DKIM: '🔑', MX: '📬' };
      const icon = iconMap[c.name] || '🔍';
      return `
        <div class="dm-check-card">
          <div class="dm-check-header">
            <span class="dm-check-icon">${icon}</span>
            <span class="dm-check-name">${esc(c.name)}</span>
            <div class="dm-check-score-bar-wrap">
              <div class="dm-check-score-bar dm-check-score-bar-${c.status === 'ok' ? 'ok' : c.status === 'warn' ? 'warn' : 'missing'}" style="width:${pct}%"></div>
            </div>
            <span class="dm-check-pts">${c.score}/${c.maxScore} pt</span>
            <div class="dm-check-status"><span class="dm-check-badge dm-check-badge-${esc(c.status)}">${esc(c.status)}</span></div>
          </div>
          ${c.detail ? `<div class="dm-check-detail">${esc(c.detail)}</div>` : ''}
          ${c.hint ? `<div class="dm-check-hint">💡 ${esc(c.hint)}</div>` : ''}
        </div>`;
    }).join('');

    resultDiv.innerHTML = `
      <div class="dm-analyse-panel">
        <div class="dm-analyse-header">
          <span class="dm-analyse-domain">${esc(data.domain)}</span>
          <div>
            <span class="dm-analyse-score-big">${data.score}</span>
            <span class="dm-analyse-score-max">/${data.maxScore}</span>
          </div>
          <span class="dm-analyse-label dm-analyse-label-${esc(labelCls)}">${esc(data.label)}</span>
        </div>
        <div class="dm-checks">${checkCards}</div>
      </div>`;
  }

  // ── Publieke ingang ──
  window.loadDomainsSection = function () {
    const tid = getTid();
    if (window._dmLastTid !== tid) { _domains = null; _analyses = {}; window._dmLastTid = tid; }
    const r = document.getElementById('dmBtnRefresh');
    if (r && !r._bound) { r._bound = true; r.addEventListener('click', () => { _domains = null; _analyses = {}; loadDomains(); }); }
    if (!_domains) loadDomains();
  };
})();
