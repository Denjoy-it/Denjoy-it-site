/**
 * Denjoy IT Portal — Bevindingen & Health module
 * Laadt /api/findings/{tid}/health en /api/findings/{tid}/list
 */

(function () {
  'use strict';

  let _loaded = false;
  let _allFindings = [];

  const DOMAIN_LABELS = {
    identity:      'Identity',
    appregs:       'App Registraties',
    exchange:      'Exchange',
    collaboration: 'Samenwerking',
    ca:            'Conditional Access',
  };

  const STATUS_ORDER = { critical: 0, warning: 1, info: 2, ok: 3 };

  function getTid() {
    return (typeof currentTenantId !== 'undefined') ? currentTenantId : null;
  }

  function getToken() {
    return localStorage.getItem('denjoy_token') || '';
  }

  async function apiFetch(path) {
    const res = await fetch(path, {
      headers: { 'Authorization': `Bearer ${getToken()}` },
      credentials: 'include',
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    return res.json();
  }

  // ── Health score ───────────────────────────────────────────────────────────

  function buildDomainBreakdown(findings) {
    const map = {};
    for (const f of findings) {
      if (!map[f.domain]) map[f.domain] = { total: 0, ok: 0, warn: 0 };
      map[f.domain].total++;
      if (f.status === 'ok') map[f.domain].ok++;
      if (f.status === 'warning') map[f.domain].warn++;
    }
    const result = {};
    for (const [d, v] of Object.entries(map)) {
      result[d] = {
        total: v.total,
        score: v.total ? Math.round((v.ok * 1.0 + v.warn * 0.5) / v.total * 100) : null,
      };
    }
    return result;
  }

  function renderHealthBar(health) {
    const scoreEl = document.getElementById('bevHealthScore');
    const pillsEl = document.getElementById('bevDomainPills');
    if (!scoreEl || !pillsEl) return;

    const score = health.score ?? null;
    const scoreClass = score === null ? '' : score >= 80 ? 'bev-score--ok' : score >= 50 ? 'bev-score--warn' : 'bev-score--crit';
    scoreEl.textContent = score !== null ? `${score}%` : '—';
    scoreEl.className = `bev-health-score ${scoreClass}`;

    const domains = buildDomainBreakdown(health.findings || []);
    pillsEl.innerHTML = Object.entries(domains).map(([domain, info]) => {
      const label = DOMAIN_LABELS[domain] || domain;
      const s = info.score ?? 0;
      const cls = s >= 80 ? 'bev-pill--ok' : s >= 50 ? 'bev-pill--warn' : 'bev-pill--crit';
      return `<span class="bev-domain-pill ${cls}" title="${label}: ${s}% (${info.total} controls)">${label} <strong>${s}%</strong></span>`;
    }).join('');
  }

  // ── Findings table ────────────────────────────────────────────────────────

  function domainLabel(d) { return DOMAIN_LABELS[d] || d; }

  function statusBadge(s) {
    const map = { ok: 'ok', warning: 'warn', critical: 'crit', info: 'info' };
    const cls = map[s] || 'neutral';
    const labels = { ok: 'OK', warning: 'Waarschuwing', critical: 'Kritiek', info: 'Info' };
    return `<span class="live-badge live-badge-${cls}">${labels[s] || s}</span>`;
  }

  function impactBadge(i) {
    const map = { high: 'crit', medium: 'warn', low: 'neutral' };
    const cls = map[i] || 'neutral';
    const labels = { high: 'Hoog', medium: 'Middel', low: 'Laag' };
    return `<span class="live-badge live-badge-${cls} live-badge-sm">${labels[i] || i}</span>`;
  }

  function formatDate(iso) {
    if (!iso) return '—';
    try {
      return new Date(iso).toLocaleString('nl-NL', { dateStyle: 'short', timeStyle: 'short' });
    } catch (_) { return iso; }
  }

  function renderTable(findings) {
    const tbody = document.getElementById('bevFindingsBody');
    const countEl = document.getElementById('bevCount');
    if (!tbody) return;

    if (!findings || findings.length === 0) {
      tbody.innerHTML = `<tr><td colspan="7" class="bev-empty">Geen bevindingen gevonden voor de geselecteerde filters.</td></tr>`;
      if (countEl) countEl.textContent = '';
      return;
    }

    const sorted = [...findings].sort((a, b) =>
      (STATUS_ORDER[a.status] ?? 99) - (STATUS_ORDER[b.status] ?? 99)
    );

    tbody.innerHTML = sorted.map(f => `
      <tr class="bev-row bev-row--${f.status || 'info'}">
        <td>${statusBadge(f.status)}</td>
        <td><span class="bev-domain-tag">${domainLabel(f.domain)}</span></td>
        <td class="bev-control">${escapeHtml(f.control || '')}</td>
        <td class="bev-finding">${escapeHtml(f.finding || f.title || '')}</td>
        <td>${impactBadge(f.impact)}</td>
        <td class="bev-recommendation">${f.recommendation ? escapeHtml(f.recommendation) : '<span class="bev-na">—</span>'}</td>
        <td class="bev-date">${formatDate(f.scanned_at)}</td>
      </tr>
    `).join('');

    if (countEl) countEl.textContent = `${findings.length} bevinding${findings.length !== 1 ? 'en' : ''}`;
  }

  function applyFilters() {
    const domain = document.getElementById('bevDomainFilter')?.value || '';
    const status = document.getElementById('bevStatusFilter')?.value || '';
    let filtered = _allFindings;
    if (domain) filtered = filtered.filter(f => f.domain === domain);
    if (status) filtered = filtered.filter(f => f.status === status);
    renderTable(filtered);
  }

  // ── Load ──────────────────────────────────────────────────────────────────

  async function loadBevindingenSection() {
    const tid = getTid();
    if (!tid) {
      renderTable([]);
      const scoreEl = document.getElementById('bevHealthScore');
      if (scoreEl) scoreEl.textContent = '—';
      const pillsEl = document.getElementById('bevDomainPills');
      if (pillsEl) pillsEl.innerHTML = '';
      const countEl = document.getElementById('bevCount');
      if (countEl) countEl.textContent = 'Geen tenant geselecteerd.';
      return;
    }

    // Loading state
    const tbody = document.getElementById('bevFindingsBody');
    if (tbody) tbody.innerHTML = `<tr><td colspan="7" class="bev-empty bev-loading">Bevindingen laden…</td></tr>`;

    try {
      const [health, list] = await Promise.all([
        apiFetch(`/api/findings/${tid}/health`),
        apiFetch(`/api/findings/${tid}/list`),
      ]);

      renderHealthBar(health);

      _allFindings = list.findings || [];
      applyFilters();
      _loaded = true;
    } catch (err) {
      if (tbody) tbody.innerHTML = `<tr><td colspan="7" class="bev-empty bev-error">Fout bij laden van bevindingen: ${escapeHtml(String(err.message || err))}</td></tr>`;
      console.error('[bevindingen] load error', err);
    }
  }

  // ── Filter wiring ─────────────────────────────────────────────────────────

  function wireFilters() {
    const domainSel = document.getElementById('bevDomainFilter');
    const statusSel = document.getElementById('bevStatusFilter');
    const refreshBtn = document.getElementById('bevRefreshBtn');

    if (domainSel && !domainSel._bevWired) {
      domainSel.addEventListener('change', applyFilters);
      domainSel._bevWired = true;
    }
    if (statusSel && !statusSel._bevWired) {
      statusSel.addEventListener('change', applyFilters);
      statusSel._bevWired = true;
    }
    if (refreshBtn && !refreshBtn._bevWired) {
      refreshBtn.addEventListener('click', () => { _loaded = false; loadBevindingenSection(); });
      refreshBtn._bevWired = true;
    }
    const importBtn = document.getElementById('bevImportBtn');
    if (importBtn && !importBtn._bevWired) {
      importBtn.addEventListener('click', importFromSnapshot);
      importBtn._bevWired = true;
    }
  }

  async function importFromSnapshot() {
    const tid = getTid();
    if (!tid) {
      if (typeof showToast === 'function') showToast('Selecteer eerst een tenant.', 'warning');
      return;
    }
    const importBtn = document.getElementById('bevImportBtn');
    if (importBtn) { importBtn.disabled = true; importBtn.textContent = '⟳ Bezig…'; }
    try {
      const res = await fetch(`/api/findings/${tid}/import-snapshot`, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${getToken()}`,
          'Content-Type': 'application/json',
          'X-CSRF-Token': document.cookie.match(/denjoy_csrf=([^;]+)/)?.[1] || '',
        },
        credentials: 'include',
        body: '{}',
      });
      const data = await res.json();
      if (data.ok) {
        if (typeof showToast === 'function') showToast(`${data.findings_written} bevindingen geïmporteerd uit assessment-snapshot.`, 'success');
        _loaded = false;
        await loadBevindingenSection();
      } else {
        if (typeof showToast === 'function') showToast(data.error || 'Import mislukt.', 'error');
      }
    } catch (err) {
      if (typeof showToast === 'function') showToast('Import mislukt: ' + String(err.message || err), 'error');
    } finally {
      if (importBtn) { importBtn.disabled = false; importBtn.textContent = '⬇ Importeer uit assessment'; }
    }
  }

  // ── Public API ────────────────────────────────────────────────────────────

  function init() {
    wireFilters();
  }

  document.addEventListener('DOMContentLoaded', init);

  window.loadBevindingenSection = loadBevindingenSection;

  // Reload when tenant switches
  document.addEventListener('tenantChanged', () => {
    _loaded = false;
    _allFindings = [];
    if (typeof window._currentSection !== 'undefined' && window._currentSection === 'bevindingen') {
      loadBevindingenSection();
    }
  });

  // Helper — may already be defined globally
  function escapeHtml(str) {
    if (typeof window.escapeHtml === 'function') return window.escapeHtml(str);
    return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  }
})();
