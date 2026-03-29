/**
 * Denjoy IT Platform — Baseline & Gold Tenant module (Fase 3)
 * Desired State Engine: beheer baselines, Gold Tenant export,
 * compliance checks en toepassing op tenants.
 */
(function () {
  'use strict';

  // ── State ────────────────────────────────────────────────────────────────

  let _baselines    = [];
  let _assignments  = [];
  let _activeTab    = 'baselines';   // 'baselines' | 'gold' | 'assignments' | 'history'

  // ── API helper ───────────────────────────────────────────────────────────

  function blApiFetch(url, opts = {}) {
    const token = localStorage.getItem('denjoy_auth_token') || '';
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

  function escHtml(s) {
    if (!s) return '';
    return String(s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function fmtDate(iso) {
    if (!iso) return '—';
    try { return new Date(iso).toLocaleDateString('nl-NL'); } catch { return iso; }
  }

  function fmtDateTime(iso) {
    if (!iso) return '—';
    try { return new Date(iso).toLocaleString('nl-NL'); } catch { return iso; }
  }

  function scoreClass(score) {
    if (score === null || score === undefined) return 'assigned';
    if (score >= 90) return 'good';
    if (score >= 60) return 'warn';
    return 'bad';
  }

  function catLabel(key) {
    const map = {
      security_defaults: 'Security Defaults',
      conditional_access: 'Conditional Access',
      auth_methods: 'Auth Methods',
      org_settings: 'Org Instellingen',
      named_locations: 'Named Locations',
      exchange: 'Exchange',
    };
    return map[key] || key;
  }

  // ── Tabs ─────────────────────────────────────────────────────────────────

  function switchTab(tab) {
    _activeTab = tab;
    document.querySelectorAll('.bl-tab[data-tab]').forEach((btn) => {
      btn.classList.toggle('active', btn.dataset.tab === tab);
    });
    document.querySelectorAll('.bl-tab-panel[data-panel]').forEach((p) => {
      p.style.display = p.dataset.panel === tab ? '' : 'none';
    });
    if (tab === 'baselines') loadBaselines();
    if (tab === 'gold')      renderGoldPanel();
    if (tab === 'assignments') loadAssignments();
    if (tab === 'history') loadBaselineHistory();
  }

  // ── Baseline bibliotheek ─────────────────────────────────────────────────

  function loadBaselines() {
    const grid = document.getElementById('blBaselineGrid');
    if (!grid) return;
    grid.innerHTML = `<div class="bl-empty"><div class="bl-empty-icon">⏳</div><div>Laden...</div></div>`;

    blApiFetch('/api/baselines')
      .then((data) => {
        _baselines = data.items || [];
        renderBaselineGrid();
      })
      .catch(() => {
        if (grid) grid.innerHTML = `<div class="bl-empty">Fout bij laden baselines.</div>`;
      });
  }

  function renderBaselineGrid() {
    const grid  = document.getElementById('blBaselineGrid');
    const count = document.getElementById('blBaselineCount');
    if (!grid) return;

    const q = (document.getElementById('blSearchInput')?.value || '').toLowerCase();
    let list = _baselines;
    if (q) list = list.filter((b) => (b.name || '').toLowerCase().includes(q) || (b.description || '').toLowerCase().includes(q));

    if (count) count.textContent = `${list.length} baseline${list.length !== 1 ? 's' : ''}`;

    if (!list.length) {
      grid.innerHTML = `<div class="bl-empty">
        <div class="bl-empty-icon">📋</div>
        <div class="bl-empty-title">${q ? 'Geen baselines gevonden' : 'Nog geen baselines'}</div>
        <div>${q ? 'Pas de zoekopdracht aan.' : 'Gebruik Gold Tenant export of maak handmatig een baseline aan.'}</div>
      </div>`;
      return;
    }

    grid.innerHTML = list.map((b) => {
      const cats = (b.categories || []).map((c) =>
        `<span class="bl-cat-chip ${escHtml(c)}">${escHtml(catLabel(c))}</span>`
      ).join('');

      return `<div class="bl-card" data-bid="${escHtml(b.id)}">
        <div class="bl-card-header">
          <div>
            <div class="bl-card-title">${escHtml(b.name)}</div>
            ${b.source_tenant_name ? `<div class="bl-card-source">Bron: ${escHtml(b.source_tenant_name)}</div>` : ''}
          </div>
        </div>
        <div class="bl-card-desc">${escHtml(b.description || 'Geen beschrijving.')}</div>
        <div class="bl-cat-chips">${cats || '<span class="bl-cat-chip">Geen categorieën</span>'}</div>
        <div class="bl-card-footer">
          <div class="bl-card-meta">${fmtDate(b.created_at)} · ${escHtml(b.created_by || '—')}</div>
          <div class="bl-card-actions">
            <button class="bl-btn bl-btn-ghost bl-btn-detail" data-bid="${escHtml(b.id)}" data-name="${escHtml(b.name)}">Bekijk</button>
            <button class="bl-btn bl-btn-ghost bl-btn-delete-bl" data-bid="${escHtml(b.id)}" data-name="${escHtml(b.name)}">✕</button>
          </div>
        </div>
      </div>`;
    }).join('');

    grid.querySelectorAll('.bl-card').forEach((card) => {
      card.addEventListener('click', (e) => {
        if (e.target.closest('.bl-card-actions')) return;
        openBaselineDetail(card.dataset.bid);
      });
    });
    grid.querySelectorAll('.bl-btn-detail').forEach((btn) => {
      btn.addEventListener('click', (e) => { e.stopPropagation(); openBaselineDetail(btn.dataset.bid); });
    });
    grid.querySelectorAll('.bl-btn-delete-bl').forEach((btn) => {
      btn.addEventListener('click', (e) => { e.stopPropagation(); confirmDeleteBaseline(btn.dataset.bid, btn.dataset.name); });
    });
  }

  // ── Baseline detail modal ─────────────────────────────────────────────────

  function openBaselineDetail(baselineId) {
    const overlay = createOverlay(closeAllModals);
    const modal   = document.createElement('div');
    modal.className = 'bl-modal bl-modal-wide';
    modal.innerHTML = `
      <div class="bl-modal-header">
        <div>
          <div class="bl-modal-title">Baseline details</div>
          <div class="bl-modal-subtitle" id="blDetSubtitle">Laden...</div>
        </div>
        <button class="bl-modal-close">✕</button>
      </div>
      <div class="bl-modal-body" id="blDetBody">
        <div style="text-align:center;padding:2rem;color:var(--text-muted)">Laden...</div>
      </div>
      <div class="bl-modal-footer">
        <button class="bl-btn bl-btn-secondary bl-close-btn">Sluiten</button>
        <button class="bl-btn bl-btn-primary bl-assign-btn">Koppelen aan tenant</button>
      </div>`;
    modal.querySelector('.bl-modal-close').onclick = closeAllModals;
    modal.querySelector('.bl-close-btn').onclick   = closeAllModals;

    overlay.appendChild(modal);
    document.body.appendChild(overlay);

    blApiFetch(`/api/baselines/${baselineId}`)
      .then((b) => {
        modal.querySelector('#blDetSubtitle').textContent = b.name || baselineId;
        const cfg = b.config || {};
        const cats = Object.keys(cfg.categories || {});

        const catSummary = cats.map((c) => {
          const data = cfg.categories[c] || {};
          let detail = '';
          if (c === 'security_defaults') detail = `isEnabled: ${data.isEnabled}`;
          if (c === 'conditional_access') detail = `${(data.policies || []).length} polic(ies)`;
          if (c === 'auth_methods') detail = `${(data.methods || []).length} methoden`;
          if (c === 'named_locations') detail = `${data.count || 0} locaties`;
          return `<div style="display:flex;justify-content:space-between;align-items:center;padding:.5rem 0;border-bottom:1px solid var(--border-color,#e2e8f0)">
            <span style="font-size:.875rem;color:var(--text-primary)">${escHtml(catLabel(c))}</span>
            <span style="font-size:.78rem;color:var(--text-muted)">${escHtml(detail)}</span>
          </div>`;
        }).join('');

        modal.querySelector('#blDetBody').innerHTML = `
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:.75rem 1.5rem;font-size:.83rem">
            <div><div style="font-size:.72rem;font-weight:600;color:var(--text-muted);text-transform:uppercase;letter-spacing:.04em">Aangemaakt</div>
              <div style="color:var(--text-primary)">${fmtDateTime(b.created_at)}</div></div>
            <div><div style="font-size:.72rem;font-weight:600;color:var(--text-muted);text-transform:uppercase;letter-spacing:.04em">Door</div>
              <div style="color:var(--text-primary)">${escHtml(b.created_by || '—')}</div></div>
            <div style="grid-column:1/-1"><div style="font-size:.72rem;font-weight:600;color:var(--text-muted);text-transform:uppercase;letter-spacing:.04em">Beschrijving</div>
              <div style="color:var(--text-secondary)">${escHtml(b.description || '—')}</div></div>
            ${b.source_tenant_name ? `<div style="grid-column:1/-1"><div style="font-size:.72rem;font-weight:600;color:var(--text-muted);text-transform:uppercase;letter-spacing:.04em">Bron tenant</div>
              <div style="color:var(--text-primary)">${escHtml(b.source_tenant_name)}</div></div>` : ''}
          </div>
          <div style="margin-top:.5rem">
            <div style="font-size:.82rem;font-weight:600;color:var(--text-secondary);margin-bottom:.5rem">Geconfigureerde categorieën</div>
            ${catSummary || '<div style="color:var(--text-muted);font-size:.83rem">Geen categorieën gevonden.</div>'}
          </div>`;

        modal.querySelector('.bl-assign-btn').onclick = () => { closeAllModals(); openAssignModal(baselineId, b.name); };
      })
      .catch((err) => {
        modal.querySelector('#blDetBody').innerHTML = `<div style="color:#dc2626">Fout: ${escHtml(String(err))}</div>`;
      });
  }

  // ── Assign modal ──────────────────────────────────────────────────────────

  function openAssignModal(baselineId, baselineName) {
    const overlay = createOverlay(closeAllModals);
    const modal   = document.createElement('div');
    modal.className = 'bl-modal';

    // Tenants ophalen voor dropdown
    blApiFetch('/api/tenants').then((data) => {
      const tenants = data.tenants || data.items || [];
      const opts = tenants.map((t) =>
        `<option value="${escHtml(t.id)}">${escHtml(t.customer_name)} (${escHtml(t.tenant_name)})</option>`
      ).join('');

      modal.innerHTML = `
        <div class="bl-modal-header">
          <div>
            <div class="bl-modal-title">Baseline koppelen</div>
            <div class="bl-modal-subtitle">${escHtml(baselineName)}</div>
          </div>
          <button class="bl-modal-close">✕</button>
        </div>
        <div class="bl-modal-body">
          <div class="bl-form">
            <div class="bl-form-group">
              <label class="bl-form-label">Tenant <span class="bl-required">*</span></label>
              <select class="bl-form-select" id="blAssignTenant">
                <option value="">Selecteer tenant...</option>
                ${opts}
              </select>
            </div>
          </div>
          <div id="blAssignResult"></div>
        </div>
        <div class="bl-modal-footer">
          <button class="bl-btn bl-btn-secondary bl-cancel-btn">Annuleren</button>
          <button class="bl-btn bl-btn-primary bl-do-assign">Koppelen</button>
        </div>`;

      modal.querySelector('.bl-modal-close').onclick = closeAllModals;
      modal.querySelector('.bl-cancel-btn').onclick  = closeAllModals;
      modal.querySelector('.bl-do-assign').onclick = () => {
        const tid = modal.querySelector('#blAssignTenant').value;
        if (!tid) return;
        const resultDiv = modal.querySelector('#blAssignResult');
        resultDiv.innerHTML = `<div style="font-size:.82rem;color:var(--text-muted)">Koppelen...</div>`;
        blApiFetch(`/api/baselines/${baselineId}/assign`, {
          method: 'POST', body: JSON.stringify({ tenant_id: tid }),
        })
          .then(() => {
            resultDiv.innerHTML = `<div class="bl-result bl-result-ok"><div class="bl-result-icon">✅</div><div class="bl-result-msg">Baseline succesvol gekoppeld</div></div>`;
            setTimeout(closeAllModals, 1200);
          })
          .catch((err) => {
            resultDiv.innerHTML = `<div class="bl-result bl-result-error"><div class="bl-result-msg">Fout: ${escHtml(String(err))}</div></div>`;
          });
      };
    });

    overlay.appendChild(modal);
    document.body.appendChild(overlay);
  }

  // ── Verwijder bevestiging ─────────────────────────────────────────────────

  function confirmDeleteBaseline(baselineId, name) {
    if (!confirm(`Baseline '${name}' verwijderen?\nAlle koppelingen worden ook verwijderd.`)) return;
    blApiFetch(`/api/baselines/${baselineId}`, { method: 'DELETE' })
      .then(() => loadBaselines())
      .catch((err) => alert(`Fout: ${err}`));
  }

  // ── Gold Tenant export panel ─────────────────────────────────────────────

  function renderGoldPanel() {
    const panel = document.getElementById('blGoldPanel');
    if (!panel) return;

    blApiFetch('/api/tenants').then((data) => {
      const tenants = data.tenants || data.items || [];
      const opts = tenants.map((t) =>
        `<option value="${escHtml(t.id)}">${escHtml(t.customer_name)} (${escHtml(t.tenant_name)})</option>`
      ).join('');

      panel.innerHTML = `
        <div class="bl-gold-card">
          <div class="bl-gold-icon">🏆</div>
          <div>
            <div class="bl-gold-title">Gold Tenant Export</div>
            <div class="bl-gold-desc">
              Selecteer een goed geconfigureerde "Gold Tenant" als referentie.<br>
              De huidige beveiligingsinstellingen worden geëxporteerd als baseline en kunnen
              vervolgens worden toegepast op andere tenants.
            </div>
          </div>
        </div>
        <div class="bl-form" style="max-width:500px">
          <div class="bl-form-group">
            <label class="bl-form-label">Gold Tenant (brontenant) <span class="bl-required">*</span></label>
            <select class="bl-form-select" id="blGoldTenant">
              <option value="">Selecteer Gold Tenant...</option>
              ${opts}
            </select>
          </div>
          <div class="bl-form-group">
            <label class="bl-form-label">Baseline naam</label>
            <input class="bl-form-input" id="blGoldName" placeholder="Bijv. Denjoy Beveiligingsstandaard 2026">
          </div>
          <div class="bl-form-group">
            <label class="bl-form-label">Beschrijving</label>
            <textarea class="bl-form-textarea" id="blGoldDesc" placeholder="Optionele beschrijving van de baseline"></textarea>
          </div>
          <div style="margin-top:.25rem">
            <button class="bl-btn bl-btn-primary" id="blBtnExport">⬇ Exporteren als baseline</button>
          </div>
        </div>
        <div id="blGoldResult" style="margin-top:1rem;max-width:500px"></div>`;

      panel.querySelector('#blBtnExport').addEventListener('click', () => {
        const tid  = panel.querySelector('#blGoldTenant').value;
        const name = (panel.querySelector('#blGoldName').value || '').trim();
        const desc = (panel.querySelector('#blGoldDesc').value || '').trim();
        if (!tid) { alert('Selecteer een Gold Tenant.'); return; }

        const resultDiv = panel.querySelector('#blGoldResult');
        resultDiv.innerHTML = `<div style="font-size:.83rem;color:var(--text-muted)">Configuratie exporteren... (kan 30s duren)</div>`;

        blApiFetch(`/api/baselines/export/${tid}`, {
          method: 'POST',
          body: JSON.stringify({ name, description: desc }),
        })
          .then((b) => {
            const cats = (b.categories || []).map((c) => catLabel(c)).join(', ');
            resultDiv.innerHTML = `<div class="bl-result bl-result-ok">
              <div class="bl-result-icon">✅</div>
              <div class="bl-result-msg">Baseline aangemaakt: ${escHtml(b.name)}</div>
              <div style="font-size:.78rem;margin-top:.25rem;color:var(--text-secondary)">${escHtml(cats)}</div>
            </div>
            <div style="margin-top:.75rem">
              <button class="bl-btn bl-btn-secondary bl-goto-baselines">Naar Baselines →</button>
            </div>`;
            resultDiv.querySelector('.bl-goto-baselines').onclick = () => switchTab('baselines');
          })
          .catch((err) => {
            resultDiv.innerHTML = `<div class="bl-result bl-result-error"><div class="bl-result-msg">Export mislukt: ${escHtml(String(err))}</div></div>`;
          });
      });
    }).catch(() => {
      panel.innerHTML = `<div style="color:var(--text-muted)">Tenants konden niet worden geladen.</div>`;
    });
  }

  // ── Assignments overzicht ────────────────────────────────────────────────

  function loadAssignments() {
    const tbody = document.getElementById('blAssignBody');
    if (!tbody) return;
    tbody.innerHTML = `<tr><td colspan="7" style="text-align:center;padding:2rem;color:var(--text-muted)">Laden...</td></tr>`;

    blApiFetch('/api/baselines/assignments/all')
      .then((data) => {
        _assignments = data.items || [];
        renderAssignmentsTable();
      })
      .catch(() => {
        tbody.innerHTML = `<tr><td colspan="7" style="text-align:center;padding:2rem;color:#dc2626">Fout bij laden.</td></tr>`;
      });
  }

  function renderAssignmentsTable() {
    const tbody = document.getElementById('blAssignBody');
    if (!tbody) return;

    if (!_assignments.length) {
      tbody.innerHTML = `<tr><td colspan="7" style="text-align:center;padding:2rem;color:var(--text-muted)">Nog geen baselines gekoppeld aan tenants.</td></tr>`;
      return;
    }

    tbody.innerHTML = _assignments.map((a) => {
      const score = a.compliance_score;
      const cls   = scoreClass(score);
      const hasScore = score !== null && score !== undefined;

      const scoreCellHtml = hasScore
        ? `<div class="bl-score-bar-wrap">
            <div class="bl-score-bar"><div class="bl-score-bar-fill ${cls}" style="width:${score}%"></div></div>
            <span class="bl-score-pct">${score}%</span>
           </div>`
        : '<span style="color:var(--text-muted);font-size:.78rem">—</span>';

      const statusMap = { compliant: 'bl-status-compliant', non_compliant: 'bl-status-non_compliant', partial: 'bl-status-partial', assigned: 'bl-status-assigned', applied: 'bl-status-applied' };
      const statusLabel = { compliant: 'Compliant', non_compliant: 'Non-compliant', partial: 'Gedeeltelijk', assigned: 'Gekoppeld', applied: 'Toegepast' };
      const statusCls   = statusMap[a.status] || 'bl-status-assigned';

      return `<tr>
        <td style="font-weight:600;color:var(--text-primary)">${escHtml(a.baseline_name || '—')}</td>
        <td>${escHtml(a.tenant_name || '—')}</td>
        <td><span class="bl-status ${statusCls}">${escHtml(statusLabel[a.status] || a.status)}</span></td>
        <td>${scoreCellHtml}</td>
        <td style="font-size:.75rem;color:var(--text-muted)">${fmtDate(a.last_checked_at)}</td>
        <td style="font-size:.75rem;color:var(--text-muted)">${fmtDate(a.last_applied_at)}</td>
        <td>
          <div style="display:flex;gap:.35rem">
            <button class="bl-btn bl-btn-ghost bl-btn-check"
              data-bid="${escHtml(a.baseline_id)}" data-tid="${escHtml(a.tenant_id)}"
              data-bname="${escHtml(a.baseline_name)}" data-tname="${escHtml(a.tenant_name)}">Check</button>
            <button class="bl-btn bl-btn-ghost bl-btn-apply"
              data-bid="${escHtml(a.baseline_id)}" data-tid="${escHtml(a.tenant_id)}"
              data-bname="${escHtml(a.baseline_name)}" data-tname="${escHtml(a.tenant_name)}">Toepassen</button>
          </div>
        </td>
      </tr>`;
    }).join('');

    tbody.querySelectorAll('.bl-btn-check').forEach((btn) => {
      btn.addEventListener('click', () => openCheckModal(btn.dataset.bid, btn.dataset.tid, btn.dataset.bname, btn.dataset.tname));
    });
    tbody.querySelectorAll('.bl-btn-apply').forEach((btn) => {
      btn.addEventListener('click', () => openApplyModal(btn.dataset.bid, btn.dataset.tid, btn.dataset.bname, btn.dataset.tname));
    });
  }

  // ── Compliance check modal ─────────────────────────────────────────────────

  function openCheckModal(baselineId, tenantId, baselineName, tenantName) {
    const overlay = createOverlay(closeAllModals);
    const modal   = document.createElement('div');
    modal.className = 'bl-modal bl-modal-wide';
    modal.innerHTML = `
      <div class="bl-modal-header">
        <div>
          <div class="bl-modal-title">Compliance check</div>
          <div class="bl-modal-subtitle">${escHtml(baselineName)} → ${escHtml(tenantName)}</div>
        </div>
        <button class="bl-modal-close">✕</button>
      </div>
      <div class="bl-modal-body" id="blCheckBody">
        <div style="text-align:center;padding:2rem;color:var(--text-muted)">Check uitvoeren... (kan 30s duren)</div>
      </div>
      <div class="bl-modal-footer">
        <button class="bl-btn bl-btn-secondary bl-close-btn">Sluiten</button>
      </div>`;
    modal.querySelector('.bl-modal-close').onclick = closeAllModals;
    modal.querySelector('.bl-close-btn').onclick   = closeAllModals;
    overlay.appendChild(modal);
    document.body.appendChild(overlay);

    blApiFetch(`/api/baselines/${baselineId}/check/${tenantId}`, { method: 'POST', body: '{}' })
      .then((data) => {
        const body    = modal.querySelector('#blCheckBody');
        const score   = data.score ?? 0;
        const cls     = scoreClass(score);
        const compliance = data.compliance || {};
        const findings   = compliance.findings || [];

        const circumference = 2 * Math.PI * 26; // r=26
        const offset = circumference - (score / 100) * circumference;

        const findingsRows = findings.map((f) => {
          const isOk = f.status === 'compliant';
          return `<tr>
            <td>${escHtml(f.category)}</td>
            <td style="font-size:.78rem;font-family:monospace">${escHtml(f.check)}</td>
            <td><span class="bl-status bl-status-${escHtml(f.status)}">${isOk ? 'Compliant' : 'Non-compliant'}</span></td>
            <td style="color:var(--text-muted);font-size:.78rem">${escHtml(String(f.want ?? '—'))}</td>
            <td style="color:var(--text-muted);font-size:.78rem">${escHtml(String(f.have ?? '—'))}</td>
            <td style="font-size:.78rem">${escHtml(f.message || '')}</td>
          </tr>`;
        }).join('');

        body.innerHTML = `
          <div class="bl-score-wrap">
            <div class="bl-score-ring">
              <svg class="bl-score-ring-svg" width="64" height="64" viewBox="0 0 64 64">
                <circle class="bl-score-ring-bg" cx="32" cy="32" r="26"/>
                <circle class="bl-score-ring-fg ${cls}" cx="32" cy="32" r="26"
                  stroke-dasharray="${circumference}" stroke-dashoffset="${offset}"/>
              </svg>
            </div>
            <div>
              <div class="bl-score-num">${score}%</div>
              <div class="bl-score-label">Compliance score</div>
            </div>
            <div class="bl-score-detail">
              <strong>${data.compliant || 0}</strong> van ${data.total_checks || 0} checks compliant
            </div>
          </div>
          <div class="bl-findings-wrap">
            <table class="bl-findings-table">
              <thead><tr>
                <th>Categorie</th><th>Check</th><th>Status</th><th>Verwacht</th><th>Huidig</th><th>Bericht</th>
              </tr></thead>
              <tbody>${findingsRows || '<tr><td colspan="6" class="bl-findings-table-empty">Geen bevindingen.</td></tr>'}</tbody>
            </table>
          </div>`;

        const footer = modal.querySelector('.bl-modal-footer');
        footer.innerHTML = `
          <button class="bl-btn bl-btn-secondary bl-close-btn">Sluiten</button>
          ${data.non_compliant > 0 ? `<button class="bl-btn bl-btn-primary bl-apply-now">Toepassen →</button>` : ''}`;
        footer.querySelector('.bl-close-btn').onclick = () => { closeAllModals(); loadAssignments(); };
        footer.querySelector('.bl-apply-now')?.addEventListener('click', () => {
          closeAllModals();
          openApplyModal(baselineId, tenantId, baselineName, tenantName);
        });
      })
      .catch((err) => {
        modal.querySelector('#blCheckBody').innerHTML = `<div class="bl-result bl-result-error"><div class="bl-result-msg">Fout: ${escHtml(String(err))}</div></div>`;
      });
  }

  // ── Apply modal ────────────────────────────────────────────────────────────

  function openApplyModal(baselineId, tenantId, baselineName, tenantName) {
    const overlay = createOverlay(closeAllModals);
    const modal   = document.createElement('div');
    modal.className = 'bl-modal';
    modal.innerHTML = `
      <div class="bl-modal-header">
        <div>
          <div class="bl-modal-title">Baseline toepassen</div>
          <div class="bl-modal-subtitle">${escHtml(baselineName)} → ${escHtml(tenantName)}</div>
        </div>
        <button class="bl-modal-close">✕</button>
      </div>
      <div class="bl-modal-body">
        <div style="background:#fef3c7;border:1px solid #fde68a;border-left:3px solid #f59e0b;color:#92400e;font-size:.83rem;padding:.65rem .9rem;border-radius:6px;line-height:1.5">
          <strong>Let op:</strong> Baseline toepassen maakt wijzigingen in de M365-configuratie van
          <strong>${escHtml(tenantName)}</strong>. Gebruik Dry-run om te previewer wat er verandert.
        </div>
        <div style="margin-top:.75rem">
          <label style="font-size:.83rem;color:var(--text-secondary)">
            <input type="checkbox" id="blApplyDry" checked style="accent-color:var(--accent)">
            Dry-run (preview, geen wijzigingen)
          </label>
        </div>
        <div id="blApplyResult"></div>
      </div>
      <div class="bl-modal-footer">
        <button class="bl-btn bl-btn-secondary bl-cancel-btn">Annuleren</button>
        <button class="bl-btn bl-btn-primary bl-do-apply">Toepassen</button>
      </div>`;
    modal.querySelector('.bl-modal-close').onclick = closeAllModals;
    modal.querySelector('.bl-cancel-btn').onclick  = closeAllModals;

    modal.querySelector('.bl-do-apply').onclick = () => {
      const dryRun   = modal.querySelector('#blApplyDry').checked;
      const resultDiv = modal.querySelector('#blApplyResult');
      const applyBtn  = modal.querySelector('.bl-do-apply');
      applyBtn.disabled = true;
      resultDiv.innerHTML = `<div style="font-size:.83rem;color:var(--text-muted);margin-top:.75rem">Toepassen... (kan 60s duren)</div>`;

      blApiFetch(`/api/baselines/${baselineId}/apply/${tenantId}`, {
        method: 'POST',
        body: JSON.stringify({ dry_run: dryRun }),
      })
        .then((data) => {
          const r       = data.result || {};
          const applied = r.applied  || [];
          const skipped = r.skipped  || [];
          const warnings = r.warnings || [];
          const isDry   = dryRun;
          const isOk    = data.ok !== false;

          const cls  = isDry ? 'bl-result-dryrun' : (isOk ? 'bl-result-ok' : 'bl-result-error');
          const icon = isDry ? 'ℹ️' : (isOk ? '✅' : '❌');

          resultDiv.innerHTML = `
            <div class="bl-result ${cls}" style="margin-top:.75rem">
              <div class="bl-result-icon">${icon}</div>
              <div class="bl-result-msg">${escHtml(r.message || (isOk ? 'Klaar' : 'Fout'))}</div>
            </div>
            ${applied.length ? `<ul class="bl-result-list" style="margin-top:.5rem">${applied.map((x) => `<li>${escHtml(x)}</li>`).join('')}</ul>` : ''}
            ${skipped.length ? `<ul class="bl-result-list skip" style="margin-top:.25rem">${skipped.map((x) => `<li>${escHtml(x)}</li>`).join('')}</ul>` : ''}
            ${warnings.length ? `<ul class="bl-result-list warn">${warnings.map((w) => `<li>${escHtml(w)}</li>`).join('')}</ul>` : ''}`;

          modal.querySelector('.bl-modal-footer').innerHTML = `<button class="bl-btn bl-btn-primary bl-done-btn">Sluiten</button>`;
          modal.querySelector('.bl-done-btn').onclick = () => { closeAllModals(); loadAssignments(); };
        })
        .catch((err) => {
          resultDiv.innerHTML = `<div class="bl-result bl-result-error" style="margin-top:.75rem"><div class="bl-result-msg">Fout: ${escHtml(String(err))}</div></div>`;
          applyBtn.disabled = false;
        });
    };

    overlay.appendChild(modal);
    document.body.appendChild(overlay);
  }

  // ── Geschiedenis ──────────────────────────────────────────────────────────

  function loadBaselineHistory() {
    const tbody = document.getElementById('blHistoryBody');
    if (!tbody) return;
    tbody.innerHTML = `<tr><td colspan="6" style="text-align:center;padding:2rem;color:var(--text-muted)">Laden...</td></tr>`;

    blApiFetch('/api/baselines/assignments/all')
      .then((data) => {
        const assignments = data.items || [];
        if (!assignments.length) {
          tbody.innerHTML = `<tr><td colspan="6" style="text-align:center;padding:2rem;color:var(--text-muted)">Geen koppelingen gevonden.</td></tr>`;
          return;
        }
        // Haal eerste baseline id op en laad zijn history als demo
        const firstBid = assignments[0].baseline_id;
        return blApiFetch(`/api/baselines/${firstBid}/history`).then((h) => {
          const items = h.items || [];
          if (!items.length) {
            tbody.innerHTML = `<tr><td colspan="6" style="text-align:center;padding:2rem;color:var(--text-muted)">Nog geen activiteit.</td></tr>`;
            return;
          }
          tbody.innerHTML = items.map((i) => {
            const statusMap = { success: 'bl-status-compliant', failed: 'bl-status-non_compliant', dry_run: 'bl-status-dryrun' };
            const label     = { success: 'Succes', failed: 'Mislukt', dry_run: 'Dry-run' };
            return `<tr>
              <td style="font-size:.75rem;color:var(--text-muted)">${fmtDateTime(i.executed_at)}</td>
              <td>${escHtml(i.action === 'check' ? 'Compliance check' : 'Toepassen')}</td>
              <td>${escHtml(i.baseline_id?.substring(0,8) + '...' || '—')}</td>
              <td>${escHtml(i.tenant_id?.substring(0,8) + '...' || '—')}</td>
              <td><span class="bl-status ${statusMap[i.status] || ''}">${escHtml(label[i.status] || i.status)}</span></td>
              <td style="font-size:.75rem">${escHtml(i.executed_by || '—')}</td>
            </tr>`;
          }).join('');
        });
      })
      .catch(() => {
        tbody.innerHTML = `<tr><td colspan="6" style="text-align:center;color:#dc2626">Fout bij laden.</td></tr>`;
      });
  }

  // ── Modal helpers ────────────────────────────────────────────────────────

  function createOverlay(onClose) {
    const el = document.createElement('div');
    el.className = 'bl-modal-overlay';
    el.addEventListener('click', (e) => { if (e.target === el) onClose(); });
    return el;
  }

  function closeAllModals() {
    document.querySelectorAll('.bl-modal-overlay').forEach((el) => el.remove());
  }

  // ── Toolbar binding ───────────────────────────────────────────────────────

  function bindToolbar() {
    const search = document.getElementById('blSearchInput');
    if (search) {
      search.addEventListener('input', () => {
        if (_activeTab === 'baselines') renderBaselineGrid();
      });
    }

    document.querySelectorAll('.bl-tab[data-tab]').forEach((btn) => {
      btn.addEventListener('click', () => switchTab(btn.dataset.tab));
    });

    const btnNew = document.getElementById('blBtnNieuw');
    if (btnNew) btnNew.addEventListener('click', openNewBaselineModal);
  }

  function openNewBaselineModal() {
    const overlay = createOverlay(closeAllModals);
    const modal   = document.createElement('div');
    modal.className = 'bl-modal';
    modal.innerHTML = `
      <div class="bl-modal-header">
        <div><div class="bl-modal-title">Nieuwe baseline</div>
        <div class="bl-modal-subtitle">Handmatig aanmaken</div></div>
        <button class="bl-modal-close">✕</button>
      </div>
      <div class="bl-modal-body">
        <div class="bl-form">
          <div class="bl-form-group">
            <label class="bl-form-label">Naam <span class="bl-required">*</span></label>
            <input class="bl-form-input" id="blNewName" placeholder="Bijv. MKB Beveiligingsstandaard">
          </div>
          <div class="bl-form-group">
            <label class="bl-form-label">Beschrijving</label>
            <textarea class="bl-form-textarea" id="blNewDesc" placeholder="Korte beschrijving van de baseline"></textarea>
          </div>
          <div style="font-size:.83rem;color:var(--text-secondary);background:var(--tag-bg,#f1f5f9);padding:.75rem;border-radius:8px;line-height:1.5">
            💡 Tip: Gebruik Gold Tenant Export om automatisch een baseline te maken van een goed geconfigureerde tenant.
          </div>
        </div>
        <div id="blNewResult"></div>
      </div>
      <div class="bl-modal-footer">
        <button class="bl-btn bl-btn-secondary bl-cancel-btn">Annuleren</button>
        <button class="bl-btn bl-btn-primary bl-create-btn">Aanmaken</button>
      </div>`;
    modal.querySelector('.bl-modal-close').onclick = closeAllModals;
    modal.querySelector('.bl-cancel-btn').onclick  = closeAllModals;
    modal.querySelector('.bl-create-btn').onclick  = () => {
      const name = (modal.querySelector('#blNewName').value || '').trim();
      const desc = (modal.querySelector('#blNewDesc').value || '').trim();
      if (!name) { modal.querySelector('#blNewName').classList.add('error'); return; }
      const resultDiv = modal.querySelector('#blNewResult');
      resultDiv.innerHTML = `<div style="font-size:.82rem;color:var(--text-muted)">Aanmaken...</div>`;
      blApiFetch('/api/baselines', {
        method: 'POST',
        body: JSON.stringify({ name, description: desc, config: { categories: {} } }),
      })
        .then(() => { closeAllModals(); loadBaselines(); })
        .catch((err) => {
          resultDiv.innerHTML = `<div class="bl-result bl-result-error"><div class="bl-result-msg">Fout: ${escHtml(String(err))}</div></div>`;
        });
    };

    overlay.appendChild(modal);
    document.body.appendChild(overlay);
  }

  // ── Publieke interface ───────────────────────────────────────────────────

  window.loadBaselineSection = function () {
    bindToolbar();
    switchTab('baselines');
  };
  window.switchBaselineTab = switchTab;

})();
