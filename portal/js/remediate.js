/**
 * Denjoy IT Platform — Herstel (Remediation) module
 * Fase 1: Active remediation via Microsoft Graph API (via backend/PowerShell)
 */

(function () {
  'use strict';

  // ── State ──────────────────────────────────────────────────────────────────
  let _currentTenantId = null;
  let _catalog = [];
  let _history = [];
  let _activeModal = null;

  // ── Helpers ────────────────────────────────────────────────────────────────
  function remEscape(v) {
    return String(v == null ? '' : v)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function remFmt(iso) {
    if (!iso) return '—';
    try { return new Date(iso).toLocaleString('nl-NL'); } catch (_) { return iso; }
  }

  function getTenantId() {
    return document.getElementById('tenantSelect')?.value || _currentTenantId || null;
  }

  async function remApiFetch(path, options = {}) {
    const token = localStorage.getItem('denjoy_token');
    const authHeader = token ? { Authorization: `Bearer ${token}` } : {};
    const res = await fetch(path, {
      credentials: 'include',
      headers: { 'Content-Type': 'application/json', ...authHeader, ...(options.headers || {}) },
      ...options,
    });
    let data = null;
    try { data = await res.json(); } catch (_) {}
    if (!res.ok) throw new Error((data && data.error) || `HTTP ${res.status}`);
    return data;
  }

  // ── Severity badge ─────────────────────────────────────────────────────────
  function severityBadge(sev) {
    const map = {
      critical: { cls: 'rem-badge-critical', label: 'Kritiek' },
      warning:  { cls: 'rem-badge-warning',  label: 'Waarschuwing' },
      info:     { cls: 'rem-badge-info',      label: 'Info' },
    };
    const d = map[sev] || map.info;
    return `<span class="rem-badge ${d.cls}">${d.label}</span>`;
  }

  function riskBadge(risk) {
    const map = {
      low:    { cls: 'rem-risk-low',    label: 'Laag risico' },
      medium: { cls: 'rem-risk-medium', label: 'Middel risico' },
      high:   { cls: 'rem-risk-high',   label: 'Hoog risico' },
    };
    const d = map[risk] || map.low;
    return `<span class="rem-risk ${d.cls}">${d.label}</span>`;
  }

  function statusBadge(status) {
    const map = {
      success: { cls: 'rem-status-success', label: 'Geslaagd' },
      failed:  { cls: 'rem-status-failed',  label: 'Mislukt' },
      dry_run: { cls: 'rem-status-dryrun',  label: 'Dry-run' },
    };
    const d = map[status] || { cls: '', label: status };
    return `<span class="rem-status ${d.cls}">${d.label}</span>`;
  }

  // ── Render catalogus ────────────────────────────────────────────────────────
  function renderCatalog(items) {
    const root = document.getElementById('remCatalogGrid');
    if (!root) return;

    if (!items.length) {
      root.innerHTML = '<p class="rem-empty">Geen remediations beschikbaar.</p>';
      return;
    }

    // Groepeer op categorie
    const groups = {};
    for (const item of items) {
      const cat = item.category_label || item.category || 'Overig';
      if (!groups[cat]) groups[cat] = [];
      groups[cat].push(item);
    }

    root.innerHTML = Object.entries(groups).map(([groupLabel, groupItems]) => `
      <div class="rem-group">
        <h3 class="rem-group-title">${remEscape(groupLabel)}</h3>
        <div class="rem-cards">
          ${groupItems.map((item) => `
            <div class="rem-card" data-rem-id="${remEscape(item.id)}">
              <div class="rem-card-header">
                <div class="rem-card-title">${remEscape(item.title)}</div>
                <div class="rem-card-badges">
                  ${severityBadge(item.severity)}
                  ${riskBadge(item.risk)}
                </div>
              </div>
              <p class="rem-card-desc">${remEscape(item.description)}</p>
              <div class="rem-card-footer">
                <code class="rem-card-endpoint">${remEscape(item.graph_endpoint || '')}</code>
                <div class="rem-card-actions">
                  ${item.dry_run_supported ? `
                    <button type="button" class="rem-btn rem-btn-preview"
                      data-action="preview" data-rem-id="${remEscape(item.id)}">
                      Voorbeeld
                    </button>
                  ` : ''}
                  <button type="button" class="rem-btn rem-btn-execute"
                    data-action="execute" data-rem-id="${remEscape(item.id)}">
                    Uitvoeren
                  </button>
                </div>
              </div>
            </div>
          `).join('')}
        </div>
      </div>
    `).join('');

    // Event listeners
    root.querySelectorAll('[data-action]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const id = btn.dataset.remId;
        const action = btn.dataset.action;
        const item = _catalog.find((c) => c.id === id);
        if (!item) return;
        if (action === 'preview') openModal(item, true);
        else openModal(item, false);
      });
    });
  }

  // ── Modal ──────────────────────────────────────────────────────────────────
  function openModal(item, isDryRun) {
    closeModal();

    const hasParams = Array.isArray(item.params_schema) && item.params_schema.length > 0;
    const paramFields = hasParams
      ? item.params_schema.map((p) => `
          <div class="rem-form-group">
            <label class="rem-form-label" for="remParam_${remEscape(p.name)}">
              ${remEscape(p.label)}${p.required ? ' <span class="rem-required">*</span>' : ''}
            </label>
            <input
              id="remParam_${remEscape(p.name)}"
              type="text"
              class="rem-form-input"
              placeholder="${remEscape(p.placeholder || '')}"
              data-param="${remEscape(p.name)}"
              ${p.required ? 'required' : ''}
            >
          </div>
        `).join('')
      : '';

    const riskWarning = item.risk === 'high'
      ? `<div class="rem-modal-risk-warn">
           Hoog risico — dit kan invloed hebben op meerdere gebruikers of services. Controleer eerst de preview.
         </div>`
      : item.risk === 'medium'
      ? `<div class="rem-modal-risk-warn rem-modal-risk-medium">
           Middel risico — controleer of dit van toepassing is op de geselecteerde tenant.
         </div>`
      : '';

    const overlay = document.createElement('div');
    overlay.className = 'rem-modal-overlay';
    overlay.innerHTML = `
      <div class="rem-modal" role="dialog" aria-modal="true" aria-label="${remEscape(item.title)}">
        <div class="rem-modal-header">
          <div>
            <div class="rem-modal-title">${remEscape(isDryRun ? '[Preview] ' : '')}${remEscape(item.title)}</div>
            <div class="rem-modal-badges">
              ${severityBadge(item.severity)} ${riskBadge(item.risk)}
              ${isDryRun ? '<span class="rem-badge rem-badge-dryrun">Dry-run</span>' : ''}
            </div>
          </div>
          <button type="button" class="rem-modal-close" aria-label="Sluiten">✕</button>
        </div>
        <div class="rem-modal-body">
          <p class="rem-modal-desc">${remEscape(item.description)}</p>
          ${riskWarning}
          ${hasParams ? `<div class="rem-form">${paramFields}</div>` : ''}
          <div id="remModalResult" class="rem-modal-result" style="display:none;"></div>
        </div>
        <div class="rem-modal-footer">
          <button type="button" class="rem-btn rem-btn-cancel" id="remModalCancel">Annuleren</button>
          <button type="button" class="rem-btn ${isDryRun ? 'rem-btn-preview' : 'rem-btn-execute'}" id="remModalConfirm">
            ${isDryRun ? 'Preview uitvoeren' : 'Bevestigen & uitvoeren'}
          </button>
        </div>
      </div>
    `;

    document.body.appendChild(overlay);
    _activeModal = overlay;

    overlay.querySelector('.rem-modal-close').addEventListener('click', closeModal);
    overlay.querySelector('#remModalCancel').addEventListener('click', closeModal);
    overlay.addEventListener('click', (e) => { if (e.target === overlay) closeModal(); });

    overlay.querySelector('#remModalConfirm').addEventListener('click', async () => {
      const params = {};
      if (hasParams) {
        let valid = true;
        overlay.querySelectorAll('[data-param]').forEach((input) => {
          const name = input.dataset.param;
          const val = input.value.trim();
          if (input.required && !val) {
            input.classList.add('rem-input-error');
            valid = false;
          } else {
            input.classList.remove('rem-input-error');
            params[name] = val;
          }
        });
        if (!valid) return;
      }
      await executeRemediation(item, isDryRun, params, overlay);
    });

    // Focus eerste input
    const firstInput = overlay.querySelector('input');
    if (firstInput) firstInput.focus();
    else overlay.querySelector('#remModalConfirm').focus();
  }

  function closeModal() {
    if (_activeModal) {
      _activeModal.remove();
      _activeModal = null;
    }
  }

  // ── Execute ─────────────────────────────────────────────────────────────────
  async function executeRemediation(item, isDryRun, params, overlay) {
    const confirmBtn = overlay.querySelector('#remModalConfirm');
    const cancelBtn  = overlay.querySelector('#remModalCancel');
    const resultDiv  = overlay.querySelector('#remModalResult');

    confirmBtn.disabled = true;
    cancelBtn.disabled  = true;
    confirmBtn.textContent = 'Bezig...';
    resultDiv.style.display = 'none';

    const tenantId = getTenantId();
    if (!tenantId) {
      showModalResult(resultDiv, false, 'Geen tenant geselecteerd.');
      confirmBtn.disabled = false;
      cancelBtn.disabled  = false;
      confirmBtn.textContent = isDryRun ? 'Preview uitvoeren' : 'Bevestigen & uitvoeren';
      return;
    }

    try {
      const payload = {
        remediation_id: item.id,
        params,
        dry_run: isDryRun,
      };
      const data = await remApiFetch(`/api/remediate/${tenantId}/execute`, {
        method: 'POST',
        body: JSON.stringify(payload),
      });

      showModalResult(resultDiv, data.ok, data.message, data.result);

      if (data.ok && !isDryRun) {
        // Sluit modal na 2.5s bij succesvolle uitvoering
        setTimeout(() => {
          closeModal();
          loadRemediationHistory();
        }, 2500);
      } else {
        confirmBtn.disabled = false;
        cancelBtn.disabled  = false;
        confirmBtn.textContent = isDryRun ? 'Preview uitvoeren' : 'Bevestigen & uitvoeren';
        // Na succesvolle dry-run: laad history bij
        if (data.ok) loadRemediationHistory();
      }
    } catch (err) {
      showModalResult(resultDiv, false, `Fout: ${err.message}`);
      confirmBtn.disabled = false;
      cancelBtn.disabled  = false;
      confirmBtn.textContent = isDryRun ? 'Preview uitvoeren' : 'Bevestigen & uitvoeren';
    }
  }

  function showModalResult(container, ok, message, details) {
    container.style.display = '';
    container.className = `rem-modal-result ${ok ? 'rem-result-ok' : 'rem-result-error'}`;

    let detailHtml = '';
    if (details && typeof details === 'object') {
      const entries = Object.entries(details)
        .filter(([k]) => !['success', 'dry_run'].includes(k))
        .map(([k, v]) => `<div class="rem-detail-row"><span class="rem-detail-key">${remEscape(k)}:</span> <span class="rem-detail-val">${remEscape(JSON.stringify(v))}</span></div>`)
        .join('');
      if (entries) detailHtml = `<div class="rem-detail-block">${entries}</div>`;
    }

    container.innerHTML = `
      <div class="rem-result-icon">${ok ? '✓' : '✗'}</div>
      <div class="rem-result-msg">${remEscape(message)}</div>
      ${detailHtml}
    `;
  }

  // ── Geschiedenis ────────────────────────────────────────────────────────────
  function renderHistory(items) {
    const container = document.getElementById('remHistoryBody');
    if (!container) return;

    if (!items.length) {
      container.innerHTML = '<tr><td colspan="6" class="rem-table-empty">Nog geen acties uitgevoerd voor deze tenant.</td></tr>';
      return;
    }

    container.innerHTML = items.map((h) => `
      <tr>
        <td>${remEscape(remFmt(h.executed_at))}</td>
        <td>${remEscape(h.title || h.remediation_id)}</td>
        <td>${statusBadge(h.status)}</td>
        <td class="rem-td-msg">${remEscape(
          (() => { try { const r = JSON.parse(h.result_json || '{}'); return r.message || ''; } catch(_) { return ''; } })()
          || h.error_message || '—'
        )}</td>
        <td>${remEscape(h.executed_by || '—')}</td>
        <td>${h.dry_run ? '<span class="rem-badge rem-badge-dryrun">Dry-run</span>' : '—'}</td>
      </tr>
    `).join('');
  }

  async function loadRemediationHistory() {
    const tenantId = getTenantId();
    if (!tenantId) return;

    const container = document.getElementById('remHistoryBody');
    if (container) container.innerHTML = '<tr><td colspan="6" class="rem-table-empty">Laden...</td></tr>';

    try {
      const data = await remApiFetch(`/api/remediate/${tenantId}/history`);
      _history = data.items || [];
      renderHistory(_history);
    } catch (err) {
      if (container) container.innerHTML = `<tr><td colspan="6" class="rem-table-empty rem-result-error">Fout: ${remEscape(err.message)}</td></tr>`;
    }
  }

  // ── Tab switching ────────────────────────────────────────────────────────────
  function switchRemTab(tab) {
    document.querySelectorAll('.rem-tab').forEach((btn) => {
      btn.classList.toggle('active', btn.dataset.remTab === tab);
    });
    document.querySelectorAll('.rem-tab-panel').forEach((panel) => {
      panel.style.display = panel.dataset.remPanel === tab ? '' : 'none';
    });
    if (tab === 'geschiedenis') loadRemediationHistory();
  }

  function bindRemTabs() {
    document.querySelectorAll('.rem-tab[data-rem-tab]').forEach((btn) => {
      btn.addEventListener('click', () => switchRemTab(btn.dataset.remTab));
    });
  }

  // ── Filter ─────────────────────────────────────────────────────────────────
  function bindFilterInput() {
    const input = document.getElementById('remFilterInput');
    if (!input) return;
    input.addEventListener('input', () => {
      const q = input.value.toLowerCase();
      const filtered = q
        ? _catalog.filter((c) =>
            c.title.toLowerCase().includes(q) ||
            c.description.toLowerCase().includes(q) ||
            (c.category_label || c.category || '').toLowerCase().includes(q)
          )
        : _catalog;
      renderCatalog(filtered);
    });
  }

  // ── Hoofdlaadfunctie ────────────────────────────────────────────────────────
  async function loadHerstellSection() {
    const tenantId = getTenantId();
    _currentTenantId = tenantId;

    bindRemTabs();

    const catalogGrid = document.getElementById('remCatalogGrid');
    if (catalogGrid) catalogGrid.innerHTML = '<p class="rem-empty">Catalogus laden...</p>';

    if (!tenantId) {
      if (catalogGrid) catalogGrid.innerHTML = '<p class="rem-empty rem-result-error">Selecteer eerst een tenant.</p>';
      return;
    }

    try {
      const catalogData = await remApiFetch(`/api/remediate/${tenantId}/catalog`);
      _catalog = catalogData.items || [];
      renderCatalog(_catalog);
      bindFilterInput();
    } catch (err) {
      if (catalogGrid) catalogGrid.innerHTML = `<p class="rem-empty rem-result-error">Fout bij laden: ${remEscape(err.message)}</p>`;
    }
  }

  // ── Publieke API ────────────────────────────────────────────────────────────
  window.loadHerstellSection    = loadHerstellSection;
  window.loadRemediationHistory = loadRemediationHistory;
  window.switchRemediationTab   = switchRemTab;

})();
