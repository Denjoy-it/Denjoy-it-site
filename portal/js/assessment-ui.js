(function () {
  const state = {
    tenantId: null,
    nav: null,
    selectedKey: 'summary',
    loading: false,
  };

  function assessmentEscapeHtml(value) {
    return String(value == null ? '' : value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function assessmentFormatDate(value) {
    if (!value) return 'Nog geen run';
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return value;
    return date.toLocaleString('nl-NL');
  }

  async function assessmentFetchJson(path) {
    const response = await fetch(path, {
      headers: { 'Content-Type': 'application/json' },
    });
    let data = null;
    try {
      data = await response.json();
    } catch (_) {
      data = null;
    }
    if (!response.ok) {
      throw new Error((data && data.error) || `HTTP ${response.status}`);
    }
    return data;
  }

  function selectedTenantId() {
    return document.getElementById('tenantSelect')?.value || null;
  }

  function currentSectionRoot() {
    return document.getElementById('assessmentExperienceRoot');
  }

  function currentLegacyRunner() {
    return document.getElementById('assessmentLegacyRunner');
  }

  function renderTopAssessmentMenu(nav) {
    const dropdown = document.getElementById('assessmentNavDropdown');
    if (!dropdown) return;
    const items = Array.isArray(nav?.items) && nav.items.length ? nav.items : [{ key: 'summary', label: 'Overzicht', count: null }];
    dropdown.innerHTML = items.map((item) => `
      <button type="button" class="nav-dropdown-link" data-assessment-navjump="${assessmentEscapeHtml(item.key)}">
        <span>${assessmentEscapeHtml(item.label)}</span>
        ${item.count != null ? `<span class="nav-dropdown-count">${assessmentEscapeHtml(item.count)}</span>` : ''}
      </button>
    `).join('');
  }

  function assessmentToneClass(tone) {
    if (tone === 'success') return 'is-success';
    if (tone === 'warn') return 'is-warn';
    return 'is-default';
  }

  function renderAssessmentTable(section) {
    const columns = Array.isArray(section.columns) ? section.columns : [];
    const rows = Array.isArray(section.rows) ? section.rows : [];
    if (!rows.length) {
      return `
        <div class="assessment-panel assessment-panel-empty">
          <p>Voor dit onderdeel is nog geen assessmentdata gevonden.</p>
        </div>
      `;
    }
    return `
      <div class="assessment-panel assessment-panel-table">
        <div class="assessment-table-wrap">
          <table class="assessment-table">
            <thead>
              <tr>${columns.map((column) => `<th>${assessmentEscapeHtml(column)}</th>`).join('')}</tr>
            </thead>
            <tbody>
              ${rows.map((row) => {
                const values = Object.values(row || {});
                return `
                  <tr>
                    ${values.map((value) => `<td>${assessmentEscapeHtml(value || '—')}</td>`).join('')}
                  </tr>
                `;
              }).join('')}
            </tbody>
          </table>
        </div>
      </div>
    `;
  }

  function renderAssessmentCards(cards) {
    if (!Array.isArray(cards) || !cards.length) return '';
    return `
      <div class="assessment-kpi-grid">
        ${cards.map((card) => `
          <article class="assessment-kpi-card ${assessmentToneClass(card.tone)}">
            <span class="assessment-kpi-label">${assessmentEscapeHtml(card.label)}</span>
            <strong class="assessment-kpi-value">${assessmentEscapeHtml(card.value)}</strong>
          </article>
        `).join('')}
      </div>
    `;
  }

  function renderAssessmentBars(bars) {
    if (!Array.isArray(bars) || !bars.length) return '';
    return `
      <div class="assessment-panel">
        <div class="assessment-panel-header">
          <div>
            <p class="assessment-panel-eyebrow">Tenant gezondheid</p>
            <h3>Gevonden onderdelen in de laatste run</h3>
          </div>
        </div>
        <div class="assessment-bars">
          ${bars.map((bar) => {
            const max = Math.max(Number(bar.max || 0), 1);
            const value = Number(bar.value || 0);
            const width = Math.max(8, Math.min(100, Math.round((value / max) * 100)));
            return `
              <div class="assessment-bar-row">
                <div class="assessment-bar-topline">
                  <span>${assessmentEscapeHtml(bar.label)}</span>
                  <strong>${assessmentEscapeHtml(value)}</strong>
                </div>
                <div class="assessment-bar-track">
                  <span class="assessment-bar-fill" style="width:${width}%"></span>
                </div>
              </div>
            `;
          }).join('')}
        </div>
      </div>
    `;
  }

  function renderAssessmentSection(section, nav) {
    const cards = renderAssessmentCards(section.cards);
    const bars = renderAssessmentBars(section.bars);
    const table = renderAssessmentTable(section);
    return `
      <div class="assessment-content-head">
        <div>
          <p class="assessment-panel-eyebrow">Assessment onderdeel</p>
          <h2>${assessmentEscapeHtml(section.title || 'Assessment')}</h2>
          <p class="assessment-content-sub">
            Laatste synchronisatie: ${assessmentEscapeHtml(assessmentFormatDate(section.generated_at || nav.generated_at))}
          </p>
        </div>
        <div class="assessment-content-actions">
          ${nav.latest_report_path ? `<a class="assessment-action-btn assessment-action-btn-secondary" href="${assessmentEscapeHtml(nav.latest_report_path)}" target="_blank" rel="noopener noreferrer">Open rapport</a>` : ''}
          <button type="button" class="assessment-action-btn" data-assessment-action="scroll-runner">Nieuwe run starten</button>
        </div>
      </div>
      ${cards}
      ${bars}
      ${section.rows ? table : ''}
    `;
  }

  function renderAssessmentShell(nav, section) {
    const root = currentSectionRoot();
    const items = Array.isArray(nav.items) ? nav.items : [];
    const score = nav.score != null && nav.score !== '' ? nav.score : '—';
    root.innerHTML = `
      <div class="assessment-experience">
        <section class="assessment-hero">
          <div class="assessment-hero-copy">
            <span class="assessment-badge">Microsoft 365 Assessment</span>
            <h1>Assessment die echt meedenkt.</h1>
            <p>
              Een Denjoy-overzicht van licenties, identiteiten, app registraties en tenantgezondheid.
              Alleen onderdelen met echte data worden hier zichtbaar.
            </p>
            <div class="assessment-hero-actions">
              <button type="button" class="assessment-action-btn" data-assessment-action="scroll-runner">Assessment uitvoeren</button>
              ${nav.latest_report_path ? `<a class="assessment-action-btn assessment-action-btn-secondary" href="${assessmentEscapeHtml(nav.latest_report_path)}" target="_blank" rel="noopener noreferrer">Volledig rapport</a>` : '<button type="button" class="assessment-action-btn assessment-action-btn-secondary" data-assessment-action="show-results">Bekijk resultaten</button>'}
            </div>
          </div>
          <div class="assessment-hero-aside">
            <div class="assessment-signal-card">
              <span>Tenant</span>
              <strong>${assessmentEscapeHtml(nav.tenant_name || 'Onbekend')}</strong>
            </div>
            <div class="assessment-signal-card">
              <span>Laatste run</span>
              <strong>${assessmentEscapeHtml(assessmentFormatDate(nav.generated_at))}</strong>
            </div>
            <div class="assessment-signal-card">
              <span>Assessment score</span>
              <strong>${assessmentEscapeHtml(score)}</strong>
            </div>
          </div>
        </section>

        <section class="assessment-shell">
          <aside class="assessment-shell-nav">
            <div class="assessment-shell-topline">
              <span class="assessment-shell-led"></span>
              <span>portal.denjoy.nl/assessment</span>
            </div>
            <div class="assessment-shell-menu">
              ${items.map((item) => `
                <button
                  type="button"
                  class="assessment-nav-item ${item.key === state.selectedKey ? 'is-active' : ''}"
                  data-assessment-key="${assessmentEscapeHtml(item.key)}"
                >
                  <span>${assessmentEscapeHtml(item.label)}</span>
                  ${item.count != null ? `<strong>${assessmentEscapeHtml(item.count)}</strong>` : ''}
                </button>
              `).join('')}
            </div>
          </aside>
          <div class="assessment-shell-body">
            ${renderAssessmentSection(section, nav)}
          </div>
        </section>
      </div>
    `;
  }

  function renderAssessmentNotice(title, message, tone) {
    const root = currentSectionRoot();
    if (!root) return;
    root.innerHTML = `
      <div class="assessment-experience">
        <section class="assessment-hero assessment-hero-notice">
          <div class="assessment-panel ${tone === 'error' ? 'assessment-panel-error' : ''}">
            <p class="assessment-panel-eyebrow">Assessment</p>
            <h2>${assessmentEscapeHtml(title)}</h2>
            <p>${assessmentEscapeHtml(message)}</p>
          </div>
        </section>
      </div>
    `;
  }

  function syncLegacyRunnerVisibility(enabled) {
    const runner = currentLegacyRunner();
    if (!runner) return;
    runner.classList.toggle('assessment-legacy-runner-enhanced', enabled);
  }

  async function loadAssessmentExperience(options = {}) {
    const force = Boolean(options.force);
    const root = currentSectionRoot();
    if (!root || state.loading) return;
    const tenantId = selectedTenantId();
    syncLegacyRunnerVisibility(true);

    if (!tenantId) {
      state.tenantId = null;
      renderAssessmentNotice('Geen tenant geselecteerd', 'Selecteer eerst een tenant om assessmentresultaten en bevindingen te laden.');
      return;
    }

    try {
      state.loading = true;
      if (force || state.tenantId !== tenantId) {
        state.tenantId = tenantId;
        state.nav = null;
        state.selectedKey = 'summary';
      }

      const nav = await assessmentFetchJson(`/api/assessment/${tenantId}/nav`);
      state.nav = nav;
      renderTopAssessmentMenu(nav);

      if (!nav.enabled) {
        renderAssessmentNotice('Nieuwe assessment-weergave staat uit', 'De huidige portal gebruikt nog de klassieke assessmentpagina. Zet assessment_ui_v1 aan om deze ervaring te tonen.');
        syncLegacyRunnerVisibility(false);
        return;
      }

      const availableKeys = (nav.items || []).map((item) => item.key);
      if (!availableKeys.includes(state.selectedKey)) {
        state.selectedKey = availableKeys[0] || 'summary';
      }

      const section = await assessmentFetchJson(`/api/assessment/${tenantId}/section/${state.selectedKey}`);
      renderAssessmentShell(nav, section);
    } catch (error) {
      console.error(error);
      renderAssessmentNotice('Assessment laden mislukt', error.message || 'Onbekende fout bij laden van assessmentdata.', 'error');
    } finally {
      state.loading = false;
    }
  }

  async function selectAssessmentSection(key) {
    if (!key || key === state.selectedKey) return;
    state.selectedKey = key;
    await loadAssessmentExperience({ force: false });
  }

  function scrollToLegacyRunner() {
    const runner = currentLegacyRunner();
    if (!runner) return;
    runner.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  function bindAssessmentUiEvents() {
    document.addEventListener('click', (event) => {
      const navButton = event.target.closest('[data-assessment-key]');
      if (navButton) {
        event.preventDefault();
        selectAssessmentSection(navButton.dataset.assessmentKey);
        return;
      }

      const actionButton = event.target.closest('[data-assessment-action]');
      if (!actionButton) return;
      event.preventDefault();
      const action = actionButton.dataset.assessmentAction;
      if (action === 'scroll-runner') {
        scrollToLegacyRunner();
      } else if (action === 'show-results' && typeof showSection === 'function') {
        showSection('results');
      }
    });

    document.addEventListener('click', (event) => {
      const jumpButton = event.target.closest('[data-assessment-navjump]');
      if (!jumpButton) return;
      event.preventDefault();
      state.selectedKey = jumpButton.dataset.assessmentNavjump || 'summary';
      const dropdownGroup = document.querySelector('.nav-dropdown');
      if (dropdownGroup) dropdownGroup.classList.remove('open');
      if (typeof showSection === 'function') showSection('assessment');
    });

    document.addEventListener('click', (event) => {
      const toggle = event.target.closest('#assessmentNavToggle');
      const dropdownGroup = document.querySelector('.nav-dropdown');
      if (!dropdownGroup) return;
      if (toggle) {
        event.preventDefault();
        dropdownGroup.classList.toggle('open');
        if (typeof showSection === 'function') showSection('assessment');
        return;
      }
      if (!event.target.closest('.nav-dropdown')) {
        dropdownGroup.classList.remove('open');
      }
    });
  }

  document.addEventListener('DOMContentLoaded', bindAssessmentUiEvents);
  window.loadAssessmentExperience = loadAssessmentExperience;
  window.selectAssessmentSection = selectAssessmentSection;
})();
