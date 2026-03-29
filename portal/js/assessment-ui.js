(function () {
  const state = {
    tenantId: null,
    nav: null,
    selectedKey: 'summary',
    loading: false,
    liveLoadingKey: null,
  };

  const LIVE_CHAPTERS = [
    {
      key: 'identity',
      title: 'Identity & Access',
      description: 'Live tenantdata voor MFA, gasten, beheerrollen en security defaults.',
      primaryAction: { label: 'Hoofdstuk ophalen', path: (tenantId) => `/api/identity/${tenantId}/mfa`, title: 'Identity / MFA' },
      actions: [
        { key: 'identity-mfa', label: 'MFA', path: (tenantId) => `/api/identity/${tenantId}/mfa`, title: 'Identity / MFA' },
        { key: 'identity-guests', label: 'Guests', path: (tenantId) => `/api/identity/${tenantId}/guests`, title: 'Identity / Guests' },
        { key: 'identity-admin-roles', label: 'Admin Roles', path: (tenantId) => `/api/identity/${tenantId}/admin-roles`, title: 'Identity / Admin Roles' },
        { key: 'identity-security-defaults', label: 'Security Defaults', path: (tenantId) => `/api/identity/${tenantId}/security-defaults`, title: 'Identity / Security Defaults' },
        { key: 'identity-legacy-auth', label: 'Legacy Auth', path: (tenantId) => `/api/identity/${tenantId}/legacy-auth`, title: 'Identity / Legacy Auth' },
      ],
    },
    {
      key: 'collaboration',
      title: 'Collaboration',
      description: 'SharePoint en Teams live uitlezen voor de geselecteerde tenant.',
      primaryAction: { label: 'Hoofdstuk ophalen', path: (tenantId) => `/api/collaboration/${tenantId}/teams`, title: 'Collaboration / Teams' },
      actions: [
        { key: 'collab-sharepoint-sites', label: 'SharePoint Sites', path: (tenantId) => `/api/collaboration/${tenantId}/sharepoint/sites`, title: 'Collaboration / SharePoint Sites' },
        { key: 'collab-sharepoint-settings', label: 'SharePoint Settings', path: (tenantId) => `/api/collaboration/${tenantId}/sharepoint/settings`, title: 'Collaboration / SharePoint Settings' },
        { key: 'collab-teams', label: 'Teams', path: (tenantId) => `/api/collaboration/${tenantId}/teams`, title: 'Collaboration / Teams' },
      ],
    },
    {
      key: 'apps',
      title: 'App Registrations',
      description: 'App-registraties, secret-verval en certificaatstatus per tenant.',
      primaryAction: { label: 'Hoofdstuk ophalen', path: (tenantId) => `/api/apps/${tenantId}/registrations`, title: 'Apps / Registrations' },
      actions: [
        { key: 'apps-registrations', label: 'Registrations', path: (tenantId) => `/api/apps/${tenantId}/registrations`, title: 'Apps / Registrations' },
      ],
    },
    {
      key: 'intune',
      title: 'Intune',
      description: 'Devices, compliance en configuratieprofielen live vanuit Graph.',
      primaryAction: { label: 'Hoofdstuk ophalen', path: (tenantId) => `/api/intune/${tenantId}/summary`, title: 'Intune / Summary' },
      actions: [
        { key: 'intune-summary', label: 'Summary', path: (tenantId) => `/api/intune/${tenantId}/summary`, title: 'Intune / Summary' },
        { key: 'intune-devices', label: 'Devices', path: (tenantId) => `/api/intune/${tenantId}/devices`, title: 'Intune / Devices' },
        { key: 'intune-compliance', label: 'Compliance', path: (tenantId) => `/api/intune/${tenantId}/compliance`, title: 'Intune / Compliance' },
        { key: 'intune-config', label: 'Config', path: (tenantId) => `/api/intune/${tenantId}/config`, title: 'Intune / Config' },
      ],
    },
    {
      key: 'backup',
      title: 'Backup',
      description: 'Microsoft 365 Backup status, policies en beschermde resources.',
      primaryAction: { label: 'Hoofdstuk ophalen', path: (tenantId) => `/api/backup/${tenantId}/summary`, title: 'Backup / Summary' },
      actions: [
        { key: 'backup-summary', label: 'Summary', path: (tenantId) => `/api/backup/${tenantId}/summary`, title: 'Backup / Summary' },
        { key: 'backup-status', label: 'Status', path: (tenantId) => `/api/backup/${tenantId}/status`, title: 'Backup / Status' },
        { key: 'backup-sharepoint', label: 'SharePoint', path: (tenantId) => `/api/backup/${tenantId}/sharepoint`, title: 'Backup / SharePoint' },
        { key: 'backup-onedrive', label: 'OneDrive', path: (tenantId) => `/api/backup/${tenantId}/onedrive`, title: 'Backup / OneDrive' },
        { key: 'backup-exchange', label: 'Exchange', path: (tenantId) => `/api/backup/${tenantId}/exchange`, title: 'Backup / Exchange' },
      ],
    },
    {
      key: 'ca',
      title: 'Conditional Access',
      description: 'Policies en named locations live ophalen voor de tenant.',
      primaryAction: { label: 'Hoofdstuk ophalen', path: (tenantId) => `/api/ca/${tenantId}/policies`, title: 'Conditional Access / Policies' },
      actions: [
        { key: 'ca-policies', label: 'Policies', path: (tenantId) => `/api/ca/${tenantId}/policies`, title: 'Conditional Access / Policies' },
        { key: 'ca-locations', label: 'Named Locations', path: (tenantId) => `/api/ca/${tenantId}/named-locations`, title: 'Conditional Access / Named Locations' },
      ],
    },
    {
      key: 'domains',
      title: 'Domains & DNS',
      description: 'Domeinen live ophalen en desgewenst direct een DNS-analyse starten.',
      primaryAction: { label: 'Hoofdstuk ophalen', path: (tenantId) => `/api/domains/${tenantId}/list`, title: 'Domains / List' },
      actions: [
        { key: 'domains-list', label: 'Domeinen', path: (tenantId) => `/api/domains/${tenantId}/list`, title: 'Domains / List' },
        { key: 'domains-analyse', label: 'Analyse domein', path: (tenantId, formValues) => `/api/domains/${tenantId}/analyse?domain=${encodeURIComponent(formValues.domain || '')}`, title: 'Domains / Analyse', requiresInput: 'domain' },
      ],
    },
    {
      key: 'alerts',
      title: 'Alerts & Audit',
      description: 'Secure Score, audit logs en risicovolle sign-ins live controleren.',
      primaryAction: { label: 'Hoofdstuk ophalen', path: (tenantId) => `/api/alerts/${tenantId}/secure-score`, title: 'Alerts / Secure Score' },
      actions: [
        { key: 'alerts-secure-score', label: 'Secure Score', path: (tenantId) => `/api/alerts/${tenantId}/secure-score`, title: 'Alerts / Secure Score' },
        { key: 'alerts-audit-logs', label: 'Audit Logs', path: (tenantId) => `/api/alerts/${tenantId}/audit-logs`, title: 'Alerts / Audit Logs' },
        { key: 'alerts-sign-ins', label: 'Sign-ins', path: (tenantId) => `/api/alerts/${tenantId}/sign-ins`, title: 'Alerts / Sign-ins' },
      ],
    },
    {
      key: 'exchange',
      title: 'Exchange',
      description: 'Mailboxen, forwarding en inbox rules live voor deze tenant.',
      primaryAction: { label: 'Hoofdstuk ophalen', path: (tenantId) => `/api/exchange/${tenantId}/mailboxes`, title: 'Exchange / Mailboxes' },
      actions: [
        { key: 'exchange-mailboxes', label: 'Mailboxes', path: (tenantId) => `/api/exchange/${tenantId}/mailboxes`, title: 'Exchange / Mailboxes' },
        { key: 'exchange-forwarding', label: 'Forwarding', path: (tenantId) => `/api/exchange/${tenantId}/forwarding`, title: 'Exchange / Forwarding' },
        { key: 'exchange-rules', label: 'Inbox Rules', path: (tenantId) => `/api/exchange/${tenantId}/mailbox-rules`, title: 'Exchange / Inbox Rules' },
      ],
    },
  ];

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

  function assessmentFormatCell(value) {
    if (value == null || value === '') return '—';
    if (Array.isArray(value)) {
      if (!value.length) return '—';
      return value.map((item) => (typeof item === 'object' ? JSON.stringify(item) : String(item))).slice(0, 3).join(', ');
    }
    if (typeof value === 'object') {
      const keys = Object.keys(value);
      if (!keys.length) return '—';
      if ('count' in value) return String(value.count);
      return JSON.stringify(value);
    }
    return String(value);
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
          <button type="button" class="assessment-action-btn assessment-action-btn-secondary" data-assessment-action="show-results">Bekijk resultaten</button>
          <button type="button" class="assessment-action-btn" data-assessment-action="scroll-runner">Nieuwe run starten</button>
        </div>
      </div>
      ${cards}
      ${bars}
      ${section.rows ? table : ''}
    `;
  }

  function renderLiveChapterCards() {
    return `
      <section class="assessment-live-panel">
        <div class="assessment-panel-header">
          <div>
            <p class="assessment-panel-eyebrow">Live tenantdata</p>
            <h3>Per hoofdstuk en subhoofdstuk direct ophalen</h3>
          </div>
        </div>
        <div class="assessment-live-grid">
          ${LIVE_CHAPTERS.map((chapter) => `
            <article class="assessment-live-card">
              <div class="assessment-live-card-head">
                <div>
                  <h4>${assessmentEscapeHtml(chapter.title)}</h4>
                  <p>${assessmentEscapeHtml(chapter.description)}</p>
                </div>
                <button
                  type="button"
                  class="assessment-live-fetch"
                  data-live-title="${assessmentEscapeHtml(chapter.primaryAction.title)}"
                  data-live-path="${assessmentEscapeHtml(chapter.primaryAction.path(state.tenantId || ''))}"
                  data-live-key="${assessmentEscapeHtml(chapter.key)}"
                >${assessmentEscapeHtml(chapter.primaryAction.label)}</button>
              </div>
              ${chapter.actions.some((action) => action.requiresInput === 'domain') ? `
                <div class="assessment-live-inline-form">
                  <input type="text" id="assessmentLiveDomainInput" class="assessment-live-input" placeholder="bijv. contoso.nl" />
                </div>
              ` : ''}
              <div class="assessment-live-actions">
                ${chapter.actions.map((action) => `
                  <button
                    type="button"
                    class="assessment-live-subaction"
                    data-live-title="${assessmentEscapeHtml(action.title)}"
                    data-live-action-key="${assessmentEscapeHtml(action.key)}"
                    data-live-chapter="${assessmentEscapeHtml(chapter.key)}"
                  >${assessmentEscapeHtml(action.label)}</button>
                `).join('')}
              </div>
            </article>
          `).join('')}
        </div>
        <div class="assessment-live-result" id="assessmentLiveResult">
          <div class="assessment-live-placeholder">
            Klik op een hoofdstuk of subhoofdstuk om live tenantdata op te halen.
          </div>
        </div>
      </section>
    `;
  }

  function normalizeLiveRows(rows) {
    if (!Array.isArray(rows) || !rows.length) return { columns: [], rows: [] };
    const sample = rows.find((row) => row && typeof row === 'object') || {};
    const columns = Object.keys(sample).slice(0, 8);
    const normalizedRows = rows.slice(0, 100).map((row) => {
      const out = {};
      columns.forEach((column) => {
        out[column] = assessmentFormatCell(row ? row[column] : null);
      });
      return out;
    });
    return { columns, rows: normalizedRows };
  }

  function extractLiveCollection(data) {
    const collectionKeys = ['users', 'guests', 'roles', 'policies', 'profiles', 'devices', 'domains', 'items', 'mailboxes', 'rules', 'forwarding', 'locations', 'sites', 'teams', 'apps'];
    for (const key of collectionKeys) {
      if (Array.isArray(data?.[key])) return { key, value: data[key] };
    }
    return null;
  }

  function renderLiveSummary(data) {
    const summaryKeys = Object.keys(data || {}).filter((key) => {
      const value = data[key];
      return !Array.isArray(value) && (typeof value !== 'object' || value === null);
    }).slice(0, 8);
    if (!summaryKeys.length) return '';
    return `
      <div class="assessment-live-summary">
        ${summaryKeys.map((key) => `
          <div class="assessment-live-summary-item">
            <span>${assessmentEscapeHtml(key)}</span>
            <strong>${assessmentEscapeHtml(assessmentFormatCell(data[key]))}</strong>
          </div>
        `).join('')}
      </div>
    `;
  }

  function renderLiveResult(title, data) {
    const resultRoot = document.getElementById('assessmentLiveResult');
    if (!resultRoot) return;
    const collection = extractLiveCollection(data);
    const summary = renderLiveSummary(data);

    if (!collection) {
      const objectRows = Object.keys(data || {}).slice(0, 20).map((key) => ({ key, value: assessmentFormatCell(data[key]) }));
      resultRoot.innerHTML = `
        <div class="assessment-live-result-head">
          <div>
            <p class="assessment-panel-eyebrow">Live resultaat</p>
            <h3>${assessmentEscapeHtml(title)}</h3>
          </div>
        </div>
        ${summary}
        <div class="assessment-table-wrap">
          <table class="assessment-table">
            <thead><tr><th>Veld</th><th>Waarde</th></tr></thead>
            <tbody>
              ${objectRows.map((row) => `
                <tr>
                  <td>${assessmentEscapeHtml(row.key)}</td>
                  <td>${assessmentEscapeHtml(row.value)}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>
      `;
      return;
    }

    const normalized = normalizeLiveRows(collection.value);
    resultRoot.innerHTML = `
      <div class="assessment-live-result-head">
        <div>
          <p class="assessment-panel-eyebrow">Live resultaat</p>
          <h3>${assessmentEscapeHtml(title)}</h3>
        </div>
        <div class="assessment-live-count">${assessmentEscapeHtml(collection.key)}: ${assessmentEscapeHtml(collection.value.length)}</div>
      </div>
      ${summary}
      <div class="assessment-table-wrap">
        <table class="assessment-table">
          <thead>
            <tr>${normalized.columns.map((column) => `<th>${assessmentEscapeHtml(column)}</th>`).join('')}</tr>
          </thead>
          <tbody>
            ${normalized.rows.map((row) => `
              <tr>
                ${normalized.columns.map((column) => `<td>${assessmentEscapeHtml(row[column])}</td>`).join('')}
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    `;
  }

  function setLiveLoading(title) {
    const resultRoot = document.getElementById('assessmentLiveResult');
    if (!resultRoot) return;
    resultRoot.innerHTML = `
      <div class="assessment-live-placeholder">
        ${assessmentEscapeHtml(title)} wordt opgehaald...
      </div>
    `;
  }

  function resolveLiveAction(actionKey, chapterKey) {
    const chapter = LIVE_CHAPTERS.find((item) => item.key === chapterKey);
    if (!chapter) return null;
    return chapter.actions.find((action) => action.key === actionKey) || null;
  }

  function getLiveFormValues() {
    return {
      domain: document.getElementById('assessmentLiveDomainInput')?.value?.trim() || '',
    };
  }

  async function fetchLiveData({ title, path, loadingKey }) {
    if (!state.tenantId || !path) return;
    state.liveLoadingKey = loadingKey || null;
    setLiveLoading(title);
    try {
      const data = await assessmentFetchJson(path);
      renderLiveResult(title, data || {});
    } catch (error) {
      const resultRoot = document.getElementById('assessmentLiveResult');
      if (resultRoot) {
        resultRoot.innerHTML = `
          <div class="assessment-live-error">
            <strong>${assessmentEscapeHtml(title)}</strong>
            <span>${assessmentEscapeHtml(error.message || 'Ophalen mislukt')}</span>
          </div>
        `;
      }
    } finally {
      state.liveLoadingKey = null;
    }
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
              <button type="button" class="assessment-action-btn assessment-action-btn-secondary" data-assessment-action="show-results">Bekijk resultaten</button>
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
      const livePrimaryButton = event.target.closest('[data-live-path]');
      if (livePrimaryButton) {
        event.preventDefault();
        fetchLiveData({
          title: livePrimaryButton.dataset.liveTitle || 'Live data',
          path: livePrimaryButton.dataset.livePath,
          loadingKey: livePrimaryButton.dataset.liveKey || '',
        });
        return;
      }

      const liveSubButton = event.target.closest('[data-live-action-key]');
      if (!liveSubButton) return;
      event.preventDefault();
      const action = resolveLiveAction(liveSubButton.dataset.liveActionKey, liveSubButton.dataset.liveChapter);
      if (!action || !state.tenantId) return;
      const formValues = getLiveFormValues();
      if (action.requiresInput === 'domain' && !formValues.domain) {
        renderLiveResult(action.title, { error: 'Vul eerst een domeinnaam in om een analyse te starten.' });
        return;
      }
      fetchLiveData({
        title: action.title,
        path: action.path(state.tenantId, formValues),
        loadingKey: action.key,
      });
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
