(function () {
  const PILLAR_META = {
    Identity: { key: 'identity', icon: 'ID', description: 'Identiteiten, toegang en sessiebeveiliging.' },
    Devices: { key: 'devices', icon: 'DV', description: 'Device posture, beheer en compliancestatus.' },
    Network: { key: 'network', icon: 'NW', description: 'Netwerksegmentatie, toegang en verkeerscontrole.' },
    Data: { key: 'data', icon: 'DT', description: 'Bescherming van data, labeling en toegangsbeleid.' },
  };

  const CHAPTERS = [
    { key: 'overview', label: 'Overzicht', title: 'Executive overzicht' },
    { key: 'identity', label: 'Identity', title: 'Identity pillar' },
    { key: 'devices', label: 'Devices', title: 'Devices pillar' },
    { key: 'network', label: 'Network', title: 'Network pillar' },
    { key: 'data', label: 'Data', title: 'Data pillar' },
    { key: 'json', label: 'JSON', title: 'Brondata' },
  ];

  const state = {
    tenantId: null,
    chapter: 'overview',
    payload: null,
    loading: false,
    activity: null,
    pollTimer: null,
  };

  function root() {
    return document.getElementById('zerotrustModuleRoot');
  }

  function getTenantId() {
    if (typeof currentTenantId !== 'undefined' && currentTenantId) return currentTenantId;
    return document.getElementById('tenantSelect')?.value || null;
  }

  function esc(value) {
    return String(value == null ? '' : value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function formatDate(value) {
    if (!value) return 'Nog geen rapport';
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return String(value);
    return date.toLocaleString('nl-NL', {
      day: '2-digit',
      month: 'short',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
    });
  }

  function scoreTone(score) {
    if (score >= 80) return 'is-good';
    if (score >= 50) return 'is-warn';
    return 'is-risk';
  }

  function statusTone(status) {
    const normalized = String(status || '').toLowerCase();
    if (normalized === 'pass') return 'is-pass';
    if (normalized === 'warning') return 'is-warning';
    if (normalized === 'fail') return 'is-fail';
    return 'is-na';
  }

  function normalizePayload(payload) {
    const moduleInfo = payload?.module || {};
    const lastReport = payload?.last_report || null;
    const results = payload?.results || {};
    const controls = Array.isArray(results?.controls) ? results.controls.filter(Boolean) : [];
    const summary = results?.summary || {};
    const pillars = results?.pillars || {};
    const backendStatus = payload?.status || {};
    const recentLogs = Array.isArray(payload?.recent_logs) ? payload.recent_logs : [];
    const authProfile = payload?.auth_profile || {};
    const linkedApp = payload?.linked_app_registration || {};
    const permissionSummary = payload?.permission_summary || {};
    const pillarCards = Object.keys(PILLAR_META).map((name) => {
      const score = Number(pillars[name]);
      const items = controls.filter((item) => String(item?.pillar || '').toLowerCase() === name.toLowerCase());
      return {
        name,
        key: PILLAR_META[name].key,
        icon: PILLAR_META[name].icon,
        description: PILLAR_META[name].description,
        score: Number.isFinite(score) ? score : null,
        total: items.length,
        pass: items.filter((item) => String(item?.status).toLowerCase() === 'pass').length,
        warning: items.filter((item) => String(item?.status).toLowerCase() === 'warning').length,
        fail: items.filter((item) => String(item?.status).toLowerCase() === 'fail').length,
      };
    });

    return {
      raw: payload,
      moduleInfo,
      lastReport,
      backendStatus,
      recentLogs,
      authProfile,
      linkedApp,
      permissionSummary,
      results,
      controls,
      summary: {
        score: Number(summary?.score || 0),
        total: Number(summary?.total || controls.length || 0),
        pass: Number(summary?.pass || 0),
        fail: Number(summary?.fail || 0),
        warning: Number(summary?.warning || 0),
      },
      pillarCards,
    };
  }

  function renderEmpty(message) {
    const host = root();
    if (!host) return;
    host.innerHTML = `<div class="zerotrust-empty">${esc(message)}</div>`;
  }

  function renderLoading() {
    const host = root();
    if (!host) return;
    host.innerHTML = `
      <div class="zerotrust-shell">
        <aside class="zerotrust-rail">
          ${Array.from({ length: 6 }, () => '<div class="zerotrust-skeleton zerotrust-skeleton-line"></div>').join('')}
        </aside>
        <section class="zerotrust-stage">
          <div class="zerotrust-skeleton zerotrust-skeleton-hero"></div>
          <div class="zerotrust-grid">
            ${Array.from({ length: 4 }, () => '<div class="zerotrust-skeleton zerotrust-skeleton-card"></div>').join('')}
          </div>
        </section>
      </div>
    `;
  }

  function renderAuthPanel(data) {
    const auth = data.authProfile || {};
    const app = data.linkedApp || {};
    const permissionSummary = data.permissionSummary || {};
    const extraPermissions = Array.isArray(permissionSummary.additional_permissions) ? permissionSummary.additional_permissions.slice(0, 8) : [];
    const hasApp = !!auth.client_id;
    const preferredAuth = auth.preferred_auth_mode === 'app' ? 'App-registratie' : 'Interactief';
    const credentialLabel = auth.has_certificate ? 'Certificaat' : (auth.has_client_secret ? 'Client secret' : 'Geen');
    const appName = app.displayName || 'Nog niet gevonden in app-registraties';
    const sourceLabel = app.source || 'configuratie';

    return `
      <section class="zerotrust-panel">
        <div class="zerotrust-section-head">
          <div>
            <span class="zerotrust-eyebrow">Authenticatie</span>
            <h4>Gekoppelde app-registratie en API-rechten</h4>
          </div>
          <span class="zerotrust-status ${auth.preferred_auth_mode === 'app' ? 'is-pass' : 'is-warning'}">${esc(preferredAuth)}</span>
        </div>
        <div class="zerotrust-detail-band">
          <article>
            <span>App ID</span>
            <strong>${esc(auth.client_id || 'Niet gekoppeld')}</strong>
          </article>
          <article>
            <span>Cred type</span>
            <strong>${esc(credentialLabel)}</strong>
          </article>
          <article>
            <span>App-auth gereed</span>
            <strong>${auth.app_auth_ready ? 'Ja' : 'Nee'}</strong>
          </article>
          <article>
            <span>App registratie</span>
            <strong>${esc(hasApp ? appName : 'Nog niet gekoppeld')}</strong>
          </article>
          <article>
            <span>Extra API-rechten</span>
            <strong>${esc(permissionSummary.additional_count || 0)}</strong>
          </article>
        </div>
        <p class="zerotrust-paragraph">
          ${esc(auth.fallback_reason || `Vergelijking op basis van ${permissionSummary.reference_count || 0} Microsoft Zero Trust-permissies. Bron appdata: ${sourceLabel}.`)}
        </p>
        ${extraPermissions.length ? `
          <div class="zerotrust-table-wrap is-compact">
            <table class="zerotrust-table">
              <thead>
                <tr>
                  <th>Resource</th>
                  <th>Permissie</th>
                  <th>Type</th>
                </tr>
              </thead>
              <tbody>
                ${extraPermissions.map((item) => `
                  <tr>
                    <td>${esc(item.resource || 'Onbekend')}</td>
                    <td>${esc(item.permission || '—')}</td>
                    <td>${esc(item.type || '—')}</td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>
        ` : '<div class="zerotrust-empty-row">Geen extra API-rechten gevonden buiten de Microsoft Zero Trust referentieset.</div>'}
      </section>
    `;
  }

  function renderActivityBanner() {
    if (!state.activity?.active) return '';
    const message = state.activity.message || 'Zero Trust verwerking actief op de backend.';
    const detail = state.activity.detail || 'Status wordt op de achtergrond ververst.';
    return `
      <section class="zerotrust-activity-banner">
        <div class="zerotrust-activity-copy">
          <strong>${esc(message)}</strong>
          <span>${esc(detail)}</span>
        </div>
        <div class="zerotrust-activity-bar" aria-hidden="true">
          <span class="zerotrust-activity-bar-fill"></span>
        </div>
      </section>
    `;
  }

  function setActivity(kind, message, detail) {
    state.activity = {
      kind,
      active: true,
      message,
      detail,
      startedAt: Date.now(),
    };
  }

  function clearActivity() {
    state.activity = null;
    if (state.pollTimer) {
      window.clearInterval(state.pollTimer);
      state.pollTimer = null;
    }
  }

  function ensurePolling() {
    if (state.pollTimer) return;
    state.pollTimer = window.setInterval(() => {
      if (!state.activity?.active) {
        window.clearInterval(state.pollTimer);
        state.pollTimer = null;
        return;
      }
      loadZeroTrustSection(state.chapter || 'overview', { silent: true });
    }, 15000);
  }

  async function fetchPayload(tenantId) {
    const token = localStorage.getItem('denjoy_token');
    const response = await fetch(`/api/compliance/${tenantId}/zerotrust`, {
      credentials: 'include',
      headers: {
        'Content-Type': 'application/json',
        ...(token ? { Authorization: `Bearer ${token}` } : {}),
      },
    });
    const data = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(data?.error || `HTTP ${response.status}`);
    }
    return data;
  }

  function chapterButton(chapter) {
    const isActive = state.chapter === chapter.key;
    return `
      <button type="button" class="zerotrust-chapter-btn${isActive ? ' active' : ''}" data-zt-chapter="${esc(chapter.key)}">
        <span>${esc(chapter.label)}</span>
        <small>${esc(chapter.title)}</small>
      </button>
    `;
  }

  function renderHero(data) {
    const score = data.summary.score || 0;
    const tone = scoreTone(score);
    const moduleLabel = data.moduleInfo.installed ? `Module ${esc(data.moduleInfo.version || 'beschikbaar')}` : 'Module ontbreekt';
    const reportLabel = data.lastReport?.date ? formatDate(data.lastReport.date) : 'Nog geen rapport';
    return `
      <section class="zerotrust-hero ${tone}">
        <div class="zerotrust-hero-copy">
          <span class="zerotrust-kicker">Microsoft Zero Trust Assessment</span>
          <h3>Eigen hoofdstuk voor rapport, controls en JSON-brondata</h3>
          <p>Geinspireerd op de opbouw van het Microsoft voorbeeld, maar uitgewerkt in de Denjoy portalstijl met hoofdstuknavigatie, pijlerkaarten en een brondata-paneel.</p>
        </div>
        <div class="zerotrust-hero-stats">
          <article class="zerotrust-score-orb">
            <span>Overall score</span>
            <strong>${esc(score)}%</strong>
            <small>${esc(moduleLabel)}</small>
          </article>
          <div class="zerotrust-hero-meta">
            <div><span>Laatste rapport</span><strong>${esc(reportLabel)}</strong></div>
            <div><span>Controls</span><strong>${esc(data.summary.total)}</strong></div>
            <div><span>Pass / Warn / Fail</span><strong>${esc(`${data.summary.pass} / ${data.summary.warning} / ${data.summary.fail}`)}</strong></div>
          </div>
        </div>
      </section>
    `;
  }

  function renderStatusPanel(data) {
    const status = data.backendStatus || {};
    const stateLabel = {
      queued: 'In wachtrij',
      running: 'Bezig',
      completed: 'Afgerond',
      failed: 'Mislukt',
    }[String(status.state || '')] || 'Onbekend';
    const authModeLabel = status.auth_mode === 'interactive' ? 'Interactief' : (status.auth_mode === 'app' ? 'App-auth' : '—');
    const statusTone = {
      queued: 'is-neutral',
      running: 'is-warning',
      completed: 'is-pass',
      failed: 'is-fail',
    }[String(status.state || '')] || 'is-neutral';
    const logs = data.recentLogs || [];
    return `
      <section class="zerotrust-panel">
        <div class="zerotrust-section-head">
          <div>
            <span class="zerotrust-eyebrow">Backend status</span>
            <h4>Runstatus en laatste logregels</h4>
          </div>
          <span class="zerotrust-status ${statusTone}">${esc(stateLabel)}</span>
        </div>
        <div class="zerotrust-detail-band">
          <article>
            <span>Actie</span>
            <strong>${esc(status.action || '—')}</strong>
          </article>
          <article>
            <span>Gestart</span>
            <strong>${esc(formatDate(status.started_at))}</strong>
          </article>
          <article>
            <span>Laatste update</span>
            <strong>${esc(formatDate(status.updated_at))}</strong>
          </article>
          <article>
            <span>Resultaat</span>
            <strong>${esc(status.message || 'Nog geen extra status.')}</strong>
          </article>
          <article>
            <span>Authenticatie</span>
            <strong>${esc(authModeLabel)}</strong>
          </article>
        </div>
        <div class="zerotrust-log-panel">
          ${logs.length ? `<pre class="zerotrust-log-viewer">${esc(logs.join('\n'))}</pre>` : '<div class="zerotrust-empty-row">Nog geen Zero Trust-logregels beschikbaar.</div>'}
        </div>
      </section>
    `;
  }

  function renderOverview(data) {
    const cards = [
      { label: 'Pass', value: data.summary.pass, tone: 'is-pass' },
      { label: 'Warning', value: data.summary.warning, tone: 'is-warning' },
      { label: 'Fail', value: data.summary.fail, tone: 'is-fail' },
      { label: 'Module', value: data.moduleInfo.installed ? (data.moduleInfo.version || 'Geinstalleerd') : 'Niet geinstalleerd', tone: 'is-neutral' },
    ];
    const focusControls = data.controls
      .filter((item) => ['fail', 'warning'].includes(String(item?.status || '').toLowerCase()))
      .slice(0, 12);

    return `
      <section class="zerotrust-panel">
        <div class="zerotrust-section-head">
          <div>
            <span class="zerotrust-eyebrow">Hoofdstuk 1</span>
            <h4>Overzicht en aanbevelingen</h4>
          </div>
          <button type="button" class="zerotrust-run-btn" data-zt-run="1">Assessment opnieuw uitvoeren</button>
        </div>
        <div class="zerotrust-kpi-grid">
          ${cards.map((card) => `
            <article class="zerotrust-kpi ${card.tone}">
              <span>${esc(card.label)}</span>
              <strong>${esc(card.value)}</strong>
            </article>
          `).join('')}
        </div>
      </section>

      ${renderStatusPanel(data)}
      ${renderAuthPanel(data)}

      <section class="zerotrust-pillar-grid">
        ${data.pillarCards.map((pillar) => `
          <article class="zerotrust-pillar-card ${pillar.score != null ? scoreTone(pillar.score) : 'is-neutral'}">
            <div class="zerotrust-pillar-top">
              <span class="zerotrust-pillar-mark">${esc(pillar.icon)}</span>
              <div>
                <h4>${esc(pillar.name)}</h4>
                <p>${esc(pillar.description)}</p>
              </div>
            </div>
            <div class="zerotrust-pillar-score">${pillar.score != null ? `${esc(pillar.score)}%` : '—'}</div>
            <div class="zerotrust-pillar-meta">
              <span>${esc(`${pillar.total} controls`)}</span>
              <span>${esc(`${pillar.fail} fail`)}</span>
              <span>${esc(`${pillar.warning} warning`)}</span>
            </div>
            <button type="button" class="zerotrust-link-btn" data-zt-chapter="${esc(pillar.key)}">Open hoofdstuk</button>
          </article>
        `).join('')}
      </section>

      <section class="zerotrust-panel">
        <div class="zerotrust-section-head">
          <div>
            <span class="zerotrust-eyebrow">Focus</span>
            <h4>Controls die aandacht vragen</h4>
          </div>
          <span class="zerotrust-inline-note">${esc(focusControls.length)} getoond</span>
        </div>
        ${renderControlsTable(focusControls, true)}
      </section>

      ${renderStatusPanel(data)}
    `;
  }

  function renderControlsTable(controls, compact) {
    if (!controls.length) {
      return '<div class="zerotrust-empty-row">Geen controls beschikbaar voor dit hoofdstuk.</div>';
    }
    return `
      <div class="zerotrust-table-wrap${compact ? ' is-compact' : ''}">
        <table class="zerotrust-table">
          <thead>
            <tr>
              <th>Control</th>
              <th>Pillar</th>
              <th>Status</th>
              <th>Risico</th>
            </tr>
          </thead>
          <tbody>
            ${controls.map((control) => `
              <tr>
                <td>
                  <strong>${esc(control.title || '—')}</strong>
                </td>
                <td>${esc(control.pillar || '—')}</td>
                <td><span class="zerotrust-status ${statusTone(control.status)}">${esc(control.status || 'NA')}</span></td>
                <td>${esc(control.riskLevel || '—')}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    `;
  }

  function renderPillarChapter(data, key) {
    const pillarName = Object.keys(PILLAR_META).find((name) => PILLAR_META[name].key === key) || 'Identity';
    const pillar = data.pillarCards.find((item) => item.key === key);
    const controls = data.controls.filter((item) => String(item?.pillar || '').toLowerCase() === pillarName.toLowerCase());
    return `
      <section class="zerotrust-panel">
        <div class="zerotrust-section-head">
          <div>
            <span class="zerotrust-eyebrow">Pillar</span>
            <h4>${esc(pillarName)}</h4>
          </div>
          <button type="button" class="zerotrust-link-btn" data-zt-chapter="json">Bekijk brondata</button>
        </div>
        <div class="zerotrust-detail-band">
          <article>
            <span>Score</span>
            <strong>${pillar?.score != null ? `${esc(pillar.score)}%` : '—'}</strong>
          </article>
          <article>
            <span>Pass</span>
            <strong>${esc(pillar?.pass || 0)}</strong>
          </article>
          <article>
            <span>Warning</span>
            <strong>${esc(pillar?.warning || 0)}</strong>
          </article>
          <article>
            <span>Fail</span>
            <strong>${esc(pillar?.fail || 0)}</strong>
          </article>
        </div>
        <p class="zerotrust-paragraph">${esc(PILLAR_META[pillarName]?.description || '')}</p>
        ${renderControlsTable(controls, false)}
      </section>
    `;
  }

  function renderJsonChapter(data) {
    return `
      <section class="zerotrust-panel">
        <div class="zerotrust-section-head">
          <div>
            <span class="zerotrust-eyebrow">Brondata</span>
            <h4>Genormaliseerde JSON-output</h4>
          </div>
          <button type="button" class="zerotrust-link-btn" data-zt-chapter="overview">Terug naar overzicht</button>
        </div>
        <p class="zerotrust-paragraph">Deze weergave laat de ruwe payload zien die uit de backend-endpoint komt, zodat JSON-bestanden ook direct uitleesbaar en controleerbaar blijven binnen het portaal.</p>
        <pre class="zerotrust-json-viewer">${esc(JSON.stringify(data.raw, null, 2))}</pre>
      </section>

      ${renderStatusPanel(data)}
    `;
  }

  function renderInstallState() {
    const host = root();
    if (!host) return;
    const data = state.payload ? normalizePayload(state.payload) : null;
    host.innerHTML = `
      <div class="zerotrust-shell">
        <aside class="zerotrust-rail">
          ${CHAPTERS.map(chapterButton).join('')}
        </aside>
        <section class="zerotrust-stage">
          ${renderActivityBanner()}
          <section class="zerotrust-panel zerotrust-panel-emphasis">
            <div class="zerotrust-section-head">
              <div>
                <span class="zerotrust-eyebrow">Module status</span>
                <h4>ZeroTrustAssessment module niet gevonden</h4>
              </div>
              <div class="zerotrust-action-row">
                <button type="button" class="zerotrust-link-btn" data-zt-install="1">Module installeren</button>
                <button type="button" class="zerotrust-run-btn" data-zt-run="1">Assessment starten</button>
              </div>
            </div>
            <p class="zerotrust-paragraph">De backend-endpoint is beschikbaar, maar op de assessmentserver is nog geen Microsoft Zero Trust Assessment module gevonden of er is nog geen rapport aangemaakt. Je kunt de module nu rechtstreeks vanuit het portaal laten installeren.</p>
            <pre class="zerotrust-json-viewer">Install-Module ZeroTrustAssessment -Scope CurrentUser
Connect-ZtAssessment
Invoke-ZtAssessment</pre>
          </section>
          ${data ? renderStatusPanel(data) : ''}
        </section>
      </div>
    `;
  }

  function renderWorkspace() {
    const host = root();
    if (!host) return;
    const data = normalizePayload(state.payload);

    if (!data.moduleInfo.installed && !data.lastReport) {
      renderInstallState();
      return;
    }

    let body = renderOverview(data);
    if (state.chapter !== 'overview') {
      body = state.chapter === 'json'
        ? renderJsonChapter(data)
        : renderPillarChapter(data, state.chapter);
    }

    host.innerHTML = `
      <div class="zerotrust-shell">
        <aside class="zerotrust-rail">
          <div class="zerotrust-rail-head">
            <span class="zerotrust-eyebrow">Hoofdstukken</span>
            <strong>${esc(getTenantId() || 'Tenant')}</strong>
          </div>
          ${CHAPTERS.map(chapterButton).join('')}
        </aside>
        <section class="zerotrust-stage">
          ${renderActivityBanner()}
          ${renderHero(data)}
          ${body}
        </section>
      </div>
    `;
  }

  async function loadZeroTrustSection(chapter, options = {}) {
    const tenantId = getTenantId();
    state.tenantId = tenantId;
    state.chapter = chapter || state.chapter || 'overview';
    const silent = !!options.silent;

    if (typeof setActiveSubnavItem === 'function') setActiveSubnavItem(state.chapter);
    if (!tenantId) {
      state.payload = null;
      clearActivity();
      renderEmpty('Selecteer eerst een tenant om Zero Trust-resultaten op te halen.');
      return;
    }

    state.loading = true;
    if (!silent || !state.payload) renderLoading();

    try {
      state.payload = await fetchPayload(tenantId);
      const serverState = String(state.payload?.status?.state || '');
      if (serverState === 'queued') {
        setActivity('queued', 'Zero Trust-taak staat in de wachtrij', state.payload?.status?.message || 'De backend heeft de taak ontvangen en wacht op uitvoering.');
        ensurePolling();
      } else if (serverState === 'running') {
        setActivity(
          String(state.payload?.status?.action || 'run'),
          state.payload?.status?.action === 'install' ? 'Zero Trust module-installatie draait' : 'Zero Trust Assessment draait',
          state.payload?.status?.message || 'Laatste status wordt automatisch ververst.'
        );
        ensurePolling();
      } else if (serverState === 'completed' || serverState === 'failed') {
        clearActivity();
      }
      if (state.activity?.active) {
        const payload = state.payload || {};
        const moduleInstalled = !!payload?.module?.installed;
        const hasReport = !!payload?.last_report;
        if (state.activity.kind === 'install' && moduleInstalled) {
          clearActivity();
        } else if (state.activity.kind === 'run' && hasReport && !payload?.error) {
          state.activity.detail = 'Laatste bekende status opgehaald. Gebruik verversen om nieuwe output direct te bekijken.';
        }
      }
      renderWorkspace();
    } catch (error) {
      renderEmpty(`Zero Trust-data laden mislukt: ${error.message || 'Onbekende fout'}`);
    } finally {
      state.loading = false;
    }
  }

  function switchZeroTrustChapter(chapter) {
    state.chapter = chapter || 'overview';
    if (typeof setActiveSubnavItem === 'function') setActiveSubnavItem(state.chapter);
    if (state.payload) {
      renderWorkspace();
      return;
    }
    loadZeroTrustSection(state.chapter);
  }

  async function runAssessment() {
    const tenantId = getTenantId();
    if (!tenantId) {
      if (typeof showToast === 'function') showToast('Selecteer eerst een tenant.', 'warning');
      return;
    }

    const token = localStorage.getItem('denjoy_token');
    const response = await fetch(`/api/compliance/${tenantId}/zerotrust/run`, {
      method: 'POST',
      credentials: 'include',
      headers: {
        'Content-Type': 'application/json',
        ...(token ? { Authorization: `Bearer ${token}` } : {}),
      },
      body: JSON.stringify({ force_interactive: false }),
    });
    const data = await response.json().catch(() => ({}));
    if (!response.ok || !data?.ok) {
      throw new Error(data?.error || 'Assessment starten mislukt');
    }
    if (typeof showToast === 'function') {
      showToast(data?.message || 'Zero Trust Assessment gestart. Controleer de status en logregels in deze pagina.', 'info', 5000);
    }
  }

  async function installModule() {
    const tenantId = getTenantId();
    if (!tenantId) {
      if (typeof showToast === 'function') showToast('Selecteer eerst een tenant.', 'warning');
      return;
    }

    const token = localStorage.getItem('denjoy_token');
    const response = await fetch(`/api/compliance/${tenantId}/zerotrust/install`, {
      method: 'POST',
      credentials: 'include',
      headers: {
        'Content-Type': 'application/json',
        ...(token ? { Authorization: `Bearer ${token}` } : {}),
      },
    });
    const data = await response.json().catch(() => ({}));
    if (!response.ok || !data?.ok) {
      throw new Error(data?.error || 'Module installatie starten mislukt');
    }
    if (typeof showToast === 'function') {
      showToast('Zero Trust module installatie gestart op de backend.', 'info', 5000);
    }
  }

  document.addEventListener('click', async (event) => {
    const chapterBtn = event.target.closest('[data-zt-chapter]');
    if (chapterBtn) {
      event.preventDefault();
      switchZeroTrustChapter(chapterBtn.dataset.ztChapter);
      return;
    }

    const runBtn = event.target.closest('[data-zt-run]');
    const installBtn = event.target.closest('[data-zt-install]');

    if (!runBtn && !installBtn) return;

    event.preventDefault();
    const actionBtn = runBtn || installBtn;
    const original = actionBtn.textContent;
    actionBtn.disabled = true;
    actionBtn.textContent = runBtn ? 'Starten...' : 'Installeren...';
    try {
      if (runBtn) {
        await runAssessment();
        setActivity(
          'run',
          'Zero Trust Assessment gestart',
          'De backend gebruikt nu eerst de gekoppelde app-registratie. Als extra login nodig is, wordt normale Microsoft browser-auth gebruikt in plaats van device-code.'
        );
        ensurePolling();
        actionBtn.textContent = 'Gestart';
      } else {
        await installModule();
        setActivity(
          'install',
          'Zero Trust module wordt geinstalleerd',
          'De backend haalt de module nu op vanuit PowerShell Gallery en controleert daarna automatisch de status.'
        );
        ensurePolling();
        actionBtn.textContent = 'Installatie gestart';
      }
      window.setTimeout(() => {
        if (getTenantId() === state.tenantId) loadZeroTrustSection(state.chapter || 'overview', { silent: true });
      }, runBtn ? 1500 : 2500);
    } catch (error) {
      actionBtn.textContent = 'Mislukt';
      if (typeof showToast === 'function') showToast(error.message || 'Assessment starten mislukt.', 'error');
      window.setTimeout(() => {
        actionBtn.disabled = false;
        actionBtn.textContent = original;
      }, 1200);
      return;
    }
    window.setTimeout(() => {
      actionBtn.disabled = false;
      actionBtn.textContent = original;
    }, 3000);
  });

  window.loadZeroTrustSection = loadZeroTrustSection;
  window.switchZeroTrustChapter = switchZeroTrustChapter;
})();
