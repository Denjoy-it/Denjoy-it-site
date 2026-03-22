// EntraFalcon — Entra ID security scan (zelfde patroon als assessment.js)

let efActiveRunId = null;
let efRunPollTimer = null;
let efLogPollTimer = null;
let efIsPolling = false;

// ── Helpers ───────────────────────────────────────────────────────────────────

function efLog(message, type = 'info') {
  const log = document.getElementById('efProgressLog');
  if (!log) return;
  const p = document.createElement('p');
  p.textContent = `[${new Date().toLocaleTimeString()}] ${message}`;
  if (type === 'error')   p.style.color = '#ffb4b4';
  if (type === 'success') p.style.color = '#b7f7c8';
  if (type === 'warning') p.style.color = '#ffe08a';
  log.appendChild(p);
  log.scrollTop = log.scrollHeight;
}

function efSetProgress(pct) {
  const fill = document.getElementById('efProgressFill');
  if (!fill) return;
  fill.style.width = `${Math.max(0, Math.min(100, pct))}%`;
  fill.textContent = `${Math.round(pct)}%`;
}

// Schat voortgang op basis van de 10 EntraFalcon-stappen in de logs
function efEstimateProgress(lines) {
  const steps = [
    'Gather Basic Data', 'Enumerating Groups', 'Enumerating Enterprise Apps',
    'Enumerating Managed', 'Enumerating App Reg', 'Enumerating Users',
    'Role Assignments', 'Conditional Access', 'PIM Role',
    'Security Findings', 'Summary Report',
  ];
  const done = steps.filter((s) => lines.some((l) => l.includes(s))).length;
  return Math.min(95, Math.round((done / steps.length) * 90) + 5);
}

function efShowConfig() {
  document.getElementById('entrafalconConfig').style.display = '';
  document.getElementById('entrafalconProgress').style.display = 'none';
}

function efShowProgress() {
  document.getElementById('entrafalconConfig').style.display = 'none';
  document.getElementById('entrafalconProgress').style.display = 'block';
}

function efStopPolling() {
  if (efRunPollTimer) clearInterval(efRunPollTimer);
  if (efLogPollTimer) clearInterval(efLogPollTimer);
  efRunPollTimer = null;
  efLogPollTimer = null;
  efIsPolling = false;
}

// ── Logs polling ──────────────────────────────────────────────────────────────

async function efPollLogs(runId) {
  try {
    const data = await fetch(`/api/runs/${runId}/logs`).then((r) => r.json());
    const log = document.getElementById('efProgressLog');
    if (!log) return data.lines || [];
    log.innerHTML = '';
    for (const line of (data.lines || [])) {
      const p = document.createElement('p');
      p.textContent = line;
      if (/failed|mislukt|error/i.test(line))            p.style.color = '#ffb4b4';
      else if (/completed|voltooid|success/i.test(line)) p.style.color = '#b7f7c8';
      else if (/device code|authenticat|interactief/i.test(line)) p.style.color = '#ffe08a';
      log.appendChild(p);
    }
    log.scrollTop = log.scrollHeight;
    return data.lines || [];
  } catch (e) {
    console.warn('EF log poll', e);
    return [];
  }
}

// ── Status polling ────────────────────────────────────────────────────────────

async function efPollStatus(runId) {
  try {
    const res = await fetch(`/api/runs/${runId}`);
    const run = await res.json();
    if (!res.ok) throw new Error(run.error || `HTTP ${res.status}`);

    const title = document.querySelector('#entrafalconProgress h3');
    if (title) title.textContent = `Entra Security scan: ${run.status}`;

    if (run.status === 'completed') {
      efSetProgress(100);
      efLog('Entra Security scan voltooid.', 'success');
      efStopPolling();
      efActiveRunId = null;
      efShowConfig();
      if (typeof refreshTenantData === 'function') await refreshTenantData();
      loadEntraFalconSection();
      if (typeof showSection === 'function') showSection('results', { resultsPanel: 'entrafalcon' });
      return { terminal: true };
    }

    if (run.status === 'failed') {
      efLog(`Scan mislukt: ${run.error_message || 'Onbekende fout'}`, 'error');
      efStopPolling();
      efActiveRunId = null;
      efShowConfig();
      if (typeof refreshTenantData === 'function') await refreshTenantData();
      return { terminal: true };
    }

    if (run.status === 'cancelled') {
      efLog('⏹ Scan gestopt.', 'warning');
      efStopPolling();
      efActiveRunId = null;
      const stopBtn = document.getElementById('efStopButton');
      if (stopBtn) { stopBtn.disabled = false; stopBtn.textContent = '⏹ Stoppen'; }
      efShowConfig();
      if (typeof refreshTenantData === 'function') await refreshTenantData();
      return { terminal: true };
    }

    const lines = await efPollLogs(runId);
    efSetProgress(efEstimateProgress(lines));
    return { terminal: false };
  } catch (e) {
    efLog(`Polling fout: ${e.message}`, 'error');
    return { terminal: false };
  }
}

function efStartPolling(runId) {
  efStopPolling();
  efIsPolling = true;

  efLogPollTimer = setInterval(() => {
    if (!efIsPolling || efActiveRunId !== runId) return;
    efPollLogs(runId);
  }, 2000);

  efRunPollTimer = setInterval(async () => {
    if (!efIsPolling || efActiveRunId !== runId) return;
    const state = await efPollStatus(runId);
    if (state?.terminal) efStopPolling();
  }, 3000);
}

// ── Start scan ────────────────────────────────────────────────────────────────

async function startEntraFalconScan() {
  if (efActiveRunId) {
    alert('Er loopt al een Entra Security scan. Wacht tot deze klaar is.');
    return;
  }

  const tenantId = document.getElementById('tenantSelect')?.value;
  if (!tenantId) {
    alert('Selecteer eerst een tenant.');
    return;
  }

  // Sla opties op (best-effort)
  try {
    const token = localStorage.getItem('denjoy_token');
    const headers = {
      'Content-Type': 'application/json',
      ...(token ? { 'Authorization': `Bearer ${token}` } : {}),
    };
    await fetch('/api/config', {
      method: 'POST',
      credentials: 'include',
      headers,
      body: JSON.stringify({
        entrafalcon_include_ms_apps: document.getElementById('efIncludeMsApps')?.checked || false,
        entrafalcon_csv:             document.getElementById('efCsvExport')?.checked || false,
      }),
    });
  } catch (_) {}

  efShowProgress();
  document.getElementById('efProgressLog').innerHTML = '';
  efSetProgress(2);
  efLog(`Entra Security scan starten voor tenant ${tenantId}`);

  try {
    const token = localStorage.getItem('denjoy_token');
    const headers = {
      'Content-Type': 'application/json',
      ...(token ? { 'Authorization': `Bearer ${token}` } : {}),
    };
    const res = await fetch('/api/runs', {
      method: 'POST',
      credentials: 'include',
      headers,
      body: JSON.stringify({ tenant_id: tenantId, scan_type: 'entrafalcon', run_mode: 'script' }),
    });
    const run = await res.json();
    if (!res.ok) throw new Error(run.error || `HTTP ${res.status}`);

    efActiveRunId = run.id;
    efIsPolling = true;
    efLog(`Run aangemaakt: ${efActiveRunId}`, 'success');

    await efPollLogs(efActiveRunId);
    const firstState = await efPollStatus(efActiveRunId);
    if (!firstState?.terminal) efStartPolling(efActiveRunId);
  } catch (err) {
    efLog(`Starten mislukt: ${err.message}`, 'error');
    alert(`Fout bij starten scan: ${err.message}`);
    efShowConfig();
    efActiveRunId = null;
    efIsPolling = false;
    efStopPolling();
  }
}

// ── Stop scan ─────────────────────────────────────────────────────────────────

async function efStopScan() {
  if (!efActiveRunId) return;
  const btn = document.getElementById('efStopButton');
  if (btn) { btn.disabled = true; btn.textContent = 'Stoppen...'; }
  try {
    const token = localStorage.getItem('denjoy_token');
    const authHeader = token ? { 'Authorization': `Bearer ${token}` } : {};
    await fetch(`/api/runs/${efActiveRunId}/stop`, {
      method: 'POST',
      credentials: 'include',
      headers: authHeader,
    });
    efLog('⏹ Stop-verzoek verstuurd...', 'warning');
  } catch (e) {
    efLog(`Stop mislukt: ${e.message}`, 'error');
    if (btn) { btn.disabled = false; btn.textContent = '⏹ Stoppen'; }
  }
}
window.efStopScan = efStopScan;

// ── Section loader (sidebar nav → "Entra Security") ──────────────────────────

async function loadEntraFalconSection() {
  const infoEl  = document.getElementById('entrafalconLatestInfo');
  const viewBtn = document.getElementById('efViewLatestButton');
  if (!infoEl) return;

  const tenantId = typeof currentTenantId !== 'undefined' ? currentTenantId : null;
  if (!tenantId) { infoEl.textContent = 'Geen tenant geselecteerd.'; return; }

  try {
    const token = localStorage.getItem('denjoy_token');
    const headers = token ? { 'Authorization': `Bearer ${token}` } : {};
    const res = await fetch(`/api/runs?tenant_id=${tenantId}&limit=50`, {
      credentials: 'include',
      headers,
    });
    const data = await res.json();
    const efRuns = (data.items || []).filter((r) => r.scan_type === 'entrafalcon');

    if (!efRuns.length) {
      infoEl.textContent = 'Nog geen Entra Security scans uitgevoerd voor deze tenant.';
      if (viewBtn) viewBtn.style.display = 'none';
      return;
    }

    const latest = efRuns[0];
    const date   = latest.completed_at || latest.started_at;
    const d      = date ? new Date(date).toLocaleString() : '-';
    infoEl.innerHTML = `Laatste scan: <strong>${d}</strong> — status: <strong>${latest.status}</strong>`;
    if (viewBtn) {
      viewBtn.style.display = '';
      viewBtn.onclick = () => {
        if (typeof showSection === 'function') showSection('results', { resultsPanel: 'entrafalcon' });
      };
    }
  } catch (e) {
    infoEl.textContent = `Laden mislukt: ${e.message}`;
  }
}

// ── Results pane (Rapporten → Entra Security tab) ─────────────────────────────

async function loadEntraFalconResultsPane() {
  const emptyEl   = document.getElementById('efResultsEmpty');
  const contentEl = document.getElementById('efResultsContent');

  function showEmpty(msg) {
    if (emptyEl)   { emptyEl.style.display = ''; emptyEl.textContent = msg; }
    if (contentEl) contentEl.style.display = 'none';
  }

  const tenantId = typeof currentTenantId !== 'undefined' ? currentTenantId : null;
  if (!tenantId) { showEmpty('Geen tenant geselecteerd.'); return; }

  try {
    const token = localStorage.getItem('denjoy_token');
    const headers = token ? { 'Authorization': `Bearer ${token}` } : {};
    const res = await fetch(`/api/runs?tenant_id=${tenantId}&limit=50`, {
      credentials: 'include',
      headers,
    });
    const data = await res.json();
    const efRuns = (data.items || []).filter((r) => r.scan_type === 'entrafalcon' && r.report_path);

    if (!efRuns.length) {
      showEmpty('Nog geen Entra Security rapporten beschikbaar. Start een scan via het menu.');
      return;
    }

    if (emptyEl)   emptyEl.style.display = 'none';
    if (contentEl) contentEl.style.display = '';

    const runSelect = document.getElementById('efRunSelect');
    if (runSelect) {
      runSelect.innerHTML = efRuns.map((r) => {
        const d     = r.completed_at || r.started_at;
        const label = d ? new Date(d).toLocaleString() : r.id.slice(0, 8);
        return `<option value="${r.id}">${label} (${r.status})</option>`;
      }).join('');
      runSelect.onchange = () => efLoadRunFiles(runSelect.value);
    }

    await efLoadRunFiles(efRuns[0].id);
  } catch (e) {
    showEmpty(`Laden mislukt: ${e.message}`);
  }
}
window.loadEntraFalconResultsPane = loadEntraFalconResultsPane;

async function efLoadRunFiles(runId) {
  const tabsBar  = document.getElementById('efReportTabsBar');
  const openLink = document.getElementById('efOpenReportLink');
  const iframe   = document.getElementById('efReportFrame');
  const loader   = document.getElementById('efIframeLoader');
  if (!tabsBar || !iframe) return;

  try {
    const token = localStorage.getItem('denjoy_token');
    const headers = token ? { 'Authorization': `Bearer ${token}` } : {};
    const res = await fetch(`/api/runs/${runId}/files`, { credentials: 'include', headers });
    const data = await res.json();
    const files = data.items || [];

    if (!files.length) {
      tabsBar.innerHTML = '<span style="font-size:0.82rem;color:var(--muted,#64748b);">Geen HTML-bestanden gevonden.</span>';
      return;
    }

    tabsBar.innerHTML = files.map((f, i) => {
      const label = f.name.replace(/_/g, ' ').replace(/\.html$/i, '');
      return `<button class="nb-btn ${i === 0 ? 'nb-btn-primary' : 'nb-btn-secondary'}"
        style="font-size:0.78rem;padding:4px 10px;"
        onclick="efSwitchTab(this, '/reports/${runId}/${f.path}')">${label}</button>`;
    }).join('');

    efLoadIframe(`/reports/${runId}/${files[0].path}`, loader, iframe, openLink);
  } catch (e) {
    if (tabsBar) tabsBar.innerHTML =
      `<span style="color:#f87171;font-size:0.82rem;">Bestanden laden mislukt: ${e.message}</span>`;
  }
}

function efSwitchTab(btn, reportPath) {
  const tabsBar = document.getElementById('efReportTabsBar');
  if (tabsBar) {
    tabsBar.querySelectorAll('button').forEach((b) => {
      b.className = b === btn ? 'nb-btn nb-btn-primary' : 'nb-btn nb-btn-secondary';
      b.style.fontSize = '0.78rem';
      b.style.padding  = '4px 10px';
    });
  }
  efLoadIframe(
    reportPath,
    document.getElementById('efIframeLoader'),
    document.getElementById('efReportFrame'),
    document.getElementById('efOpenReportLink'),
  );
}
window.efSwitchTab = efSwitchTab;

function efLoadIframe(path, loader, iframe, openLink) {
  if (loader)   { loader.style.display = 'block'; loader.textContent = 'Rapport laden...'; }
  if (openLink) { openLink.href = path; openLink.style.display = ''; }
  if (!iframe) return;
  iframe.onload  = () => { if (loader) loader.style.display = 'none'; };
  iframe.onerror = () => {
    if (loader) { loader.style.display = 'block'; loader.textContent = 'Rapport kon niet worden geladen.'; }
  };
  iframe.src = path;
}

// ── Bootstrap ─────────────────────────────────────────────────────────────────

function bindEntraFalconButton() {
  const btn = document.getElementById('startEntraFalconButton');
  if (!btn) return;
  btn.removeEventListener('click', startEntraFalconScan);
  btn.addEventListener('click', startEntraFalconScan);
}

document.addEventListener('DOMContentLoaded', bindEntraFalconButton);
