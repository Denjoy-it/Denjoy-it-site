// Local assessment execution (backend-driven)

let activeRunId = null;
let activeRunPollTimer = null;
let activeLogPollTimer = null;
let isRunPollingActive = false;

function getSelectedPhases() {
  const checkboxes = document.querySelectorAll('input[name="phase"]:checked');
  return Array.from(checkboxes).map((cb) => cb.value);
}

function phaseLabel(phase) {
  const labels = {
    phase1: 'Phase 1: Users & Licensing',
    phase2: 'Phase 2: Collaboration & Storage',
    phase3: 'Phase 3: Compliance & Security',
    phase4: 'Phase 4: Advanced Security',
    phase5: 'Phase 5: Intune',
    phase6: 'Phase 6: Azure',
  };
  return labels[phase] || phase;
}

function addProgressLog(message, type = 'info') {
  const log = document.getElementById('progressLog');
  if (!log) return;
  const p = document.createElement('p');
  p.textContent = `[${new Date().toLocaleTimeString()}] ${message}`;
  if (type === 'error') p.style.color = '#ffb4b4';
  if (type === 'success') p.style.color = '#b7f7c8';
  if (type === 'warning') p.style.color = '#ffe08a';
  log.appendChild(p);
  log.scrollTop = log.scrollHeight;
}

function setProgress(percent, label) {
  const fill = document.getElementById('progressFill');
  if (!fill) return;
  fill.style.width = `${Math.max(0, Math.min(100, percent))}%`;
  fill.textContent = label || `${Math.round(percent)}%`;
}

function estimateProgressFromLogs(logText, selectedPhases) {
  if (!logText) return 5;
  const lines = logText.split(/\r?\n/).filter(Boolean);
  let done = 0;
  for (const phase of selectedPhases) {
    const phaseNum = phase.replace('phase', '');
    const phaseRegex = new RegExp(`(Completed\\s+${phase}|Voltooid\\s+${phase}|SkipPhase${phaseNum})`, 'i');
    if (lines.some((l) => phaseRegex.test(l))) done += 1;
  }
  if (!selectedPhases.length) return 10;
  return Math.min(95, Math.round((done / selectedPhases.length) * 90) + 5);
}

async function pollRunStatus(runId, selectedPhases) {
  try {
    const res = await fetch(`/api/runs/${runId}`);
    const run = await res.json();
    if (!res.ok) throw new Error(run.error || `HTTP ${res.status}`);
    const title = document.querySelector('#assessmentProgress h3');
    if (title) title.textContent = `Assessment status: ${run.status}`;

    if (run.status === 'completed') {
      setProgress(100, '100%');
      addProgressLog('Assessment voltooid.', 'success');
      if (run.snapshot_path) {
        addProgressLog('Assessmentdata is gesynchroniseerd naar de portal-weergave.', 'success');
      } else if (run.report_path) {
        addProgressLog('Legacy rapportbestand beschikbaar voor export/doelarchief.', 'success');
      }
      stopRunPolling();
      activeRunId = null;
      isRunPollingActive = false;
      document.querySelector('.assessment-config').style.display = 'block';
      document.getElementById('assessmentProgress').style.display = 'none';
      // Vervolgstappen-banner tonen
      const existing = document.getElementById('assessmentDoneNotice');
      if (existing) existing.remove();
      const configEl = document.querySelector('.assessment-config');
      if (configEl) {
        const banner = document.createElement('div');
        banner.id = 'assessmentDoneNotice';
        banner.style.cssText = 'margin:1rem 0;padding:1rem 1.25rem;background:rgba(34,197,94,0.1);border:1px solid rgba(34,197,94,0.3);border-radius:10px;display:flex;align-items:center;gap:1rem;flex-wrap:wrap;';
        banner.innerHTML = `
          <div style="flex:1;min-width:180px;">
            <strong style="color:#16a34a;display:block;margin-bottom:0.2rem;">Assessment voltooid</strong>
            <span style="font-size:0.82rem;color:var(--text-muted)">Wat wil je nu doen?</span>
          </div>
          <div style="display:flex;gap:0.5rem;flex-wrap:wrap;">
            <button type="button" onclick="document.querySelector('[data-section=results][data-results-panel=viewer]')?.click()" style="padding:0.45rem 0.9rem;background:var(--orange);color:#fff;border:none;border-radius:7px;font-size:0.82rem;font-weight:600;cursor:pointer;">Rapport bekijken →</button>
            <button type="button" onclick="document.querySelector('[data-section=results][data-results-panel=actions]')?.click()" style="padding:0.45rem 0.9rem;background:transparent;color:var(--text);border:1.5px solid var(--border);border-radius:7px;font-size:0.82rem;font-weight:600;cursor:pointer;">Acties bekijken</button>
            <button type="button" onclick="document.getElementById('assessmentDoneNotice')?.remove()" style="padding:0.45rem 0.7rem;background:transparent;color:var(--text-muted);border:none;font-size:0.82rem;cursor:pointer;">✕</button>
          </div>`;
        configEl.insertAdjacentElement('beforebegin', banner);
      }
      if (typeof refreshTenantData === 'function') await refreshTenantData();
      if (document.getElementById('autoOpenReport')?.checked) {
        setTimeout(() => {
          if (typeof showSection === 'function') showSection('assessment');
        }, 500);
      }
      return { terminal: true, status: run.status };
    }

    if (run.status === 'failed') {
      addProgressLog(`Assessment mislukt: ${run.error_message || 'Onbekende fout'}`, 'error');
      stopRunPolling();
      activeRunId = null;
      isRunPollingActive = false;
      document.querySelector('.assessment-config').style.display = 'block';
      document.getElementById('assessmentProgress').style.display = 'none';
      if (typeof refreshTenantData === 'function') await refreshTenantData();
      return { terminal: true, status: run.status };
    }

    if (run.status === 'cancelled') {
      addProgressLog('Assessment gestopt.', 'warning');
      stopRunPolling();
      activeRunId = null;
      isRunPollingActive = false;
      const stopBtn = document.getElementById('stopAssessmentButton');
      if (stopBtn) { stopBtn.disabled = false; stopBtn.textContent = '⏹ Stoppen'; }
      document.querySelector('.assessment-config').style.display = 'block';
      document.getElementById('assessmentProgress').style.display = 'none';
      if (typeof refreshTenantData === 'function') await refreshTenantData();
      return { terminal: true, status: run.status };
    }

    const logs = await fetch(`/api/runs/${runId}/logs`).then((r) => r.json());
    setProgress(estimateProgressFromLogs(logs.text || '', selectedPhases));
    return { terminal: false, status: run.status };
  } catch (e) {
    console.error(e);
    addProgressLog(`Polling fout: ${e.message}`, 'error');
    return { terminal: false, status: 'error' };
  }
}

async function pollRunLogs(runId) {
  try {
    const data = await fetch(`/api/runs/${runId}/logs`).then((r) => r.json());
    const log = document.getElementById('progressLog');
    if (!log) return;
    const text = (data.lines || []).join('\n');
    log.innerHTML = '';
    for (const line of (data.lines || [])) {
      const p = document.createElement('p');
      p.textContent = line;
      if (/failed|mislukt|error/i.test(line)) p.style.color = '#ffb4b4';
      else if (/completed|voltooid|success/i.test(line)) p.style.color = '#b7f7c8';
      log.appendChild(p);
    }
    log.scrollTop = log.scrollHeight;
    return text;
  } catch (e) {
    console.warn('Log polling fout', e);
  }
}

function stopRunPolling() {
  if (activeRunPollTimer) clearInterval(activeRunPollTimer);
  if (activeLogPollTimer) clearInterval(activeLogPollTimer);
  activeRunPollTimer = null;
  activeLogPollTimer = null;
  isRunPollingActive = false;
}

function startRunPollingLoop(runId, selectedPhases) {
  stopRunPolling();
  isRunPollingActive = true;

  // Logs poll every 2s
  activeLogPollTimer = setInterval(() => {
    if (!isRunPollingActive || !activeRunId || activeRunId !== runId) return;
    pollRunLogs(runId);
  }, 2000);

  // Status poll every 3s (single source of truth)
  activeRunPollTimer = setInterval(async () => {
    if (!isRunPollingActive || !activeRunId || activeRunId !== runId) return;
    const state = await pollRunStatus(runId, selectedPhases);
    if (state && state.terminal) {
      stopRunPolling();
    }
  }, 3000);
}

async function startAssessment() {
  if (activeRunId) {
    alert('Er draait al een assessment. Wacht tot deze is afgerond.');
    return;
  }

  const selectedPhases = getSelectedPhases();
  if (!selectedPhases.length) {
    alert('Selecteer minimaal een fase.');
    return;
  }

  const tenantSelect = document.getElementById('tenantSelect');
  const tenantId = tenantSelect?.value;
  if (!tenantId) {
    alert('Selecteer eerst een tenant.');
    return;
  }

  const runMode = document.getElementById('runModeSelect')?.value || 'demo';
  const authTenantId = (
    sessionStorage.getItem('denjoy_msal_tenant_id')
    || localStorage.getItem('m365_tenantId')
    || ''
  ).trim();

  document.querySelector('.assessment-config').style.display = 'none';
  document.getElementById('assessmentProgress').style.display = 'block';
  document.getElementById('progressLog').innerHTML = '';
  setProgress(2, 'Start...');
  addProgressLog(`Run wordt gestart voor tenant ${tenantId}`);
  addProgressLog(`Mode: ${runMode}`);
  addProgressLog(`Fasen: ${selectedPhases.map(phaseLabel).join(', ')}`);

  try {
    const token = localStorage.getItem('denjoy_token') || sessionStorage.getItem('denjoy_token');
    const authHeader = token ? { Authorization: `Bearer ${token}` } : {};
    const res = await fetch('/api/runs', {
      method: 'POST',
      credentials: 'include',
      headers: { 'Content-Type': 'application/json', ...authHeader },
      body: JSON.stringify({
        tenant_id: tenantId,
        auth_tenant_id: authTenantId || null,
        phases: selectedPhases,
        run_mode: runMode,
        scan_type: 'full',
      }),
    });
    const run = await res.json();
    if (!res.ok) throw new Error(run.error || `HTTP ${res.status}`);

    activeRunId = run.id;
    isRunPollingActive = true;
    addProgressLog(`Run aangemaakt: ${activeRunId}`, 'success');

    await pollRunLogs(activeRunId);
    const firstState = await pollRunStatus(activeRunId, selectedPhases);
    if (!firstState || !firstState.terminal) {
      startRunPollingLoop(activeRunId, selectedPhases);
    }
  } catch (error) {
    console.error(error);
    addProgressLog(`Starten assessment mislukt: ${error.message}`, 'error');
    alert(`Fout bij starten assessment: ${error.message}`);
    document.querySelector('.assessment-config').style.display = 'block';
    document.getElementById('assessmentProgress').style.display = 'none';
    activeRunId = null;
    isRunPollingActive = false;
    stopRunPolling();
  }
}

async function stopAssessment() {
  if (!activeRunId) return;
  const btn = document.getElementById('stopAssessmentButton');
  if (btn) { btn.disabled = true; btn.textContent = 'Stoppen...'; }
  try {
    const token = localStorage.getItem('denjoy_token');
    const authHeader = token ? { 'Authorization': `Bearer ${token}` } : {};
    await fetch(`/api/runs/${activeRunId}/stop`, {
      method: 'POST',
      credentials: 'include',
      headers: authHeader,
    });
    addProgressLog('⏹ Stop-verzoek verstuurd...', 'warning');
  } catch (e) {
    addProgressLog(`Stop mislukt: ${e.message}`, 'error');
    if (btn) { btn.disabled = false; btn.textContent = '⏹ Stoppen'; }
  }
}

function bindAssessmentButton() {
  const btn = document.getElementById('startAssessmentButton');
  if (!btn) return;
  btn.removeEventListener('click', startAssessment);
  btn.addEventListener('click', startAssessment);
}

document.addEventListener('DOMContentLoaded', () => {
  bindAssessmentButton();
  loadScheduledRuns();
});

async function scheduleAssessmentRun() {
  const dtInput = document.getElementById('scheduleDateTime');
  const noteInput = document.getElementById('scheduleNote');
  const btn = document.getElementById('scheduleRunBtn');
  const tid = typeof currentTenantId !== 'undefined' ? currentTenantId : (window.currentTenantId || '');
  if (!tid) { _scheduleMsg('Selecteer eerst een tenant.', 'error'); return; }
  const dt = dtInput?.value;
  if (!dt) { _scheduleMsg('Kies een datum en tijd.', 'error'); return; }
  const scheduledAt = new Date(dt).toISOString();
  const phases = Array.from(document.querySelectorAll('input[name="phase"]:checked')).map(el => el.value);
  if (btn) { btn.disabled = true; btn.textContent = 'Plannen...'; }
  try {
    const authHeader = {};
    const token = localStorage.getItem('denjoy_token') || sessionStorage.getItem('denjoy_token');
    if (token) authHeader['Authorization'] = `Bearer ${token}`;
    const res = await fetch('/api/scheduled-runs', {
      method: 'POST', credentials: 'include',
      headers: { 'Content-Type': 'application/json', ...authHeader },
      body: JSON.stringify({ tenant_id: tid, scheduled_at: scheduledAt, phases, note: noteInput?.value || '' }),
    });
    const data = await res.json();
    if (!res.ok || !data.ok) throw new Error(data.error || `HTTP ${res.status}`);
    _scheduleMsg('Assessment ingepland.', 'success');
    if (dtInput) dtInput.value = '';
    if (noteInput) noteInput.value = '';
    loadScheduledRuns();
  } catch (e) {
    _scheduleMsg('Fout: ' + e.message, 'error');
  } finally {
    if (btn) { btn.disabled = false; btn.textContent = 'Inplannen →'; }
  }
}

function _scheduleMsg(msg, type) {
  const el = document.getElementById('scheduledRunsList');
  if (!el) return;
  const div = document.createElement('div');
  div.style.cssText = `padding:0.5rem 0.75rem;margin-bottom:0.5rem;border-radius:7px;font-size:0.82rem;
    background:${type === 'error' ? 'rgba(239,68,68,0.1)' : 'rgba(34,197,94,0.1)'};
    border:1px solid ${type === 'error' ? 'rgba(239,68,68,0.3)' : 'rgba(34,197,94,0.3)'};
    color:${type === 'error' ? '#dc2626' : '#16a34a'};`;
  div.textContent = msg;
  el.prepend(div);
  setTimeout(() => div.remove(), 5000);
}

async function loadScheduledRuns() {
  const el = document.getElementById('scheduledRunsList');
  if (!el) return;
  try {
    const token = localStorage.getItem('denjoy_token') || sessionStorage.getItem('denjoy_token');
    const res = await fetch('/api/scheduled-runs', { credentials: 'include', headers: token ? { 'Authorization': `Bearer ${token}` } : {} });
    const data = await res.json();
    const items = (data.items || []).filter(j => j.status === 'pending');
    if (!items.length) { el.innerHTML = '<p style="color:var(--text-muted);font-size:0.8rem">Geen geplande assessments.</p>'; return; }
    el.innerHTML = `<table style="width:100%;border-collapse:collapse;font-size:0.82rem;">
      <thead><tr style="border-bottom:1px solid var(--border)"><th style="text-align:left;padding:0.4rem 0.5rem">Tenant</th><th style="padding:0.4rem 0.5rem">Gepland op</th><th style="padding:0.4rem 0.5rem">Status</th></tr></thead>
      <tbody>${items.map(j => `<tr style="border-bottom:1px solid var(--border)">
        <td style="padding:0.4rem 0.5rem">${j.tenant_id || '—'}</td>
        <td style="padding:0.4rem 0.5rem;font-family:var(--mono,monospace)">${j.scheduled_at ? new Date(j.scheduled_at).toLocaleString('nl-NL') : '—'}</td>
        <td style="padding:0.4rem 0.5rem"><span style="padding:0.15rem 0.5rem;border-radius:999px;background:rgba(234,179,8,0.12);color:#a16207;font-size:0.75rem;font-weight:600">${j.status}</span></td>
      </tr>`).join('')}</tbody></table>`;
  } catch (_) { el.innerHTML = '<p style="color:var(--text-muted);font-size:0.8rem">Kon geplande jobs niet laden.</p>'; }
}
