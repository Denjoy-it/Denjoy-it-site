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
      if (run.report_path) {
        addProgressLog(`Rapport beschikbaar: ${run.report_path}`, 'success');
      }
      stopRunPolling();
      activeRunId = null;
      isRunPollingActive = false;
      document.querySelector('.assessment-config').style.display = 'block';
      document.getElementById('assessmentProgress').style.display = 'none';
      if (typeof refreshTenantData === 'function') await refreshTenantData();
      if (document.getElementById('autoOpenReport')?.checked) {
        setTimeout(() => {
          if (typeof showSection === 'function') showSection('results');
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

  document.querySelector('.assessment-config').style.display = 'none';
  document.getElementById('assessmentProgress').style.display = 'block';
  document.getElementById('progressLog').innerHTML = '';
  setProgress(2, 'Start...');
  addProgressLog(`Run wordt gestart voor tenant ${tenantId}`);
  addProgressLog(`Mode: ${runMode}`);
  addProgressLog(`Fasen: ${selectedPhases.map(phaseLabel).join(', ')}`);

  try {
    const res = await fetch('/api/runs', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        tenant_id: tenantId,
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

function bindAssessmentButton() {
  const btn = document.getElementById('startAssessmentButton');
  if (!btn) return;
  btn.removeEventListener('click', startAssessment);
  btn.addEventListener('click', startAssessment);
}

document.addEventListener('DOMContentLoaded', bindAssessmentButton);
