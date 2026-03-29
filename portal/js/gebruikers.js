/**
 * Denjoy IT Platform — Gebruikersbeheer module (Fase 2)
 * User management: overzicht, provisioning wizard, offboarding wizard,
 * gebruiker detail panel.
 */
(function () {
  'use strict';

  // ── State ────────────────────────────────────────────────────────────────

  let _users        = [];
  let _licenses     = [];
  let _filterStatus = 'all';   // 'all' | 'enabled' | 'disabled'
  let _searchQ      = '';
  let _loadingUsers = false;
  let _usersSource  = 'live';
  let _userCountsOverride = null;
  let _currentCapability = null;

  // ── API helper ───────────────────────────────────────────────────────────

  function gbApiFetch(url, opts = {}) {
    const token = localStorage.getItem('denjoy_token') || localStorage.getItem('denjoy_auth_token') || '';
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

  // ── Weergave helpers ─────────────────────────────────────────────────────

  function initials(name) {
    if (!name) return '?';
    const parts = name.trim().split(/\s+/);
    if (parts.length === 1) return parts[0][0].toUpperCase();
    return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
  }

  function fmtDate(iso) {
    if (!iso) return '—';
    try { return new Date(iso).toLocaleDateString('nl-NL'); } catch { return iso; }
  }

  function escHtml(s) {
    if (!s) return '';
    return String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function getCapabilitySubsection(tab) {
    const map = {
      gebruikers: 'users',
      licenties: 'licenses',
      geschiedenis: 'history',
    };
    return map[tab] || 'users';
  }

  async function renderCapabilityBanner(tab, forceRefresh = false) {
    const tid = getTenantId();
    const wrap = document.getElementById('gbCapabilityBanner');
    if (!wrap || !tid || typeof window.denjoyFetchCapabilityStatus !== 'function' || typeof window.denjoyDescribeCapabilityStatus !== 'function') {
      if (wrap) wrap.innerHTML = '';
      return null;
    }
    try {
      const capability = await window.denjoyFetchCapabilityStatus(tid, 'gebruikers', getCapabilitySubsection(tab), { forceRefresh });
      _currentCapability = capability;
      const info = window.denjoyDescribeCapabilityStatus(capability);
      const roles = (capability.extra_roles || []).slice(0, 3).join(', ');
      wrap.innerHTML = `
        <div class="live-module-source">
          <span class="live-module-source-pill ${escHtml(info.className || '')}">${escHtml(info.label)}</span>
          <span>${escHtml(info.detail)}</span>
        </div>
        <div class="gb-capability-meta">${escHtml(roles || 'Geen extra rollen gespecificeerd')}</div>`;
      const liveButtons = ['gbBtnLiveScan', 'gbBtnRefresh', 'gbBtnRefreshLic']
        .map((id) => document.getElementById(id))
        .filter(Boolean);
      const blocked = capability.status === 'config_required' || capability.status === 'not_implemented' || !capability.supports_live;
      liveButtons.forEach((btn) => { btn.disabled = blocked; });
      return capability;
    } catch (_) {
      wrap.innerHTML = '';
      return null;
    }
  }

  function renderOverviewStats() {
    const computedTotal = _users.length;
    const computedActive = _users.filter((u) => u.accountEnabled).length;
    const computedDisabled = computedTotal - computedActive;
    const computedGuests = _users.filter((u) => {
      const userType = String(u.userType || '').toLowerCase();
      const upn = String(u.userPrincipalName || '').toLowerCase();
      return userType === 'guest' || upn.includes('#ext#');
    }).length;
    const total = Number.isFinite(_userCountsOverride?.total) ? _userCountsOverride.total : computedTotal;
    const active = Number.isFinite(_userCountsOverride?.active) ? _userCountsOverride.active : computedActive;
    const disabled = Number.isFinite(_userCountsOverride?.disabled) ? _userCountsOverride.disabled : computedDisabled;
    const guests = Number.isFinite(_userCountsOverride?.guest) ? _userCountsOverride.guest : computedGuests;

    const updateValue = (id, value) => {
      const el = document.getElementById(id);
      if (el) el.textContent = String(value);
    };

    updateValue('gbStatTotal', total);
    updateValue('gbStatActive', active);
    updateValue('gbStatDisabled', disabled);
    updateValue('gbStatGuests', guests);
  }

  function applySourceState() {
    const disabledFilter = document.querySelector('.gb-filter-tab[data-filter="disabled"]');
    if (!disabledFilter) return;

    const isSnapshot = _usersSource === 'assessment_snapshot';
    disabledFilter.disabled = isSnapshot;
    disabledFilter.title = isSnapshot
      ? 'Uitgeschakelde gebruikers zijn niet betrouwbaar beschikbaar in assessment snapshots.'
      : '';

    if (isSnapshot && _filterStatus === 'disabled') {
      _filterStatus = 'all';
      document.querySelectorAll('.gb-filter-tab').forEach((btn) => {
        btn.classList.toggle('active', btn.dataset.filter === 'all');
      });
    }
  }

  // ── Hoofd render: gebruikerstabel ────────────────────────────────────────

  function renderUsersTable() {
    const tbody = document.getElementById('gbUserTableBody');
    const info  = document.getElementById('gbUserCount');
    if (!tbody) return;

    let filtered = _users;
    if (_filterStatus === 'enabled')  filtered = filtered.filter((u) => u.accountEnabled);
    if (_filterStatus === 'disabled') filtered = filtered.filter((u) => !u.accountEnabled);
    if (_searchQ) {
      const q = _searchQ.toLowerCase();
      filtered = filtered.filter(
        (u) =>
          (u.displayName || '').toLowerCase().includes(q) ||
          (u.userPrincipalName || '').toLowerCase().includes(q) ||
          (u.department || '').toLowerCase().includes(q) ||
          (u.jobTitle || '').toLowerCase().includes(q)
      );
    }

    if (info) {
      const total = _users.length;
      const shown = filtered.length;
      info.textContent = shown === total ? `${total} gebruiker${total !== 1 ? 's' : ''}` : `${shown} van ${total}`;
    }

    if (filtered.length === 0) {
      const snapshotDisabledMessage = _usersSource === 'assessment_snapshot' && _filterStatus === 'disabled'
        ? 'Uitgeschakelde gebruikers zijn niet beschikbaar in assessment snapshotdata. Gebruik live data via Verversen.'
        : null;
      tbody.innerHTML = `<tr><td colspan="6" class="gb-table-empty">${
        snapshotDisabledMessage || _searchQ || _filterStatus !== 'all'
          ? (snapshotDisabledMessage || 'Geen gebruikers gevonden voor deze filter.')
          : 'Geen gebruikers geladen. Klik op Verversen.'
      }</td></tr>`;
      return;
    }

    tbody.innerHTML = filtered.map((u) => {
      const enabled = u.accountEnabled;
      const statusHtml = enabled
        ? `<span class="gb-status gb-status-enabled"><span class="gb-dot"></span>Actief</span>`
        : `<span class="gb-status gb-status-disabled"><span class="gb-dot"></span>Uitgeschakeld</span>`;
      const licHtml = u.licenseCount
        ? `<span class="gb-lic-count has-lic">${u.licenseCount}</span>`
        : `<span class="gb-lic-count">0</span>`;
      const lastSign = u.lastSignIn ? fmtDate(u.lastSignIn) : '—';

      return `<tr data-uid="${escHtml(u.id)}" data-upn="${escHtml(u.userPrincipalName)}">
        <td>
          <div class="gb-user-cell">
            <div class="gb-avatar">${initials(u.displayName)}</div>
            <div>
              <div class="gb-user-name">${escHtml(u.displayName)}</div>
              <div class="gb-user-upn">${escHtml(u.userPrincipalName)}</div>
            </div>
          </div>
        </td>
        <td>${escHtml(u.department || '—')}</td>
        <td>${escHtml(u.jobTitle || '—')}</td>
        <td>${statusHtml}</td>
        <td>${licHtml}</td>
        <td>
          <div class="gb-row-actions">
            <button class="gb-btn gb-btn-ghost gb-btn-detail" data-uid="${escHtml(u.id)}" data-name="${escHtml(u.displayName)}">Detail</button>
            ${enabled
              ? `<button class="gb-btn gb-btn-ghost gb-btn-danger gb-btn-offboard" data-uid="${escHtml(u.id)}" data-upn="${escHtml(u.userPrincipalName)}" data-name="${escHtml(u.displayName)}">Offboard</button>`
              : ''}
          </div>
        </td>
      </tr>`;
    }).join('');

    // Rij klikken → detail
    tbody.querySelectorAll('tr[data-uid]').forEach((row) => {
      row.addEventListener('click', (e) => {
        if (e.target.closest('.gb-row-actions')) return;
        openDetailPanel(row.dataset.uid, row.dataset.upn);
      });
    });
    tbody.querySelectorAll('.gb-btn-detail').forEach((btn) => {
      btn.addEventListener('click', (e) => { e.stopPropagation(); openDetailPanel(btn.dataset.uid, btn.dataset.name); });
    });
    tbody.querySelectorAll('.gb-btn-offboard').forEach((btn) => {
      btn.addEventListener('click', (e) => { e.stopPropagation(); openOffboardWizard(btn.dataset.uid, btn.dataset.upn, btn.dataset.name); });
    });
  }

  // ── Gebruikers laden ─────────────────────────────────────────────────────

  function loadUsers(options = {}) {
    const tid = getTenantId();
    if (!tid) { showStatus('Selecteer eerst een tenant.'); return; }
    if (_loadingUsers) return;
    _loadingUsers = true;
    const strictLive = !!options.strictLive;
    const previousUsers = Array.isArray(_users) ? [..._users] : [];
    const previousSource = _usersSource;
    const previousCounts = _userCountsOverride ? { ..._userCountsOverride } : null;

    const tbody = document.getElementById('gbUserTableBody');
    if (tbody) tbody.innerHTML = `<tr><td colspan="6" class="gb-loading">Gebruikers laden...</td></tr>`;
    const banner = document.getElementById('gbSnapshotBanner');
    const btnLiveScan = document.getElementById('gbBtnLiveScan');
    const btnRefresh = document.getElementById('gbBtnRefresh');
    const setLoadingButtons = (isLoading) => {
      if (btnLiveScan) {
        btnLiveScan.disabled = isLoading;
        btnLiveScan.textContent = isLoading ? 'Live scan...' : 'Live scan';
      }
      if (btnRefresh) btnRefresh.disabled = isLoading;
    };
    setLoadingButtons(true);
    if (strictLive && banner) {
      banner.style.display = '';
      banner.textContent = 'Gerichte live scan voor gebruikers wordt uitgevoerd...';
    }

    const url = strictLive ? `${API.m365.users(tid)}?strict_live=1` : API.m365.users(tid);
    gbApiFetch(url)
      .then((data) => {
        _users = data.users || [];
        _usersSource = data._source || 'live';
        _userCountsOverride = data.counts && typeof data.counts === 'object' ? data.counts : null;
        renderOverviewStats();
        applySourceState();
        if (data._source === 'assessment_snapshot') {
          const info = document.getElementById('gbUserCount');
          if (info) info.title = 'Data uit laatste assessment — klik Vernieuwen voor live data';
          if (banner) { banner.style.display = ''; banner.textContent = 'Gegevens uit laatste assessment. Live data vereist actieve verbinding.'; }
        } else {
          if (banner) {
            banner.style.display = '';
            banner.textContent = strictLive
              ? 'Gerichte live scan voor gebruikers succesvol afgerond.'
              : 'Live gebruikersdata succesvol opgehaald.';
          }
        }
        renderUsersTable();
        if (strictLive && typeof showToast === 'function') showToast('Live scan voor gebruikers afgerond.', 'success');
      })
      .catch((err) => {
        _users = strictLive ? previousUsers : [];
        _usersSource = strictLive ? previousSource : 'live';
        _userCountsOverride = strictLive ? previousCounts : null;
        renderOverviewStats();
        applySourceState();
        if (banner) {
          banner.style.display = '';
          banner.textContent = strictLive
            ? `Gerichte live scan voor gebruikers mislukt: ${escHtml(String(err))}`
            : `Fout: ${escHtml(String(err))}`;
        }
        if (strictLive) {
          renderUsersTable();
        } else if (tbody) {
          tbody.innerHTML = `<tr><td colspan="6" class="gb-table-empty">Fout: ${escHtml(String(err))}</td></tr>`;
        }
        if (strictLive && typeof showToast === 'function') showToast(String(err), 'error');
      })
      .finally(() => {
        _loadingUsers = false;
        setLoadingButtons(false);
      });
  }

  function loadLicenses(tid) {
    return gbApiFetch(API.m365.licenses(tid))
      .then((data) => { _licenses = data.licenses || []; return _licenses; })
      .catch(() => { _licenses = []; return []; });
  }

  function findCachedUser(userId, fallbackName) {
    return _users.find((u) =>
      u.id === userId ||
      u.userPrincipalName === userId ||
      u.userPrincipalName === fallbackName ||
      u.displayName === fallbackName
    ) || null;
  }

  // ── Detail panel ─────────────────────────────────────────────────────────

  function openDetailPanel(userId, fallbackName) {
    const tid = getTenantId();
    if (!tid) return;
    const cachedUser = findCachedUser(userId, fallbackName);

    // Open het Inzichten-paneel
    if (typeof window.openSideRailDetail === 'function') {
      window.openSideRailDetail('Gebruiker', fallbackName || userId);
    }

    gbApiFetch(API.m365.user(tid, userId))
      .then((data) => {
        const isSnapshot = data._source === 'assessment_snapshot';
        const u = Object.assign({}, cachedUser || {}, data.user || {});
        const licenses = Array.isArray(u.licenses) ? u.licenses : [];
        const licChips = licenses.map((l) => `<span class="gb-chip gb-chip-lic">${escHtml(l)}</span>`).join('');
        const mfaMethods = u.mfaMethods || [];
        const mfaChips = mfaMethods.map((m) => `<span class="gb-chip gb-chip-mfa">${escHtml(m)}</span>`).join('');
        const mfaEmpty = isSnapshot
          ? '<span style="color:var(--text-muted)">Geen MFA geregistreerd</span>'
          : '<span style="color:var(--color-danger,#e53)">Geen MFA (!)</span>';
        const grpList = u.groups || [];
        const grpChips = grpList.slice(0, 8).map((g) => `<span class="gb-chip gb-chip-grp">${escHtml(g)}</span>`).join('') +
          (grpList.length > 8 ? `<span class="gb-chip">+${grpList.length - 8}</span>` : '');
        const grpEmpty = isSnapshot
          ? '<span style="color:var(--text-muted)" title="Groepslidmaatschap is niet beschikbaar in de snapshot. Stel app-authenticatie in voor live data.">Niet beschikbaar in snapshot</span>'
          : '<span style="color:var(--text-muted)">Geen groepen</span>';
        const snapshotBadge = isSnapshot
          ? `<div style="margin-bottom:.75rem;padding:.35rem .6rem;background:var(--surface-raised,#1e2430);border-radius:6px;font-size:.75rem;color:var(--text-muted)">📋 Snapshot — groepslidmaatschap vereist live app-auth</div>`
          : '';
        const offboardBtn = u.accountEnabled
          ? `<button class="gb-btn gb-btn-danger" id="gbRailOffboardBtn" style="margin-top:1rem;width:100%">Offboard gebruiker</button>`
          : '';
        const bodyHtml = `
          ${snapshotBadge}
          <div class="gb-detail-grid">
            <div class="gb-detail-item"><div class="gb-detail-key">UPN</div><div class="gb-detail-val" style="word-break:break-all">${escHtml(u.userPrincipalName || '—')}</div></div>
            <div class="gb-detail-item"><div class="gb-detail-key">Status</div><div class="gb-detail-val">${u.accountEnabled ? '<span class="gb-status gb-status-enabled"><span class="gb-dot"></span>Actief</span>' : '<span class="gb-status gb-status-disabled"><span class="gb-dot"></span>Uitgeschakeld</span>'}</div></div>
            <div class="gb-detail-item"><div class="gb-detail-key">Functie</div><div class="gb-detail-val">${escHtml(u.jobTitle || '—')}</div></div>
            <div class="gb-detail-item"><div class="gb-detail-key">Afdeling</div><div class="gb-detail-val">${escHtml(u.department || '—')}</div></div>
            <div class="gb-detail-item"><div class="gb-detail-key">Locatie</div><div class="gb-detail-val">${escHtml(u.officeLocation || '—')}</div></div>
            <div class="gb-detail-item"><div class="gb-detail-key">Taal</div><div class="gb-detail-val">${escHtml(u.preferredLanguage || '—')}</div></div>
            <div class="gb-detail-item"><div class="gb-detail-key">Aangemaakt</div><div class="gb-detail-val">${fmtDate(u.createdDateTime)}</div></div>
            <div class="gb-detail-item"><div class="gb-detail-key">On-Prem sync</div><div class="gb-detail-val">${u.onPremisesSyncEnabled ? 'Ja' : 'Nee'}</div></div>
            <div class="gb-detail-item gb-detail-full" style="grid-column:1/-1"><div class="gb-detail-key">Licenties</div><div class="gb-chip-list">${licChips || '<span style="color:var(--text-muted)">Geen licenties</span>'}</div></div>
            <div class="gb-detail-item" style="grid-column:1/-1"><div class="gb-detail-key">MFA methoden</div><div class="gb-chip-list">${mfaChips || mfaEmpty}</div></div>
            <div class="gb-detail-item" style="grid-column:1/-1"><div class="gb-detail-key">Groepen (top 8)</div><div class="gb-chip-list">${grpChips || grpEmpty}</div></div>
          </div>
          ${offboardBtn}`;
        if (typeof window.updateSideRailDetail === 'function') {
          window.updateSideRailDetail(u.displayName || fallbackName || userId, bodyHtml);
        }
        // Offboard-knop activeren na renderen
        setTimeout(() => {
          document.getElementById('gbRailOffboardBtn')?.addEventListener('click', () => {
            openOffboardWizard(userId, u.userPrincipalName, u.displayName);
          });
        }, 50);
      })
      .catch((err) => {
        if (typeof window.updateSideRailDetail === 'function') {
          if (cachedUser) {
            const cachedLicenses = Array.isArray(cachedUser.licenses) ? cachedUser.licenses : [];
            window.updateSideRailDetail(cachedUser.displayName || fallbackName, `
              <div class="gb-detail-grid">
                <div class="gb-detail-item"><div class="gb-detail-key">UPN</div><div class="gb-detail-val">${escHtml(cachedUser.userPrincipalName || '—')}</div></div>
                <div class="gb-detail-item"><div class="gb-detail-key">Status</div><div class="gb-detail-val">${cachedUser.accountEnabled ? 'Actief' : 'Uitgeschakeld'}</div></div>
                <div class="gb-detail-item"><div class="gb-detail-key">Functie</div><div class="gb-detail-val">${escHtml(cachedUser.jobTitle || '—')}</div></div>
                <div class="gb-detail-item"><div class="gb-detail-key">Afdeling</div><div class="gb-detail-val">${escHtml(cachedUser.department || '—')}</div></div>
                <div class="gb-detail-item gb-detail-full" style="grid-column:1/-1"><div class="gb-detail-key">Licenties</div><div class="gb-chip-list">${cachedLicenses.map((l) => `<span class="gb-chip gb-chip-lic">${escHtml(l)}</span>`).join('') || '<span style="color:var(--text-muted)">Geen licenties</span>'}</div></div>
              </div>
              <div class="gb-empty" style="margin-top:.75rem">Live detail ophalen mislukt. Basisinformatie getoond.</div>`);
          } else {
            window.updateSideRailDetail('Fout', `<div class="gb-empty">Fout: ${escHtml(String(err))}</div>`);
          }
        }
      });
  }

  // ── Offboarding wizard ───────────────────────────────────────────────────

  function openOffboardWizard(userId, upn, displayName) {
    const tid = getTenantId();
    if (!tid) return;

    const overlay = createOverlay('gb-modal-overlay', closeAllModals);
    const modal = document.createElement('div');
    modal.className = 'gb-modal';

    function render(step, state) {
      modal.innerHTML = `
        <div class="gb-modal-header">
          <div>
            <div class="gb-modal-title">Gebruiker offboarden</div>
            <div class="gb-modal-subtitle">${escHtml(displayName)} — ${escHtml(upn)}</div>
          </div>
          <button class="gb-modal-close" title="Sluiten">✕</button>
        </div>
        <div class="gb-modal-body">
          <div class="gb-wizard-steps">
            <div class="gb-step ${step >= 1 ? 'done' : ''} ${step === 0 ? 'active' : ''}">
              <span class="gb-step-num">${step > 0 ? '✓' : '1'}</span><span>Opties</span>
            </div>
            <div class="gb-step ${step >= 2 ? 'done' : ''} ${step === 1 ? 'active' : ''}">
              <span class="gb-step-num">${step > 1 ? '✓' : '2'}</span><span>Bevestig</span>
            </div>
            <div class="gb-step ${step === 2 ? 'active' : ''}">
              <span class="gb-step-num">3</span><span>Resultaat</span>
            </div>
          </div>
          <div id="gbOffStepContent"></div>
        </div>
        <div class="gb-modal-footer" id="gbOffFooter"></div>`;

      modal.querySelector('.gb-modal-close').onclick = closeAllModals;

      const content = modal.querySelector('#gbOffStepContent');
      const footer  = modal.querySelector('#gbOffFooter');

      if (step === 0) {
        // Stap 1: opties kiezen
        content.innerHTML = `
          <div class="gb-checkbox-group">
            <label class="gb-checkbox-item">
              <input type="checkbox" id="gbOpt_revoke" checked>
              <div><div class="gb-checkbox-label">Sessies & tokens intrekken</div>
              <div class="gb-checkbox-desc">Beëindigt alle actieve sessies direct</div></div>
            </label>
            <label class="gb-checkbox-item">
              <input type="checkbox" id="gbOpt_disable" checked>
              <div><div class="gb-checkbox-label">Account uitschakelen</div>
              <div class="gb-checkbox-desc">Blokkeert inloggen voor deze gebruiker</div></div>
            </label>
            <label class="gb-checkbox-item">
              <input type="checkbox" id="gbOpt_lic" checked>
              <div><div class="gb-checkbox-label">Licenties verwijderen</div>
              <div class="gb-checkbox-desc">Verwijdert alle M365 licentietoewijzingen</div></div>
            </label>
            <label class="gb-checkbox-item">
              <input type="checkbox" id="gbOpt_ooo">
              <div><div class="gb-checkbox-label">Out-of-Office instellen</div>
              <div class="gb-checkbox-desc">Automatisch antwoord inschakelen (vereist MailboxSettings.ReadWrite)</div></div>
            </label>
          </div>
          <div id="gbOooMsgWrap" style="display:none;margin-top:.75rem">
            <label class="gb-form-label">OOO bericht</label>
            <textarea class="gb-form-input" id="gbOooMsg" rows="3" style="resize:vertical"
              placeholder="Deze medewerker is niet meer werkzaam. Neem contact op via info@uw-domein.nl">Deze medewerker is niet meer werkzaam bij ons bedrijf. Neem contact op via info@uw-domein.nl</textarea>
          </div>`;

        modal.querySelector('#gbOpt_ooo').addEventListener('change', (e) => {
          modal.querySelector('#gbOooMsgWrap').style.display = e.target.checked ? '' : 'none';
        });

        footer.innerHTML = `
          <button class="gb-btn gb-btn-secondary gb-cancel">Annuleren</button>
          <button class="gb-btn gb-btn-primary gb-next">Volgende →</button>`;
        footer.querySelector('.gb-cancel').onclick = closeAllModals;
        footer.querySelector('.gb-next').onclick = () => {
          const opts = {
            revoke_tokens:    modal.querySelector('#gbOpt_revoke')?.checked ?? true,
            disable_account:  modal.querySelector('#gbOpt_disable')?.checked ?? true,
            remove_licenses:  modal.querySelector('#gbOpt_lic')?.checked ?? true,
            set_out_of_office: modal.querySelector('#gbOpt_ooo')?.checked ?? false,
            ooo_message:      (modal.querySelector('#gbOooMsg')?.value || '').trim(),
          };
          render(1, opts);
        };

      } else if (step === 1) {
        // Stap 2: bevestiging
        const checks = [
          state.revoke_tokens    && '✓ Sessies & tokens intrekken',
          state.disable_account  && '✓ Account uitschakelen',
          state.remove_licenses  && '✓ Licenties verwijderen',
          state.set_out_of_office && '✓ Out-of-Office instellen',
        ].filter(Boolean);

        content.innerHTML = `
          <div class="gb-warn-box danger">
            <strong>Bevestig offboarding van ${escHtml(displayName)}</strong><br>
            De volgende acties worden <strong>direct</strong> uitgevoerd:
          </div>
          <ul style="margin:.75rem 0 0;padding-left:1.25rem;font-size:.875rem;color:var(--text-secondary)">
            ${checks.map((c) => `<li>${escHtml(c)}</li>`).join('')}
          </ul>
          <div style="margin-top:.75rem">
            <label class="gb-form-label" for="gbDryRunToggle">
              <input type="checkbox" id="gbDryRunToggle" style="accent-color:var(--accent)">
              Dry-run (preview zonder uitvoering)
            </label>
          </div>`;

        footer.innerHTML = `
          <button class="gb-btn gb-btn-secondary gb-back">← Terug</button>
          <button class="gb-btn gb-btn-danger gb-execute">Offboard uitvoeren</button>`;
        footer.querySelector('.gb-back').onclick = () => render(0, state);
        footer.querySelector('.gb-execute').onclick = () => {
          const dryRun = modal.querySelector('#gbDryRunToggle')?.checked ?? false;
          executeOffboard(userId, upn, displayName, state, dryRun, modal, overlay);
        };

      } else if (step === 2) {
        // Stap 3: resultaat — gevuld door executeOffboard
      }
    }

    overlay.appendChild(modal);
    document.body.appendChild(overlay);
    render(0, {});
  }

  function executeOffboard(userId, upn, displayName, opts, dryRun, modal, overlay) {
    const tid = getTenantId();
    const footer = modal.querySelector('#gbOffFooter');
    const content = modal.querySelector('#gbOffStepContent');

    if (footer) footer.innerHTML = `<span style="font-size:.82rem;color:var(--text-muted)">Bezig met offboarding...</span>`;

    const payload = Object.assign({}, opts, { display_name: displayName, dry_run: dryRun });

    gbApiFetch(API.m365.offboard(tid, userId), {
      method: 'POST',
      body: JSON.stringify(payload),
    })
      .then((data) => {
        const actions  = data.actions  || [];
        const warnings = data.warnings || [];
        const isOk     = data.ok !== false;
        const isDry    = dryRun;

        if (content) {
          const resultClass = isDry ? 'gb-result-dryrun' : (isOk ? 'gb-result-ok' : 'gb-result-error');
          const icon = isDry ? 'ℹ️' : (isOk ? '✅' : '❌');
          content.innerHTML = `
            <div class="gb-result ${resultClass}">
              <div class="gb-result-icon">${icon}</div>
              <div class="gb-result-msg">${isDry ? 'Dry-run voltooid — geen wijzigingen gemaakt' : (isOk ? 'Offboarding succesvol uitgevoerd' : 'Offboarding gedeeltelijk mislukt')}</div>
            </div>
            ${actions.length ? `<ul class="gb-result-list">${actions.map((a) => `<li>${escHtml(a)}</li>`).join('')}</ul>` : ''}
            ${warnings.length ? `<ul class="gb-result-list warnings">${warnings.map((w) => `<li>${escHtml(w)}</li>`).join('')}</ul>` : ''}`;
        }
        if (footer) {
          footer.innerHTML = `<button class="gb-btn gb-btn-primary gb-done">Sluiten</button>`;
          footer.querySelector('.gb-done').onclick = () => { closeAllModals(); loadUsers(); };
        }
      })
      .catch((err) => {
        if (content) content.innerHTML = `<div class="gb-result gb-result-error"><div class="gb-result-msg">Fout: ${escHtml(String(err))}</div></div>`;
        if (footer) {
          footer.innerHTML = `<button class="gb-btn gb-btn-secondary gb-done">Sluiten</button>`;
          footer.querySelector('.gb-done').onclick = closeAllModals;
        }
      });
  }

  // ── Provisioning wizard ──────────────────────────────────────────────────

  function openProvisioningWizard() {
    const tid = getTenantId();
    if (!tid) { showStatus('Selecteer eerst een tenant.'); return; }

    const overlay = createOverlay('gb-modal-overlay', closeAllModals);
    const modal   = document.createElement('div');
    modal.className = 'gb-modal gb-modal-wide';

    const state = { step: 0, formData: {}, selectedLicenses: [] };

    function renderStep() {
      modal.innerHTML = `
        <div class="gb-modal-header">
          <div>
            <div class="gb-modal-title">Nieuwe gebruiker aanmaken</div>
            <div class="gb-modal-subtitle">Provisioning wizard</div>
          </div>
          <button class="gb-modal-close">✕</button>
        </div>
        <div class="gb-modal-body">
          <div class="gb-wizard-steps">
            <div class="gb-step ${state.step > 0 ? 'done' : ''} ${state.step === 0 ? 'active' : ''}">
              <span class="gb-step-num">${state.step > 0 ? '✓' : '1'}</span><span>Gegevens</span>
            </div>
            <div class="gb-step ${state.step > 1 ? 'done' : ''} ${state.step === 1 ? 'active' : ''}">
              <span class="gb-step-num">${state.step > 1 ? '✓' : '2'}</span><span>Licenties</span>
            </div>
            <div class="gb-step ${state.step === 2 ? 'active' : ''}">
              <span class="gb-step-num">3</span><span>Resultaat</span>
            </div>
          </div>
          <div id="gbProvContent"></div>
        </div>
        <div class="gb-modal-footer" id="gbProvFooter"></div>`;

      modal.querySelector('.gb-modal-close').onclick = closeAllModals;

      const content = modal.querySelector('#gbProvContent');
      const footer  = modal.querySelector('#gbProvFooter');

      if (state.step === 0) {
        // Stap 1: gebruikersgegevens
        const d = state.formData;
        content.innerHTML = `
          <div class="gb-form">
            <div class="gb-form-row">
              <div class="gb-form-group">
                <label class="gb-form-label">Voornaam <span class="gb-required">*</span></label>
                <input class="gb-form-input" id="gbFldFirst" value="${escHtml(d.givenName || '')}" placeholder="Jan">
              </div>
              <div class="gb-form-group">
                <label class="gb-form-label">Achternaam <span class="gb-required">*</span></label>
                <input class="gb-form-input" id="gbFldLast" value="${escHtml(d.surname || '')}" placeholder="de Vries">
              </div>
            </div>
            <div class="gb-form-group">
              <label class="gb-form-label">Weergavenaam <span class="gb-required">*</span></label>
              <input class="gb-form-input" id="gbFldDisplay" value="${escHtml(d.displayName || '')}" placeholder="Jan de Vries">
            </div>
            <div class="gb-form-group">
              <label class="gb-form-label">UPN (e-mail) <span class="gb-required">*</span></label>
              <input class="gb-form-input" id="gbFldUpn" value="${escHtml(d.userPrincipalName || '')}" placeholder="jan.devries@bedrijf.onmicrosoft.com">
              <div class="gb-form-hint">Gebruik het .onmicrosoft.com domein of een geverifieerd domein</div>
            </div>
            <div class="gb-form-row">
              <div class="gb-form-group">
                <label class="gb-form-label">Functie</label>
                <input class="gb-form-input" id="gbFldJob" value="${escHtml(d.jobTitle || '')}" placeholder="Medewerker">
              </div>
              <div class="gb-form-group">
                <label class="gb-form-label">Afdeling</label>
                <input class="gb-form-input" id="gbFldDept" value="${escHtml(d.department || '')}" placeholder="ICT">
              </div>
            </div>
            <div class="gb-form-row">
              <div class="gb-form-group">
                <label class="gb-form-label">Gebruikslocatie</label>
                <select class="gb-form-select" id="gbFldLocale">
                  <option value="NL" ${d.usageLocation === 'NL' || !d.usageLocation ? 'selected' : ''}>Nederland (NL)</option>
                  <option value="BE" ${d.usageLocation === 'BE' ? 'selected' : ''}>België (BE)</option>
                  <option value="DE" ${d.usageLocation === 'DE' ? 'selected' : ''}>Duitsland (DE)</option>
                  <option value="GB" ${d.usageLocation === 'GB' ? 'selected' : ''}>Verenigd Koninkrijk (GB)</option>
                  <option value="US" ${d.usageLocation === 'US' ? 'selected' : ''}>Verenigde Staten (US)</option>
                </select>
              </div>
              <div class="gb-form-group">
                <label class="gb-form-label">Tijdelijk wachtwoord <span class="gb-required">*</span></label>
                <input class="gb-form-input" id="gbFldPwd" type="text" value="${escHtml(d.password || '')}"
                  placeholder="Minimaal 8 tekens, hoofdletter + cijfer">
                <div class="gb-form-hint">Gebruiker moet wachtwoord wijzigen bij eerste inlog</div>
              </div>
            </div>
          </div>`;

        // Auto-fill displayName
        ['gbFldFirst','gbFldLast'].forEach((id) => {
          modal.querySelector(`#${id}`)?.addEventListener('input', () => {
            const first = modal.querySelector('#gbFldFirst')?.value.trim() || '';
            const last  = modal.querySelector('#gbFldLast')?.value.trim()  || '';
            const disp  = modal.querySelector('#gbFldDisplay');
            if (disp && (!disp.value || disp.value === state.formData.displayName)) {
              disp.value = [first, last].filter(Boolean).join(' ');
            }
          });
        });

        footer.innerHTML = `
          <button class="gb-btn gb-btn-secondary gb-cancel">Annuleren</button>
          <button class="gb-btn gb-btn-primary gb-next">Licenties →</button>`;
        footer.querySelector('.gb-cancel').onclick = closeAllModals;
        footer.querySelector('.gb-next').onclick = () => {
          const fd = {
            givenName:         (modal.querySelector('#gbFldFirst')?.value || '').trim(),
            surname:           (modal.querySelector('#gbFldLast')?.value  || '').trim(),
            displayName:       (modal.querySelector('#gbFldDisplay')?.value || '').trim(),
            userPrincipalName: (modal.querySelector('#gbFldUpn')?.value || '').trim(),
            jobTitle:          (modal.querySelector('#gbFldJob')?.value  || '').trim(),
            department:        (modal.querySelector('#gbFldDept')?.value || '').trim(),
            usageLocation:     modal.querySelector('#gbFldLocale')?.value || 'NL',
            password:          (modal.querySelector('#gbFldPwd')?.value  || '').trim(),
          };
          if (!fd.displayName || !fd.userPrincipalName || !fd.password) {
            showFieldErrors(modal, ['gbFldDisplay','gbFldUpn','gbFldPwd'], fd);
            return;
          }
          state.formData = fd;
          // Licenties laden als nog niet gedaan
          const licPromise = _licenses.length ? Promise.resolve(_licenses) : loadLicenses(tid);
          licPromise.then(() => { state.step = 1; renderStep(); });
        };

      } else if (state.step === 1) {
        // Stap 2: licenties kiezen
        const licRows = _licenses.length
          ? _licenses.map((l) => {
              const availClass = l.available === 0 ? 'none' : l.available < 5 ? 'low' : '';
              const checked = state.selectedLicenses.includes(l.skuId) ? 'checked' : '';
              return `<label class="gb-license-item">
                <input type="checkbox" data-skuid="${escHtml(l.skuId)}" ${checked} ${l.available === 0 ? 'disabled' : ''}>
                <span class="gb-license-name">${escHtml(l.displayName)}</span>
                <span class="gb-license-avail ${availClass}">${l.consumed}/${l.enabled}</span>
              </label>`;
            }).join('')
          : `<div class="gb-empty">Geen licenties beschikbaar of laden mislukt.</div>`;

        content.innerHTML = `
          <p style="font-size:.875rem;color:var(--text-secondary);margin:0 0 .75rem">
            Kies de licenties voor <strong>${escHtml(state.formData.displayName)}</strong>.<br>
            Licenties kunnen later ook worden aangepast.
          </p>
          <div class="gb-license-list">${licRows}</div>
          <div style="margin-top:.75rem">
            <label style="font-size:.82rem;color:var(--text-secondary)">
              <input type="checkbox" id="gbDryRunProv" style="accent-color:var(--accent)">
              Dry-run (preview zonder aanmaken)
            </label>
          </div>`;

        footer.innerHTML = `
          <button class="gb-btn gb-btn-secondary gb-back">← Terug</button>
          <button class="gb-btn gb-btn-primary gb-create">Gebruiker aanmaken</button>`;
        footer.querySelector('.gb-back').onclick = () => { state.step = 0; renderStep(); };
        footer.querySelector('.gb-create').onclick = () => {
          const checked = [...modal.querySelectorAll('.gb-license-item input:checked')]
            .map((el) => el.dataset.skuid).filter(Boolean);
          state.selectedLicenses = checked;
          const dryRun = modal.querySelector('#gbDryRunProv')?.checked ?? false;
          executeProvision(state, dryRun, modal);
        };
      }
    }

    overlay.appendChild(modal);
    document.body.appendChild(overlay);
    renderStep();
  }

  function showFieldErrors(modal, fieldIds, data) {
    fieldIds.forEach((id) => {
      const el = modal.querySelector(`#${id}`);
      if (el && !el.value.trim()) el.classList.add('error');
    });
    // Verwijder error klasse bij input
    fieldIds.forEach((id) => {
      modal.querySelector(`#${id}`)?.addEventListener('input', (e) => e.target.classList.remove('error'), { once: true });
    });
  }

  function executeProvision(state, dryRun, modal) {
    const tid = getTenantId();
    const footer  = modal.querySelector('#gbProvFooter');
    const content = modal.querySelector('#gbProvContent');
    if (footer) footer.innerHTML = `<span style="font-size:.82rem;color:var(--text-muted)">Gebruiker aanmaken...</span>`;

    const payload = Object.assign({}, state.formData, {
      licenseSkuIds: state.selectedLicenses,
      dry_run: dryRun,
    });

    gbApiFetch(API.m365.users(tid), {
      method: 'POST',
      body: JSON.stringify(payload),
    })
      .then((data) => {
        const isOk   = data.ok !== false;
        const isDry  = dryRun;
        const msg    = data.message || (isOk ? 'Gebruiker aangemaakt' : 'Aanmaken mislukt');
        const licAssigned = data.licenses_assigned || data.preview?.licenseCount;

        if (content) {
          const cls = isDry ? 'gb-result-dryrun' : (isOk ? 'gb-result-ok' : 'gb-result-error');
          const icon = isDry ? 'ℹ️' : (isOk ? '✅' : '❌');
          content.innerHTML = `
            <div class="gb-result ${cls}">
              <div class="gb-result-icon">${icon}</div>
              <div class="gb-result-msg">${escHtml(msg)}</div>
            </div>
            ${data.upn ? `<p style="font-size:.83rem;color:var(--text-secondary);margin:.5rem 0 0">UPN: <code>${escHtml(data.upn)}</code></p>` : ''}
            ${(data.preview?.licenseCount !== undefined) ? `<p style="font-size:.83rem;color:var(--text-secondary);margin:.25rem 0 0">Licenties: ${data.preview.licenseCount}</p>` : ''}`;
        }
        if (footer) {
          footer.innerHTML = `<button class="gb-btn gb-btn-primary gb-done">Sluiten</button>`;
          footer.querySelector('.gb-done').onclick = () => { closeAllModals(); if (!dryRun && isOk) loadUsers(); };
        }
      })
      .catch((err) => {
        if (content) content.innerHTML = `<div class="gb-result gb-result-error"><div class="gb-result-msg">Fout: ${escHtml(String(err))}</div></div>`;
        if (footer) {
          footer.innerHTML = `<button class="gb-btn gb-btn-secondary gb-done">Sluiten</button>`;
          footer.querySelector('.gb-done').onclick = closeAllModals;
        }
      });
  }

  // ── Provisioning-geschiedenis ────────────────────────────────────────────

  function loadProvisioningHistory() {
    const tid = getTenantId();
    if (!tid) return;
    const tbody = document.getElementById('gbHistoryBody');
    if (!tbody) return;
    tbody.innerHTML = `<tr><td colspan="6" class="gb-loading">Laden...</td></tr>`;

    gbApiFetch(API.m365.provisioningHistory(tid))
      .then((data) => {
        const items = data.items || [];
        if (!items.length) {
          tbody.innerHTML = `<tr><td colspan="6" class="gb-table-empty">Nog geen activiteit gelogd.</td></tr>`;
          return;
        }
        tbody.innerHTML = items.map((h) => {
          const statusBadge = h.status === 'success'
            ? `<span class="gb-status gb-status-enabled"><span class="gb-dot"></span>Succes</span>`
            : h.status === 'dry_run'
              ? `<span class="gb-status" style="background:#f0f9ff;color:#0369a1;border:1px solid #bae6fd"><span class="gb-dot" style="background:#0369a1"></span>Dry-run</span>`
              : `<span class="gb-status gb-status-disabled"><span class="gb-dot"></span>Mislukt</span>`;
          const action = h.action === 'create-user' ? 'Aangemaakt' : h.action === 'offboard-user' ? 'Offboarded' : escHtml(h.action);
          return `<tr>
            <td>${fmtDate(h.executed_at)}</td>
            <td>${action}</td>
            <td>${escHtml(h.target_display_name || h.target_upn || '—')}</td>
            <td>${statusBadge}</td>
            <td>${escHtml(h.executed_by || '—')}</td>
            <td style="max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${escHtml(h.error_message || '—')}</td>
          </tr>`;
        }).join('');
      })
      .catch(() => {
        tbody.innerHTML = `<tr><td colspan="6" class="gb-table-empty">Fout bij laden geschiedenis.</td></tr>`;
      });
  }

  // ── Modal helpers ────────────────────────────────────────────────────────

  function createOverlay(cls, onClose) {
    const el = document.createElement('div');
    el.className = cls;
    el.addEventListener('click', (e) => { if (e.target === el) onClose(); });
    return el;
  }

  function closeAllModals() {
    document.querySelectorAll('.gb-modal-overlay').forEach((el) => el.remove());
  }

  function showStatus(msg) {
    // Hergebruik bestaand toast systeem als aanwezig
    if (typeof window.showToast === 'function') { window.showToast(msg); return; }
    console.info('[Gebruikers]', msg);
  }

  // ── Tab switching ────────────────────────────────────────────────────────

  function switchGbTab(tab) {
    document.querySelectorAll('.gb-tab').forEach((btn) => {
      btn.classList.toggle('active', btn.dataset.gbTab === tab);
    });
    document.querySelectorAll('.gb-tab-panel').forEach((panel) => {
      panel.style.display = panel.dataset.gbPanel === tab ? '' : 'none';
    });
    renderCapabilityBanner(tab);
    if (tab === 'licenties') loadLicensesTab();
    if (tab === 'geschiedenis') loadProvisioningHistory();
  }

  function bindGbTabs() {
    document.querySelectorAll('.gb-tab[data-gb-tab]').forEach((btn) => {
      btn.addEventListener('click', () => switchGbTab(btn.dataset.gbTab));
    });
  }

  // ── Licenties tabblad ─────────────────────────────────────────────────────

  function loadLicensesTab() {
    const tid = getTenantId();
    if (!tid) {
      const grid = document.getElementById('gbLicGrid');
      if (grid) grid.innerHTML = '<p class="gb-empty">Selecteer eerst een tenant.</p>';
      return;
    }

    const grid    = document.getElementById('gbLicGrid');
    const counter = document.getElementById('gbLicCount');
    if (grid) grid.innerHTML = '<p class="gb-empty gb-loading">Licenties laden...</p>';

    gbApiFetch(API.m365.licenses(tid))
      .then((data) => {
        const lics = data.licenses || [];
        if (counter) counter.textContent = `${lics.length} licentie${lics.length !== 1 ? 's' : ''}`;
        if (!lics.length) {
          if (grid) grid.innerHTML = '<p class="gb-empty">Geen licenties gevonden voor deze tenant.</p>';
          return;
        }
        if (grid) {
          grid.innerHTML = lics.map((l) => {
            const pct = l.enabled > 0 ? Math.round((l.consumed / l.enabled) * 100) : 0;
            const barClass = pct >= 90 ? 'full' : pct >= 70 ? 'warn' : '';
            const avail = l.enabled - l.consumed;
            const wasteAlert = l.enabled >= 5 && pct < 10
              ? `<span class="gb-lic-badge gb-lic-badge--waste" title="Weinig gebruik — mogelijk onnodig betaald">Laag gebruik</span>` : '';
            const overAlert = pct >= 90
              ? `<span class="gb-lic-badge gb-lic-badge--over" title="Bijna vol — overweeg uitbreiding">Bijna vol</span>` : '';
            return `<div class="gb-lic-card" data-skuid="${escHtml(l.skuId)}" title="Klik om gekoppelde gebruikers te zien">
              <div class="gb-lic-name">${escHtml(l.displayName || l.skuPartNumber || l.skuId)}${wasteAlert}${overAlert}</div>
              <div class="gb-lic-stats">
                <span class="gb-lic-stat"><strong>${l.consumed}</strong> in gebruik</span>
                <span class="gb-lic-stat"><strong>${l.enabled}</strong> totaal</span>
                <span class="gb-lic-stat ${avail === 0 ? 'none' : avail < 5 ? 'low' : ''}">${avail} beschikbaar</span>
              </div>
              <div class="gb-lic-bar-wrap">
                <div class="gb-lic-bar ${barClass}" style="width:${pct}%"></div>
              </div>
              <div class="gb-lic-pct">${pct}% gebruikt</div>
            </div>`;
          }).join('');

          grid.querySelectorAll('.gb-lic-card[data-skuid]').forEach((card) => {
            card.addEventListener('click', () => openLicenseUsersModal(card.dataset.skuid, lics));
          });
        }
      })
      .catch((err) => {
        if (grid) grid.innerHTML = `<p class="gb-empty">Fout: ${escHtml(String(err))}</p>`;
      });
  }

  function openLicenseUsersModal(skuId, lics) {
    const lic = lics.find((l) => l.skuId === skuId);
    const licName = lic ? (lic.displayName || lic.skuPartNumber || skuId) : skuId;

    // Filter gebruikers met deze licentie op basis van licenseSkuIds als beschikbaar,
    // anders val terug op de geladen gebruikers vergelijken
    const matched = _users.filter((u) => {
      if (Array.isArray(u.licenseSkuIds)) return u.licenseSkuIds.includes(skuId);
      // Fallback: als licenseCount > 0 maar geen skuIds, toon alle met licenties
      return false;
    });

    const overlay = createOverlay('gb-modal-overlay', closeAllModals);
    const modal   = document.createElement('div');
    modal.className = 'gb-modal gb-modal-wide';
    modal.innerHTML = `
      <div class="gb-modal-header">
        <div>
          <div class="gb-modal-title">${escHtml(licName)}</div>
          <div class="gb-modal-subtitle">${matched.length} gebruiker${matched.length !== 1 ? 's' : ''} gekoppeld</div>
        </div>
        <button class="gb-modal-close">✕</button>
      </div>
      <div class="gb-modal-body">
        ${matched.length === 0
          ? `<p style="color:var(--text-secondary);font-size:.875rem">
               Geen gebruikers gevonden met deze licentie in de huidige gebruikerslijst.<br>
               <em>Tip: ververs de gebruikerslijst zodat skuIds worden meegeladen.</em>
             </p>`
          : `<div class="gb-table-wrap"><table class="gb-table">
               <thead><tr><th>Naam</th><th>UPN</th><th>Afdeling</th><th>Status</th></tr></thead>
               <tbody>${matched.map((u) => `
                 <tr>
                   <td><div class="gb-user-cell">
                     <div class="gb-avatar">${initials(u.displayName)}</div>
                     <span>${escHtml(u.displayName)}</span>
                   </div></td>
                   <td>${escHtml(u.userPrincipalName)}</td>
                   <td>${escHtml(u.department || '—')}</td>
                   <td>${u.accountEnabled
                     ? '<span class="gb-status gb-status-enabled"><span class="gb-dot"></span>Actief</span>'
                     : '<span class="gb-status gb-status-disabled"><span class="gb-dot"></span>Uitgeschakeld</span>'}</td>
                 </tr>`).join('')}
               </tbody>
             </table></div>`
        }
      </div>
      <div class="gb-modal-footer">
        <button class="gb-btn gb-btn-secondary gb-modal-cancel">Sluiten</button>
      </div>`;
    overlay.appendChild(modal);
    document.body.appendChild(overlay);
    modal.querySelector('.gb-modal-close').onclick = closeAllModals;
    modal.querySelector('.gb-modal-cancel').onclick = closeAllModals;
  }

  // ── Filter & zoek binding ────────────────────────────────────────────────

  function bindToolbar() {
    const search = document.getElementById('gbSearchInput');
    if (search) {
      search.addEventListener('input', () => {
        _searchQ = search.value.trim();
        renderUsersTable();
      });
    }

    document.querySelectorAll('.gb-filter-tab[data-filter]').forEach((btn) => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('.gb-filter-tab').forEach((b) => b.classList.remove('active'));
        btn.classList.add('active');
        _filterStatus = btn.dataset.filter;
        renderUsersTable();
      });
    });

    const btnRefresh = document.getElementById('gbBtnRefresh');
    if (btnRefresh) btnRefresh.addEventListener('click', loadUsers);

    const btnLiveScan = document.getElementById('gbBtnLiveScan');
    if (btnLiveScan) btnLiveScan.addEventListener('click', () => loadUsers({ strictLive: true }));

    const btnNew = document.getElementById('gbBtnNieuw');
    if (btnNew) btnNew.addEventListener('click', openProvisioningWizard);
  }

  // ── Publieke interface ───────────────────────────────────────────────────

  /**
   * Laad de Gebruikers sectie.
   * Wordt aangeroepen vanuit dashboard.js showSection('gebruikers')
   */
  window.loadGebruikersSection = function () {
    renderOverviewStats();
    bindGbTabs();
    bindToolbar();
    applySourceState();
    const btnRefreshLic = document.getElementById('gbBtnRefreshLic');
    if (btnRefreshLic) btnRefreshLic.onclick = loadLicensesTab;
    renderCapabilityBanner('gebruikers');
    loadUsers();
  };

  window.loadGebruikersHistory = function () {
    loadProvisioningHistory();
  };
  window.switchGebruikersTab = switchGbTab;
  window.scanGebruikersLive = function () {
    loadUsers({ strictLive: true });
  };

})();
