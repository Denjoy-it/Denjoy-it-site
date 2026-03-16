/**
 * Netwerk Inventaris — NetBox-inspired UI
 * Depends on apiFetch() and currentTenantId from dashboard.js
 */

// ---------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------
let kbCurrentTab   = 'overview';
let kbAssetTypes   = [];
let kbAssetsCache  = [];
let kbVlansCache   = [];
let kbPagesCache   = [];
let kbContactsCache= [];
let kbPasswordsCache = [];
let kbSoftwareCache = [];
let kbDomainsCache = [];
let kbChangelogCache = [];
let kbM365Cache = null;

// Cache-TTL: herlaad uiterlijk elke 5 minuten of bij tenantwissel
const _KB_CACHE_TTL = 5 * 60 * 1000;
let _kbCacheTime = 0;
let _kbCacheTid  = null;

/** Forceer herlaad bij volgende kbRefreshCounts-aanroep (gebruik na mutaties). */
function kbInvalidateCache() {
  _kbCacheTime = 0;
}
let kbEditingPageId= null;
let kbSelectedAssetId  = null;
let kbSelectedVlanId   = null;
let kbSelectedContactId= null;
let kbSelectedPasswordId = null;
let kbSelectedSoftwareId = null;
let kbSelectedDomainId = null;
let kbSelectedChangelogId = null;
let kbSelectedSwitchNode = null;
let kbSwitchTabAssets = [];
let kbSoftwareSourceFilter = 'all';
let kbContactSourceFilter = 'all';

// ---------------------------------------------------------------------------
// Sub-nav switching
// ---------------------------------------------------------------------------
function kbSwitchTab(tabName) {
  kbCurrentTab = tabName;

  document.querySelectorAll('.nb-tab').forEach((item) => {
    item.classList.toggle('active', item.dataset.kbTab === tabName);
  });
  document.querySelectorAll('.nb-pane').forEach((pane) => {
    pane.classList.toggle('active', pane.dataset.kbPane === tabName);
  });

  // Sync portal sub-nav actief item
  if (typeof setActiveSubnavItem === 'function') setActiveSubnavItem(tabName);

  const tid = currentTenantId;
  if (!tid) return;
  if (tabName === 'overview') kbLoadOverview(tid);
  if (tabName === 'assets')   kbLoadAssets(tid);
  if (tabName === 'vlans')    kbLoadVlans(tid);
  if (tabName === 'pages')    kbLoadPages(tid);
  if (tabName === 'contacts') kbLoadContacts(tid);
  if (tabName === 'passwords') kbLoadPasswords(tid);
  if (tabName === 'software') kbLoadSoftware(tid);
  if (tabName === 'domains') kbLoadDomains(tid);
  if (tabName === 'm365') kbLoadM365(tid);
  if (tabName === 'changelog') kbLoadChangelog(tid);
  if (tabName.startsWith('switch-')) kbRenderSwitchTab(tabName);

  // Ververs tellers na laden
  setTimeout(() => { if (typeof refreshSubnavCounts === 'function') refreshSubnavCounts(); }, 800);
}
window.kbSwitchTab = kbSwitchTab;

// ---------------------------------------------------------------------------
// Init
// ---------------------------------------------------------------------------
async function kbInit() {
  const tid = currentTenantId;
  if (!tid) {
    document.querySelectorAll('.nb-empty').forEach((el) => {
      el.textContent = 'Selecteer eerst een tenant.';
    });
    return;
  }

  if (!kbAssetTypes.length) {
    try {
      kbAssetTypes = await apiFetch(`/api/kb/${tid}/asset-types`);
      _populateTypeDropdowns(kbAssetTypes);
    } catch (_) {}
  }

  await kbRefreshCounts(tid);
  kbSwitchTab(kbCurrentTab);
}

async function kbRefreshCounts(tid, { force = false } = {}) {
  const now = Date.now();
  // Sla herlaad over als cache nog vers is voor dezelfde tenant
  if (!force && _kbCacheTid === tid && (now - _kbCacheTime) < _KB_CACHE_TTL) {
    _updateCountBadges();
    return;
  }
  try {
    const [assets, vlans, pages, contacts, passwords, software, domains, m365, changelog] = await Promise.all([
      apiFetch(`/api/kb/${tid}/assets`),
      apiFetch(`/api/kb/${tid}/vlans`),
      apiFetch(`/api/kb/${tid}/pages`),
      apiFetch(`/api/kb/${tid}/contacts`),
      apiFetch(`/api/kb/${tid}/passwords`),
      apiFetch(`/api/kb/${tid}/software`),
      apiFetch(`/api/kb/${tid}/domains`),
      apiFetch(`/api/kb/${tid}/m365`),
      apiFetch(`/api/kb/${tid}/changelog`),
    ]);
    kbAssetsCache    = assets;
    kbVlansCache     = vlans;
    kbPagesCache     = pages;
    kbContactsCache  = contacts;
    kbPasswordsCache = passwords;
    kbSoftwareCache  = software;
    kbDomainsCache   = domains;
    kbM365Cache      = m365;
    kbChangelogCache = changelog;
    kbSwitchTabAssets = assets.filter((a) => _isSwitchAsset(a));

    _kbCacheTid  = tid;
    _kbCacheTime = Date.now();

    _updateCountBadges();
    kbRenderSwitchTabs();
    if (kbCurrentTab === 'overview') kbLoadOverview(tid);
  } catch (_) {}
}

function _updateCountBadges() {
  const m365 = kbM365Cache;
  _setCount('nbCountAssets',    kbAssetsCache.length);
  _setCount('nbCountVlans',     kbVlansCache.length);
  _setCount('nbCountPages',     kbPagesCache.length);
  _setCount('nbCountContacts',  kbContactsCache.length + ((m365?.assessment_user_mailboxes || []).filter((item) => _looksLikePersonName(item.DisplayName || item.display_name)).length));
  _setCount('nbCountPasswords', kbPasswordsCache.length);
  _setCount('nbCountSoftware',  kbSoftwareCache.length + ((m365?.assessment_app_registrations || []).length));
  _setCount('nbCountDomains',   kbDomainsCache.length);
  _setCount('nbCountChangelog', kbChangelogCache.length);
}

function _setCount(id, n) {
  const el = document.getElementById(id);
  if (el) el.textContent = n;
}

function _populateTypeDropdowns(types) {
  const filterSel = document.getElementById('kbAssetTypeFilter');
  const formSel   = document.getElementById('kbAssetType');
  if (filterSel) {
    filterSel.innerHTML = '<option value="">Alle typen</option>' +
      types.map((t) => `<option value="${t.id}">${t.icon} ${t.name}</option>`).join('');
  }
  if (formSel) {
    formSel.innerHTML = '<option value="">— geen —</option>' +
      types.map((t) => `<option value="${t.id}">${t.icon} ${t.name}</option>`).join('');
  }
  _toggleAssetSwitchFields();
}

// ---------------------------------------------------------------------------
// Utility
// ---------------------------------------------------------------------------
function esc(v) {
  if (v == null || v === '') return '';
  return String(v)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function _dash(v) { return (v == null || v === '') ? '—' : esc(v); }

function _toNumber(v) {
  if (v == null || v === '') return null;
  const n = Number(String(v).replace(',', '.'));
  return Number.isFinite(n) ? n : null;
}

function _fmtCurrency(v) {
  const n = _toNumber(v);
  if (n == null) return '—';
  return new Intl.NumberFormat('nl-NL', { style: 'currency', currency: 'EUR' }).format(n);
}

function _daysUntil(dateValue) {
  if (!dateValue) return null;
  const target = new Date(dateValue);
  if (Number.isNaN(target.getTime())) return null;
  const now = new Date();
  target.setHours(0, 0, 0, 0);
  now.setHours(0, 0, 0, 0);
  return Math.round((target - now) / 86400000);
}

function _daysUntilLoose(dateValue) {
  if (!dateValue) return null;
  const native = _daysUntil(dateValue);
  if (native != null) return native;
  const m = String(dateValue).match(/^(\d{2})-(\d{2})-(\d{4})$/);
  if (!m) return null;
  return _daysUntil(`${m[3]}-${m[2]}-${m[1]}`);
}

function _softwareItems() {
  const manual = kbSoftwareCache.map((item) => ({
    ...item,
    _key: `manual-${item.id}`,
    _source: 'Handmatig',
    _readonly: false,
  }));
  const discovered = (kbM365Cache?.assessment_app_registrations || []).map((app, index) => {
    const secretDate = app.SecretExpiration || app.secret_expiration || null;
    const certDate = app.CertificateExpiration || app.certificate_expiration || null;
    const datedEntries = [secretDate, certDate]
      .map((v) => ({ raw: v, days: _daysUntilLoose(v) }))
      .filter((v) => v.raw && v.days != null)
      .sort((a, b) => a.days - b.days);
    const nextExpiry = datedEntries[0]?.raw || null;
    const nextDays = datedEntries[0]?.days ?? null;
    return {
      id: `assessment-${index}`,
      name: app.DisplayName || app.display_name || `App ${index + 1}`,
      vendor: 'Microsoft 365 / Azure',
      software_type: 'App-registratie',
      licenses: null,
      unit_price: null,
      total_price: null,
      cost: null,
      expiry: nextExpiry,
      status: nextDays == null ? 'geen vervaldatum' : (nextDays < 0 ? 'verlopen' : 'assessment'),
      ref: app.AppId || app.app_id || null,
      notes: `Secret: ${app.SecretExpirationStatus || app.secret_expiration_status || 'Geen secret'}\nCertificaat: ${app.CertificateExpirationStatus || app.certificate_expiration_status || 'Geen certificaat'}`,
      assessment_secret_status: app.SecretExpirationStatus || app.secret_expiration_status || 'Geen secret',
      assessment_cert_status: app.CertificateExpirationStatus || app.certificate_expiration_status || 'Geen certificaat',
      assessment_secret_expiry: secretDate,
      assessment_cert_expiry: certDate,
      _key: `assessment-${index}`,
      _source: 'Assessment',
      _readonly: true,
    };
  });
  return [...manual, ...discovered];
}

function _looksLikePersonName(name) {
  if (!name) return false;
  const cleaned = String(name).trim();
  if (!cleaned || cleaned.includes('|') || cleaned.includes('@')) return false;
  const parts = cleaned.split(/\s+/).filter(Boolean);
  if (parts.length < 2) return false;
  return parts.every((part) => /[A-Za-zÀ-ÿ]/.test(part));
}

function _contactItems() {
  const manual = kbContactsCache.map((item) => ({
    ...item,
    _key: `manual-${item.id}`,
    _source: 'Handmatig',
    _readonly: false,
  }));
  const discovered = (kbM365Cache?.assessment_user_mailboxes || [])
    .filter((item) => _looksLikePersonName(item.DisplayName || item.display_name))
    .map((item, index) => ({
      id: `assessment-contact-${index}`,
      name: item.DisplayName || item.display_name,
      role: 'M365 gebruiker',
      phone: null,
      email: item.PrimarySmtpAddress || item.primary_smtp_address || null,
      is_primary_contact: 0,
      notes: item.WhenCreated || item.when_created ? `Aangemaakt: ${item.WhenCreated || item.when_created}` : null,
      created_at: item.WhenCreated || item.when_created || null,
      _key: `assessment-contact-${index}`,
      _source: 'Assessment',
      _readonly: true,
    }));
  return [...manual, ...discovered];
}


const PURPOSE_LABELS = {
  user: 'Gebruikers', server: 'Servers', mgmt: 'Management',
  guest: 'Gasten', iot: 'IoT', dmz: 'DMZ',
};

const PAGE_CAT_LABELS = {
  network: '🌐 Netwerk', security: '🛡️ Beveiliging',
  procedures: '📋 Procedures', contacts: '👥 Contacten',
};

// ---------------------------------------------------------------------------
// OVERVIEW
// ---------------------------------------------------------------------------
function kbLoadOverview() {
  const wrap = document.getElementById('kbOverviewPane');
  if (!wrap) return;
  const switchCount = kbAssetsCache.filter((asset) => _isSwitchAsset(asset)).length;
  const softwareItems = _softwareItems();
  const expiringSoftware = softwareItems
    .map((item) => ({ ...item, days: _daysUntil(item.expiry) }))
    .filter((item) => item.days != null && item.days >= 0 && item.days <= 90)
    .sort((a, b) => a.days - b.days)
    .slice(0, 6);
  const appCredentialAlerts = (kbM365Cache?.assessment_app_registrations || [])
    .flatMap((app) => {
      const items = [];
      const secretDays = _daysUntilLoose(app.SecretExpiration || app.secret_expiration);
      const certDays = _daysUntilLoose(app.CertificateExpiration || app.certificate_expiration);
      if (secretDays != null && secretDays <= 90) {
        items.push({
          name: app.DisplayName || app.display_name || 'App-registratie',
          kind: 'App secret',
          date: app.SecretExpiration || app.secret_expiration,
          days: secretDays,
        });
      }
      if (certDays != null && certDays <= 90) {
        items.push({
          name: app.DisplayName || app.display_name || 'App-registratie',
          kind: 'App certificaat',
          date: app.CertificateExpiration || app.certificate_expiration,
          days: certDays,
        });
      }
      return items;
    })
    .sort((a, b) => a.days - b.days)
    .slice(0, 6);
  const expiringDomains = kbDomainsCache
    .flatMap((item) => ([
      item.expiry ? { name: item.domain, source: 'Domein', date: item.expiry, days: _daysUntil(item.expiry) } : null,
      item.ssl_expiry ? { name: item.domain, source: 'SSL', date: item.ssl_expiry, days: _daysUntil(item.ssl_expiry) } : null,
    ]))
    .filter((item) => item && item.days != null && item.days >= 0 && item.days <= 90)
    .sort((a, b) => a.days - b.days)
    .slice(0, 6);
  const tenantDomainAlerts = (kbM365Cache?.assessment_domain_dns_checks || [])
    .map((item) => {
      const spf = String(item.SPF || item.spf || 'Unknown');
      const dmarc = String(item.DMARC || item.dmarc || 'Unknown');
      const dkim = String(item.DKIM || item.dkim || 'Unknown');
      const issues = [
        spf !== 'Present' ? 'SPF' : null,
        dmarc !== 'Present' ? 'DMARC' : null,
        dkim !== 'Present' && dkim !== 'Present (Selector1)' ? 'DKIM' : null,
      ].filter(Boolean);
      return issues.length ? {
        name: item.Domain || item.domain || 'Domein',
        source: 'M365 DNS',
        issue: issues.join(', ontbreekt'),
      } : null;
    })
    .filter(Boolean)
    .slice(0, 6);
  const recentChanges = [...kbChangelogCache]
    .sort((a, b) => String(b.change_date || '').localeCompare(String(a.change_date || '')))
    .slice(0, 6);

  wrap.innerHTML = `
    <div class="nb-overview-grid">
      <section class="nb-overview-card nb-overview-hero">
        <div>
          <div class="nb-overview-eyebrow">Kennisbank</div>
          <h3>Centraal NetDocs-overzicht</h3>
          <p>Alle netwerkdocumentatie, licenties, domeinen, M365 en wijzigingen per tenant op een plek.</p>
        </div>
        <div class="nb-overview-actions">
          <button type="button" class="nb-btn nb-btn-primary" onclick="kbSwitchTab('assets')">Apparaten</button>
          <button type="button" class="nb-btn nb-btn-ghost" onclick="kbSwitchTab('software')">Licenties</button>
          <button type="button" class="nb-btn nb-btn-ghost" onclick="kbSwitchTab('domains')">Domeinen</button>
        </div>
      </section>

      <section class="nb-overview-stats">
        <article class="nb-overview-stat"><span class="nb-overview-stat-value">${kbAssetsCache.length}</span><span class="nb-overview-stat-label">Apparaten</span><span class="nb-overview-stat-sub">${switchCount} switches</span></article>
        <article class="nb-overview-stat"><span class="nb-overview-stat-value">${kbVlansCache.length}</span><span class="nb-overview-stat-label">VLANs</span><span class="nb-overview-stat-sub">${kbPagesCache.length} documenten</span></article>
        <article class="nb-overview-stat"><span class="nb-overview-stat-value">${kbPasswordsCache.length}</span><span class="nb-overview-stat-label">Passwords</span><span class="nb-overview-stat-sub">${kbContactsCache.length} contacten</span></article>
        <article class="nb-overview-stat"><span class="nb-overview-stat-value">${softwareItems.length}</span><span class="nb-overview-stat-label">Software</span><span class="nb-overview-stat-sub">${kbDomainsCache.length} domeinen</span></article>
      </section>

      <section class="nb-overview-card">
        <div class="nb-overview-card-head">
          <h3>Aankomende vervaldatums</h3>
          <button type="button" class="nb-btn nb-btn-xs nb-btn-ghost" onclick="kbSwitchTab('software')">Software</button>
        </div>
        ${(expiringSoftware.length || appCredentialAlerts.length) ? `
          <div class="nb-overview-list">
            ${expiringSoftware.map((item) => `<div class="nb-overview-list-item"><strong>${esc(item.name)}</strong><span>${esc(item.expiry)} · ${item.days} dagen</span></div>`).join('')}
            ${appCredentialAlerts.map((item) => `<div class="nb-overview-list-item"><strong>${esc(item.name)}</strong><span>${esc(item.kind)} · ${esc(item.date)} · ${item.days < 0 ? 'verlopen' : `${item.days} dagen`}</span></div>`).join('')}
          </div>
        ` : '<p class="nb-empty nb-empty-inline">Geen softwarelicenties of app-credentials die binnen 90 dagen verlopen.</p>'}
      </section>

      <section class="nb-overview-card">
        <div class="nb-overview-card-head">
          <h3>Domein & SSL alerts</h3>
          <button type="button" class="nb-btn nb-btn-xs nb-btn-ghost" onclick="kbSwitchTab('domains')">Domeinen</button>
        </div>
        ${(expiringDomains.length || tenantDomainAlerts.length) ? `
          <div class="nb-overview-list">
            ${expiringDomains.map((item) => `<div class="nb-overview-list-item"><strong>${esc(item.name)}</strong><span>${esc(item.source)} · ${esc(item.date)} · ${item.days} dagen</span></div>`).join('')}
            ${tenantDomainAlerts.map((item) => `<div class="nb-overview-list-item"><strong>${esc(item.name)}</strong><span>${esc(item.source)} · ${esc(item.issue)}</span></div>`).join('')}
          </div>
        ` : '<p class="nb-empty nb-empty-inline">Geen domein-, SSL- of tenant DNS-waarschuwingen gevonden.</p>'}
      </section>

      <section class="nb-overview-card">
        <div class="nb-overview-card-head">
          <h3>Microsoft 365</h3>
          <button type="button" class="nb-btn nb-btn-xs nb-btn-ghost" onclick="kbSwitchTab('m365')">M365</button>
        </div>
        <div class="nb-overview-m365">
          <div><span>Tenant</span><strong>${_dash(kbM365Cache?.tenant_name)}</strong></div>
          <div><span>Licenties</span><strong>${_dash(kbM365Cache?.licenses_used)} / ${_dash(kbM365Cache?.licenses_total)}</strong></div>
          <div><span>MFA</span><strong>${_dash(kbM365Cache?.mfa)}</strong></div>
          <div><span>Conditional Access</span><strong>${kbM365Cache?.conditional_access ? 'Actief' : 'Niet ingesteld'}</strong></div>
        </div>
      </section>

      <section class="nb-overview-card">
        <div class="nb-overview-card-head">
          <h3>Recente wijzigingen</h3>
          <button type="button" class="nb-btn nb-btn-xs nb-btn-ghost" onclick="kbSwitchTab('changelog')">Wijzigingslog</button>
        </div>
        ${recentChanges.length ? `
          <div class="nb-overview-list">
            ${recentChanges.map((item) => `<div class="nb-overview-list-item"><strong>${esc(item.action)}</strong><span>${_dash(item.user_name)} · ${_dash(item.change_date)}</span></div>`).join('')}
          </div>
        ` : '<p class="nb-empty nb-empty-inline">Nog geen wijzigingen vastgelegd.</p>'}
      </section>
    </div>`;
}

const SWITCH_ROLE_LABELS = {
  client: 'Client device',
  server: 'Server',
  ap: 'Access Point',
  camera: 'Camera',
  trunk: 'Trunk / uplink',
  empty: 'Leeg',
};

function _isSwitchTypeId(typeId) {
  const t = kbAssetTypes.find((item) => String(item.id) === String(typeId));
  return !!t && t.name === 'switch';
}

function _isSwitchAsset(asset) {
  return !!asset && (asset.type_name === 'switch' || _isSwitchTypeId(asset.asset_type_id) || !!asset.switch_config);
}

function _createDefaultSwitchPort(number) {
  return {
    number,
    label: '',
    role: 'empty',
    vlan: '',
    patch_panel: '',
    connected_to: '',
    note: '',
  };
}

function _createDefaultSfp(number) {
  return {
    number,
    label: '',
    speed: '10G',
    connected_to: '',
    note: '',
  };
}

function _normalizeSwitchEntries(items, count, factory) {
  const result = [];
  for (let i = 0; i < count; i += 1) {
    const base = factory(i + 1);
    const current = items?.[i] || {};
    result.push({ ...base, ...current, number: i + 1 });
  }
  return result;
}

function _buildSwitchConfig(asset) {
  const raw = (asset && typeof asset.switch_config === 'object' && asset.switch_config) || {};
  const portCount = [24, 48].includes(Number(raw.port_count)) ? Number(raw.port_count) : 24;
  const sfpCount = Math.max(0, Math.min(8, Number(raw.sfp_count) || 2));
  return {
    port_count: portCount,
    sfp_count: sfpCount,
    ports: _normalizeSwitchEntries(raw.ports, portCount, _createDefaultSwitchPort),
    sfps: _normalizeSwitchEntries(raw.sfps, sfpCount, _createDefaultSfp),
  };
}

function _toggleAssetSwitchFields() {
  const wrap = document.getElementById('kbSwitchFields');
  if (!wrap) return;
  wrap.style.display = _isSwitchTypeId(document.getElementById('kbAssetType')?.value) ? '' : 'none';
}

function _switchTabName(assetId) {
  return `switch-${assetId}`;
}

function kbRenderSwitchTabs() {
  const tabbar = document.getElementById('kbTabbar');
  const panesWrap = document.getElementById('kbDynamicSwitchPanes');
  if (!tabbar || !panesWrap) return;

  tabbar.querySelectorAll('.nb-tab[data-kb-dynamic="switch"]').forEach((el) => el.remove());
  panesWrap.innerHTML = '';

  kbSwitchTabAssets.forEach((asset) => {
    const tabName = _switchTabName(asset.id);
    const tab = document.createElement('a');
    tab.href = '#';
    tab.className = 'nb-tab';
    tab.dataset.kbTab = tabName;
    tab.dataset.kbDynamic = 'switch';
    tab.innerHTML = `
      <span class="nb-tab-icon">🔀</span>
      <span class="nb-tab-label">${esc(asset.name || 'Switch')}</span>`;
    tab.addEventListener('click', (e) => {
      e.preventDefault();
      kbSwitchTab(tabName);
    });
    tabbar.appendChild(tab);

    const pane = document.createElement('div');
    pane.className = 'nb-pane';
    pane.dataset.kbPane = tabName;
    pane.dataset.kbDynamic = 'switch';
    panesWrap.appendChild(pane);
  });

  if (kbCurrentTab.startsWith('switch-') && !kbSwitchTabAssets.find((asset) => _switchTabName(asset.id) === kbCurrentTab)) {
    kbCurrentTab = 'assets';
  }
}

function kbRenderSwitchTab(tabName) {
  const pane = document.querySelector(`.nb-pane[data-kb-pane="${tabName}"]`);
  const asset = kbSwitchTabAssets.find((item) => _switchTabName(item.id) === tabName);
  if (!pane || !asset) return;
  pane.innerHTML = _renderSwitchStandalonePane(asset);
  _bindSwitchInteractive(pane, asset.id);
}

function _collectSwitchConfigFromModal(existingAsset) {
  if (!_isSwitchTypeId(document.getElementById('kbAssetType').value)) return null;
  const current = _buildSwitchConfig(existingAsset || {});
  const portCount = Number(document.getElementById('kbAssetSwitchPortCount').value || current.port_count || 24);
  const sfpCount = Number(document.getElementById('kbAssetSwitchSfpCount').value || current.sfp_count || 2);
  return {
    port_count: [24, 48].includes(portCount) ? portCount : 24,
    sfp_count: Math.max(0, Math.min(8, sfpCount)),
    ports: _normalizeSwitchEntries(current.ports, [24, 48].includes(portCount) ? portCount : 24, _createDefaultSwitchPort),
    sfps: _normalizeSwitchEntries(current.sfps, Math.max(0, Math.min(8, sfpCount)), _createDefaultSfp),
  };
}

// ---------------------------------------------------------------------------
// ASSETS
// ---------------------------------------------------------------------------
async function kbLoadAssets(tid) {
  const wrap = document.getElementById('kbAssetsTable');
  if (!wrap) return;
  if (!kbAssetsCache.length) {
    wrap.innerHTML = '<p class="nb-empty">Laden…</p>';
    try { kbAssetsCache = await apiFetch(`/api/kb/${tid}/assets`); }
    catch (e) { wrap.innerHTML = `<p class="nb-empty">${esc(e.message)}</p>`; return; }
  }
  _renderAssetsTable();
}

function _renderAssetsTable() {
  const wrap   = document.getElementById('kbAssetsTable');
  const search = (document.getElementById('kbAssetSearch')?.value || '').toLowerCase();
  const type   = document.getElementById('kbAssetTypeFilter')?.value || '';
  const status = document.getElementById('kbAssetStatusFilter')?.value;

  const filtered = kbAssetsCache.filter((a) => {
    const matchSearch = !search ||
      (a.name        || '').toLowerCase().includes(search) ||
      (a.ip_address  || '').toLowerCase().includes(search) ||
      (a.hostname    || '').toLowerCase().includes(search) ||
      (a.location    || '').toLowerCase().includes(search) ||
      (a.vendor      || '').toLowerCase().includes(search);
    const matchType   = !type   || String(a.asset_type_id) === type;
    const matchStatus = status === '' || status == null || String(a.is_active) === status;
    return matchSearch && matchType && matchStatus;
  });

  const countEl = document.getElementById('nbAssetListCount');
  if (countEl) countEl.textContent = `(${filtered.length})`;

  if (!filtered.length) {
    wrap.innerHTML = '<p class="nb-empty">Geen apparaten gevonden.</p>';
    return;
  }

  wrap.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>Naam</th>
          <th>Type</th>
          <th>IP-adres</th>
          <th>Locatie</th>
          <th>Vendor / Model</th>
          <th>Status</th>
        </tr>
      </thead>
      <tbody>
        ${filtered.map((a) => `
          <tr data-id="${a.id}" class="${kbSelectedAssetId === a.id ? 'nb-row-selected' : ''}">
            <td class="nb-td-name">${esc(a.name)}${a.hostname ? `<br><small style="font-weight:400;color:var(--text-secondary,#6c757d)">${esc(a.hostname)}</small>` : ''}</td>
            <td><span class="nb-badge nb-badge-type">${esc(a.type_icon || '')} ${esc(a.type_name)}</span></td>
            <td><code>${_dash(a.ip_address)}</code></td>
            <td>${_dash(a.location)}</td>
            <td>${_dash(a.vendor)}${a.model ? ` <span style="color:var(--text-secondary,#6c757d)">${esc(a.model)}</span>` : ''}</td>
            <td><span class="nb-badge ${a.is_active ? 'nb-badge-active' : 'nb-badge-inactive'}">${a.is_active ? 'Actief' : 'Inactief'}</span></td>
          </tr>`).join('')}
      </tbody>
    </table>`;

  wrap.querySelectorAll('tbody tr').forEach((row) => {
    row.addEventListener('click', () => {
      const id = parseInt(row.dataset.id);
      const asset = kbAssetsCache.find((a) => a.id === id);
      if (_isSwitchAsset(asset)) {
        kbSelectedAssetId = id;
        wrap.querySelectorAll('tr').forEach((r) => r.classList.toggle('nb-row-selected', parseInt(r.dataset.id) === id));
        kbSelectedSwitchNode = { assetId: id, kind: 'port', index: 0 };
        kbSwitchTab(_switchTabName(id));
        return;
      }
      if (kbSelectedAssetId === id) {
        kbSelectedAssetId = null;
        document.getElementById('kbAssetDetail').style.display = 'none';
        wrap.querySelectorAll('tr').forEach((r) => r.classList.remove('nb-row-selected'));
      } else {
        kbSelectedAssetId = id;
        wrap.querySelectorAll('tr').forEach((r) => r.classList.toggle('nb-row-selected', parseInt(r.dataset.id) === id));
        _showAssetDetail(asset);
      }
    });
  });
}

function _showAssetDetail(a) {
  const panel = document.getElementById('kbAssetDetail');
  if (!a || !panel) return;
  const switchSection = _isSwitchAsset(a) ? _renderSwitchSection(a) : '';
  panel.classList.toggle('nb-detail-panel-switch', _isSwitchAsset(a));
  panel.style.display = 'block';
  panel.innerHTML = `
    <div class="nb-detail-header">
      <h3 class="nb-detail-title">
        <span class="nb-badge nb-badge-type">${esc(a.type_icon || '')} ${esc(a.type_name)}</span>
        ${esc(a.name)}
        <span class="nb-badge ${a.is_active ? 'nb-badge-active' : 'nb-badge-inactive'}">${a.is_active ? 'Actief' : 'Inactief'}</span>
      </h3>
      <div class="nb-detail-actions">
        <button class="nb-btn nb-btn-ghost nb-btn-sm" onclick="kbOpenEditAsset(${a.id})">✏️ Bewerken</button>
        <button class="nb-btn nb-btn-ghost nb-btn-sm" onclick="document.getElementById('kbAssetDetail').style.display='none';kbSelectedAssetId=null;">✕ Sluiten</button>
      </div>
    </div>
    <div class="nb-detail-grid">
      <div class="nb-detail-field"><span class="nb-detail-field-label">Hostname</span><span class="nb-detail-field-value">${_dash(a.hostname)}</span></div>
      <div class="nb-detail-field"><span class="nb-detail-field-label">IP-adres</span><span class="nb-detail-field-value">${_dash(a.ip_address)}</span></div>
      <div class="nb-detail-field"><span class="nb-detail-field-label">Locatie</span><span class="nb-detail-field-value">${_dash(a.location)}</span></div>
      <div class="nb-detail-field"><span class="nb-detail-field-label">Vendor</span><span class="nb-detail-field-value">${_dash(a.vendor)}</span></div>
      <div class="nb-detail-field"><span class="nb-detail-field-label">Model</span><span class="nb-detail-field-value">${_dash(a.model)}</span></div>
      <div class="nb-detail-field"><span class="nb-detail-field-label">Firmware</span><span class="nb-detail-field-value">${_dash(a.firmware)}</span></div>
      <div class="nb-detail-field"><span class="nb-detail-field-label">Serienummer</span><span class="nb-detail-field-value">${_dash(a.serial)}</span></div>
      <div class="nb-detail-field"><span class="nb-detail-field-label">Aangemaakt</span><span class="nb-detail-field-value">${_dash(a.created_at?.substring(0,10))}</span></div>
    </div>
    ${a.notes ? `<div class="nb-detail-notes"><strong>Notities</strong><br>${esc(a.notes)}</div>` : ''}
    ${switchSection}`;
  _bindSwitchInteractive(panel, a.id);
  panel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function _bindSwitchInteractive(root, assetId) {
  root.querySelectorAll('.nb-switch-port, .nb-switch-sfp').forEach((btn) => {
    btn.addEventListener('click', () => kbSelectSwitchNode(assetId, btn.dataset.kind, parseInt(btn.dataset.index, 10)));
  });
}

function _renderAssetMetaBlock(a) {
  return `
    <div class="nb-detail-grid">
      <div class="nb-detail-field"><span class="nb-detail-field-label">Hostname</span><span class="nb-detail-field-value">${_dash(a.hostname)}</span></div>
      <div class="nb-detail-field"><span class="nb-detail-field-label">IP-adres</span><span class="nb-detail-field-value">${_dash(a.ip_address)}</span></div>
      <div class="nb-detail-field"><span class="nb-detail-field-label">Locatie</span><span class="nb-detail-field-value">${_dash(a.location)}</span></div>
      <div class="nb-detail-field"><span class="nb-detail-field-label">Vendor</span><span class="nb-detail-field-value">${_dash(a.vendor)}</span></div>
      <div class="nb-detail-field"><span class="nb-detail-field-label">Model</span><span class="nb-detail-field-value">${_dash(a.model)}</span></div>
      <div class="nb-detail-field"><span class="nb-detail-field-label">Firmware</span><span class="nb-detail-field-value">${_dash(a.firmware)}</span></div>
      <div class="nb-detail-field"><span class="nb-detail-field-label">Serienummer</span><span class="nb-detail-field-value">${_dash(a.serial)}</span></div>
      <div class="nb-detail-field"><span class="nb-detail-field-label">Aangemaakt</span><span class="nb-detail-field-value">${_dash(a.created_at?.substring(0,10))}</span></div>
    </div>`;
}

function _renderSwitchStandalonePane(a) {
  return `
    <div class="nb-breadcrumb">
      <span>Kennisbank</span><span class="nb-bc-sep">›</span>
      <span>Switches</span><span class="nb-bc-sep">›</span>
      <span class="nb-bc-current">${esc(a.name)}</span>
    </div>
    <div class="nb-page-header">
      <h2 class="nb-page-title">${esc(a.name)} <span class="nb-page-count">${esc(a.vendor || '')}${a.model ? ` / ${esc(a.model)}` : ''}</span></h2>
      <div class="nb-page-actions">
        <button class="nb-btn nb-btn-secondary" type="button" onclick="kbOpenEditAsset(${a.id})">Bewerken</button>
      </div>
    </div>
    <div class="nb-switch-pane-body">
      <div class="nb-switch-pane-summary">
        <div class="nb-detail-header" style="margin-bottom:0.9rem;">
          <h3 class="nb-detail-title">
            <span class="nb-badge nb-badge-type">${esc(a.type_icon || '')} ${esc(a.type_name)}</span>
            ${esc(a.name)}
            <span class="nb-badge ${a.is_active ? 'nb-badge-active' : 'nb-badge-inactive'}">${a.is_active ? 'Actief' : 'Inactief'}</span>
          </h3>
        </div>
        ${_renderAssetMetaBlock(a)}
        ${a.notes ? `<div class="nb-detail-notes"><strong>Notities</strong><br>${esc(a.notes)}</div>` : ''}
      </div>
      <div class="nb-switch-pane-layout">
        ${_renderSwitchSection(a)}
      </div>
    </div>`;
}

function kbOpenAddAsset() {
  _resetAssetModal('Apparaat toevoegen', null);
  document.getElementById('kbAssetModal').style.display = 'flex';
}
window.kbOpenAddAsset = kbOpenAddAsset;

async function kbOpenEditAsset(assetId) {
  const a = kbAssetsCache.find((x) => x.id === assetId)
         || await apiFetch(`/api/kb/${currentTenantId}/assets/${assetId}`).catch(() => null);
  if (!a) return;
  _resetAssetModal('Apparaat bewerken', a);
  document.getElementById('kbAssetModal').style.display = 'flex';
}
window.kbOpenEditAsset = kbOpenEditAsset;

function _resetAssetModal(title, a) {
  document.getElementById('kbAssetModalTitle').textContent = title;
  document.getElementById('kbAssetId').value        = a?.id ?? '';
  document.getElementById('kbAssetName').value      = a?.name ?? '';
  document.getElementById('kbAssetType').value      = a?.asset_type_id ?? '';
  document.getElementById('kbAssetHostname').value  = a?.hostname ?? '';
  document.getElementById('kbAssetIP').value        = a?.ip_address ?? '';
  document.getElementById('kbAssetLocation').value  = a?.location ?? '';
  document.getElementById('kbAssetVendor').value    = a?.vendor ?? '';
  document.getElementById('kbAssetModel').value     = a?.model ?? '';
  document.getElementById('kbAssetFirmware').value  = a?.firmware ?? '';
  document.getElementById('kbAssetSerial').value    = a?.serial ?? '';
  document.getElementById('kbAssetNotes').value     = a?.notes ?? '';
  document.getElementById('kbAssetActive').checked  = a ? !!a.is_active : true;
  const switchConfig = _buildSwitchConfig(a || {});
  document.getElementById('kbAssetSwitchPortCount').value = String(switchConfig.port_count);
  document.getElementById('kbAssetSwitchSfpCount').value = String(switchConfig.sfp_count);
  document.getElementById('kbDeleteAssetBtn').style.display = a ? '' : 'none';
  _toggleAssetSwitchFields();
}

async function kbSaveAsset() {
  const tid = currentTenantId;
  const id  = document.getElementById('kbAssetId').value;
  const existingAsset = kbAssetsCache.find((item) => String(item.id) === String(id));
  const payload = {
    name: document.getElementById('kbAssetName').value.trim(),
    asset_type_id: document.getElementById('kbAssetType').value || null,
    hostname:  document.getElementById('kbAssetHostname').value.trim() || null,
    ip_address:document.getElementById('kbAssetIP').value.trim() || null,
    location:  document.getElementById('kbAssetLocation').value.trim() || null,
    vendor:    document.getElementById('kbAssetVendor').value.trim() || null,
    model:     document.getElementById('kbAssetModel').value.trim() || null,
    firmware:  document.getElementById('kbAssetFirmware').value.trim() || null,
    serial:    document.getElementById('kbAssetSerial').value.trim() || null,
    notes:     document.getElementById('kbAssetNotes').value.trim() || null,
    is_active: document.getElementById('kbAssetActive').checked ? 1 : 0,
    switch_config: _collectSwitchConfigFromModal(existingAsset),
  };
  if (!payload.name) { alert('Naam is verplicht.'); return; }
  if (!payload.asset_type_id) { alert('Type is verplicht.'); return; }
  try {
    let saved;
    if (id) {
      saved = await apiFetch(`/api/kb/${tid}/assets/${id}`, { method: 'PUT', body: JSON.stringify(payload) });
    } else {
      saved = await apiFetch(`/api/kb/${tid}/assets`, { method: 'POST', body: JSON.stringify(payload) });
    }
    document.getElementById('kbAssetModal').style.display = 'none';
    kbAssetsCache = []; kbInvalidateCache();
    kbSelectedAssetId = saved?.id || null;
    kbSelectedSwitchNode = saved && _isSwitchAsset(saved) ? { assetId: saved.id, kind: 'port', index: 0 } : null;
    await kbRefreshCounts(tid);
    await kbLoadAssets(tid);
    if (saved?.id) {
      const current = kbAssetsCache.find((item) => item.id === saved.id) || saved;
      if (_isSwitchAsset(current)) {
        kbSwitchTab(_switchTabName(current.id));
        document.querySelector(`.nb-pane[data-kb-pane="${_switchTabName(current.id)}"] .nb-switch-card`)?.scrollIntoView({ behavior: 'smooth', block: 'start' });
      } else {
        _renderAssetsTable();
        _showAssetDetail(current);
      }
    } else {
      document.getElementById('kbAssetDetail').style.display = 'none';
    }
  } catch (e) { alert('Opslaan mislukt: ' + e.message); }
}

function _renderSwitchSection(asset) {
  const cfg = _buildSwitchConfig(asset);
  const selected = (kbSelectedSwitchNode && kbSelectedSwitchNode.assetId === asset.id)
    ? kbSelectedSwitchNode
    : { assetId: asset.id, kind: 'port', index: 0 };
  const activePort = selected.kind === 'sfp' ? null : cfg.ports[selected.index] || cfg.ports[0];
  const activeSfp = selected.kind === 'sfp' ? (cfg.sfps[selected.index] || cfg.sfps[0] || null) : null;
  kbSelectedSwitchNode = { assetId: asset.id, kind: activeSfp ? 'sfp' : 'port', index: activeSfp ? selected.index : (activePort ? activePort.number - 1 : 0) };

  return `
    <section class="nb-switch-card">
      <div class="nb-switch-card-head">
        <div>
          <h4 class="nb-switch-title">Switchdocumentatie</h4>
          <p class="nb-switch-subtitle">${cfg.port_count}-poort layout met ${cfg.sfp_count} SFP uplink${cfg.sfp_count === 1 ? '' : 's'}</p>
        </div>
        <div class="nb-switch-actions">
          <button class="nb-btn nb-btn-secondary nb-btn-sm" type="button" onclick="kbResetSwitchLayout(${asset.id})">Layout resetten</button>
        </div>
      </div>
      <div class="nb-switch-legend">
        ${Object.entries(SWITCH_ROLE_LABELS).map(([key, label]) => `
          <span class="nb-switch-legend-item"><span class="nb-switch-color nb-switch-color-${key}"></span>${esc(label)}</span>
        `).join('')}
      </div>
      <div class="nb-switch-visual-wrap">
        <div class="nb-switch-visual-grid">
          ${cfg.ports.map((port) => `
            <button type="button" class="nb-switch-port nb-switch-port-${esc(port.role || 'empty')} ${(activePort && activePort.number === port.number) ? 'is-selected' : ''}" data-kind="port" data-index="${port.number - 1}">
              <span class="nb-switch-port-no">${port.number}</span>
              <span class="nb-switch-port-label">${esc(port.label || port.connected_to || 'Leeg')}</span>
            </button>
          `).join('')}
        </div>
        ${cfg.sfp_count ? `
          <div class="nb-switch-sfp-wrap">
            <div class="nb-switch-sfp-title">SFP uplinks</div>
            <div class="nb-switch-sfp-grid">
              ${cfg.sfps.map((sfp) => `
                <button type="button" class="nb-switch-sfp ${(activeSfp && activeSfp.number === sfp.number) ? 'is-selected' : ''}" data-kind="sfp" data-index="${sfp.number - 1}">
                  <span class="nb-switch-sfp-name">SFP${sfp.number}</span>
                  <span class="nb-switch-sfp-target">${esc(sfp.connected_to || sfp.label || 'Niet ingevuld')}</span>
                </button>
              `).join('')}
            </div>
          </div>` : ''}
      </div>
      <div class="nb-switch-editor">
        ${activePort ? _renderSwitchPortEditor(activePort) : ''}
        ${activeSfp ? _renderSwitchSfpEditor(activeSfp) : ''}
      </div>
      <div class="nb-switch-table-wrap">
        <table class="nb-switch-table">
          <thead>
            <tr>
              <th>Poort</th>
              <th>Type</th>
              <th>Device</th>
              <th>VLAN</th>
              <th>Patchpanel</th>
              <th>Verbinding</th>
              <th>Opmerking</th>
            </tr>
          </thead>
          <tbody>
            ${cfg.ports.map((port) => `
              <tr>
                <td><code>${port.number}</code></td>
                <td><span class="nb-badge nb-switch-role-badge nb-switch-role-badge-${esc(port.role || 'empty')}">${esc(SWITCH_ROLE_LABELS[port.role || 'empty'] || port.role || 'Leeg')}</span></td>
                <td>${_dash(port.label)}</td>
                <td>${_dash(port.vlan)}</td>
                <td>${_dash(port.patch_panel)}</td>
                <td>${_dash(port.connected_to)}</td>
                <td>${_dash(port.note)}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    </section>`;
}

function _renderSwitchPortEditor(port) {
  return `
    <div class="nb-switch-editor-card">
      <div class="nb-switch-editor-head">
        <h5>Poort ${port.number}</h5>
        <span class="nb-badge nb-switch-role-badge nb-switch-role-badge-${esc(port.role || 'empty')}">${esc(SWITCH_ROLE_LABELS[port.role || 'empty'] || 'Leeg')}</span>
      </div>
      <div class="nb-switch-editor-grid">
        <label>Device / label<input type="text" id="kbSwitchPortLabel" class="nb-input" value="${esc(port.label)}"></label>
        <label>Type
          <select id="kbSwitchPortRole" class="nb-input">
            ${Object.entries(SWITCH_ROLE_LABELS).map(([key, label]) => `<option value="${key}" ${port.role === key ? 'selected' : ''}>${esc(label)}</option>`).join('')}
          </select>
        </label>
        <label>VLAN<input type="text" id="kbSwitchPortVlan" class="nb-input" value="${esc(port.vlan)}" placeholder="10 / trunk"></label>
        <label>Patchpanel<input type="text" id="kbSwitchPortPatch" class="nb-input" value="${esc(port.patch_panel)}" placeholder="PP1-01"></label>
        <label>Verbinding<input type="text" id="kbSwitchPortConnected" class="nb-input" value="${esc(port.connected_to)}" placeholder="AP-02 / Core-SW-01"></label>
        <label class="nb-form-full">Opmerking<textarea id="kbSwitchPortNote" class="nb-input nb-textarea" rows="2">${esc(port.note)}</textarea></label>
      </div>
      <div class="nb-switch-editor-actions">
        <button class="nb-btn nb-btn-primary nb-btn-sm" type="button" onclick="kbSaveSelectedSwitchNode()">Poort opslaan</button>
      </div>
    </div>`;
}

function _renderSwitchSfpEditor(sfp) {
  return `
    <div class="nb-switch-editor-card">
      <div class="nb-switch-editor-head">
        <h5>SFP${sfp.number}</h5>
        <span class="nb-badge nb-badge-active">${esc(sfp.speed || '10G')}</span>
      </div>
      <div class="nb-switch-editor-grid">
        <label>Naam<input type="text" id="kbSwitchSfpLabel" class="nb-input" value="${esc(sfp.label)}" placeholder="Uplink A"></label>
        <label>Snelheid<input type="text" id="kbSwitchSfpSpeed" class="nb-input" value="${esc(sfp.speed)}" placeholder="10G"></label>
        <label>Verbinding<input type="text" id="kbSwitchSfpConnected" class="nb-input" value="${esc(sfp.connected_to)}" placeholder="Core-SW-01"></label>
        <label class="nb-form-full">Opmerking<textarea id="kbSwitchSfpNote" class="nb-input nb-textarea" rows="2">${esc(sfp.note)}</textarea></label>
      </div>
      <div class="nb-switch-editor-actions">
        <button class="nb-btn nb-btn-primary nb-btn-sm" type="button" onclick="kbSaveSelectedSwitchNode()">SFP opslaan</button>
      </div>
    </div>`;
}

async function _saveSwitchConfig(assetId, config) {
  const asset = kbAssetsCache.find((item) => item.id === assetId);
  if (!asset) return;
  const payload = {
    name: asset.name,
    asset_type_id: asset.asset_type_id,
    hostname: asset.hostname,
    ip_address: asset.ip_address,
    location: asset.location,
    vendor: asset.vendor,
    model: asset.model,
    firmware: asset.firmware,
    serial: asset.serial,
    notes: asset.notes,
    is_active: asset.is_active ? 1 : 0,
    switch_config: config,
  };
  const updated = await apiFetch(`/api/kb/${currentTenantId}/assets/${assetId}`, { method: 'PUT', body: JSON.stringify(payload) });
  const idx = kbAssetsCache.findIndex((item) => item.id === assetId);
  if (idx >= 0) kbAssetsCache[idx] = updated;
  kbSwitchTabAssets = kbAssetsCache.filter((item) => _isSwitchAsset(item));
  kbRenderSwitchTabs();
  kbSelectedAssetId = assetId;
  _showAssetDetail(updated);
  _renderAssetsTable();
  if (kbCurrentTab === _switchTabName(assetId)) kbRenderSwitchTab(kbCurrentTab);
}

function kbSelectSwitchNode(assetId, kind, index) {
  kbSelectedSwitchNode = { assetId, kind, index };
  const asset = kbAssetsCache.find((item) => item.id === assetId);
  if (!asset) return;
  if (kbCurrentTab === _switchTabName(assetId)) {
    kbRenderSwitchTab(kbCurrentTab);
  } else {
    _showAssetDetail(asset);
  }
}
window.kbSelectSwitchNode = kbSelectSwitchNode;

async function kbSaveSelectedSwitchNode() {
  const selected = kbSelectedSwitchNode;
  if (!selected) return;
  const asset = kbAssetsCache.find((item) => item.id === selected.assetId);
  if (!asset) return;
  const config = _buildSwitchConfig(asset);
  const activePane = (kbCurrentTab === _switchTabName(asset.id))
    ? document.querySelector(`.nb-pane[data-kb-pane="${_switchTabName(asset.id)}"]`)
    : document.getElementById('kbAssetDetail');
  const getField = (selector) => activePane?.querySelector(selector);
  if (selected.kind === 'sfp') {
    const item = config.sfps[selected.index];
    if (!item) return;
    item.label = getField('#kbSwitchSfpLabel')?.value.trim() || '';
    item.speed = getField('#kbSwitchSfpSpeed')?.value.trim() || '10G';
    item.connected_to = getField('#kbSwitchSfpConnected')?.value.trim() || '';
    item.note = getField('#kbSwitchSfpNote')?.value.trim() || '';
  } else {
    const item = config.ports[selected.index];
    if (!item) return;
    item.label = getField('#kbSwitchPortLabel')?.value.trim() || '';
    item.role = getField('#kbSwitchPortRole')?.value || 'empty';
    item.vlan = getField('#kbSwitchPortVlan')?.value.trim() || '';
    item.patch_panel = getField('#kbSwitchPortPatch')?.value.trim() || '';
    item.connected_to = getField('#kbSwitchPortConnected')?.value.trim() || '';
    item.note = getField('#kbSwitchPortNote')?.value.trim() || '';
  }
  try {
    await _saveSwitchConfig(asset.id, config);
  } catch (e) {
    alert('Poortinformatie opslaan mislukt: ' + e.message);
  }
}
window.kbSaveSelectedSwitchNode = kbSaveSelectedSwitchNode;

async function kbResetSwitchLayout(assetId) {
  if (!confirm('Switchlayout resetten naar een lege standaardindeling?')) return;
  const asset = kbAssetsCache.find((item) => item.id === assetId);
  if (!asset) return;
  const cfg = _buildSwitchConfig(asset);
  try {
    await _saveSwitchConfig(assetId, {
      port_count: cfg.port_count,
      sfp_count: cfg.sfp_count,
      ports: _normalizeSwitchEntries([], cfg.port_count, _createDefaultSwitchPort),
      sfps: _normalizeSwitchEntries([], cfg.sfp_count, _createDefaultSfp),
    });
  } catch (e) {
    alert('Layout resetten mislukt: ' + e.message);
  }
}
window.kbResetSwitchLayout = kbResetSwitchLayout;

async function kbDeleteAsset() {
  const tid = currentTenantId;
  const id  = document.getElementById('kbAssetId').value;
  if (!id || !confirm('Apparaat verwijderen?')) return;
  try {
    await apiFetch(`/api/kb/${tid}/assets/${id}`, { method: 'DELETE' });
    document.getElementById('kbAssetModal').style.display = 'none';
    kbAssetsCache = [];
    kbSelectedAssetId = null;
    if (kbCurrentTab === _switchTabName(Number(id))) kbCurrentTab = 'assets';
    document.getElementById('kbAssetDetail').style.display = 'none';
    await kbRefreshCounts(tid);
    kbLoadAssets(tid);
  } catch (e) { alert('Verwijderen mislukt: ' + e.message); }
}

// ---------------------------------------------------------------------------
// VLANS
// ---------------------------------------------------------------------------
async function kbLoadVlans(tid) {
  const wrap = document.getElementById('kbVlansTable');
  if (!wrap) return;
  if (!kbVlansCache.length) {
    wrap.innerHTML = '<p class="nb-empty">Laden…</p>';
    try { kbVlansCache = await apiFetch(`/api/kb/${tid}/vlans`); }
    catch (e) { wrap.innerHTML = `<p class="nb-empty">${esc(e.message)}</p>`; return; }
  }
  _renderVlansTable();
}

function _renderVlansTable() {
  const wrap    = document.getElementById('kbVlansTable');
  const search  = (document.getElementById('kbVlanSearch')?.value || '').toLowerCase();
  const purpose = document.getElementById('kbVlanPurposeFilter')?.value || '';

  const filtered = kbVlansCache.filter((v) => {
    const matchSearch = !search ||
      String(v.vlan_id).includes(search) ||
      (v.name    || '').toLowerCase().includes(search) ||
      (v.subnet  || '').toLowerCase().includes(search);
    const matchPurpose = !purpose || v.purpose === purpose;
    return matchSearch && matchPurpose;
  });

  const countEl = document.getElementById('nbVlanListCount');
  if (countEl) countEl.textContent = `(${filtered.length})`;

  if (!filtered.length) {
    wrap.innerHTML = '<p class="nb-empty">Geen VLANs gevonden.</p>';
    return;
  }

  wrap.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>VLAN-ID</th>
          <th>Naam</th>
          <th>Subnet</th>
          <th>Gateway</th>
          <th>Doel</th>
          <th>Beschrijving</th>
        </tr>
      </thead>
      <tbody>
        ${filtered.map((v) => `
          <tr data-id="${v.id}" class="${kbSelectedVlanId === v.id ? 'nb-row-selected' : ''}">
            <td class="nb-td-name"><code>${v.vlan_id}</code></td>
            <td>${esc(v.name)}</td>
            <td><code>${_dash(v.subnet)}</code></td>
            <td><code>${_dash(v.gateway)}</code></td>
            <td><span class="nb-badge nb-badge-${v.purpose}">${PURPOSE_LABELS[v.purpose] || esc(v.purpose)}</span></td>
            <td>${_dash(v.description)}</td>
          </tr>`).join('')}
      </tbody>
    </table>`;

  wrap.querySelectorAll('tbody tr').forEach((row) => {
    row.addEventListener('click', () => {
      const id = parseInt(row.dataset.id);
      if (kbSelectedVlanId === id) {
        kbSelectedVlanId = null;
        document.getElementById('kbVlanDetail').style.display = 'none';
        wrap.querySelectorAll('tr').forEach((r) => r.classList.remove('nb-row-selected'));
      } else {
        kbSelectedVlanId = id;
        wrap.querySelectorAll('tr').forEach((r) => r.classList.toggle('nb-row-selected', parseInt(r.dataset.id) === id));
        _showVlanDetail(kbVlansCache.find((v) => v.id === id));
      }
    });
  });
}

function _showVlanDetail(v) {
  const panel = document.getElementById('kbVlanDetail');
  if (!v || !panel) return;
  panel.style.display = 'block';
  panel.innerHTML = `
    <div class="nb-detail-header">
      <h3 class="nb-detail-title">
        VLAN ${v.vlan_id} — ${esc(v.name)}
        <span class="nb-badge nb-badge-${v.purpose}">${PURPOSE_LABELS[v.purpose] || esc(v.purpose)}</span>
      </h3>
      <div class="nb-detail-actions">
        <button class="nb-btn nb-btn-ghost nb-btn-sm" onclick="kbOpenEditVlan(${v.id})">✏️ Bewerken</button>
        <button class="nb-btn nb-btn-ghost nb-btn-sm" onclick="document.getElementById('kbVlanDetail').style.display='none';kbSelectedVlanId=null;">✕ Sluiten</button>
      </div>
    </div>
    <div class="nb-detail-grid">
      <div class="nb-detail-field"><span class="nb-detail-field-label">Subnet</span><span class="nb-detail-field-value"><code>${_dash(v.subnet)}</code></span></div>
      <div class="nb-detail-field"><span class="nb-detail-field-label">Gateway</span><span class="nb-detail-field-value"><code>${_dash(v.gateway)}</code></span></div>
      <div class="nb-detail-field"><span class="nb-detail-field-label">Beschrijving</span><span class="nb-detail-field-value">${_dash(v.description)}</span></div>
      <div class="nb-detail-field"><span class="nb-detail-field-label">Aangemaakt</span><span class="nb-detail-field-value">${_dash(v.created_at?.substring(0,10))}</span></div>
    </div>
    ${v.notes ? `<div class="nb-detail-notes"><strong>Notities</strong><br>${esc(v.notes)}</div>` : ''}`;
  panel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function kbOpenAddVlan() {
  _resetVlanModal('VLAN toevoegen', null);
  document.getElementById('kbVlanModal').style.display = 'flex';
}
window.kbOpenAddVlan = kbOpenAddVlan;

async function kbOpenEditVlan(dbId) {
  const v = kbVlansCache.find((x) => x.id === dbId);
  if (!v) return;
  _resetVlanModal('VLAN bewerken', v);
  document.getElementById('kbVlanModal').style.display = 'flex';
}
window.kbOpenEditVlan = kbOpenEditVlan;

function _resetVlanModal(title, v) {
  document.getElementById('kbVlanModalTitle').textContent = title;
  document.getElementById('kbVlanDbId').value      = v?.id ?? '';
  document.getElementById('kbVlanId').value        = v?.vlan_id ?? '';
  document.getElementById('kbVlanName').value      = v?.name ?? '';
  document.getElementById('kbVlanSubnet').value    = v?.subnet ?? '';
  document.getElementById('kbVlanGateway').value   = v?.gateway ?? '';
  document.getElementById('kbVlanPurpose').value   = v?.purpose ?? 'user';
  document.getElementById('kbVlanDescription').value = v?.description ?? '';
  document.getElementById('kbVlanNotes').value     = v?.notes ?? '';
  document.getElementById('kbDeleteVlanBtn').style.display = v ? '' : 'none';
}

async function kbSaveVlan() {
  const tid  = currentTenantId;
  const dbId = document.getElementById('kbVlanDbId').value;
  const payload = {
    vlan_id:     document.getElementById('kbVlanId').value,
    name:        document.getElementById('kbVlanName').value.trim(),
    subnet:      document.getElementById('kbVlanSubnet').value.trim() || null,
    gateway:     document.getElementById('kbVlanGateway').value.trim() || null,
    purpose:     document.getElementById('kbVlanPurpose').value,
    description: document.getElementById('kbVlanDescription').value.trim() || null,
    notes:       document.getElementById('kbVlanNotes').value.trim() || null,
  };
  if (!payload.vlan_id || !payload.name) { alert('VLAN-ID en naam zijn verplicht.'); return; }
  try {
    if (dbId) {
      await apiFetch(`/api/kb/${tid}/vlans/${dbId}`, { method: 'PUT', body: JSON.stringify(payload) });
    } else {
      await apiFetch(`/api/kb/${tid}/vlans`, { method: 'POST', body: JSON.stringify(payload) });
    }
    document.getElementById('kbVlanModal').style.display = 'none';
    kbVlansCache = []; kbInvalidateCache();
    kbSelectedVlanId = null;
    document.getElementById('kbVlanDetail').style.display = 'none';
    await kbRefreshCounts(tid);
    kbLoadVlans(tid);
  } catch (e) { alert('Opslaan mislukt: ' + e.message); }
}

async function kbDeleteVlan() {
  const tid  = currentTenantId;
  const dbId = document.getElementById('kbVlanDbId').value;
  if (!dbId || !confirm('VLAN verwijderen?')) return;
  try {
    await apiFetch(`/api/kb/${tid}/vlans/${dbId}`, { method: 'DELETE' });
    document.getElementById('kbVlanModal').style.display = 'none';
    kbVlansCache = []; kbInvalidateCache();
    kbSelectedVlanId = null;
    document.getElementById('kbVlanDetail').style.display = 'none';
    await kbRefreshCounts(tid);
    kbLoadVlans(tid);
  } catch (e) { alert('Verwijderen mislukt: ' + e.message); }
}

// ---------------------------------------------------------------------------
// KB PAGES (Markdown docs)
// ---------------------------------------------------------------------------
async function kbLoadPages(tid) {
  const list = document.getElementById('kbPagesList');
  if (!list) return;
  if (!kbPagesCache.length) {
    list.innerHTML = '<p class="nb-empty">Laden…</p>';
    try { kbPagesCache = await apiFetch(`/api/kb/${tid}/pages`); }
    catch (e) { list.innerHTML = `<p class="nb-empty">${esc(e.message)}</p>`; return; }
  }
  _renderPagesList();
  const countEl = document.getElementById('nbPageListCount');
  if (countEl) countEl.textContent = `(${kbPagesCache.length})`;
}

function _renderPagesList() {
  const list = document.getElementById('kbPagesList');
  if (!kbPagesCache.length) {
    list.innerHTML = '<p class="nb-empty">Geen documenten. Maak een nieuw document aan.</p>';
    return;
  }
  const byCategory = kbPagesCache.reduce((acc, p) => {
    (acc[p.category || 'network'] = acc[p.category || 'network'] || []).push(p);
    return acc;
  }, {});

  list.innerHTML = Object.entries(byCategory).map(([cat, items]) => `
    <div class="nb-pages-category">
      <div class="nb-pages-category-label">${PAGE_CAT_LABELS[cat] || cat}</div>
      ${items.map((p) => `
        <div class="nb-pages-item ${kbEditingPageId === p.id ? 'active' : ''}" onclick="kbOpenPage(${p.id})">
          <span class="nb-pages-item-title">${esc(p.title)}</span>
          <span class="nb-pages-item-date">${(p.updated_at || '').substring(0, 10)}</span>
        </div>`).join('')}
    </div>`).join('');
}

async function kbOpenPage(pageId) {
  const tid = currentTenantId;
  try {
    const p = await apiFetch(`/api/kb/${tid}/pages/${pageId}`);
    kbEditingPageId = p.id;
    document.getElementById('kbPageTitleInput').value    = p.title || '';
    document.getElementById('kbPageCategorySelect').value= p.category || 'network';
    document.getElementById('kbPageContent').value       = p.content || '';
    document.getElementById('kbDeletePageBtn').style.display = '';
    document.getElementById('kbPageEditor').style.display   = 'flex';
    kbUpdatePreview();
    _renderPagesList();
  } catch (e) { alert('Laden mislukt: ' + e.message); }
}
window.kbOpenPage = kbOpenPage;

function kbOpenNewPage() {
  kbEditingPageId = null;
  document.getElementById('kbPageTitleInput').value     = '';
  document.getElementById('kbPageCategorySelect').value = 'network';
  document.getElementById('kbPageContent').value        = '';
  document.getElementById('kbDeletePageBtn').style.display = 'none';
  document.getElementById('kbPageEditor').style.display    = 'flex';
  document.getElementById('kbPagePreview').innerHTML       = '';
  document.getElementById('kbPageTitleInput').focus();
  _renderPagesList();
}

async function kbSavePage() {
  const tid = currentTenantId;
  const payload = {
    title:    document.getElementById('kbPageTitleInput').value.trim(),
    content:  document.getElementById('kbPageContent').value,
    category: document.getElementById('kbPageCategorySelect').value,
  };
  if (!payload.title) { alert('Titel is verplicht.'); return; }
  try {
    let saved;
    if (kbEditingPageId) {
      saved = await apiFetch(`/api/kb/${tid}/pages/${kbEditingPageId}`, { method: 'PUT', body: JSON.stringify(payload) });
    } else {
      saved = await apiFetch(`/api/kb/${tid}/pages`, { method: 'POST', body: JSON.stringify(payload) });
    }
    kbEditingPageId = saved.id;
    kbPagesCache = [];
    await kbRefreshCounts(tid, { force: true });
    kbLoadPages(tid);
  } catch (e) { alert('Opslaan mislukt: ' + e.message); }
}

async function kbDeletePage() {
  const tid = currentTenantId;
  if (!kbEditingPageId || !confirm('Document verwijderen?')) return;
  try {
    await apiFetch(`/api/kb/${tid}/pages/${kbEditingPageId}`, { method: 'DELETE' });
    kbEditingPageId = null;
    document.getElementById('kbPageEditor').style.display = 'none';
    kbPagesCache = [];
    await kbRefreshCounts(tid, { force: true });
    kbLoadPages(tid);
  } catch (e) { alert('Verwijderen mislukt: ' + e.message); }
}

function kbUpdatePreview() {
  const content = document.getElementById('kbPageContent')?.value || '';
  const preview = document.getElementById('kbPagePreview');
  if (preview) preview.innerHTML = _mdToHtml(content);
}

function _safeLinkUrl(url) {
  // Sta alleen http/https/mailto toe — blokkeer javascript: en data: URLs
  return /^(https?:|mailto:)/i.test(url.trim()) ? url : '#';
}

function _mdToHtml(md) {
  let h = md
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/```[\w]*\n([\s\S]*?)```/g, (_, c) => `<pre><code>${c}</code></pre>`)
    .replace(/^#{4} (.+)$/gm, '<h4>$1</h4>')
    .replace(/^#{3} (.+)$/gm, '<h3>$1</h3>')
    .replace(/^#{2} (.+)$/gm, '<h2>$1</h2>')
    .replace(/^# (.+)$/gm,   '<h1>$1</h1>')
    .replace(/^---+$/gm, '<hr>')
    .replace(/\*\*\*(.+?)\*\*\*/g, '<strong><em>$1</em></strong>')
    .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
    .replace(/\*(.+?)\*/g, '<em>$1</em>')
    .replace(/`([^`]+)`/g, '<code>$1</code>')
    .replace(/\[([^\]]+)\]\(([^)]+)\)/g, (_, label, url) =>
      `<a href="${_safeLinkUrl(url)}" target="_blank" rel="noopener noreferrer">${label}</a>`)
    .replace(/^[-*] (.+)$/gm, '<li>$1</li>')
    .replace(/\n\n+/g, '</p><p>');
  return `<p>${h}</p>`;
}

// ---------------------------------------------------------------------------
// CONTACTS
// ---------------------------------------------------------------------------
async function kbLoadContacts(tid) {
  const wrap = document.getElementById('kbContactsTable');
  if (!wrap) return;
  if (!kbContactsCache.length) {
    wrap.innerHTML = '<p class="nb-empty">Laden…</p>';
    try { kbContactsCache = await apiFetch(`/api/kb/${tid}/contacts`); }
    catch (e) { wrap.innerHTML = `<p class="nb-empty">${esc(e.message)}</p>`; return; }
  }
  _renderContactsTable();
}

function _renderContactsTable() {
  const wrap = document.getElementById('kbContactsTable');
  const countEl = document.getElementById('nbContactListCount');
  const items = _contactItems().filter((item) => (
    kbContactSourceFilter === 'all' ||
    (kbContactSourceFilter === 'manual' && item._source === 'Handmatig') ||
    (kbContactSourceFilter === 'assessment' && item._source === 'Assessment')
  ));
  if (countEl) countEl.textContent = `(${items.length})`;

  if (!items.length) {
    wrap.innerHTML = '<p class="nb-empty">Geen contacten gevonden.</p>';
    return;
  }

  wrap.innerHTML = `
    <table>
      <thead>
        <tr><th></th><th>Bron</th><th>Naam</th><th>Rol</th><th>Telefoon</th><th>E-mail</th></tr>
      </thead>
      <tbody>
        ${items.map((c) => `
          <tr data-key="${c._key}" class="${kbSelectedContactId === c._key ? 'nb-row-selected' : ''}">
            <td style="width:28px;text-align:center">${c.is_primary_contact ? '⭐' : ''}</td>
            <td>${esc(c._source)}</td>
            <td class="nb-td-name">${esc(c.name)}</td>
            <td>${_dash(c.role)}</td>
            <td>${c.phone ? `<a href="tel:${esc(c.phone)}">${esc(c.phone)}</a>` : '—'}</td>
            <td>${c.email ? `<a href="mailto:${esc(c.email)}">${esc(c.email)}</a>` : '—'}</td>
          </tr>`).join('')}
      </tbody>
    </table>`;

  wrap.querySelectorAll('tbody tr').forEach((row) => {
    row.addEventListener('click', () => {
      const key = row.dataset.key;
      if (kbSelectedContactId === key) {
        kbSelectedContactId = null;
        document.getElementById('kbContactDetail').style.display = 'none';
        wrap.querySelectorAll('tr').forEach((r) => r.classList.remove('nb-row-selected'));
      } else {
        kbSelectedContactId = key;
        wrap.querySelectorAll('tr').forEach((r) => r.classList.toggle('nb-row-selected', r.dataset.key === key));
        _showContactDetail(items.find((c) => c._key === key));
      }
    });
  });
}

function _showContactDetail(c) {
  const panel = document.getElementById('kbContactDetail');
  if (!c || !panel) return;
  panel.style.display = 'block';
  panel.innerHTML = `
    <div class="nb-detail-header">
      <h3 class="nb-detail-title">
        ${c.is_primary_contact ? '⭐ ' : ''}${esc(c.name)}
        ${c.is_primary_contact ? '<span class="nb-badge nb-badge-active">Primair</span>' : ''}
      </h3>
      <div class="nb-detail-actions">
        ${c._readonly ? '<span class="nb-badge nb-badge-type">Automatisch uit assessment</span>' : `<button class="nb-btn nb-btn-ghost nb-btn-sm" onclick="kbOpenEditContact(${c.id})">✏️ Bewerken</button>`}
        <button class="nb-btn nb-btn-ghost nb-btn-sm" onclick="document.getElementById('kbContactDetail').style.display='none';kbSelectedContactId=null;">✕ Sluiten</button>
      </div>
    </div>
    <div class="nb-detail-grid">
      <div class="nb-detail-field"><span class="nb-detail-field-label">Bron</span><span class="nb-detail-field-value">${esc(c._source || 'Handmatig')}</span></div>
      <div class="nb-detail-field"><span class="nb-detail-field-label">Rol</span><span class="nb-detail-field-value">${_dash(c.role)}</span></div>
      <div class="nb-detail-field"><span class="nb-detail-field-label">Telefoon</span><span class="nb-detail-field-value">${c.phone ? `<a href="tel:${esc(c.phone)}">${esc(c.phone)}</a>` : '—'}</span></div>
      <div class="nb-detail-field"><span class="nb-detail-field-label">E-mail</span><span class="nb-detail-field-value">${c.email ? `<a href="mailto:${esc(c.email)}">${esc(c.email)}</a>` : '—'}</span></div>
    </div>
    ${c.notes ? `<div class="nb-detail-notes"><strong>Notities</strong><br>${esc(c.notes)}</div>` : ''}`;
}

function kbOpenAddContact() {
  _resetContactModal('Contact toevoegen', null);
  document.getElementById('kbContactModal').style.display = 'flex';
}
window.kbOpenAddContact = kbOpenAddContact;

async function kbOpenEditContact(contactId) {
  const c = kbContactsCache.find((x) => x.id === contactId);
  if (!c) return;
  _resetContactModal('Contact bewerken', c);
  document.getElementById('kbContactModal').style.display = 'flex';
}
window.kbOpenEditContact = kbOpenEditContact;

function _resetContactModal(title, c) {
  document.getElementById('kbContactModalTitle').textContent = title;
  document.getElementById('kbContactId').value      = c?.id ?? '';
  document.getElementById('kbContactName').value    = c?.name ?? '';
  document.getElementById('kbContactRole').value    = c?.role ?? '';
  document.getElementById('kbContactPhone').value   = c?.phone ?? '';
  document.getElementById('kbContactEmail').value   = c?.email ?? '';
  document.getElementById('kbContactPrimary').checked = c ? !!c.is_primary_contact : false;
  document.getElementById('kbContactNotes').value   = c?.notes ?? '';
  document.getElementById('kbDeleteContactBtn').style.display = c ? '' : 'none';
}

async function kbSaveContact() {
  const tid = currentTenantId;
  const id  = document.getElementById('kbContactId').value;
  const payload = {
    name:               document.getElementById('kbContactName').value.trim(),
    role:               document.getElementById('kbContactRole').value.trim() || null,
    phone:              document.getElementById('kbContactPhone').value.trim() || null,
    email:              document.getElementById('kbContactEmail').value.trim() || null,
    is_primary_contact: document.getElementById('kbContactPrimary').checked ? 1 : 0,
    notes:              document.getElementById('kbContactNotes').value.trim() || null,
  };
  if (!payload.name) { alert('Naam is verplicht.'); return; }
  try {
    if (id) {
      await apiFetch(`/api/kb/${tid}/contacts/${id}`, { method: 'PUT', body: JSON.stringify(payload) });
    } else {
      await apiFetch(`/api/kb/${tid}/contacts`, { method: 'POST', body: JSON.stringify(payload) });
    }
    document.getElementById('kbContactModal').style.display = 'none';
    kbContactsCache = []; kbInvalidateCache();
    kbSelectedContactId = null;
    document.getElementById('kbContactDetail').style.display = 'none';
    await kbRefreshCounts(tid);
    kbLoadContacts(tid);
  } catch (e) { alert('Opslaan mislukt: ' + e.message); }
}

async function kbDeleteContact() {
  const tid = currentTenantId;
  const id  = document.getElementById('kbContactId').value;
  if (!id || !confirm('Contact verwijderen?')) return;
  try {
    await apiFetch(`/api/kb/${tid}/contacts/${id}`, { method: 'DELETE' });
    document.getElementById('kbContactModal').style.display = 'none';
    kbContactsCache = []; kbInvalidateCache();
    kbSelectedContactId = null;
    document.getElementById('kbContactDetail').style.display = 'none';
    await kbRefreshCounts(tid);
    kbLoadContacts(tid);
  } catch (e) { alert('Verwijderen mislukt: ' + e.message); }
}

// ---------------------------------------------------------------------------
// PASSWORDS
// ---------------------------------------------------------------------------
async function kbLoadPasswords(tid) {
  const wrap = document.getElementById('kbPasswordsTable');
  if (!wrap) return;
  if (!kbPasswordsCache.length) {
    wrap.innerHTML = '<p class="nb-empty">Laden…</p>';
    try { kbPasswordsCache = await apiFetch(`/api/kb/${tid}/passwords`); }
    catch (e) { wrap.innerHTML = `<p class="nb-empty">${esc(e.message)}</p>`; return; }
  }
  _renderPasswordsTable();
}

function _renderPasswordsTable() {
  const wrap = document.getElementById('kbPasswordsTable');
  const countEl = document.getElementById('nbPasswordListCount');
  if (countEl) countEl.textContent = `(${kbPasswordsCache.length})`;
  if (!kbPasswordsCache.length) { wrap.innerHTML = '<p class="nb-empty">Geen passwords gevonden.</p>'; return; }
  wrap.innerHTML = `<table><thead><tr><th>Naam</th><th>Categorie</th><th>Gebruikersnaam</th><th>Vault</th><th>Sterkte</th></tr></thead><tbody>
    ${kbPasswordsCache.map((p) => `<tr data-id="${p.id}" class="${kbSelectedPasswordId === p.id ? 'nb-row-selected' : ''}">
      <td class="nb-td-name">${esc(p.name)}</td><td>${_dash(p.category)}</td><td>${_dash(p.username)}</td><td>${_dash(p.secret_ref)}</td><td>${_dash(p.strength)}</td>
    </tr>`).join('')}</tbody></table>`;
  wrap.querySelectorAll('tbody tr').forEach((row) => row.addEventListener('click', () => {
    const id = parseInt(row.dataset.id, 10);
    kbSelectedPasswordId = id;
    wrap.querySelectorAll('tr').forEach((r) => r.classList.toggle('nb-row-selected', parseInt(r.dataset.id, 10) === id));
    _showPasswordDetail(kbPasswordsCache.find((p) => p.id === id));
  }));
}

function _showPasswordDetail(p) {
  const panel = document.getElementById('kbPasswordDetail');
  if (!p || !panel) return;
  panel.style.display = 'block';
  panel.innerHTML = `<div class="nb-detail-header"><h3 class="nb-detail-title">${esc(p.name)}</h3><div class="nb-detail-actions">
    <button class="nb-btn nb-btn-ghost nb-btn-sm" onclick="kbOpenEditPassword(${p.id})">✏️ Bewerken</button>
  </div></div>
  <div class="nb-detail-grid">
    <div class="nb-detail-field"><span class="nb-detail-field-label">Categorie</span><span class="nb-detail-field-value">${_dash(p.category)}</span></div>
    <div class="nb-detail-field"><span class="nb-detail-field-label">Gebruikersnaam</span><span class="nb-detail-field-value">${_dash(p.username)}</span></div>
    <div class="nb-detail-field"><span class="nb-detail-field-label">Vault</span><span class="nb-detail-field-value">${_dash(p.secret_ref)}</span></div>
    <div class="nb-detail-field"><span class="nb-detail-field-label">Sterkte</span><span class="nb-detail-field-value">${_dash(p.strength)}</span></div>
    <div class="nb-detail-field"><span class="nb-detail-field-label">Laatst bijgewerkt</span><span class="nb-detail-field-value">${_dash(p.last_updated)}</span></div>
  </div>${p.notes ? `<div class="nb-detail-notes"><strong>Notities</strong><br>${esc(p.notes)}</div>` : ''}`;
}

function kbOpenAddPassword() { _resetPasswordModal('Password toevoegen', null); document.getElementById('kbPasswordModal').style.display = 'flex'; }
window.kbOpenAddPassword = kbOpenAddPassword;
function kbOpenEditPassword(id) { const p = kbPasswordsCache.find((x) => x.id === id); if (!p) return; _resetPasswordModal('Password bewerken', p); document.getElementById('kbPasswordModal').style.display = 'flex'; }
window.kbOpenEditPassword = kbOpenEditPassword;
function _resetPasswordModal(title, p) {
  document.getElementById('kbPasswordModalTitle').textContent = title;
  document.getElementById('kbPasswordId').value = p?.id ?? '';
  document.getElementById('kbPasswordName').value = p?.name ?? '';
  document.getElementById('kbPasswordCategory').value = p?.category ?? '';
  document.getElementById('kbPasswordUsername').value = p?.username ?? '';
  document.getElementById('kbPasswordSecretRef').value = p?.secret_ref ?? '';
  document.getElementById('kbPasswordStrength').value = p?.strength ?? '';
  document.getElementById('kbPasswordLastUpdated').value = p?.last_updated ?? '';
  document.getElementById('kbPasswordNotes').value = p?.notes ?? '';
  document.getElementById('kbDeletePasswordBtn').style.display = p ? '' : 'none';
}
async function kbSavePassword() {
  const tid = currentTenantId; const id = document.getElementById('kbPasswordId').value;
  const payload = {
    name: document.getElementById('kbPasswordName').value.trim(),
    category: document.getElementById('kbPasswordCategory').value.trim() || null,
    username: document.getElementById('kbPasswordUsername').value.trim() || null,
    secret_ref: document.getElementById('kbPasswordSecretRef').value.trim() || null,
    strength: document.getElementById('kbPasswordStrength').value || 0,
    last_updated: document.getElementById('kbPasswordLastUpdated').value || null,
    notes: document.getElementById('kbPasswordNotes').value.trim() || null,
  };
  if (!payload.name) { alert('Naam is verplicht.'); return; }
  try {
    if (id) await apiFetch(`/api/kb/${tid}/passwords/${id}`, { method: 'PUT', body: JSON.stringify(payload) });
    else await apiFetch(`/api/kb/${tid}/passwords`, { method: 'POST', body: JSON.stringify(payload) });
    document.getElementById('kbPasswordModal').style.display = 'none'; kbPasswordsCache = []; await kbRefreshCounts(tid, { force: true }); kbLoadPasswords(tid);
  } catch (e) { alert('Opslaan mislukt: ' + e.message); }
}
async function kbDeletePassword() {
  const tid = currentTenantId; const id = document.getElementById('kbPasswordId').value;
  if (!id || !confirm('Password verwijderen?')) return;
  try { await apiFetch(`/api/kb/${tid}/passwords/${id}`, { method: 'DELETE' }); document.getElementById('kbPasswordModal').style.display = 'none'; kbPasswordsCache = []; await kbRefreshCounts(tid, { force: true }); kbLoadPasswords(tid); } catch (e) { alert('Verwijderen mislukt: ' + e.message); }
}

// ---------------------------------------------------------------------------
// SOFTWARE
// ---------------------------------------------------------------------------
async function kbLoadSoftware(tid) {
  const wrap = document.getElementById('kbSoftwareTable');
  if (!wrap) return;
  if (!kbSoftwareCache.length) {
    wrap.innerHTML = '<p class="nb-empty">Laden…</p>';
    try { kbSoftwareCache = await apiFetch(`/api/kb/${tid}/software`); }
    catch (e) { wrap.innerHTML = `<p class="nb-empty">${esc(e.message)}</p>`; return; }
  }
  _renderSoftwareTable();
}
function _renderSoftwareTable() {
  const wrap = document.getElementById('kbSoftwareTable');
  const countEl = document.getElementById('nbSoftwareListCount');
  const items = _softwareItems().filter((item) => (
    kbSoftwareSourceFilter === 'all' ||
    (kbSoftwareSourceFilter === 'manual' && item._source === 'Handmatig') ||
    (kbSoftwareSourceFilter === 'assessment' && item._source === 'Assessment')
  ));
  if (countEl) countEl.textContent = `(${items.length})`;
  if (!items.length) { wrap.innerHTML = '<p class="nb-empty">Geen software gevonden.</p>'; return; }
  wrap.innerHTML = `<table><thead><tr><th>Naam</th><th>Bron</th><th>Vendor</th><th>Type</th><th>Aantal</th><th>Totaal</th><th>Expiry</th><th>Status</th></tr></thead><tbody>
    ${items.map((s) => `<tr data-key="${s._key}" class="${kbSelectedSoftwareId === s._key ? 'nb-row-selected' : ''}">
      <td class="nb-td-name">${esc(s.name)}</td><td>${esc(s._source)}</td><td>${_dash(s.vendor)}</td><td>${_dash(s.software_type)}</td><td>${_dash(s.licenses)}</td><td>${s.total_price != null ? _fmtCurrency(s.total_price) : _dash(s.cost)}</td><td>${_dash(s.expiry)}</td><td>${_dash(s.status)}</td>
    </tr>`).join('')}</tbody></table>`;
  wrap.querySelectorAll('tbody tr').forEach((row) => row.addEventListener('click', () => {
    const key = row.dataset.key;
    kbSelectedSoftwareId = key;
    wrap.querySelectorAll('tr').forEach((r) => r.classList.toggle('nb-row-selected', r.dataset.key === key));
    _showSoftwareDetail(items.find((s) => s._key === key));
  }));
}
function _showSoftwareDetail(s) {
  const panel = document.getElementById('kbSoftwareDetail'); if (!s || !panel) return;
  panel.style.display = 'block';
  panel.innerHTML = `<div class="nb-detail-header"><h3 class="nb-detail-title">${esc(s.name)}</h3><div class="nb-detail-actions">${s._readonly ? '<span class="nb-badge nb-badge-type">Automatisch uit assessment</span>' : `<button class="nb-btn nb-btn-ghost nb-btn-sm" onclick="kbOpenEditSoftware(${s.id})">✏️ Bewerken</button>`}</div></div>
  <div class="nb-detail-grid">
    <div class="nb-detail-field"><span class="nb-detail-field-label">Bron</span><span class="nb-detail-field-value">${esc(s._source || 'Handmatig')}</span></div>
    <div class="nb-detail-field"><span class="nb-detail-field-label">Vendor</span><span class="nb-detail-field-value">${_dash(s.vendor)}</span></div>
    <div class="nb-detail-field"><span class="nb-detail-field-label">Type</span><span class="nb-detail-field-value">${_dash(s.software_type)}</span></div>
    <div class="nb-detail-field"><span class="nb-detail-field-label">Aantal licenties</span><span class="nb-detail-field-value">${_dash(s.licenses)}</span></div>
    <div class="nb-detail-field"><span class="nb-detail-field-label">Bedrag per licentie</span><span class="nb-detail-field-value">${s.unit_price != null ? _fmtCurrency(s.unit_price) : '—'}</span></div>
    <div class="nb-detail-field"><span class="nb-detail-field-label">Totaal bedrag</span><span class="nb-detail-field-value">${s.total_price != null ? _fmtCurrency(s.total_price) : _dash(s.cost)}</span></div>
    <div class="nb-detail-field"><span class="nb-detail-field-label">Expiry</span><span class="nb-detail-field-value">${_dash(s.expiry)}</span></div>
    <div class="nb-detail-field"><span class="nb-detail-field-label">Referentie</span><span class="nb-detail-field-value">${_dash(s.ref)}</span></div>
    ${s._readonly ? `<div class="nb-detail-field"><span class="nb-detail-field-label">Secret status</span><span class="nb-detail-field-value">${_dash(s.assessment_secret_status)}</span></div>
    <div class="nb-detail-field"><span class="nb-detail-field-label">Certificaat status</span><span class="nb-detail-field-value">${_dash(s.assessment_cert_status)}</span></div>` : ''}
  </div>${s.notes ? `<div class="nb-detail-notes"><strong>Notities</strong><br>${esc(s.notes)}</div>` : ''}`;
}
function kbOpenAddSoftware() { _resetSoftwareModal('Software toevoegen', null); document.getElementById('kbSoftwareModal').style.display = 'flex'; }
window.kbOpenAddSoftware = kbOpenAddSoftware;
function kbOpenEditSoftware(id) { const s = kbSoftwareCache.find((x) => x.id === id); if (!s) return; _resetSoftwareModal('Software bewerken', s); document.getElementById('kbSoftwareModal').style.display = 'flex'; }
window.kbOpenEditSoftware = kbOpenEditSoftware;
function _resetSoftwareModal(title, s) {
  document.getElementById('kbSoftwareModalTitle').textContent = title;
  document.getElementById('kbSoftwareId').value = s?.id ?? '';
  document.getElementById('kbSoftwareName').value = s?.name ?? '';
  document.getElementById('kbSoftwareVendor').value = s?.vendor ?? '';
  document.getElementById('kbSoftwareType').value = s?.software_type ?? '';
  document.getElementById('kbSoftwareLicenses').value = s?.licenses ?? '';
  document.getElementById('kbSoftwareUnitPrice').value = s?.unit_price ?? '';
  document.getElementById('kbSoftwareTotalPrice').value = s?.total_price != null ? _fmtCurrency(s.total_price) : '';
  document.getElementById('kbSoftwareCost').value = s?.cost ?? '';
  document.getElementById('kbSoftwareExpiry').value = s?.expiry ?? '';
  document.getElementById('kbSoftwareStatus').value = s?.status ?? 'active';
  document.getElementById('kbSoftwareRef').value = s?.ref ?? '';
  document.getElementById('kbSoftwareNotes').value = s?.notes ?? '';
  document.getElementById('kbDeleteSoftwareBtn').style.display = s ? '' : 'none';
  kbSyncSoftwareTotal();
}

function kbSyncSoftwareTotal() {
  const licenses = _toNumber(document.getElementById('kbSoftwareLicenses')?.value);
  const unitPrice = _toNumber(document.getElementById('kbSoftwareUnitPrice')?.value);
  const totalInput = document.getElementById('kbSoftwareTotalPrice');
  if (!totalInput) return;
  if (licenses == null || unitPrice == null) {
    totalInput.value = '';
    return;
  }
  totalInput.value = _fmtCurrency(licenses * unitPrice);
}
async function kbSaveSoftware() {
  const tid = currentTenantId; const id = document.getElementById('kbSoftwareId').value;
  const licenses = _toNumber(document.getElementById('kbSoftwareLicenses').value);
  const unitPrice = _toNumber(document.getElementById('kbSoftwareUnitPrice').value);
  const totalPrice = (licenses != null && unitPrice != null) ? Number((licenses * unitPrice).toFixed(2)) : null;
  const payload = {
    name: document.getElementById('kbSoftwareName').value.trim(),
    vendor: document.getElementById('kbSoftwareVendor').value.trim() || null,
    software_type: document.getElementById('kbSoftwareType').value.trim() || null,
    licenses,
    unit_price: unitPrice,
    total_price: totalPrice,
    cost: document.getElementById('kbSoftwareCost').value.trim() || null,
    expiry: document.getElementById('kbSoftwareExpiry').value || null,
    status: document.getElementById('kbSoftwareStatus').value,
    ref: document.getElementById('kbSoftwareRef').value.trim() || null,
    notes: document.getElementById('kbSoftwareNotes').value.trim() || null,
  };
  if (!payload.name) { alert('Naam is verplicht.'); return; }
  try { if (id) await apiFetch(`/api/kb/${tid}/software/${id}`, { method: 'PUT', body: JSON.stringify(payload) }); else await apiFetch(`/api/kb/${tid}/software`, { method: 'POST', body: JSON.stringify(payload) }); document.getElementById('kbSoftwareModal').style.display = 'none'; kbSoftwareCache = []; await kbRefreshCounts(tid, { force: true }); kbLoadSoftware(tid); } catch (e) { alert('Opslaan mislukt: ' + e.message); }
}
async function kbDeleteSoftware() {
  const tid = currentTenantId; const id = document.getElementById('kbSoftwareId').value;
  if (!id || !confirm('Software verwijderen?')) return;
  try { await apiFetch(`/api/kb/${tid}/software/${id}`, { method: 'DELETE' }); document.getElementById('kbSoftwareModal').style.display = 'none'; kbSoftwareCache = []; await kbRefreshCounts(tid, { force: true }); kbLoadSoftware(tid); } catch (e) { alert('Verwijderen mislukt: ' + e.message); }
}

// ---------------------------------------------------------------------------
// DOMAINS
// ---------------------------------------------------------------------------
async function kbLoadDomains(tid) {
  const wrap = document.getElementById('kbDomainsTable');
  if (!wrap) return;
  if (!kbDomainsCache.length) {
    wrap.innerHTML = '<p class="nb-empty">Laden…</p>';
    try { kbDomainsCache = await apiFetch(`/api/kb/${tid}/domains`); }
    catch (e) { wrap.innerHTML = `<p class="nb-empty">${esc(e.message)}</p>`; return; }
  }
  _renderDomainsTable();
}
function _renderDomainsTable() {
  const wrap = document.getElementById('kbDomainsTable');
  const countEl = document.getElementById('nbDomainListCount');
  if (countEl) countEl.textContent = `(${kbDomainsCache.length})`;
  if (!kbDomainsCache.length) { wrap.innerHTML = '<p class="nb-empty">Geen domeinen gevonden.</p>'; return; }
  wrap.innerHTML = `<table><thead><tr><th>Domein</th><th>Type</th><th>Registrar</th><th>Expiry</th><th>SSL</th><th>Status</th></tr></thead><tbody>
    ${kbDomainsCache.map((d) => `<tr data-id="${d.id}" class="${kbSelectedDomainId === d.id ? 'nb-row-selected' : ''}">
      <td class="nb-td-name">${esc(d.domain)}</td><td>${_dash(d.domain_type)}</td><td>${_dash(d.registrar)}</td><td>${_dash(d.expiry)}</td><td>${_dash(d.ssl_expiry)}</td><td>${_dash(d.status)}</td>
    </tr>`).join('')}</tbody></table>`;
  wrap.querySelectorAll('tbody tr').forEach((row) => row.addEventListener('click', () => {
    const id = parseInt(row.dataset.id, 10);
    kbSelectedDomainId = id;
    wrap.querySelectorAll('tr').forEach((r) => r.classList.toggle('nb-row-selected', parseInt(r.dataset.id, 10) === id));
    _showDomainDetail(kbDomainsCache.find((d) => d.id === id));
  }));
}
function _showDomainDetail(d) {
  const panel = document.getElementById('kbDomainDetail'); if (!d || !panel) return;
  panel.style.display = 'block';
  panel.innerHTML = `<div class="nb-detail-header"><h3 class="nb-detail-title">${esc(d.domain)}</h3><div class="nb-detail-actions"><button class="nb-btn nb-btn-ghost nb-btn-sm" onclick="kbOpenEditDomain(${d.id})">✏️ Bewerken</button></div></div>
  <div class="nb-detail-grid">
    <div class="nb-detail-field"><span class="nb-detail-field-label">Type</span><span class="nb-detail-field-value">${_dash(d.domain_type)}</span></div>
    <div class="nb-detail-field"><span class="nb-detail-field-label">Registrar</span><span class="nb-detail-field-value">${_dash(d.registrar)}</span></div>
    <div class="nb-detail-field"><span class="nb-detail-field-label">Expiry</span><span class="nb-detail-field-value">${_dash(d.expiry)}</span></div>
    <div class="nb-detail-field"><span class="nb-detail-field-label">SSL expiry</span><span class="nb-detail-field-value">${_dash(d.ssl_expiry)}</span></div>
    <div class="nb-detail-field"><span class="nb-detail-field-label">SSL issuer</span><span class="nb-detail-field-value">${_dash(d.ssl_issuer)}</span></div>
    <div class="nb-detail-field"><span class="nb-detail-field-label">Nameservers</span><span class="nb-detail-field-value">${_dash(d.nameservers)}</span></div>
  </div>${d.notes ? `<div class="nb-detail-notes"><strong>Notities</strong><br>${esc(d.notes)}</div>` : ''}`;
}
function kbOpenAddDomain() { _resetDomainModal('Domein toevoegen', null); document.getElementById('kbDomainModal').style.display = 'flex'; }
window.kbOpenAddDomain = kbOpenAddDomain;
function kbOpenEditDomain(id) { const d = kbDomainsCache.find((x) => x.id === id); if (!d) return; _resetDomainModal('Domein bewerken', d); document.getElementById('kbDomainModal').style.display = 'flex'; }
window.kbOpenEditDomain = kbOpenEditDomain;
function _resetDomainModal(title, d) {
  document.getElementById('kbDomainModalTitle').textContent = title;
  document.getElementById('kbDomainId').value = d?.id ?? '';
  document.getElementById('kbDomainName').value = d?.domain ?? '';
  document.getElementById('kbDomainType').value = d?.domain_type ?? '';
  document.getElementById('kbDomainRegistrar').value = d?.registrar ?? '';
  document.getElementById('kbDomainExpiry').value = d?.expiry ?? '';
  document.getElementById('kbDomainSslExpiry').value = d?.ssl_expiry ?? '';
  document.getElementById('kbDomainSslIssuer').value = d?.ssl_issuer ?? '';
  document.getElementById('kbDomainStatus').value = d?.status ?? 'active';
  document.getElementById('kbDomainAutoRenew').checked = !!d?.auto_renew;
  document.getElementById('kbDomainNameservers').value = d?.nameservers ?? '';
  document.getElementById('kbDomainNotes').value = d?.notes ?? '';
  document.getElementById('kbDeleteDomainBtn').style.display = d ? '' : 'none';
}
async function kbSaveDomain() {
  const tid = currentTenantId; const id = document.getElementById('kbDomainId').value;
  const payload = {
    domain: document.getElementById('kbDomainName').value.trim(),
    domain_type: document.getElementById('kbDomainType').value.trim() || null,
    registrar: document.getElementById('kbDomainRegistrar').value.trim() || null,
    expiry: document.getElementById('kbDomainExpiry').value || null,
    ssl_expiry: document.getElementById('kbDomainSslExpiry').value || null,
    ssl_issuer: document.getElementById('kbDomainSslIssuer').value.trim() || null,
    status: document.getElementById('kbDomainStatus').value,
    auto_renew: document.getElementById('kbDomainAutoRenew').checked ? 1 : 0,
    nameservers: document.getElementById('kbDomainNameservers').value.trim() || null,
    notes: document.getElementById('kbDomainNotes').value.trim() || null,
  };
  if (!payload.domain) { alert('Domein is verplicht.'); return; }
  try { if (id) await apiFetch(`/api/kb/${tid}/domains/${id}`, { method: 'PUT', body: JSON.stringify(payload) }); else await apiFetch(`/api/kb/${tid}/domains`, { method: 'POST', body: JSON.stringify(payload) }); document.getElementById('kbDomainModal').style.display = 'none'; kbDomainsCache = []; await kbRefreshCounts(tid, { force: true }); kbLoadDomains(tid); } catch (e) { alert('Opslaan mislukt: ' + e.message); }
}
async function kbDeleteDomain() {
  const tid = currentTenantId; const id = document.getElementById('kbDomainId').value;
  if (!id || !confirm('Domein verwijderen?')) return;
  try { await apiFetch(`/api/kb/${tid}/domains/${id}`, { method: 'DELETE' }); document.getElementById('kbDomainModal').style.display = 'none'; kbDomainsCache = []; await kbRefreshCounts(tid, { force: true }); kbLoadDomains(tid); } catch (e) { alert('Verwijderen mislukt: ' + e.message); }
}

// ---------------------------------------------------------------------------
// M365
// ---------------------------------------------------------------------------
async function kbLoadM365(tid) {
  try { kbM365Cache = await apiFetch(`/api/kb/${tid}/m365`); } catch (_) {}
  const p = kbM365Cache || {};
  const set = (id, v) => { const el = document.getElementById(id); if (!el) return; if (el.type === 'checkbox') el.checked = !!v; else el.value = v ?? ''; };
  set('kbM365TenantName', p.tenant_name); set('kbM365TenantId', p.tenant_id); set('kbM365GlobalAdmin', p.global_admin);
  set('kbM365LicenseType', p.license_type); set('kbM365LicensesTotal', p.licenses_total); set('kbM365LicensesUsed', p.licenses_used);
  set('kbM365Mfa', p.mfa); set('kbM365Mdm', p.mdm); set('kbM365AdConnect', p.ad_connect);
  set('kbM365SharedMailboxes', p.shared_mailboxes); set('kbM365GuestUsers', p.guest_users); set('kbM365Notes', p.notes);
  set('kbM365ConditionalAccess', p.conditional_access); set('kbM365Defender', p.defender); set('kbM365Purview', p.purview); set('kbM365Hybrid', p.hybrid); set('kbM365ExchangeHybrid', p.exchange_hybrid);
  _renderM365AssessmentInfo(p);
  _renderM365LicenseOverview(p.assessment_licenses || []);
}

function _renderM365AssessmentInfo(profile) {
  const el = document.getElementById('kbM365AssessmentInfo');
  if (!el) return;
  if (!profile?.assessment_generated_at) {
    el.style.display = 'none';
    el.innerHTML = '';
    return;
  }
  el.style.display = 'block';
  el.innerHTML = `Laatste assessment synchronisatie: <strong>${esc(profile.assessment_generated_at)}</strong>${profile.assessment_report_id ? ` · Run ${esc(profile.assessment_report_id)}` : ''}. Tenant- en licentievelden zijn hiermee automatisch aangevuld.`;
}

function _renderM365LicenseOverview(licenses) {
  const wrap = document.getElementById('kbM365LicenseOverview');
  if (!wrap) return;
  if (!licenses.length) {
    wrap.innerHTML = '<p class="nb-empty">Nog geen licentiegegevens uit assessment beschikbaar.</p>';
    return;
  }
  wrap.innerHTML = `<table><thead><tr><th>SKU</th><th>Totaal</th><th>Gebruikt</th><th>Beschikbaar</th><th>Benutting</th></tr></thead><tbody>
    ${licenses.map((lic) => `<tr>
      <td class="nb-td-name">${esc(lic.SkuPartNumber || lic.sku_part_number || 'Onbekend')}</td>
      <td>${_dash(lic.Total ?? lic.total)}</td>
      <td>${_dash(lic.Consumed ?? lic.consumed)}</td>
      <td>${_dash(lic.Available ?? lic.available)}</td>
      <td>${lic.Utilization != null || lic.utilization != null ? `${esc(String(lic.Utilization ?? lic.utilization))}%` : '—'}</td>
    </tr>`).join('')}
  </tbody></table>`;
}
async function kbSaveM365() {
  const tid = currentTenantId;
  const payload = {
    tenant_name: document.getElementById('kbM365TenantName').value.trim() || null,
    tenant_id: document.getElementById('kbM365TenantId').value.trim() || null,
    global_admin: document.getElementById('kbM365GlobalAdmin').value.trim() || null,
    license_type: document.getElementById('kbM365LicenseType').value.trim() || null,
    licenses_total: document.getElementById('kbM365LicensesTotal').value || null,
    licenses_used: document.getElementById('kbM365LicensesUsed').value || null,
    mfa: document.getElementById('kbM365Mfa').value.trim() || null,
    mdm: document.getElementById('kbM365Mdm').value.trim() || null,
    ad_connect: document.getElementById('kbM365AdConnect').value.trim() || null,
    shared_mailboxes: document.getElementById('kbM365SharedMailboxes').value || 0,
    guest_users: document.getElementById('kbM365GuestUsers').value || 0,
    notes: document.getElementById('kbM365Notes').value.trim() || null,
    conditional_access: document.getElementById('kbM365ConditionalAccess').checked ? 1 : 0,
    defender: document.getElementById('kbM365Defender').checked ? 1 : 0,
    purview: document.getElementById('kbM365Purview').checked ? 1 : 0,
    hybrid: document.getElementById('kbM365Hybrid').checked ? 1 : 0,
    exchange_hybrid: document.getElementById('kbM365ExchangeHybrid').checked ? 1 : 0,
  };
  try { kbM365Cache = await apiFetch(`/api/kb/${tid}/m365`, { method: 'PUT', body: JSON.stringify(payload) }); alert('M365 profiel opgeslagen.'); } catch (e) { alert('Opslaan mislukt: ' + e.message); }
}

// ---------------------------------------------------------------------------
// CHANGELOG
// ---------------------------------------------------------------------------
async function kbLoadChangelog(tid) {
  const wrap = document.getElementById('kbChangelogTable');
  if (!wrap) return;
  if (!kbChangelogCache.length) {
    wrap.innerHTML = '<p class="nb-empty">Laden…</p>';
    try { kbChangelogCache = await apiFetch(`/api/kb/${tid}/changelog`); }
    catch (e) { wrap.innerHTML = `<p class="nb-empty">${esc(e.message)}</p>`; return; }
  }
  _renderChangelogTable();
}
function _renderChangelogTable() {
  const wrap = document.getElementById('kbChangelogTable');
  const countEl = document.getElementById('nbChangelogListCount');
  if (countEl) countEl.textContent = `(${kbChangelogCache.length})`;
  if (!kbChangelogCache.length) { wrap.innerHTML = '<p class="nb-empty">Geen wijzigingen gevonden.</p>'; return; }
  wrap.innerHTML = `<table><thead><tr><th>Datum</th><th>Gebruiker</th><th>Categorie</th><th>Actie</th><th>Ref</th></tr></thead><tbody>
    ${kbChangelogCache.map((c) => `<tr data-id="${c.id}" class="${kbSelectedChangelogId === c.id ? 'nb-row-selected' : ''}">
      <td>${esc(c.change_date)}</td><td>${_dash(c.user_name)}</td><td>${_dash(c.category)}</td><td class="nb-td-name">${esc(c.action)}</td><td>${_dash(c.ref)}</td>
    </tr>`).join('')}</tbody></table>`;
  wrap.querySelectorAll('tbody tr').forEach((row) => row.addEventListener('click', () => {
    const id = parseInt(row.dataset.id, 10); kbSelectedChangelogId = id; wrap.querySelectorAll('tr').forEach((r) => r.classList.toggle('nb-row-selected', parseInt(r.dataset.id, 10) === id)); _showChangelogDetail(kbChangelogCache.find((c) => c.id === id));
  }));
}
function _showChangelogDetail(c) {
  const panel = document.getElementById('kbChangelogDetail'); if (!c || !panel) return;
  panel.style.display = 'block';
  panel.innerHTML = `<div class="nb-detail-header"><h3 class="nb-detail-title">${esc(c.action)}</h3><div class="nb-detail-actions"><button class="nb-btn nb-btn-ghost nb-btn-sm" onclick="kbOpenEditChangelog(${c.id})">✏️ Bewerken</button></div></div>
  <div class="nb-detail-grid">
    <div class="nb-detail-field"><span class="nb-detail-field-label">Datum</span><span class="nb-detail-field-value">${_dash(c.change_date)}</span></div>
    <div class="nb-detail-field"><span class="nb-detail-field-label">Gebruiker</span><span class="nb-detail-field-value">${_dash(c.user_name)}</span></div>
    <div class="nb-detail-field"><span class="nb-detail-field-label">Categorie</span><span class="nb-detail-field-value">${_dash(c.category)}</span></div>
    <div class="nb-detail-field"><span class="nb-detail-field-label">Referentie</span><span class="nb-detail-field-value">${_dash(c.ref)}</span></div>
  </div>${c.notes ? `<div class="nb-detail-notes"><strong>Notities</strong><br>${esc(c.notes)}</div>` : ''}`;
}
function kbOpenAddChangelog() { _resetChangelogModal('Wijziging toevoegen', null); document.getElementById('kbChangelogModal').style.display = 'flex'; }
window.kbOpenAddChangelog = kbOpenAddChangelog;
function kbOpenEditChangelog(id) { const c = kbChangelogCache.find((x) => x.id === id); if (!c) return; _resetChangelogModal('Wijziging bewerken', c); document.getElementById('kbChangelogModal').style.display = 'flex'; }
window.kbOpenEditChangelog = kbOpenEditChangelog;
function _resetChangelogModal(title, c) {
  document.getElementById('kbChangelogModalTitle').textContent = title;
  document.getElementById('kbChangelogId').value = c?.id ?? '';
  document.getElementById('kbChangelogDate').value = c?.change_date ?? '';
  document.getElementById('kbChangelogUser').value = c?.user_name ?? '';
  document.getElementById('kbChangelogCategory').value = c?.category ?? '';
  document.getElementById('kbChangelogRef').value = c?.ref ?? '';
  document.getElementById('kbChangelogAction').value = c?.action ?? '';
  document.getElementById('kbChangelogNotes').value = c?.notes ?? '';
  document.getElementById('kbDeleteChangelogBtn').style.display = c ? '' : 'none';
}
async function kbSaveChangelog() {
  const tid = currentTenantId; const id = document.getElementById('kbChangelogId').value;
  const payload = {
    change_date: document.getElementById('kbChangelogDate').value || null,
    user_name: document.getElementById('kbChangelogUser').value.trim() || null,
    category: document.getElementById('kbChangelogCategory').value.trim() || null,
    ref: document.getElementById('kbChangelogRef').value.trim() || null,
    action: document.getElementById('kbChangelogAction').value.trim(),
    notes: document.getElementById('kbChangelogNotes').value.trim() || null,
  };
  if (!payload.change_date || !payload.action) { alert('Datum en actie zijn verplicht.'); return; }
  try { if (id) await apiFetch(`/api/kb/${tid}/changelog/${id}`, { method: 'PUT', body: JSON.stringify(payload) }); else await apiFetch(`/api/kb/${tid}/changelog`, { method: 'POST', body: JSON.stringify(payload) }); document.getElementById('kbChangelogModal').style.display = 'none'; kbChangelogCache = []; await kbRefreshCounts(tid, { force: true }); kbLoadChangelog(tid); } catch (e) { alert('Opslaan mislukt: ' + e.message); }
}
async function kbDeleteChangelog() {
  const tid = currentTenantId; const id = document.getElementById('kbChangelogId').value;
  if (!id || !confirm('Wijziging verwijderen?')) return;
  try { await apiFetch(`/api/kb/${tid}/changelog/${id}`, { method: 'DELETE' }); document.getElementById('kbChangelogModal').style.display = 'none'; kbChangelogCache = []; await kbRefreshCounts(tid, { force: true }); kbLoadChangelog(tid); } catch (e) { alert('Verwijderen mislukt: ' + e.message); }
}

// ---------------------------------------------------------------------------
// Event wiring
// ---------------------------------------------------------------------------
document.addEventListener('DOMContentLoaded', () => {
  // Sub-nav
  document.querySelectorAll('.nb-tab[data-kb-tab]').forEach((item) => {
    item.addEventListener('click', (e) => {
      e.preventDefault();
      kbSwitchTab(item.dataset.kbTab);
    });
  });

  // Assets
  document.getElementById('kbAddAssetBtn')?.addEventListener('click', kbOpenAddAsset);
  document.getElementById('kbSaveAssetBtn')?.addEventListener('click', kbSaveAsset);
  document.getElementById('kbCancelAssetBtn')?.addEventListener('click', () => { document.getElementById('kbAssetModal').style.display = 'none'; });
  document.getElementById('kbDeleteAssetBtn')?.addEventListener('click', kbDeleteAsset);
  document.getElementById('kbAssetSearch')?.addEventListener('input', _renderAssetsTable);
  document.getElementById('kbAssetTypeFilter')?.addEventListener('change', _renderAssetsTable);
  document.getElementById('kbAssetStatusFilter')?.addEventListener('change', _renderAssetsTable);
  document.getElementById('kbAssetType')?.addEventListener('change', _toggleAssetSwitchFields);

  // VLANs
  document.getElementById('kbAddVlanBtn')?.addEventListener('click', kbOpenAddVlan);
  document.getElementById('kbSaveVlanBtn')?.addEventListener('click', kbSaveVlan);
  document.getElementById('kbCancelVlanBtn')?.addEventListener('click', () => { document.getElementById('kbVlanModal').style.display = 'none'; });
  document.getElementById('kbDeleteVlanBtn')?.addEventListener('click', kbDeleteVlan);
  document.getElementById('kbVlanSearch')?.addEventListener('input', _renderVlansTable);
  document.getElementById('kbVlanPurposeFilter')?.addEventListener('change', _renderVlansTable);

  // Pages
  document.getElementById('kbAddPageBtn')?.addEventListener('click', kbOpenNewPage);
  document.getElementById('kbSavePageBtn')?.addEventListener('click', kbSavePage);
  document.getElementById('kbDeletePageBtn')?.addEventListener('click', kbDeletePage);
  document.getElementById('kbCancelPageBtn')?.addEventListener('click', () => {
    document.getElementById('kbPageEditor').style.display = 'none';
    kbEditingPageId = null;
    _renderPagesList();
  });
  document.getElementById('kbPageContent')?.addEventListener('input', kbUpdatePreview);

  // Contacts
  document.getElementById('kbAddContactBtn')?.addEventListener('click', kbOpenAddContact);
  document.getElementById('kbSaveContactBtn')?.addEventListener('click', kbSaveContact);
  document.getElementById('kbCancelContactBtn')?.addEventListener('click', () => { document.getElementById('kbContactModal').style.display = 'none'; });
  document.getElementById('kbDeleteContactBtn')?.addEventListener('click', kbDeleteContact);
  document.getElementById('kbContactSourceFilter')?.addEventListener('change', (e) => {
    kbContactSourceFilter = e.target.value || 'all';
    _renderContactsTable();
  });

  // Passwords
  document.getElementById('kbAddPasswordBtn')?.addEventListener('click', kbOpenAddPassword);
  document.getElementById('kbSavePasswordBtn')?.addEventListener('click', kbSavePassword);
  document.getElementById('kbCancelPasswordBtn')?.addEventListener('click', () => { document.getElementById('kbPasswordModal').style.display = 'none'; });
  document.getElementById('kbDeletePasswordBtn')?.addEventListener('click', kbDeletePassword);

  // Software
  document.getElementById('kbAddSoftwareBtn')?.addEventListener('click', kbOpenAddSoftware);
  document.getElementById('kbSaveSoftwareBtn')?.addEventListener('click', kbSaveSoftware);
  document.getElementById('kbCancelSoftwareBtn')?.addEventListener('click', () => { document.getElementById('kbSoftwareModal').style.display = 'none'; });
  document.getElementById('kbDeleteSoftwareBtn')?.addEventListener('click', kbDeleteSoftware);
  document.getElementById('kbSoftwareSourceFilter')?.addEventListener('change', (e) => {
    kbSoftwareSourceFilter = e.target.value || 'all';
    _renderSoftwareTable();
  });
  document.getElementById('kbSoftwareLicenses')?.addEventListener('input', kbSyncSoftwareTotal);
  document.getElementById('kbSoftwareUnitPrice')?.addEventListener('input', kbSyncSoftwareTotal);

  // Domains
  document.getElementById('kbAddDomainBtn')?.addEventListener('click', kbOpenAddDomain);
  document.getElementById('kbSaveDomainBtn')?.addEventListener('click', kbSaveDomain);
  document.getElementById('kbCancelDomainBtn')?.addEventListener('click', () => { document.getElementById('kbDomainModal').style.display = 'none'; });
  document.getElementById('kbDeleteDomainBtn')?.addEventListener('click', kbDeleteDomain);

  // M365
  document.getElementById('kbSaveM365Btn')?.addEventListener('click', kbSaveM365);

  // Changelog
  document.getElementById('kbAddChangelogBtn')?.addEventListener('click', kbOpenAddChangelog);
  document.getElementById('kbSaveChangelogBtn')?.addEventListener('click', kbSaveChangelog);
  document.getElementById('kbCancelChangelogBtn')?.addEventListener('click', () => { document.getElementById('kbChangelogModal').style.display = 'none'; });
  document.getElementById('kbDeleteChangelogBtn')?.addEventListener('click', kbDeleteChangelog);

  // Close modals on backdrop click
  ['kbAssetModal', 'kbVlanModal', 'kbContactModal', 'kbPasswordModal', 'kbSoftwareModal', 'kbDomainModal', 'kbChangelogModal'].forEach((id) => {
    document.getElementById(id)?.addEventListener('click', (e) => {
      if (e.target === e.currentTarget) e.currentTarget.style.display = 'none';
    });
  });
});

// ---------------------------------------------------------------------------
// KB Settings (beheer categorieën, asset typen, VLAN doelen)
// ---------------------------------------------------------------------------

async function kbSettingsLoad() {
  const tid = currentTenantId;
  const infoEl = document.getElementById('kbSettingsTenantInfo');
  if (!tid) {
    if (infoEl) infoEl.textContent = 'Selecteer eerst een tenant in de topbar om kennisbank-instellingen te beheren.';
    ['kbSettingsAssetTypesList','kbSettingsCategoriesList','kbSettingsVlanPurposesList']
      .forEach(id => { const el = document.getElementById(id); if (el) el.innerHTML = '<span style="color:#aaa">Geen tenant geselecteerd.</span>'; });
    return;
  }
  if (infoEl) infoEl.textContent = `Instellingen voor tenant: ${tid}`;
  try {
    const [types, meta] = await Promise.all([
      apiFetch(`/api/kb/${tid}/asset-types`),
      apiFetch(`/api/kb/${tid}/meta`),
    ]);
    _kbSettingsRenderAssetTypes(types);
    _kbSettingsRenderCategories(meta.categories || []);
    _kbSettingsRenderVlanPurposes(meta.vlan_purposes || []);
  } catch (e) {
    console.error('KB settings load failed', e);
  }
}
window.kbSettingsLoad = kbSettingsLoad;

function _kbSettingsRenderAssetTypes(types) {
  const el = document.getElementById('kbSettingsAssetTypesList');
  if (!el) return;
  if (!types.length) { el.innerHTML = '<span class="kb-settings-empty">Nog geen typen.</span>'; return; }
  el.innerHTML = types.map(t =>
    `<span class="kb-settings-tag">${t.icon || ''} ${t.name}
      <button class="kb-settings-tag-del" title="Verwijderen" onclick="kbSettingsDeleteAssetType(${t.id})">×</button>
    </span>`
  ).join('');
}

function _kbSettingsRenderCategories(cats) {
  const el = document.getElementById('kbSettingsCategoriesList');
  if (!el) return;
  if (!cats.length) { el.innerHTML = '<span class="kb-settings-empty">Nog geen categorieën.</span>'; return; }
  el.innerHTML = cats.map((c, i) =>
    `<span class="kb-settings-tag">${c}
      <button class="kb-settings-tag-del" title="Verwijderen" onclick="kbSettingsDeleteCategory(${i})">×</button>
    </span>`
  ).join('');
}

function _kbSettingsRenderVlanPurposes(purposes) {
  const el = document.getElementById('kbSettingsVlanPurposesList');
  if (!el) return;
  if (!purposes.length) { el.innerHTML = '<span class="kb-settings-empty">Nog geen doelen.</span>'; return; }
  el.innerHTML = purposes.map((p, i) =>
    `<span class="kb-settings-tag"><code>${p.key}</code> — ${p.label}
      <button class="kb-settings-tag-del" title="Verwijderen" onclick="kbSettingsDeleteVlanPurpose(${i})">×</button>
    </span>`
  ).join('');
}

async function kbSettingsAddAssetType() {
  const tid = currentTenantId;
  if (!tid) return;
  const name = (document.getElementById('kbNewAssetTypeName').value || '').trim();
  const icon = (document.getElementById('kbNewAssetTypeIcon').value || '🖥️').trim();
  if (!name) { alert('Voer een naam in.'); return; }
  try {
    await apiFetch(`/api/kb/${tid}/asset-types`, { method: 'POST', body: JSON.stringify({ name, icon }) });
    document.getElementById('kbNewAssetTypeName').value = '';
    document.getElementById('kbNewAssetTypeIcon').value = '';
    const types = await apiFetch(`/api/kb/${tid}/asset-types`);
    _kbSettingsRenderAssetTypes(types);
    kbAssetTypes = types; // keep cache in sync
  } catch (e) { alert('Toevoegen mislukt: ' + e.message); }
}
window.kbSettingsAddAssetType = kbSettingsAddAssetType;

async function kbSettingsDeleteAssetType(id) {
  const tid = currentTenantId;
  if (!tid || !confirm('Type verwijderen?')) return;
  try {
    await apiFetch(`/api/kb/${tid}/asset-types/${id}`, { method: 'DELETE' });
    const types = await apiFetch(`/api/kb/${tid}/asset-types`);
    _kbSettingsRenderAssetTypes(types);
    kbAssetTypes = types;
  } catch (e) { alert('Verwijderen mislukt: ' + e.message); }
}
window.kbSettingsDeleteAssetType = kbSettingsDeleteAssetType;

async function _kbSettingsGetMeta() {
  return apiFetch(`/api/kb/${currentTenantId}/meta`);
}

async function _kbSettingsSaveMeta(patch) {
  const tid = currentTenantId;
  const current = await _kbSettingsGetMeta();
  return apiFetch(`/api/kb/${tid}/meta`, { method: 'PUT', body: JSON.stringify({ ...current, ...patch }) });
}

async function kbSettingsAddCategory() {
  const tid = currentTenantId;
  if (!tid) return;
  const val = (document.getElementById('kbNewCategory').value || '').trim().toLowerCase();
  if (!val) { alert('Voer een categorie in.'); return; }
  try {
    const meta = await _kbSettingsGetMeta();
    const cats = meta.categories || [];
    if (cats.includes(val)) { alert('Categorie bestaat al.'); return; }
    cats.push(val);
    await _kbSettingsSaveMeta({ categories: cats });
    document.getElementById('kbNewCategory').value = '';
    _kbSettingsRenderCategories(cats);
  } catch (e) { alert('Toevoegen mislukt: ' + e.message); }
}
window.kbSettingsAddCategory = kbSettingsAddCategory;

async function kbSettingsDeleteCategory(index) {
  const tid = currentTenantId;
  if (!tid || !confirm('Categorie verwijderen?')) return;
  try {
    const meta = await _kbSettingsGetMeta();
    const cats = (meta.categories || []).filter((_, i) => i !== index);
    await _kbSettingsSaveMeta({ categories: cats });
    _kbSettingsRenderCategories(cats);
  } catch (e) { alert('Verwijderen mislukt: ' + e.message); }
}
window.kbSettingsDeleteCategory = kbSettingsDeleteCategory;

async function kbSettingsAddVlanPurpose() {
  const tid = currentTenantId;
  if (!tid) return;
  const key   = (document.getElementById('kbNewVlanPurposeKey').value || '').trim().toLowerCase();
  const label = (document.getElementById('kbNewVlanPurposeLabel').value || '').trim();
  if (!key || !label) { alert('Voer zowel sleutel als label in.'); return; }
  try {
    const meta = await _kbSettingsGetMeta();
    const purposes = meta.vlan_purposes || [];
    if (purposes.find(p => p.key === key)) { alert('Doel bestaat al.'); return; }
    purposes.push({ key, label });
    await _kbSettingsSaveMeta({ vlan_purposes: purposes });
    document.getElementById('kbNewVlanPurposeKey').value = '';
    document.getElementById('kbNewVlanPurposeLabel').value = '';
    _kbSettingsRenderVlanPurposes(purposes);
  } catch (e) { alert('Toevoegen mislukt: ' + e.message); }
}
window.kbSettingsAddVlanPurpose = kbSettingsAddVlanPurpose;

async function kbSettingsDeleteVlanPurpose(index) {
  const tid = currentTenantId;
  if (!tid || !confirm('Doel verwijderen?')) return;
  try {
    const meta = await _kbSettingsGetMeta();
    const purposes = (meta.vlan_purposes || []).filter((_, i) => i !== index);
    await _kbSettingsSaveMeta({ vlan_purposes: purposes });
    _kbSettingsRenderVlanPurposes(purposes);
  } catch (e) { alert('Verwijderen mislukt: ' + e.message); }
}
window.kbSettingsDeleteVlanPurpose = kbSettingsDeleteVlanPurpose;

// Hook into showSection — toggle body scroll lock for KB
const _origShowSection = window.showSection;
window.showSection = function (sectionName) {
  _origShowSection(sectionName);
  document.body.classList.toggle('kb-active', sectionName === 'kb');
  if (sectionName === 'kb') kbInit();
};
