(function () {
  const LIVE_MODULES = {
    teams: {
      rootId: 'teamsModuleRoot',
      defaultTab: 'teams',
      tabs: {
        teams: {
          title: 'Teams',
          description: 'Alle Teams met zicht op leden, owners en zichtbaarheid.',
          endpoint: (tenantId) => `/api/collaboration/${tenantId}/teams`,
        },
        groepen: {
          title: 'Groepen',
          description: 'Alle M365-groepen, beveiligingsgroepen en distributielijsten met leden en owners.',
          endpoint: (tenantId) => `/api/collaboration/${tenantId}/groups`,
        },
      },
    },
    sharepoint: {
      rootId: 'sharepointModuleRoot',
      defaultTab: 'sharepoint-sites',
      tabs: {
        'sharepoint-sites': {
          title: 'SharePoint Sites',
          description: 'Live overzicht van sites, opslag en laatste wijzigingen.',
          endpoint: (tenantId) => `/api/collaboration/${tenantId}/sharepoint/sites`,
        },
        'sharepoint-settings': {
          title: 'SharePoint Settings',
          description: 'Tenant-brede sharinginstellingen en linkdefaults.',
          endpoint: (tenantId) => `/api/collaboration/${tenantId}/sharepoint/settings`,
        },
        'sharepoint-backup': {
          title: 'SharePoint Backup',
          description: 'Beschermde SharePoint-sites en policies binnen Microsoft 365 Backup.',
          endpoint: (tenantId) => `/api/backup/${tenantId}/sharepoint`,
        },
      },
    },
    identity: {
      rootId: 'identityModuleRoot',
      defaultTab: 'mfa',
      tabs: {
        mfa: {
          title: 'MFA',
          description: 'Registratiestatus en dekking van MFA binnen de tenant.',
          endpoint: (tenantId) => `/api/identity/${tenantId}/mfa`,
        },
        guests: {
          title: 'Guests',
          description: 'Gastaccounts, status en laatste gebruik.',
          endpoint: (tenantId) => `/api/identity/${tenantId}/guests`,
        },
        'admin-roles': {
          title: 'Admin Roles',
          description: 'Directory-rollen en hun leden live uitgelezen.',
          endpoint: (tenantId) => `/api/identity/${tenantId}/admin-roles`,
        },
        'security-defaults': {
          title: 'Security Defaults',
          description: 'Status van Security Defaults en relatie tot CA policies.',
          endpoint: (tenantId) => `/api/identity/${tenantId}/security-defaults`,
        },
        'legacy-auth': {
          title: 'Legacy Auth',
          description: 'Gebruikers met legacy-auth activiteit in recente sign-ins.',
          endpoint: (tenantId) => `/api/identity/${tenantId}/legacy-auth`,
        },
      },
    },
    apps: {
      rootId: 'appsModuleRoot',
      defaultTab: 'registrations',
      tabs: {
        registrations: {
          title: 'Registrations',
          description: 'App-registraties, secrets, certificaten en vervalstatus.',
          endpoint: (tenantId) => `/api/apps/${tenantId}/registrations`,
        },
      },
    },
    domains: {
      rootId: 'domainsModuleRoot',
      defaultTab: 'domains-list',
      tabs: {
        'domains-list': {
          title: 'Domeinen',
          description: 'Overzicht van alle tenantdomeinen die via Graph beschikbaar zijn.',
          endpoint: (tenantId) => `/api/domains/${tenantId}/list`,
        },
        'domains-analyse': {
          title: 'DNS Analyse',
          description: 'Analyseer SPF, DKIM, DMARC en MX voor een specifiek domein.',
          endpoint: (tenantId, inputs) => `/api/domains/${tenantId}/analyse?domain=${encodeURIComponent(inputs.domain || '')}`,
          input: {
            key: 'domain',
            label: 'Domein',
            placeholder: 'bijv. contoso.nl',
          },
        },
      },
    },
    exchange: {
      rootId: 'exchangeModuleRoot',
      defaultTab: 'mailboxen',
      tabs: {
        mailboxen: {
          title: 'Mailboxen',
          description: 'Overzicht van alle mailboxen in de tenant.',
          endpoint: (tenantId) => `/api/exchange/${tenantId}/mailboxes`,
        },
        forwarding: {
          title: 'Forwarding',
          description: 'Controleer actieve forwardings per mailbox.',
          endpoint: (tenantId) => `/api/exchange/${tenantId}/forwarding`,
        },
        regels: {
          title: 'Inbox Regels',
          description: 'Analyseer verdachte en normale inboxregels tenant-breed.',
          endpoint: (tenantId) => `/api/exchange/${tenantId}/mailbox-rules`,
        },
      },
    },
    intune: {
      rootId: 'intuneModuleRoot',
      defaultTab: 'overzicht',
      tabs: {
        overzicht: {
          title: 'Overzicht',
          description: 'Samenvatting van compliance en device posture in Intune.',
          endpoint: (tenantId) => `/api/intune/${tenantId}/summary`,
        },
        apparaten: {
          title: 'Apparaten',
          description: 'Live overzicht van managed devices in de tenant.',
          endpoint: (tenantId) => `/api/intune/${tenantId}/devices`,
        },
        compliance: {
          title: 'Compliance',
          description: 'Compliance policies en hun basisinformatie.',
          endpoint: (tenantId) => `/api/intune/${tenantId}/compliance`,
        },
        configuratie: {
          title: 'Configuratie',
          description: 'Configuratieprofielen en settings catalog items.',
          endpoint: (tenantId) => `/api/intune/${tenantId}/config`,
        },
        geschiedenis: {
          title: 'Geschiedenis',
          description: 'Historie van Intune-activiteiten in het portaal.',
          endpoint: (tenantId) => `/api/intune/${tenantId}/history`,
        },
      },
    },
    backup: {
      rootId: 'backupModuleRoot',
      defaultTab: 'overzicht',
      tabs: {
        overzicht: {
          title: 'Overzicht',
          description: 'Algemene backupstatus en beschermde resources per workload.',
          endpoint: (tenantId) => `/api/backup/${tenantId}/summary`,
        },
        onedrive: {
          title: 'OneDrive',
          description: 'OneDrive protection policies en beschermde drives.',
          endpoint: (tenantId) => `/api/backup/${tenantId}/onedrive`,
        },
        exchange: {
          title: 'Exchange',
          description: 'Exchange protection policies en beschermde mailboxen.',
          endpoint: (tenantId) => `/api/backup/${tenantId}/exchange`,
        },
        geschiedenis: {
          title: 'Geschiedenis',
          description: 'Historie van backupacties en statussen.',
          endpoint: (tenantId) => `/api/backup/${tenantId}/history`,
        },
      },
    },
    compliance: {
      rootId: 'complianceModuleRoot',
      defaultTab: 'cis',
      tabs: {
        cis: {
          title: 'CIS M365 Benchmark',
          description: 'CIS M365 Foundations Benchmark v3.0 — pass/fail per control met framework-mapping.',
          endpoint: (tenantId) => `/api/compliance/${tenantId}/cis`,
        },
        zerotrust: {
          title: 'Zero Trust Assessment',
          description: 'Microsoft Zero Trust Assessment — identiteiten, apparaten, netwerk en data getoetst aan SFI-pilaren.',
          endpoint: (tenantId) => `/api/compliance/${tenantId}/zerotrust`,
        },
      },
    },
    hybrid: {
      rootId: 'hybridModuleRoot',
      defaultTab: 'sync',
      tabs: {
        sync: {
          title: 'AD Connect Sync',
          description: 'Synchronisatiestatus, authenticatietype en domeinen voor hybrid-tenants.',
          endpoint: (tenantId) => `/api/hybrid/${tenantId}/sync`,
        },
      },
    },
    alerts: {
      rootId: 'alertsModuleRoot',
      defaultTab: 'auditlog',
      tabs: {
        auditlog: {
          title: 'Audit Log',
          description: 'Directory audit events en tenantwijzigingen.',
          endpoint: (tenantId) => `/api/alerts/${tenantId}/audit-logs`,
        },
        securescr: {
          title: 'Secure Score',
          description: 'Microsoft Secure Score met aanbevelingen.',
          endpoint: (tenantId) => `/api/alerts/${tenantId}/secure-score`,
        },
        signins: {
          title: 'Aanmeldingen',
          description: 'Recente sign-ins en risico-indicatoren.',
          endpoint: (tenantId) => `/api/alerts/${tenantId}/sign-ins`,
        },
        config: {
          title: 'Notificaties',
          description: 'Webhook- en e-mailinstellingen voor alerts beheren.',
          endpoint: (tenantId) => `/api/alerts/${tenantId}/config`,
          customType: 'alerts-config',
        },
      },
    },
  };

  const liveState = {
    section: null,
    tab: null,
  };
  const capabilityState = {
    byKey: {},
  };
  const appRegState = {
    items: [],
  };

  function getStoredToken() {
    return localStorage.getItem('denjoy_auth_token') || localStorage.getItem('denjoy_token') || '';
  }

  function liveEscapeHtml(value) {
    return String(value == null ? '' : value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function selectedTenantId() {
    try {
      if (typeof currentTenantId !== 'undefined' && currentTenantId) return currentTenantId;
    } catch (_) {}
    return document.getElementById('tenantSelect')?.value || null;
  }

  // TTL voor live-module tabs (ms) — endpoints die zelden wijzigen krijgen langere cache
  const LIVE_TTL = {
    '/api/apps/':          5 * 60 * 1000,
    '/api/identity/':      3 * 60 * 1000,
    '/api/collaboration/': 3 * 60 * 1000,
    '/api/domains/':       5 * 60 * 1000,
    '/api/ca/':            5 * 60 * 1000,
    '/api/exchange/':      2 * 60 * 1000,
    '/api/alerts/':        1 * 60 * 1000,
    '/api/intune/':        5 * 60 * 1000,
    '/api/backup/':        5 * 60 * 1000,
  };

  function liveTtlFor(path) {
    for (const [prefix, ttl] of Object.entries(LIVE_TTL)) {
      if (path.startsWith(prefix)) return ttl;
    }
    return 60 * 1000; // 1 min standaard
  }

  async function liveFetchJson(path, { skipCache = false } = {}) {
    if (!skipCache && window.cacheGet) {
      const hit = window.cacheGet(path);
      if (hit !== null) return hit;
    }
    const token = getStoredToken();
    const res = await fetch(path, {
      credentials: 'include',
      headers: {
        'Content-Type': 'application/json',
        ...(token ? { Authorization: `Bearer ${token}` } : {}),
      },
    });
    let data = null;
    try {
      data = await res.json();
    } catch (_) {
      data = null;
    }
    if (!res.ok) {
      throw new Error((data && data.error) || `HTTP ${res.status}`);
    }
    if (data !== null && window.cacheSet) {
      window.cacheSet(path, data, liveTtlFor(path));
    }
    return data;
  }

  async function liveApiRequest(path, options = {}) {
    const token = getStoredToken();
    const res = await fetch(path, {
      credentials: 'include',
      headers: {
        'Content-Type': 'application/json',
        ...(token ? { Authorization: `Bearer ${token}` } : {}),
        ...(options.headers || {}),
      },
      ...options,
    });
    let data = null;
    try {
      data = await res.json();
    } catch (_) {
      data = null;
    }
    if (!res.ok) {
      throw new Error((data && data.error) || `HTTP ${res.status}`);
    }
    return data;
  }

  function getModuleConfig(sectionName) {
    return LIVE_MODULES[sectionName] || null;
  }

  function getModuleRoot(sectionName) {
    const config = getModuleConfig(sectionName);
    if (!config) return null;
    return document.getElementById(config.rootId);
  }

  function getTabConfig(sectionName, tabKey) {
    const config = getModuleConfig(sectionName);
    return config?.tabs?.[tabKey] || null;
  }

  function normalizeScalar(value) {
    if (value == null || value === '') return '—';
    if (typeof value === 'boolean') return value ? 'Ja' : 'Nee';
    if (Array.isArray(value)) return value.length ? value.slice(0, 3).map(normalizeScalar).join(', ') : '—';
    if (typeof value === 'object') return JSON.stringify(value);
    return String(value);
  }

  function formatDate(value) {
    if (!value) return '—';
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return normalizeScalar(value);
    return date.toLocaleDateString('nl-NL');
  }

  function formatNumber(value, digits = 0) {
    const number = Number(value);
    if (!Number.isFinite(number)) return '—';
    return number.toLocaleString('nl-NL', {
      minimumFractionDigits: digits,
      maximumFractionDigits: digits,
    });
  }

  function formatCompactNumber(value) {
    const number = Number(value);
    if (!Number.isFinite(number)) return '—';
    return new Intl.NumberFormat('nl-NL', { notation: 'compact', maximumFractionDigits: 1 }).format(number);
  }

  function formatStorageGb(value) {
    const number = Number(value);
    if (!Number.isFinite(number)) return '—';
    return `${formatNumber(number, number >= 100 ? 0 : 2)} GB`;
  }

  function formatPercent(value, digits = 1) {
    const number = Number(value);
    if (!Number.isFinite(number)) return '—';
    return `${formatNumber(number, digits)}%`;
  }

  function buildProgressTone(percent) {
    if (!Number.isFinite(percent)) return 'ok';
    if (percent >= 100) return 'crit';
    if (percent >= 85) return 'warn';
    return 'ok';
  }

  function extractCollection(data) {
    const keys = ['users', 'guests', 'roles', 'policies', 'profiles', 'devices', 'domains', 'items', 'mailboxes', 'rules', 'forwarding', 'locations', 'sites', 'teams', 'apps'];
    for (const key of keys) {
      if (Array.isArray(data?.[key])) return { key, items: data[key] };
    }
    return null;
  }

  function renderSummary(data) {
    const keys = Object.keys(data || {}).filter((key) => {
      const value = data[key];
      return !Array.isArray(value) && (typeof value !== 'object' || value === null);
    }).slice(0, 8);
    if (!keys.length) return '';
    return `
      <div class="live-module-summary">
        ${keys.map((key) => `
          <article class="live-module-summary-card">
            <span>${liveEscapeHtml(key)}</span>
            <strong>${liveEscapeHtml(normalizeScalar(data[key]))}</strong>
          </article>
        `).join('')}
      </div>
    `;
  }

  function renderServiceOverview(sectionName, tabKey, data, collection) {
    const cards = [];
    const total = Number(data?.total || data?.count || data?.roleCount || (collection?.items?.length || 0));

    if (sectionName === 'teams' && tabKey === 'teams') {
      cards.push({ label: 'Teams', value: String(Number(data?.count || data?.teams?.length || 0) || 0), meta: 'werkruimtes' });
      cards.push({ label: 'Publiek', value: String(Number(data?.publicCount || 0) || 0), meta: 'zichtbaar' });
      const memberCount = (data?.teams || []).reduce((sum, item) => sum + (Number(item?.memberCount) || 0), 0);
      cards.push({ label: 'Leden', value: String(memberCount || 0), meta: 'geteld' });
      const guestCount = (data?.teams || []).reduce((sum, item) => sum + (Number(item?.guestCount) || 0), 0);
      cards.push({ label: 'Guests', value: String(guestCount || 0), meta: 'binnen teams' });
    } else if (sectionName === 'teams' && tabKey === 'groepen') {
      const stats = data?.stats || {};
      cards.push({ label: 'Groepen', value: String(Number(data?.count || data?.groups?.length || 0) || 0), meta: 'totaal' });
      cards.push({ label: 'M365', value: String(Number(stats.microsoft365 || 0) || 0), meta: 'Microsoft 365' });
      cards.push({ label: 'Security', value: String(Number(stats.security || 0) || 0), meta: 'beveiligingsgroepen' });
      cards.push({ label: 'Distributie', value: String(Number(stats.distribution || 0) || 0), meta: 'distributielijsten' });
    } else if (sectionName === 'sharepoint' && tabKey === 'sharepoint-sites') {
      cards.push({ label: 'Sites', value: String(Number(data?.count || data?.sites?.length || 0) || 0), meta: 'gevonden' });
      cards.push({ label: 'Storage', value: Number.isFinite(Number(data?.totalStorageUsedGB)) ? formatStorageGb(data.totalStorageUsedGB) : '—', meta: 'totaal gebruikt' });
      cards.push({ label: 'Inactief', value: String(Number(data?.inactiveSites || 0) || 0), meta: '> 90 dagen' });
      cards.push({ label: 'Quota', value: Number.isFinite(Number(data?.storageUsedPct)) ? formatPercent(data.storageUsedPct) : '—', meta: 'van capaciteit' });
    } else if (sectionName === 'sharepoint' && tabKey === 'sharepoint-settings') {
      cards.push({ label: 'Sharing', value: data?.sharingCapability ? String(data.sharingCapability) : '—', meta: 'tenantbreed' });
      cards.push({ label: 'Guest sharing', value: typeof data?.guestSharingEnabled === 'boolean' ? (data.guestSharingEnabled ? 'Ja' : 'Nee') : '—', meta: 'status' });
      cards.push({ label: 'Default link', value: data?.defaultSharingLinkType ? String(data.defaultSharingLinkType) : '—', meta: 'standaard' });
    } else if (sectionName === 'identity' && tabKey === 'mfa') {
      cards.push({ label: 'Gebruikers', value: String(Number(data?.total || data?.users?.length || 0) || 0), meta: 'geanalyseerd' });
      cards.push({ label: 'MFA', value: `${Number(data?.mfaPercentage || 0)}%`, meta: 'dekking' });
      cards.push({ label: 'Geregistreerd', value: String(Number(data?.mfaRegistered || 0) || 0), meta: 'accounts' });
    } else if (sectionName === 'identity' && tabKey === 'guests') {
      cards.push({ label: 'Guests', value: String(Number(data?.count || data?.guests?.length || 0) || 0), meta: 'accounts' });
      const enabled = (data?.guests || []).filter((item) => item?.accountEnabled !== false).length;
      cards.push({ label: 'Actief', value: String(enabled || 0), meta: 'enabled' });
    } else if (sectionName === 'identity' && tabKey === 'admin-roles') {
      cards.push({ label: 'Rollen', value: String(Number(data?.roleCount || data?.roles?.length || 0) || 0), meta: 'actief' });
      cards.push({ label: 'Admins', value: String(Number(data?.totalAdmins || 0) || 0), meta: 'uniek' });
    } else if (sectionName === 'identity' && tabKey === 'security-defaults') {
      cards.push({ label: 'Security Defaults', value: typeof data?.securityDefaultsEnabled === 'boolean' ? (data.securityDefaultsEnabled ? 'Aan' : 'Uit') : '—', meta: 'status' });
      cards.push({ label: 'CA Policies', value: String(Number(data?.caEnabledPolicies || 0) || 0), meta: 'enabled' });
    } else if (sectionName === 'identity' && tabKey === 'legacy-auth') {
      cards.push({ label: 'Gebruikers', value: String(Number(data?.affectedUsers || data?.users?.length || 0) || 0), meta: 'geraakt' });
      cards.push({ label: 'Periode', value: String(Number(data?.daysChecked || 30) || 30), meta: 'dagen' });
    } else if (sectionName === 'apps') {
      cards.push({ label: 'Apps', value: String(Number(data?.total || data?.apps?.length || 0) || 0), meta: 'registraties' });
      cards.push({ label: 'Expired', value: String(Number(data?.expired || 0) || 0), meta: 'kritiek' });
      cards.push({ label: 'Critical', value: String(Number(data?.critical || 0) || 0), meta: 'binnenkort' });
      cards.push({ label: 'Warning', value: String(Number(data?.warning || 0) || 0), meta: 'attentie' });
    } else if (sectionName === 'domains' && tabKey === 'domains-list') {
      cards.push({ label: 'Domeinen', value: String(Number(data?.count || data?.domains?.length || 0) || 0), meta: 'tenantbreed' });
      const initial = (data?.domains || []).filter((item) => item?.isInitial).length;
      cards.push({ label: 'OnMicrosoft', value: String(initial || 0), meta: 'initieel' });
    } else if (sectionName === 'domains' && tabKey === 'domains-analyse') {
      cards.push({ label: 'Score', value: Number.isFinite(Number(data?.score)) ? String(Number(data.score)) : '—', meta: `/ ${Number(data?.maxScore || 100) || 100}` });
      cards.push({ label: 'Label', value: data?.label ? String(data.label) : '—', meta: 'beoordeling' });
      const okChecks = (data?.checks || []).filter((item) => String(item?.status || '').toLowerCase() === 'ok').length;
      cards.push({ label: 'Checks OK', value: String(okChecks || 0), meta: 'DNS' });
    } else {
      if (Number.isFinite(total) && total > 0) {
        cards.push({ label: 'Totaal', value: String(total), meta: collection?.key || 'records' });
      }
      if (Number.isFinite(Number(data?.publicCount || 0)) && Number(data?.publicCount || 0) > 0) {
        cards.push({ label: 'Publiek', value: String(Number(data.publicCount)), meta: 'zichtbaar' });
      }
      if (Number.isFinite(Number(data?.currentScore ?? data?.score))) {
        const score = Number(data.currentScore ?? data.score);
        const max = Number(data?.maxScore || 100);
        cards.push({ label: 'Score', value: String(score), meta: `/ ${max}` });
      }
      if (Number.isFinite(Number(data?.guestCount || 0)) && Number(data?.guestCount || 0) >= 0 && (data?.guestCount || data?.guests)) {
        cards.push({ label: 'Guests', value: String(Number(data.guestCount || data.guests?.length || 0)), meta: 'accounts' });
      }
    }
    if (cards.length === 0 && collection?.items?.length) {
      cards.push({ label: 'Records', value: String(collection.items.length), meta: collection.key || 'items' });
    }
    if (!cards.length) return '';
    return `
      <div class="workspace-service-overview">
        ${cards.slice(0, 4).map((card) => `
          <article class="workspace-service-card">
            <span class="workspace-service-label">${liveEscapeHtml(card.label)}</span>
            <strong class="workspace-service-value">${liveEscapeHtml(card.value)}</strong>
            <span class="workspace-service-meta">${liveEscapeHtml(card.meta || '')}</span>
          </article>
        `).join('')}
      </div>
    `;
  }

  const _EMPTY_STATE_MSG = {
    'intune:overzicht':      'Geen Intune-overzicht beschikbaar. Controleer of de tenant Intune-licenties heeft.',
    'intune:apparaten':      'Geen apparaten gevonden. Voer een assessment uit of verifieer dat apparaten zijn ingeschreven.',
    'intune:compliance':     'Geen compliancebeleid gevonden voor deze tenant.',
    'intune:configuratie':   'Geen configuratieprofielen gevonden voor deze tenant.',
    'intune:geschiedenis':   'Geen Intune-historiedata beschikbaar.',
    'backup:overzicht':      'Geen Microsoft 365 Backup-data beschikbaar. Controleer of Backup actief is voor deze tenant.',
    'backup:onedrive':       'Geen OneDrive-backup instellingen gevonden.',
    'backup:exchange':       'Geen Exchange-backup instellingen gevonden.',
    'domains:domains-list':  'Geen domeinen gevonden. Voer een assessment uit om domeindata te laden.',
    'domains:domains-analyse':'Geen domeinanalyse beschikbaar. Ververs de domeinlijst eerst.',
    'exchange:mailboxen':    'Geen mailboxen gevonden. Voer een assessment uit of controleer de Exchange-verbinding.',
    'exchange:forwarding':   'Geen actieve e-mail forwarding gevonden — dit is een goed teken.',
    'exchange:regels':       'Geen inbox-regels gevonden of alle regels zijn normaal.',
    'alerts:auditlog':       'Geen auditloggebeurtenissen gevonden in de geselecteerde periode.',
    'alerts:securescr':      'Geen Secure Score data beschikbaar. Ververs of controleer de verbinding.',
    'alerts:signins':        'Geen aanmeldingsactiviteit gevonden in de geselecteerde periode.',
    'compliance:zerotrust':  'Zero Trust Assessment nog niet uitgevoerd. Gebruik de knop om een assessment te starten.',
  };

  function renderTable(collection, sectionName, tabKey) {
    if (!collection?.items?.length) {
      const key = sectionName && tabKey ? `${sectionName}:${tabKey}` : null;
      const msg = (key && _EMPTY_STATE_MSG[key]) || 'Geen records gevonden voor dit onderdeel.';
      return `<p class="live-module-empty">${liveEscapeHtml(msg)}</p>`;
    }
    const sample = collection.items.find((item) => item && typeof item === 'object') || {};
    const columns = Object.keys(sample).slice(0, 8);
    return `
      <div class="assessment-table-wrap">
        <table class="assessment-table">
          <thead>
            <tr>${columns.map((column) => `<th>${liveEscapeHtml(column)}</th>`).join('')}</tr>
          </thead>
          <tbody>
            ${collection.items.slice(0, 100).map((item) => `
              <tr>
                ${columns.map((column) => `<td>${liveEscapeHtml(normalizeScalar(item[column]))}</td>`).join('')}
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    `;
  }

  function renderTeamsBody(data) {
    const teams = (data?.teams || []).filter((item) => item && typeof item === 'object');
    if (!teams.length) return '<p class="live-module-empty">Geen Teams-data beschikbaar voor deze tenant.</p>';
    const guestTotal = teams.reduce((sum, item) => sum + (Number(item.guestCount) || 0), 0);
    const ownerTotal = teams.reduce((sum, item) => sum + (Number(item.ownerCount) || 0), 0);
    const privateCount = teams.filter((item) => String(item.visibility || '').toLowerCase() !== 'public').length;
    const dynamicCount = teams.filter((item) => item.isDynamic).length;
    return `
      <div class="live-insight-strip">
        <article class="live-insight-card">
          <span class="live-insight-label">Privaat</span>
          <strong>${liveEscapeHtml(formatNumber(privateCount))}</strong>
          <span class="live-insight-meta">teams zonder publieke zichtbaarheid</span>
        </article>
        <article class="live-insight-card">
          <span class="live-insight-label">Owners</span>
          <strong>${liveEscapeHtml(formatNumber(ownerTotal))}</strong>
          <span class="live-insight-meta">owners over alle teams</span>
        </article>
        <article class="live-insight-card">
          <span class="live-insight-label">Guests</span>
          <strong>${liveEscapeHtml(formatNumber(guestTotal))}</strong>
          <span class="live-insight-meta">gastleden binnen teams</span>
        </article>
        <article class="live-insight-card">
          <span class="live-insight-label">Dynamisch</span>
          <strong>${liveEscapeHtml(formatNumber(dynamicCount))}</strong>
          <span class="live-insight-meta">teams met membership rule</span>
        </article>
      </div>
      <div class="assessment-table-wrap live-entity-table-wrap">
        <table class="assessment-table live-entity-table">
          <thead>
            <tr>
              <th>Team</th>
              <th>Zichtbaarheid</th>
              <th>Leden</th>
              <th>Owners</th>
              <th>Gasten</th>
              <th>Aangemaakt</th>
            </tr>
          </thead>
          <tbody>
            ${teams.slice(0, 200).map((team) => {
              const visibility = String(team.visibility || 'Private');
              const visibilityClass = visibility.toLowerCase() === 'public' ? 'live-badge-warn' : 'live-badge-ok';
              return `
                <tr>
                  <td>
                    <div class="live-entity-main">
                      <strong>${liveEscapeHtml(team.displayName || team.mail || 'Onbekend team')}</strong>
                      <span>${liveEscapeHtml(team.mail || 'Geen teamadres')}</span>
                      ${team.description ? `<p>${liveEscapeHtml(team.description)}</p>` : ''}
                    </div>
                  </td>
                  <td>
                    <div class="live-pill-stack">
                      <span class="live-badge ${visibilityClass}">${liveEscapeHtml(visibility)}</span>
                      ${team.isDynamic ? '<span class="live-badge live-badge-info">Dynamisch</span>' : ''}
                    </div>
                  </td>
                  <td>${liveEscapeHtml(formatNumber(team.memberCount || 0))}</td>
                  <td>${liveEscapeHtml(formatNumber(team.ownerCount || 0))}</td>
                  <td>${liveEscapeHtml(formatNumber(team.guestCount || 0))}</td>
                  <td>${liveEscapeHtml(formatDate(team.createdAt))}</td>
                </tr>
              `;
            }).join('')}
          </tbody>
        </table>
      </div>
    `;
  }

  const GROUP_TYPE_LABELS = {
    Microsoft365:       'M365 Groep',
    Security:           'Beveiligingsgroep',
    Distribution:       'Distributielijst',
    MailEnabledSecurity:'Mail-beveiligd',
    Other:              'Overig',
  };
  const GROUP_TYPE_BADGE = {
    Microsoft365:       'live-badge-info',
    Security:           'live-badge-ok',
    Distribution:       'live-badge-warn',
    MailEnabledSecurity:'live-badge-warn',
    Other:              '',
  };

  function renderGroupsBody(data) {
    const groups = (data?.groups || []).filter((item) => item && typeof item === 'object');
    if (!groups.length) return '<p class="live-module-empty">Geen groepen beschikbaar voor deze tenant.</p>';
    const stats = data?.stats || {};
    const dynamicCount = groups.filter((g) => g.isDynamic).length;
    return `
      <div class="live-insight-strip">
        <article class="live-insight-card">
          <span class="live-insight-label">Microsoft 365</span>
          <strong>${liveEscapeHtml(formatNumber(stats.microsoft365 || 0))}</strong>
          <span class="live-insight-meta">unified groepen</span>
        </article>
        <article class="live-insight-card">
          <span class="live-insight-label">Security</span>
          <strong>${liveEscapeHtml(formatNumber(stats.security || 0))}</strong>
          <span class="live-insight-meta">beveiligingsgroepen</span>
        </article>
        <article class="live-insight-card">
          <span class="live-insight-label">Distributie</span>
          <strong>${liveEscapeHtml(formatNumber(stats.distribution || 0))}</strong>
          <span class="live-insight-meta">distributielijsten</span>
        </article>
        <article class="live-insight-card">
          <span class="live-insight-label">Dynamisch</span>
          <strong>${liveEscapeHtml(formatNumber(dynamicCount))}</strong>
          <span class="live-insight-meta">membership-regel</span>
        </article>
      </div>
      <div class="assessment-table-wrap live-entity-table-wrap">
        <table class="assessment-table live-entity-table">
          <thead>
            <tr>
              <th>Groep</th>
              <th>Type</th>
              <th>Leden</th>
              <th>Owners</th>
              <th>Gasten</th>
              <th>Aangemaakt</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            ${groups.slice(0, 300).map((g) => {
              const typeLabel = GROUP_TYPE_LABELS[g.groupType] || g.groupType || 'Onbekend';
              const typeBadge = GROUP_TYPE_BADGE[g.groupType] || '';
              return `
                <tr>
                  <td>
                    <div class="live-entity-main">
                      <strong>${liveEscapeHtml(g.displayName || 'Onbekende groep')}</strong>
                      <span>${liveEscapeHtml(g.mail || 'Geen e-mailadres')}</span>
                      ${g.description ? `<p>${liveEscapeHtml(g.description)}</p>` : ''}
                    </div>
                  </td>
                  <td>
                    <div class="live-pill-stack">
                      <span class="live-badge ${typeBadge}">${liveEscapeHtml(typeLabel)}</span>
                      ${g.isDynamic ? '<span class="live-badge live-badge-info">Dynamisch</span>' : ''}
                    </div>
                  </td>
                  <td>${liveEscapeHtml(formatNumber(g.memberCount || 0))}</td>
                  <td>${liveEscapeHtml(formatNumber(g.ownerCount || 0))}</td>
                  <td>${liveEscapeHtml(formatNumber(g.guestCount || 0))}</td>
                  <td>${liveEscapeHtml(formatDate(g.createdAt))}</td>
                  <td>
                    <button class="live-detail-btn grp-detail-btn"
                      data-group-id="${liveEscapeHtml(g.id || '')}"
                      data-group-name="${liveEscapeHtml(g.displayName || '')}"
                      data-group-type="${liveEscapeHtml(typeLabel)}"
                      data-group-mail="${liveEscapeHtml(g.mail || '')}"
                      data-member-count="${liveEscapeHtml(String(g.memberCount || 0))}"
                      data-owner-count="${liveEscapeHtml(String(g.ownerCount || 0))}"
                      data-guest-count="${liveEscapeHtml(String(g.guestCount || 0))}"
                      data-is-dynamic="${g.isDynamic ? 'true' : 'false'}"
                      data-created-at="${liveEscapeHtml(g.createdAt || '')}"
                      data-description="${liveEscapeHtml(g.description || '')}">Detail</button>
                  </td>
                </tr>
              `;
            }).join('')}
          </tbody>
        </table>
      </div>
    `;
  }

  function renderSharePointSitesBody(data) {
    const sites = (data?.sites || []).filter((item) => item && typeof item === 'object');
    if (!sites.length) return '<p class="live-module-empty">Geen SharePoint-sites beschikbaar voor deze tenant.</p>';
    const usedPercent = Number(data?.storageUsedPct);
    const progressTone = buildProgressTone(usedPercent);
    const quotaPanel = Number.isFinite(usedPercent) ? `
      <div class="live-storage-panel">
        <div class="live-storage-panel-head">
          <div>
            <span class="live-storage-kicker">SharePoint Storage Quota</span>
            <h4>Capaciteit & gebruik</h4>
            <p>${liveEscapeHtml(data?.storageCapacityLabel || 'Quotaformule niet beschikbaar')}</p>
          </div>
          <div class="live-storage-percent live-storage-percent--${progressTone}">${liveEscapeHtml(formatPercent(usedPercent))}</div>
        </div>
        <div class="live-storage-grid">
          <article class="live-storage-card">
            <span>Capaciteit</span>
            <strong>${liveEscapeHtml(formatStorageGb(data?.totalCapacityGB))}</strong>
            <small>beschikbaar volgens tenantformule</small>
          </article>
          <article class="live-storage-card">
            <span>Gebruikt</span>
            <strong>${liveEscapeHtml(formatStorageGb(data?.totalStorageUsedGB))}</strong>
            <small>${liveEscapeHtml(formatPercent(usedPercent))} van capaciteit</small>
          </article>
          <article class="live-storage-card">
            <span>Beschikbaar</span>
            <strong>${liveEscapeHtml(formatStorageGb(data?.storageRemainingGB))}</strong>
            <small>${liveEscapeHtml(formatNumber(data?.licenseUnitsForQuota || 0))} licenties meegenomen</small>
          </article>
          <article class="live-storage-card">
            <span>Gemiddeld per site</span>
            <strong>${liveEscapeHtml(formatStorageGb(data?.avgStoragePerSiteGB))}</strong>
            <small>${liveEscapeHtml(formatNumber(data?.sitesWithStorage || 0))} sites met data</small>
          </article>
        </div>
        <div class="live-storage-progress">
          <div class="live-storage-progress-bar">
            <span class="live-storage-progress-fill live-storage-progress-fill--${progressTone}" style="width:${Math.max(0, Math.min(100, usedPercent))}%"></span>
          </div>
          <div class="live-storage-progress-meta">
            <span>${liveEscapeHtml(formatStorageGb(data?.totalStorageUsedGB))} gebruikt</span>
            <span>${liveEscapeHtml(formatStorageGb(data?.totalCapacityGB))} totaal</span>
          </div>
        </div>
      </div>
    ` : '';

    return `
      ${quotaPanel}
      <div class="assessment-table-wrap live-entity-table-wrap">
        <table class="assessment-table live-entity-table">
          <thead>
            <tr>
              <th>Site</th>
              <th>Opslag</th>
              <th>Status</th>
              <th>Laatst gewijzigd</th>
              <th>Type</th>
            </tr>
          </thead>
          <tbody>
            ${sites.slice(0, 200).map((site) => {
              const statusText = String(site.status || 'Onbekend');
              const statusClass = statusText.toLowerCase() === 'inactief' ? 'live-badge-warn' : 'live-badge-ok';
              const storageLabel = site.storageLabel && site.storageLabel !== '—'
                ? String(site.storageLabel)
                : formatStorageGb(site.storageUsed);
              return `
                <tr>
                  <td>
                    <div class="live-entity-main">
                      <strong>${liveEscapeHtml(site.displayName || site.webUrl || 'Onbekende site')}</strong>
                      <span>${site.webUrl ? `<a href="${liveEscapeHtml(site.webUrl)}" target="_blank" rel="noopener noreferrer">${liveEscapeHtml(site.webUrl)}</a>` : 'Geen URL beschikbaar'}</span>
                    </div>
                  </td>
                  <td>${liveEscapeHtml(storageLabel || '—')}</td>
                  <td><span class="live-badge ${statusClass}">${liveEscapeHtml(statusText)}</span></td>
                  <td>${liveEscapeHtml(formatDate(site.lastModified || site.lastModifiedDateTime))}</td>
                  <td>${site.isRootSite ? '<span class="live-badge live-badge-info">Root site</span>' : '<span class="live-badge live-badge-neutral">Site</span>'}</td>
                </tr>
              `;
            }).join('')}
          </tbody>
        </table>
      </div>
    `;
  }

  function renderAppRegistrationsBody(data) {
    const apps = (data?.apps || data?.items || []).filter((item) => item && typeof item === 'object');
    appRegState.items = apps;
    if (!apps.length) return '<p class="live-module-empty">Geen app-registraties beschikbaar voor deze tenant.</p>';
    return `
      <div class="assessment-table-wrap live-entity-table-wrap">
        <table class="assessment-table live-entity-table">
          <thead>
            <tr>
              <th>Applicatie</th>
              <th>Status</th>
              <th>Secrets</th>
              <th>Certificaten</th>
              <th>Permissies</th>
              <th>Actie</th>
            </tr>
          </thead>
          <tbody>
            ${apps.slice(0, 200).map((app) => {
              const secretStatus = String(app.secretExpirationStatus || 'Onbekend');
              const certificateStatus = String(app.certificateExpirationStatus || 'Onbekend');
              const statusText = `${secretStatus} / ${certificateStatus}`.toLowerCase();
              const statusClass = statusText.includes('verlopen') || statusText.includes('expired')
                ? 'live-badge-crit'
                : (statusText.includes('14 dagen') || statusText.includes('critical') || statusText.includes('warning')
                  ? 'live-badge-warn'
                  : 'live-badge-ok');
              return `
                <tr>
                  <td>
                    <div class="live-entity-main">
                      <strong>${liveEscapeHtml(app.displayName || 'Onbekende app')}</strong>
                      <span>${liveEscapeHtml(app.appId || 'Geen appId')}</span>
                      ${app.hasEnterpriseApp ? '<p>Enterprise app aanwezig</p>' : '<p>Alleen app-registratie zichtbaar</p>'}
                    </div>
                  </td>
                  <td><span class="live-badge ${statusClass}">${liveEscapeHtml(secretStatus)}</span></td>
                  <td>${liveEscapeHtml(formatNumber(app.secretCount || 0))}</td>
                  <td>${liveEscapeHtml(formatNumber(app.certificateCount || 0))}</td>
                  <td>${liveEscapeHtml(formatNumber(app.permissionCount || 0))}</td>
                  <td><button type="button" class="live-module-refresh live-module-inline-btn" data-appreg-id="${liveEscapeHtml(app.id || app.appId || '')}">Details</button></td>
                </tr>
              `;
            }).join('')}
          </tbody>
        </table>
      </div>
    `;
  }

  async function openAppRegistrationModal(appId) {
    const tenantId = selectedTenantId();
    if (!tenantId || !appId) return;

    // Gebruik het Inzichten-paneel in plaats van een modale popup
    if (typeof window.openSideRailDetail === 'function') {
      window.openSideRailDetail('App registratie', 'Laden…');
    }

    try {
      const data = await liveFetchJson(`/api/apps/${tenantId}/registrations/${encodeURIComponent(appId)}`);
      const appTitle = data.displayName || data.appId || 'App registratie';
      const secrets = Array.isArray(data.secrets) ? data.secrets : [];
      const certs = Array.isArray(data.certs) ? data.certs : [];
      const redirects = Array.isArray(data.redirectUris) ? data.redirectUris : [];
      const identifiers = Array.isArray(data.identifierUris) ? data.identifierUris : [];
      const access = Array.isArray(data.requiredResourceAccess) ? data.requiredResourceAccess : [];
      const bodyHtml = `
          <div class="live-detail-grid">
            <div class="live-detail-card"><span>App ID</span><strong>${liveEscapeHtml(data.appId || '—')}</strong></div>
            <div class="live-detail-card"><span>Audience</span><strong>${liveEscapeHtml(data.signInAudience || '—')}</strong></div>
            <div class="live-detail-card"><span>Aangemaakt</span><strong>${liveEscapeHtml(formatDate(data.createdAt))}</strong></div>
            <div class="live-detail-card"><span>Enterprise app</span><strong>${data.hasEnterpriseApp ? 'Ja' : 'Nee'}</strong></div>
          </div>
          <div class="live-detail-section">
            <h5>Secrets</h5>
            ${secrets.length ? secrets.map((secret) => `<p>${liveEscapeHtml(secret.hint || secret.keyId || 'Secret')} · ${liveEscapeHtml(secret.statusLabel || '—')}</p>`).join('') : '<p class="live-module-empty">Geen secrets gevonden.</p>'}
          </div>
          <div class="live-detail-section">
            <h5>Certificaten</h5>
            ${certs.length ? certs.map((cert) => `<p>${liveEscapeHtml(cert.type || cert.keyId || 'Certificaat')} · ${liveEscapeHtml(cert.statusLabel || '—')}</p>`).join('') : '<p class="live-module-empty">Geen certificaten gevonden.</p>'}
          </div>
          <div class="live-detail-section">
            <h5>Redirect URI's</h5>
            ${redirects.length ? redirects.map((uri) => `<p>${liveEscapeHtml(uri)}</p>`).join('') : "<p class=\"live-module-empty\">Geen redirect URI's.</p>"}
          </div>
          <div class="live-detail-section">
            <h5>Identifier URI's</h5>
            ${identifiers.length ? identifiers.map((uri) => `<p>${liveEscapeHtml(uri)}</p>`).join('') : "<p class=\"live-module-empty\">Geen identifier URI's.</p>"}
          </div>
          <div class="live-detail-section">
            <h5>API-rechten</h5>
            ${(() => {
              const resolvedPerms = Array.isArray(data.permissions) ? data.permissions : [];
              if (resolvedPerms.length) {
                const grouped = {};
                resolvedPerms.forEach((p) => {
                  const res = p.Resource || p.resource || 'Onbekend';
                  if (!grouped[res]) grouped[res] = [];
                  grouped[res].push({ type: p.Type || p.type || '', name: p.Permission || p.permission || p.value || '?' });
                });
                return Object.entries(grouped).map(([res, perms]) =>
                  `<div style="margin-bottom:.5rem"><strong style="font-size:.8rem">${liveEscapeHtml(res)}</strong><ul style="margin:.2rem 0 0 1rem;padding:0;list-style:disc;font-size:.8rem">${perms.map((p) => `<li>${liveEscapeHtml(p.name)} <span style="color:var(--text-muted);font-size:.75rem">(${liveEscapeHtml(p.type)})</span></li>`).join('')}</ul></div>`
                ).join('');
              }
              if (access.length) {
                return access.map((item) => `<p>${liveEscapeHtml(item.resourceAppId || 'resource')} · ${liveEscapeHtml(String((item.resourceAccess || []).length || 0))} rechten</p>`).join('');
              }
              return '<p class="live-module-empty">Geen API-rechten gevonden.</p>';
            })()}
          </div>
      `;
      if (typeof window.updateSideRailDetail === 'function') {
        window.updateSideRailDetail(appTitle, bodyHtml);
      }
    } catch (error) {
      if (typeof window.updateSideRailDetail === 'function') {
        window.updateSideRailDetail('Fout', `<div class="live-module-error">${liveEscapeHtml(error.message || 'Details laden mislukt.')}</div>`);
      }
    }
  }

  function renderAlertsConfigForm(sectionName, tabKey, data) {
    const cfg = data?.config || {};
    return `
      <div class="live-module-config-card">
        <div class="live-module-config-title">Webhook & e-mail notificaties</div>
        <div class="live-module-form-group">
          <label for="alertsWebhookUrl">Webhook URL</label>
          <input type="url" id="alertsWebhookUrl" class="live-module-input" value="${liveEscapeHtml(cfg.webhook_url || '')}" placeholder="https://outlook.office.com/webhook/...">
        </div>
        <div class="live-module-form-group">
          <label for="alertsWebhookType">Webhook type</label>
          <select id="alertsWebhookType" class="live-module-select">
            <option value="teams"${cfg.webhook_type === 'teams' ? ' selected' : ''}>Microsoft Teams</option>
            <option value="slack"${cfg.webhook_type === 'slack' ? ' selected' : ''}>Slack</option>
            <option value="generic"${cfg.webhook_type === 'generic' ? ' selected' : ''}>Generiek JSON</option>
          </select>
        </div>
        <div class="live-module-form-group">
          <label for="alertsEmailAddr">E-mailadres</label>
          <input type="email" id="alertsEmailAddr" class="live-module-input" value="${liveEscapeHtml(cfg.email_addr || '')}" placeholder="alerts@bedrijf.nl">
        </div>
        <div class="live-module-action-row">
          <button type="button" class="live-module-refresh" data-alerts-action="save">Opslaan</button>
          <button type="button" class="live-module-refresh" data-alerts-action="test">Test webhook</button>
        </div>
        <div id="alertsConfigResult" class="live-module-config-result"></div>
      </div>
    `;
  }

  /* ── Custom renderers per identity tab ── */

  function renderMfaBody(data) {
    const users = data?.users || [];
    if (!users.length) {
      return '<p class="live-module-empty">Geen MFA-data beschikbaar. Controleer UserAuthenticationMethod.Read.All permissie.</p>';
    }
    const adminNoMfa = users.filter((u) => u.isAdmin && !u.isMfaRegistered).length;
    const alertRow = adminNoMfa > 0
      ? `<div class="live-module-alert-row">Let op: ${adminNoMfa} admin${adminNoMfa > 1 ? 's' : ''} zonder MFA-registratie</div>`
      : '';
    return alertRow + `
      <div class="assessment-table-wrap">
        <table class="assessment-table">
          <thead>
            <tr>
              <th>Gebruiker</th>
              <th>UPN</th>
              <th>MFA</th>
              <th>Standaard methode</th>
              <th>Passwordless</th>
              <th>Admin</th>
            </tr>
          </thead>
          <tbody>
            ${users.slice(0, 200).map((u) => `
              <tr>
                <td>${liveEscapeHtml(u.displayName || '—')}</td>
                <td class="live-small">${liveEscapeHtml(u.upn || '—')}</td>
                <td><span class="live-badge ${u.isMfaRegistered ? 'live-badge-ok' : 'live-badge-warn'}">${u.isMfaRegistered ? 'Ja' : 'Nee'}</span></td>
                <td class="live-small">${liveEscapeHtml(u.defaultMfaMethod && u.defaultMfaMethod !== 'none' ? u.defaultMfaMethod : (u.methodsRegistered?.join(', ') || '—'))}</td>
                <td><span class="live-badge ${u.isPasswordless ? 'live-badge-ok' : 'live-badge-neutral'}">${u.isPasswordless ? 'Ja' : 'Nee'}</span></td>
                <td>${u.isAdmin ? '<span class="live-badge live-badge-info">Admin</span>' : ''}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    `;
  }

  function renderGuestsBody(data) {
    const guests = data?.guests || [];
    if (!guests.length) return '<p class="live-module-empty">Geen gastaccounts gevonden in deze tenant.</p>';
    return `
      <div class="assessment-table-wrap">
        <table class="assessment-table">
          <thead>
            <tr>
              <th>Naam</th>
              <th>E-mail</th>
              <th>Account</th>
              <th>Uitnodiging</th>
              <th>Aangemaakt</th>
              <th>Laatste aanmelding</th>
            </tr>
          </thead>
          <tbody>
            ${guests.slice(0, 200).map((g) => {
              const created = g.createdAt ? new Date(g.createdAt).toLocaleDateString('nl-NL') : '—';
              const lastSignIn = g.lastSignIn ? new Date(g.lastSignIn).toLocaleDateString('nl-NL') : '—';
              return `
                <tr>
                  <td>${liveEscapeHtml(g.displayName || '—')}</td>
                  <td class="live-small">${liveEscapeHtml(g.mail || g.upn || '—')}</td>
                  <td><span class="live-badge ${g.accountEnabled ? 'live-badge-ok' : 'live-badge-warn'}">${g.accountEnabled ? 'Actief' : 'Uitgeschakeld'}</span></td>
                  <td class="live-small">${liveEscapeHtml(g.inviteStatus || '—')}</td>
                  <td>${liveEscapeHtml(created)}</td>
                  <td>${liveEscapeHtml(lastSignIn)}</td>
                </tr>
              `;
            }).join('')}
          </tbody>
        </table>
      </div>
    `;
  }

  function renderAdminRolesBody(data) {
    const roles = data?.roles || [];
    if (!roles.length) return '<p class="live-module-empty">Geen beheerdersrollen met leden gevonden.</p>';
    return `
      <div class="live-roles-grid">
        ${roles.map((role) => `
          <div class="live-role-card">
            <div class="live-role-header">
              <span>${liveEscapeHtml(role.roleName || '—')}</span>
              <span class="live-badge live-badge-info">${role.memberCount || 0} ${role.memberCount === 1 ? 'lid' : 'leden'}</span>
            </div>
            <div class="live-role-members">
              ${(role.members || []).map((m) => `
                <div class="live-role-member">
                  <span>${liveEscapeHtml(m.displayName || '—')}</span>
                  <span class="live-role-member-upn">${liveEscapeHtml(m.upn || '')}</span>
                </div>
              `).join('')}
            </div>
          </div>
        `).join('')}
      </div>
    `;
  }

  function renderSecurityDefaultsBody(data) {
    const enabled = data?.securityDefaultsEnabled;
    const rec = data?.recommendation || '';
    const lastMod = data?.lastModifiedAt ? new Date(data.lastModifiedAt).toLocaleDateString('nl-NL') : '—';
    const isWarn = rec.startsWith('Waarschuwing') || rec.startsWith('Let op');
    const statusLabel = enabled === true ? 'Ingeschakeld' : enabled === false ? 'Uitgeschakeld' : '—';
    const statusClass = enabled === true ? 'live-badge-ok' : enabled === false ? 'live-badge-warn' : 'live-badge-neutral';
    return `
      <div class="live-security-defaults-card">
        <div class="live-sd-status">
          <span>Security Defaults</span>
          <span class="live-badge ${statusClass} live-badge-lg">${liveEscapeHtml(statusLabel)}</span>
        </div>
        <div class="live-sd-row">
          <span class="live-sd-label">Laatste wijziging</span>
          <span>${liveEscapeHtml(lastMod)}</span>
        </div>
        <div class="live-sd-row">
          <span class="live-sd-label">Actieve CA-policies</span>
          <span>${data?.caEnabledPolicies ?? '—'}</span>
        </div>
        ${rec ? `<div class="live-sd-recommendation ${isWarn ? 'live-sd-warn' : 'live-sd-ok'}">${liveEscapeHtml(rec)}</div>` : ''}
      </div>
    `;
  }

  function renderLegacyAuthBody(data) {
    const users = data?.users || [];
    const note = data?.note;
    if (!users.length) {
      return `<p class="live-module-empty">${liveEscapeHtml(note || 'Geen legacy-auth activiteit gevonden in de afgelopen 30 dagen.')}</p>`;
    }
    return `
      <div class="assessment-table-wrap">
        <table class="assessment-table">
          <thead>
            <tr>
              <th>Gebruiker</th>
              <th>UPN</th>
              <th>Legacy clients</th>
              <th>Aanmeldingen</th>
              <th>Laatste aanmelding</th>
            </tr>
          </thead>
          <tbody>
            ${users.slice(0, 150).map((u) => {
              const lastSignIn = u.lastSignIn ? new Date(u.lastSignIn).toLocaleDateString('nl-NL') : '—';
              return `
                <tr>
                  <td>${liveEscapeHtml(u.displayName || '—')}</td>
                  <td class="live-small">${liveEscapeHtml(u.upn || '—')}</td>
                  <td><span class="live-badge live-badge-warn">${liveEscapeHtml(u.clients || '—')}</span></td>
                  <td>${u.signInCount || 0}</td>
                  <td>${liveEscapeHtml(lastSignIn)}</td>
                </tr>
              `;
            }).join('')}
          </tbody>
        </table>
      </div>
    `;
  }

  function renderCisBenchmarkBody(data) {
    const items = (data?.items || []).filter((item) => item && typeof item === 'object');
    const summary = data?.summary || {};
    if (!items.length && !summary.total) {
      return '<p class="live-module-empty">Geen CIS-data beschikbaar. Voer een assessment uit met de "-ExportJson" optie.</p>';
    }
    const statusBadge = (status) => {
      const map = { Pass: 'live-badge-ok', Fail: 'live-badge-crit', Warning: 'live-badge-warn', NA: 'live-badge-neutral' };
      const labels = { Pass: '✓ Pass', Fail: '✗ Fail', Warning: '⚠ Warning', NA: 'N/A' };
      return `<span class="live-badge ${map[status] || 'live-badge-neutral'}">${liveEscapeHtml(labels[status] || status)}</span>`;
    };
    const score = Number(summary.score) || 0;
    const progressTone = score >= 70 ? 'ok' : score >= 50 ? 'warn' : 'crit';
    return `
      <div class="live-insight-strip">
        <article class="live-insight-card">
          <span class="live-insight-label">Score</span>
          <strong class="live-storage-percent--${liveEscapeHtml(progressTone)}">${liveEscapeHtml(String(score))}%</strong>
          <span class="live-insight-meta">${liveEscapeHtml(formatNumber(summary.pass || 0))} / ${liveEscapeHtml(formatNumber(summary.total || 0))} controls</span>
        </article>
        <article class="live-insight-card">
          <span class="live-insight-label">Pass</span>
          <strong>${liveEscapeHtml(formatNumber(summary.pass || 0))}</strong>
          <span class="live-insight-meta">voldaan</span>
        </article>
        <article class="live-insight-card">
          <span class="live-insight-label">Fail</span>
          <strong>${liveEscapeHtml(formatNumber(summary.fail || 0))}</strong>
          <span class="live-insight-meta">niet voldaan</span>
        </article>
        <article class="live-insight-card">
          <span class="live-insight-label">Warning</span>
          <strong>${liveEscapeHtml(formatNumber(summary.warning || 0))}</strong>
          <span class="live-insight-meta">aandachtspunt</span>
        </article>
      </div>
      <div class="live-storage-progress" style="margin-bottom:1.5rem">
        <div class="live-storage-progress-bar">
          <span class="live-storage-progress-fill live-storage-progress-fill--${liveEscapeHtml(progressTone)}" style="width:${Math.max(0, Math.min(100, score))}%"></span>
        </div>
        <div class="live-storage-progress-meta">
          <span>${liveEscapeHtml(String(score))}% compliant</span>
          <span>CIS M365 Foundations Benchmark v3.0</span>
        </div>
      </div>
      <div class="assessment-table-wrap live-entity-table-wrap">
        <table class="assessment-table live-entity-table">
          <thead>
            <tr>
              <th>Control</th>
              <th>Level</th>
              <th>Categorie</th>
              <th>Status</th>
              <th>Detail</th>
              <th>NIST</th>
              <th>ISO 27001</th>
            </tr>
          </thead>
          <tbody>
            ${items.map((ctrl) => `
              <tr>
                <td>
                  <div class="live-entity-main">
                    <strong>${liveEscapeHtml(ctrl.Title || '—')}</strong>
                    <span>CIS ${liveEscapeHtml(ctrl.Id || '—')}</span>
                  </div>
                </td>
                <td><span class="live-badge live-badge-info">L${liveEscapeHtml(String(ctrl.Level || '?'))}</span></td>
                <td class="live-small">${liveEscapeHtml(ctrl.Category || '—')}</td>
                <td>${statusBadge(ctrl.Status)}</td>
                <td class="live-small">${liveEscapeHtml(ctrl.Detail || '—')}</td>
                <td class="live-small">${liveEscapeHtml(ctrl.NIST || '—')}</td>
                <td class="live-small">${liveEscapeHtml(ctrl.ISO27001 || '—')}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    `;
  }

  function renderHybridSyncBody(data) {
    const summary = data?.summary || {};
    const domains = (data?.items || []).filter((item) => item && typeof item === 'object');
    const isHybrid = !!summary.isHybrid;
    const syncStatus = String(summary.lastSyncStatus || 'Unknown');
    const syncStatusClass = syncStatus === 'OK' ? 'live-badge-ok' : syncStatus === 'Warning' ? 'live-badge-warn' : syncStatus === 'Critical' ? 'live-badge-crit' : 'live-badge-neutral';
    const syncAge = summary.lastSyncAgeHours != null ? `${Number(summary.lastSyncAgeHours).toFixed(1)} uur geleden` : '—';
    return `
      <div class="live-insight-strip">
        <article class="live-insight-card">
          <span class="live-insight-label">Type</span>
          <strong>${liveEscapeHtml(isHybrid ? 'Hybrid' : 'Cloud Only')}</strong>
          <span class="live-insight-meta">${liveEscapeHtml(summary.authType || '—')}</span>
        </article>
        <article class="live-insight-card">
          <span class="live-insight-label">Sync status</span>
          <strong><span class="live-badge ${liveEscapeHtml(syncStatusClass)}">${liveEscapeHtml(syncStatus)}</span></strong>
          <span class="live-insight-meta">${liveEscapeHtml(syncAge)}</span>
        </article>
        <article class="live-insight-card">
          <span class="live-insight-label">Gesynchroniseerd</span>
          <strong>${liveEscapeHtml(formatNumber(summary.syncedUsers || 0))}</strong>
          <span class="live-insight-meta">${liveEscapeHtml(formatNumber(summary.syncedUsersPercent || 0, 1))}% van ${liveEscapeHtml(formatNumber(summary.totalUsers || 0))}</span>
        </article>
        <article class="live-insight-card">
          <span class="live-insight-label">Cloud-only</span>
          <strong>${liveEscapeHtml(formatNumber(summary.cloudOnlyUsers || 0))}</strong>
          <span class="live-insight-meta">gebruikers zonder on-prem account</span>
        </article>
      </div>
      ${!isHybrid ? '<div class="live-module-empty" style="margin-bottom:1rem">Deze tenant is pure cloud — geen AD Connect configuratie aanwezig.</div>' : ''}
      ${domains.length ? `
        <div class="assessment-table-wrap live-entity-table-wrap">
          <table class="assessment-table live-entity-table">
            <thead>
              <tr>
                <th>Domein</th>
                <th>Auth type</th>
                <th>Geverifieerd</th>
                <th>Standaard</th>
              </tr>
            </thead>
            <tbody>
              ${domains.map((d) => `
                <tr>
                  <td><strong>${liveEscapeHtml(d.Domain || d.domain || '—')}</strong></td>
                  <td><span class="live-badge ${d.AuthType === 'Federated' ? 'live-badge-info' : 'live-badge-neutral'}">${liveEscapeHtml(d.AuthType || 'Managed')}</span></td>
                  <td><span class="live-badge ${d.IsVerified ? 'live-badge-ok' : 'live-badge-warn'}">${d.IsVerified ? 'Ja' : 'Nee'}</span></td>
                  <td>${d.IsDefault ? '<span class="live-badge live-badge-info">Standaard</span>' : ''}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>
      ` : ''}
    `;
  }

  // ── Zero Trust Assessment renderer ──────────────────────────────────────────

  const ZT_PILLAR_ICONS = { Identity: '🪪', Devices: '💻', Network: '🌐', Data: '🗄️' };

  function ztScoreColor(pct) {
    if (pct >= 80) return 'var(--color-ok, #22c55e)';
    if (pct >= 50) return 'var(--color-warn, #f59e0b)';
    return 'var(--color-danger, #ef4444)';
  }

  function renderZeroTrustBody(data) {
    const mod     = data?.module || {};
    const results = data?.results || {};
    const report  = data?.last_report || null;

    // Module not installed
    if (!mod.installed && !report) {
      return `
        <div class="live-module-notice" style="max-width:640px;margin:2rem auto">
          <h3 style="margin:0 0 .5rem">Zero Trust Assessment module niet gevonden</h3>
          <p style="color:var(--text-muted);margin:0 0 1.25rem">
            De Microsoft Zero Trust Assessment PowerShell-module is niet geïnstalleerd op de assessmentserver.
            Installeer de module en start daarna een assessment via onderstaande knop.
          </p>
          <pre style="background:var(--surface-raised,#1e2430);padding:.75rem 1rem;border-radius:8px;font-size:.8rem;overflow-x:auto">Install-Module ZeroTrustAssessment -Scope CurrentUser
Connect-ZtAssessment
Invoke-ZtAssessment</pre>
          <button class="live-btn live-btn-primary zt-run-btn" style="margin-top:1.25rem">
            Assessment starten (achtergrond)
          </button>
          <p style="font-size:.75rem;color:var(--text-muted);margin-top:.5rem">
            ⚠️ De assessment kan meerdere uren duren. De pagina kan worden verlaten.
          </p>
        </div>`;
    }

    const summary  = results?.summary || {};
    const pillars  = results?.pillars || {};
    const controls = (results?.controls || []).filter(Boolean);
    const total    = summary.total || 0;
    const score    = summary.score || 0;
    const reportDate = report?.date ? new Date(report.date).toLocaleDateString('nl-NL', { day:'2-digit', month:'short', year:'numeric', hour:'2-digit', minute:'2-digit' }) : '—';

    const PILLARS = ['Identity', 'Devices', 'Network', 'Data'];

    const pillarCards = PILLARS.map((name) => {
      const pct  = pillars[name] ?? null;
      const icon = ZT_PILLAR_ICONS[name] || '🔒';
      const color = pct !== null ? ztScoreColor(pct) : 'var(--text-muted)';
      return `
        <div class="live-stat-card" style="min-width:140px">
          <div style="font-size:1.5rem;margin-bottom:.25rem">${icon}</div>
          <div class="live-stat-value" style="color:${color}">${pct !== null ? pct + '%' : '—'}</div>
          <div class="live-stat-label">${liveEscapeHtml(name)}</div>
        </div>`;
    }).join('');

    const statusBadge = (s) => {
      const map = { Pass: 'live-badge-ok', Fail: 'live-badge-crit', Warning: 'live-badge-warn', NA: 'live-badge-neutral' };
      return `<span class="live-badge ${map[s] || 'live-badge-neutral'}">${liveEscapeHtml(s || 'NA')}</span>`;
    };

    const controlRows = controls.length
      ? controls.map((c) => `
          <tr>
            <td><div class="live-entity-main"><strong>${liveEscapeHtml(c.title || '—')}</strong></div></td>
            <td><span class="live-badge live-badge-info">${liveEscapeHtml(c.pillar || '—')}</span></td>
            <td>${statusBadge(c.status)}</td>
            <td class="live-small">${liveEscapeHtml(c.riskLevel || '—')}</td>
          </tr>`).join('')
      : `<tr><td colspan="4" class="live-empty-row">Geen controls beschikbaar — assessment nog niet uitgevoerd of rapport kon niet worden geparsed.</td></tr>`;

    const overallColor = ztScoreColor(score);

    return `
      <div class="live-insight-strip" style="margin-bottom:1.25rem;display:flex;align-items:center;gap:1.25rem;flex-wrap:wrap">
        <div style="display:contents">
          <div>
            <div style="font-size:.75rem;color:var(--text-muted);text-transform:uppercase;letter-spacing:.05em">Overall score</div>
            <div style="font-size:2rem;font-weight:700;color:${overallColor}">${score}%</div>
          </div>
          <div class="live-storage-progress" style="flex:1;min-width:160px">
            <div class="live-storage-progress-bar">
              <span class="live-storage-progress-fill" style="width:${score}%;background:${overallColor}"></span>
            </div>
            <div class="live-storage-progress-meta">
              <span>${summary.pass || 0} pass · ${summary.fail || 0} fail · ${summary.warning || 0} warning</span>
              <span>Laatste run: ${liveEscapeHtml(reportDate)}</span>
            </div>
          </div>
          <button class="live-btn live-btn-secondary zt-run-btn" style="white-space:nowrap">↺ Opnieuw uitvoeren</button>
        </div>
      </div>

      <div class="live-stats-row" style="margin-bottom:1.5rem">
        ${pillarCards}
      </div>

      ${controls.length ? `
      <div class="assessment-table-wrap live-entity-table-wrap">
        <table class="assessment-table live-entity-table">
          <thead><tr>
            <th>Control</th>
            <th>Pillar</th>
            <th>Status</th>
            <th>Risico</th>
          </tr></thead>
          <tbody>${controlRows}</tbody>
        </table>
      </div>` : `
      <div class="live-module-empty">
        <p>Geen controlresultaten beschikbaar.</p>
        ${!controls.length && report ? '<p style="font-size:.8rem;color:var(--text-muted)">Het rapport is gevonden maar kon niet automatisch worden geparsed. Open het volledige rapport via de rapportenpagina.</p>' : ''}
        <button class="live-btn live-btn-primary zt-run-btn" style="margin-top:1rem">Assessment uitvoeren</button>
      </div>`}`;
  }

  // Wire up run button (event delegation)
  document.addEventListener('click', (e) => {
    if (!e.target.closest('.zt-run-btn')) return;
    const tid = (typeof currentTenantId !== 'undefined') ? currentTenantId : null;
    if (!tid) return;
    const btn = e.target.closest('.zt-run-btn');
    btn.disabled = true;
    btn.textContent = 'Starten…';
    fetch(`/api/compliance/${tid}/zerotrust/run`, {
      method: 'POST',
      credentials: 'include',
      headers: { 'Content-Type': 'application/json', ...(localStorage.getItem('denjoy_token') ? { 'Authorization': `Bearer ${localStorage.getItem('denjoy_token')}` } : {}) },
    }).then((r) => r.json()).then((d) => {
      btn.textContent = d.ok ? '✓ Assessment gestart' : '✗ Mislukt';
      if (d.ok) {
        setTimeout(() => { btn.disabled = false; btn.textContent = '↺ Opnieuw uitvoeren'; }, 5000);
        if (typeof window.showToast === 'function') window.showToast('Zero Trust Assessment gestart als achtergrondtaak. Dit kan meerdere uren duren.', 'info');
      }
    }).catch(() => { btn.textContent = '✗ Fout'; btn.disabled = false; });
  });

  function renderSectionBody(sectionName, tabKey, data, collection) {
    if (sectionName === 'teams' && tabKey === 'teams') return renderTeamsBody(data);
    if (sectionName === 'teams' && tabKey === 'groepen') return renderGroupsBody(data);
    if (sectionName === 'sharepoint' && tabKey === 'sharepoint-sites') return renderSharePointSitesBody(data);
    if (sectionName === 'apps' && tabKey === 'registrations') return renderAppRegistrationsBody(data);
    if (sectionName === 'compliance' && tabKey === 'cis') return renderCisBenchmarkBody(data);
    if (sectionName === 'compliance' && tabKey === 'zerotrust') return renderZeroTrustBody(data);
    if (sectionName === 'hybrid' && tabKey === 'sync') return renderHybridSyncBody(data);
    if (sectionName === 'identity') {
      if (tabKey === 'mfa') return renderMfaBody(data);
      if (tabKey === 'guests') return renderGuestsBody(data);
      if (tabKey === 'admin-roles') return renderAdminRolesBody(data);
      if (tabKey === 'security-defaults') return renderSecurityDefaultsBody(data);
      if (tabKey === 'legacy-auth') return renderLegacyAuthBody(data);
    }
    // Generic fallback: summary + table
    const summary = renderSummary(data);
    const body = collection ? renderTable(collection, sectionName, tabKey) : renderObjectTable(data);
    return summary + body;
  }

  function renderObjectTable(data) {
    const rows = Object.keys(data || {}).slice(0, 20);
    if (!rows.length) return '<p class="live-module-empty">Geen data beschikbaar.</p>';
    return `
      <div class="assessment-table-wrap">
        <table class="assessment-table">
          <thead>
            <tr><th>Veld</th><th>Waarde</th></tr>
          </thead>
          <tbody>
            ${rows.map((key) => `
              <tr>
                <td>${liveEscapeHtml(key)}</td>
                <td>${liveEscapeHtml(normalizeScalar(data[key]))}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    `;
  }

  function _formatRelativeTime(isoStr) {
    if (!isoStr) return null;
    try {
      const diff = Math.floor((Date.now() - new Date(isoStr).getTime()) / 1000);
      if (diff < 60)   return 'zojuist';
      if (diff < 3600) return `${Math.floor(diff / 60)} min geleden`;
      if (diff < 86400) return `${Math.floor(diff / 3600)} uur geleden`;
      return `${Math.floor(diff / 86400)} dag(en) geleden`;
    } catch (_) { return null; }
  }

  function renderSourceInfo(data) {
    const describe = window.denjoyDescribeSourceMeta;
    if (typeof describe !== 'function') return '';
    const info = describe(data || {});
    const ts = data?._generated_at || data?.generated_at || data?.assessment_generated_at || null;
    const rel = _formatRelativeTime(ts);
    const stale = !!data?._stale;
    const syncHtml = rel
      ? `<span class="live-module-sync-time${stale ? ' live-module-sync-stale' : ''}">Sync: ${liveEscapeHtml(rel)}${stale ? ' ⚠ verouderd' : ''}</span>`
      : '';
    return `
      <div class="live-module-source">
        <span class="live-module-source-pill ${liveEscapeHtml(info.className || '')}">${liveEscapeHtml(info.label)}</span>
        <span>${liveEscapeHtml(info.detail)}</span>
        ${syncHtml}
      </div>
    `;
  }

  function renderCapabilityInfo(capability) {
    const describe = window.denjoyDescribeCapabilityStatus;
    if (typeof describe !== 'function' || !capability) return '';
    const info = describe(capability);
    const roles = (capability.extra_roles || []).slice(0, 3).join(', ');
    const consent = (capability.extra_consent || []).slice(0, 3).join(', ');
    return `
      <div class="live-module-capability">
        <div class="live-module-source">
          <span class="live-module-source-pill ${liveEscapeHtml(info.className || '')}">${liveEscapeHtml(info.label)}</span>
          <span>${liveEscapeHtml(info.detail)}</span>
        </div>
        <div class="live-module-capability-grid">
          <article class="live-module-capability-card">
            <span class="live-module-capability-label">Engine</span>
            <strong>${liveEscapeHtml(capability.engine || '—')}</strong>
            <span class="live-module-capability-meta">${liveEscapeHtml(capability.access_method || '—')}</span>
          </article>
          <article class="live-module-capability-card">
            <span class="live-module-capability-label">Rollen</span>
            <strong>${liveEscapeHtml(roles || 'Geen extra rollen')}</strong>
            <span class="live-module-capability-meta">${capability.gdap_required ? 'GDAP betrokken' : 'geen GDAP vereist'}</span>
          </article>
          <article class="live-module-capability-card">
            <span class="live-module-capability-label">Consent</span>
            <strong>${liveEscapeHtml(consent || 'Geen extra consent')}</strong>
            <span class="live-module-capability-meta">${capability.supports_live ? 'live-capable' : 'snapshot-only'}</span>
          </article>
        </div>
      </div>
    `;
  }

  function syncModuleContext(sectionName, tabKey, capability = null, data = null) {
    const setter = window.denjoySetLiveModuleContext;
    if (typeof setter !== 'function') return;
    setter({
      section: sectionName,
      tab: tabKey,
      capability,
      source: data?._source || null,
      stale: !!data?._stale,
    });
  }

  function renderLegacyIntuneBody() {
    return `
      <div class="it-module-shell">
        <div id="itWorkspaceSource"></div>
        <div id="itServiceOverview"></div>
        <div class="it-tab-panel" data-it-panel="overzicht">
          <div class="it-topbar">
            <div class="it-counter" id="itSummaryCounter">Live tenantoverzicht</div>
            <button type="button" class="live-module-refresh" id="itBtnRefreshSummary">Overzicht vernieuwen</button>
          </div>
          <div id="itSummaryWrap"></div>
        </div>

        <div class="it-tab-panel" data-it-panel="apparaten" style="display:none">
          <div class="it-topbar">
            <div class="it-search">
              <input type="search" id="itSearchInput" placeholder="Zoek op apparaat, gebruiker of model">
            </div>
            <div class="it-filter-row">
              <button type="button" class="it-filter-tab active" data-filter-os="all">Alle OS</button>
              <button type="button" class="it-filter-tab" data-filter-os="windows">Windows</button>
              <button type="button" class="it-filter-tab" data-filter-os="ios">iOS</button>
              <button type="button" class="it-filter-tab" data-filter-os="android">Android</button>
              <button type="button" class="it-filter-tab" data-filter-os="macos">macOS</button>
              <button type="button" class="it-filter-tab active" data-filter-state="all">Alle statussen</button>
              <button type="button" class="it-filter-tab" data-filter-state="compliant">Compliant</button>
              <button type="button" class="it-filter-tab" data-filter-state="noncompliant">Non-compliant</button>
              <button type="button" class="it-filter-tab" data-filter-state="inGracePeriod">Grace period</button>
              <button type="button" class="live-module-refresh" id="itBtnRefresh">Apparaten vernieuwen</button>
            </div>
          </div>
          <div class="it-subtle-meta" id="itDeviceCount">0 apparaten</div>
          <div class="assessment-table-wrap">
            <table class="assessment-table">
              <thead>
                <tr>
                  <th>Apparaat</th>
                  <th>OS</th>
                  <th>Gebruiker</th>
                  <th>Compliance</th>
                  <th>Laatst gezien</th>
                  <th>Actie</th>
                </tr>
              </thead>
              <tbody id="itDeviceTableBody">
                <tr><td colspan="6" class="it-table-empty">Nog geen apparaten geladen.</td></tr>
              </tbody>
            </table>
          </div>
        </div>

        <div class="it-tab-panel" data-it-panel="compliance" style="display:none">
          <div class="it-topbar">
            <div class="it-counter" id="itComplianceCount">Compliance policies</div>
            <button type="button" class="live-module-refresh" id="itBtnRefreshCompliance">Compliance vernieuwen</button>
          </div>
          <div id="itComplianceGrid"></div>
        </div>

        <div class="it-tab-panel" data-it-panel="configuratie" style="display:none">
          <div class="it-topbar">
            <div class="it-counter" id="itConfigCount">Configuratieprofielen</div>
            <button type="button" class="live-module-refresh" id="itBtnRefreshConfig">Configuratie vernieuwen</button>
          </div>
          <div id="itConfigGrid"></div>
        </div>

        <div class="it-tab-panel" data-it-panel="geschiedenis" style="display:none">
          <div class="assessment-table-wrap">
            <table class="assessment-table">
              <thead>
                <tr>
                  <th>Tijdstip</th>
                  <th>Actie</th>
                  <th>Status</th>
                  <th>Uitgevoerd door</th>
                  <th>Details</th>
                </tr>
              </thead>
              <tbody id="itHistoryBody">
                <tr><td colspan="5" class="it-table-empty">Nog geen geschiedenis geladen.</td></tr>
              </tbody>
            </table>
          </div>
        </div>
      </div>
    `;
  }

  function renderLegacyBackupBody() {
    return `
      <div class="bk-module-shell">
        <div id="bkWorkspaceSource"></div>
        <div id="bkServiceOverview"></div>
        <div class="bk-tab-panel" data-bk-panel="overzicht">
          <div class="it-topbar">
            <div class="it-counter">Backupsamenvatting</div>
            <button type="button" class="live-module-refresh" id="bkBtnRefreshSummary">Overzicht vernieuwen</button>
          </div>
          <div id="bkSummaryWrap"></div>
        </div>

        <div class="bk-tab-panel" data-bk-panel="sharepoint" style="display:none">
          <div class="it-topbar">
            <div class="it-counter" id="bkSPCount">— policies</div>
            <button type="button" class="live-module-refresh" id="bkBtnRefreshSP">SharePoint vernieuwen</button>
          </div>
          <div id="bkSPList"></div>
        </div>

        <div class="bk-tab-panel" data-bk-panel="onedrive" style="display:none">
          <div class="it-topbar">
            <div class="it-counter" id="bkODCount">— policies</div>
            <button type="button" class="live-module-refresh" id="bkBtnRefreshOD">OneDrive vernieuwen</button>
          </div>
          <div id="bkODList"></div>
        </div>

        <div class="bk-tab-panel" data-bk-panel="exchange" style="display:none">
          <div class="it-topbar">
            <div class="it-counter" id="bkEXCount">— policies</div>
            <button type="button" class="live-module-refresh" id="bkBtnRefreshEX">Exchange vernieuwen</button>
          </div>
          <div id="bkEXList"></div>
        </div>

        <div class="bk-tab-panel" data-bk-panel="geschiedenis" style="display:none">
          <div class="assessment-table-wrap">
            <table class="assessment-table">
              <thead>
                <tr>
                  <th>Tijdstip</th>
                  <th>Workload</th>
                  <th>Status</th>
                  <th>Uitgevoerd door</th>
                  <th>Resultaat</th>
                </tr>
              </thead>
              <tbody id="bkHistoryBody">
                <tr><td colspan="5" class="bk-empty">Nog geen geschiedenis geladen.</td></tr>
              </tbody>
            </table>
          </div>
        </div>
      </div>
    `;
  }

  function renderLegacyAlertsBody() {
    return `
      <div class="al-module-shell">
        <div id="alWorkspaceSource"></div>
        <div id="alServiceOverview"></div>
        <div id="alSnapshotBanner" class="snapshot-banner" style="display:none"></div>

        <div class="al-tab-panel" data-al-panel="auditlog">
          <div class="it-topbar">
            <div class="it-counter" id="alAuditCount">0 events</div>
            <button type="button" class="live-module-refresh" id="alBtnRefreshAudit">Auditlog vernieuwen</button>
          </div>
          <div id="alAuditWrap"></div>
        </div>

        <div class="al-tab-panel" data-al-panel="securescr" style="display:none">
          <div class="it-topbar">
            <div class="it-counter">Microsoft Secure Score</div>
            <button type="button" class="live-module-refresh" id="alBtnRefreshScore">Secure Score vernieuwen</button>
          </div>
          <div id="alScoreWrap"></div>
        </div>

        <div class="al-tab-panel" data-al-panel="signins" style="display:none">
          <div class="it-topbar">
            <div class="it-counter">Recente aanmeldingen</div>
            <button type="button" class="live-module-refresh" id="alBtnRefreshSignIns">Aanmeldingen vernieuwen</button>
          </div>
          <div id="alSignInsWrap"></div>
        </div>

        <div class="al-tab-panel" data-al-panel="config" style="display:none">
          <div class="live-module-config-card">
            <div class="live-module-config-title">Webhook & e-mail notificaties</div>
            <div class="live-module-form-group">
              <label for="alWebhookUrl">Webhook URL</label>
              <input type="url" id="alWebhookUrl" class="live-module-input" placeholder="https://outlook.office.com/webhook/...">
            </div>
            <div class="live-module-form-group">
              <label for="alWebhookType">Webhook type</label>
              <select id="alWebhookType" class="live-module-select">
                <option value="teams">Microsoft Teams</option>
                <option value="slack">Slack</option>
                <option value="generic">Generiek JSON</option>
              </select>
            </div>
            <div class="live-module-form-group">
              <label for="alEmailAddr">E-mailadres</label>
              <input type="email" id="alEmailAddr" class="live-module-input" placeholder="alerts@bedrijf.nl">
            </div>
            <div class="live-module-action-row">
              <button type="button" class="live-module-refresh" id="alBtnSaveConfig">Opslaan</button>
              <button type="button" class="live-module-refresh" id="alBtnTestWebhook">Test webhook</button>
            </div>
            <div id="alConfigResult" class="al-test-result" style="display:none"></div>
          </div>
        </div>
      </div>
    `;
  }

  function renderLegacyExchangeBody() {
    return `
      <div class="ex-module-shell">
        <div id="exWorkspaceSource"></div>
        <div id="exServiceOverview"></div>
        <div class="ex-tab-panel" data-ex-panel="mailboxen">
          <div class="it-topbar">
            <div class="it-search">
              <input type="search" id="exSearchInput" placeholder="Zoek op naam, UPN of e-mail">
            </div>
            <div class="it-filter-row">
              <div class="it-counter" id="exMbxCount">0 mailboxen</div>
              <button type="button" class="live-module-refresh" id="exBtnRefreshMbx">Mailboxen vernieuwen</button>
            </div>
          </div>
          <div class="assessment-table-wrap">
            <table class="assessment-table ex-table">
              <thead>
                <tr>
                  <th>Mailbox</th>
                  <th>E-mail</th>
                  <th>Status</th>
                  <th>Tijdzone</th>
                  <th>Actie</th>
                </tr>
              </thead>
              <tbody id="exMailboxTableBody">
                <tr><td colspan="5" class="ex-table-empty">Nog geen mailboxen geladen.</td></tr>
              </tbody>
            </table>
          </div>
        </div>

        <div class="ex-tab-panel" data-ex-panel="forwarding" style="display:none">
          <div class="it-topbar">
            <div class="it-counter" id="exFwdCount">0 actieve forwardings</div>
            <button type="button" class="live-module-refresh" id="exBtnRefreshFwd">Forwarding vernieuwen</button>
          </div>
          <div id="exFwdWrap"></div>
        </div>

        <div class="ex-tab-panel" data-ex-panel="regels" style="display:none">
          <div class="it-topbar">
            <div class="it-counter" id="exRulesCount">0 regels</div>
            <button type="button" class="live-module-refresh" id="exBtnRefreshRules">Regels vernieuwen</button>
          </div>
          <div id="exRulesWrap"></div>
        </div>
      </div>
    `;
  }

  const LEGACY_MODULES = {
    intune: {
      renderBody: renderLegacyIntuneBody,
      load: () => window.loadIntuneSection?.(),
      switchTab: (tabKey) => window.switchIntuneTab?.(tabKey),
    },
    backup: {
      renderBody: renderLegacyBackupBody,
      load: () => window.loadBackupSection?.(),
      switchTab: (tabKey) => window.switchBackupTab?.(tabKey),
    },
    alerts: {
      renderBody: renderLegacyAlertsBody,
      load: () => window.loadAlertsSection?.(),
      switchTab: (tabKey) => window.switchAlertsTab?.(tabKey),
    },
    exchange: {
      renderBody: renderLegacyExchangeBody,
      load: () => window.loadExchangeSection?.(),
      switchTab: (tabKey) => window.switchExchangeTab?.(tabKey),
    },
  };

  function renderModuleShell(sectionName, tabKey, innerHtml, noticeHtml = '', shellOpts = {}) {
    const config = getModuleConfig(sectionName);
    const tab = config?.tabs?.[tabKey];
    const root = getModuleRoot(sectionName);
    if (!root || !tab) return;
    const buttonLabel = shellOpts.buttonLabel || 'Data ophalen';
    const buttonDisabled = shellOpts.buttonDisabled ? ' disabled aria-disabled="true"' : '';
    const toolbarMeta = shellOpts.toolbarMetaHtml || '';
    const inputHtml = tab.input ? `
      <div class="live-module-input-row">
        <label class="live-module-input-label" for="liveModuleInput-${liveEscapeHtml(sectionName)}-${liveEscapeHtml(tabKey)}">${liveEscapeHtml(tab.input.label)}</label>
        <input
          type="text"
          id="liveModuleInput-${liveEscapeHtml(sectionName)}-${liveEscapeHtml(tabKey)}"
          class="live-module-input"
          placeholder="${liveEscapeHtml(tab.input.placeholder || '')}"
        />
      </div>
    ` : '';
    const _sectionMeta = window.SECTION_META || {};
    const _kickerLabel = (_sectionMeta[sectionName] && _sectionMeta[sectionName].title) || sectionName;
    root.innerHTML = `
      <div class="live-module-shell">
        <div class="live-module-toolbar">
          <div>
            <div class="live-module-kicker">${liveEscapeHtml(_kickerLabel)}</div>
            <h3>${liveEscapeHtml(tab.title)}</h3>
            <p>${liveEscapeHtml(tab.description)}</p>
            ${toolbarMeta}
          </div>
          <button type="button" class="live-module-refresh" data-live-section="${liveEscapeHtml(sectionName)}" data-live-subtab="${liveEscapeHtml(tabKey)}"${buttonDisabled}>
            ${liveEscapeHtml(buttonLabel)}
          </button>
        </div>
        ${inputHtml}
        ${noticeHtml}
        <div class="live-module-body">
          ${innerHtml}
        </div>
      </div>
    `;
  }

  function renderLegacyModule(sectionName, tabKey) {
    const legacy = LEGACY_MODULES[sectionName];
    if (!legacy) return false;
    const capability = capabilityState.byKey[`${sectionName}:${tabKey}`] || null;

    if (sectionName === 'backup') window._bkLastTenantId = null;
    if (sectionName === 'alerts') window._alLastTid = null;
    if (sectionName === 'exchange') window._exLastTid = null;

    const noticeHtml = `
      <div class="live-module-banner">
        Live module met bestaande detailweergave voor dit subhoofdstuk.
      </div>
    `;
    renderModuleShell(sectionName, tabKey, legacy.renderBody(), noticeHtml, {
      buttonDisabled: !!(capability && !capability.supports_live),
    });
    legacy.load();
    legacy.switchTab(tabKey);
    if (sectionName === 'intune') {
      const summaryButton = document.getElementById('itBtnRefreshSummary');
      if (summaryButton) {
        summaryButton.addEventListener('click', () => {
          window.loadIntuneSection?.();
          window.switchIntuneTab?.('overzicht');
        });
      }
    }
    return true;
  }

  function renderLoading(sectionName, tabKey) {
    const skBody = window.skeletonCards ? window.skeletonCards(4) : '<p class="live-module-empty">Tenantdata wordt opgehaald...</p>';
    const capability = capabilityState.byKey[`${sectionName}:${tabKey}`] || null;
    renderModuleShell(sectionName, tabKey, skBody, '', {
      buttonDisabled: !!(capability && !capability.supports_live),
    });
  }

  function renderError(sectionName, tabKey, message) {
    const capability = capabilityState.byKey[`${sectionName}:${tabKey}`] || null;
    renderModuleShell(
      sectionName,
      tabKey,
      `<div class="live-module-error">${liveEscapeHtml(message || 'Ophalen mislukt.')}</div>`,
      '<div class="snapshot-banner">Live data ophalen is mislukt. Controleer tenantrechten of verbinding.</div>',
      {
        buttonDisabled: !!(capability && !capability.supports_live),
      }
    );
  }

  function renderData(sectionName, tabKey, data) {
    const capability = capabilityState.byKey[`${sectionName}:${tabKey}`] || null;
    if (data && data.ok === false) {
      renderError(sectionName, tabKey, data.error || 'API-fout bij ophalen tenantdata.');
      return;
    }
    const tab = getTabConfig(sectionName, tabKey);
    if (tab?.customType === 'alerts-config') {
      const noticeHtml = '<div class="live-module-banner">Beheer notificatie-uitvoer direct vanuit deze workspace.</div>';
      renderModuleShell(sectionName, tabKey, renderAlertsConfigForm(sectionName, tabKey, data), noticeHtml, {
        buttonDisabled: !!(capability && !capability.supports_live),
      });
      return;
    }
    const collection = extractCollection(data);
    const serviceOverview = renderServiceOverview(sectionName, tabKey, data, collection);
    const body = renderSectionBody(sectionName, tabKey, data, collection);
    const sourceHtml = renderSourceInfo(data);
    const noticeHtml = data && data._source === 'assessment_snapshot'
      ? '<div class="snapshot-banner">Gegevens uit laatste assessment. Live data vereist actieve verbinding.</div>'
      : '<div class="live-module-banner">Live tenantdata succesvol opgehaald voor het geselecteerde subhoofdstuk.</div>';
    renderModuleShell(sectionName, tabKey, `${renderCapabilityInfo(capability)}${sourceHtml}${serviceOverview}${body}`, noticeHtml, {
      buttonDisabled: !!(capability && !capability.supports_live),
    });
  }

  function renderCapabilityBlocked(sectionName, tabKey, capability) {
    const roles = (capability?.extra_roles || []).join(', ');
    const consent = (capability?.extra_consent || []).slice(0, 4).join(', ');
    const reqHtml = (roles || consent) ? `
      <div class="live-module-req-info">
        ${roles ? `<div class="live-module-req-row"><span class="live-module-req-label">Vereiste rollen</span><span class="live-module-req-val">${liveEscapeHtml(roles)}</span></div>` : ''}
        ${consent ? `<div class="live-module-req-row"><span class="live-module-req-label">Graph consent</span><span class="live-module-req-val">${liveEscapeHtml(consent)}</span></div>` : ''}
      </div>` : '';
    const notice = capability?.assessment_available
      ? '<div class="snapshot-banner">Assessment fallback is beschikbaar, maar live ophalen is nog niet gereed.</div>'
      : '<div class="snapshot-banner">Live ophalen is nog niet gereed en er is geen assessment fallback gevonden.</div>';
    renderModuleShell(
      sectionName,
      tabKey,
      `<div class="live-module-empty">${liveEscapeHtml(capability?.status_reason || 'Live data ophalen is nog niet mogelijk voor dit subhoofdstuk.')}</div>${reqHtml}`,
      notice,
      {
        buttonDisabled: true,
        buttonLabel: capability?.assessment_available ? 'Live nog niet gereed' : 'Niet gereed',
      }
    );
  }

  function getInputValues(sectionName, tabKey) {
    const tab = getTabConfig(sectionName, tabKey);
    if (!tab?.input) return {};
    const inputId = `liveModuleInput-${sectionName}-${tabKey}`;
    const value = document.getElementById(inputId)?.value?.trim() || '';
    return { [tab.input.key]: value };
  }

  async function loadLiveModuleSection(sectionName, tabKey, { forceRefresh = false } = {}) {
    const config = getModuleConfig(sectionName);
    if (!config) return;
    const tenantId = selectedTenantId();
    const activeTab = tabKey || config.defaultTab;
    liveState.section = sectionName;
    liveState.tab = activeTab;

    if (typeof setActiveSubnavItem === 'function') setActiveSubnavItem(activeTab);

    if (!tenantId) {
      syncModuleContext(sectionName, activeTab, null, null);
      renderModuleShell(sectionName, activeTab, '<p class="live-module-empty">Selecteer eerst een tenant om data op te halen.</p>');
      return;
    }

    const tab = config.tabs[activeTab];
    if (!tab) return;

    const fetchCapability = window.denjoyFetchCapabilityStatus;
    if (typeof fetchCapability === 'function') {
      try {
        capabilityState.byKey[`${sectionName}:${activeTab}`] = await fetchCapability(tenantId, sectionName, activeTab, { forceRefresh });
      } catch (_) {}
    }
    const capability = capabilityState.byKey[`${sectionName}:${activeTab}`] || null;
    syncModuleContext(sectionName, activeTab, capability, null);

    if (capability && (!capability.supports_live || capability.status === 'config_required' || capability.status === 'not_implemented')) {
      renderCapabilityBlocked(sectionName, activeTab, capability);
      return;
    }

    if (renderLegacyModule(sectionName, activeTab)) {
      return;
    }

    const inputs = getInputValues(sectionName, activeTab);
    if (tab.input && !inputs[tab.input.key]) {
      syncModuleContext(sectionName, activeTab, capability, null);
      renderModuleShell(sectionName, activeTab, '<p class="live-module-empty">Vul eerst een waarde in om dit subhoofdstuk live op te halen.</p>');
      return;
    }

    renderLoading(sectionName, activeTab);
    try {
      const apiPath = tab.endpoint(tenantId, inputs);
      if (forceRefresh && window.cacheClear) window.cacheClear(apiPath);
      const data = await liveFetchJson(apiPath, { skipCache: forceRefresh });
      syncModuleContext(sectionName, activeTab, capability, data || null);
      renderData(sectionName, activeTab, data || {});
    } catch (error) {
      renderError(sectionName, activeTab, error.message || 'Ophalen mislukt.');
      if (typeof showToast === 'function') showToast(error.message || 'Ophalen mislukt.', 'error');
    }
  }

  function switchLiveModuleTab(sectionName, tabKey) {
    loadLiveModuleSection(sectionName, tabKey);
  }

  document.addEventListener('click', (event) => {
    const button = event.target.closest('[data-live-section][data-live-subtab]');
    if (!button) return;
    event.preventDefault();
    // Refresh-knoppen (class live-module-refresh) slaan de cache over
    const isRefresh = button.classList.contains('live-module-refresh');
    loadLiveModuleSection(button.dataset.liveSection, button.dataset.liveSubtab, { forceRefresh: isRefresh });
  });

  document.addEventListener('click', async (event) => {
    const button = event.target.closest('[data-alerts-action]');
    if (!button) return;
    event.preventDefault();
    const tenantId = selectedTenantId();
    if (!tenantId) return;
    const result = document.getElementById('alertsConfigResult');
    const webhook_url = document.getElementById('alertsWebhookUrl')?.value?.trim() || '';
    const webhook_type = document.getElementById('alertsWebhookType')?.value || 'teams';
    const email_addr = document.getElementById('alertsEmailAddr')?.value?.trim() || '';
    try {
      if (button.dataset.alertsAction === 'save') {
        await liveApiRequest(`/api/alerts/${tenantId}/config`, {
          method: 'POST',
          body: JSON.stringify({ webhook_url, webhook_type, email_addr }),
        });
        if (result) result.textContent = 'Configuratie opgeslagen.';
      } else if (button.dataset.alertsAction === 'test') {
        const data = await liveApiRequest(`/api/alerts/${tenantId}/test-webhook`, {
          method: 'POST',
          body: JSON.stringify({ webhook_url, webhook_type }),
        });
        if (result) result.textContent = data.ok ? 'Testbericht verzonden.' : (data.error || 'Test mislukt.');
      }
      if (result) result.className = 'live-module-config-result is-visible';
    } catch (error) {
      if (result) {
        result.textContent = error.message || 'Actie mislukt.';
        result.className = 'live-module-config-result is-visible is-error';
      }
    }
  });

  document.addEventListener('click', (event) => {
    const button = event.target.closest('[data-appreg-id]');
    if (!button) return;
    event.preventDefault();
    openAppRegistrationModal(button.dataset.appregId);
  });

  document.addEventListener('click', (event) => {
    const button = event.target.closest('.grp-detail-btn');
    if (!button) return;
    event.preventDefault();
    const d = button.dataset;
    const name = d.groupName || d.groupId || 'Groep';
    if (typeof window.openSideRailDetail === 'function') {
      window.openSideRailDetail(d.groupType || 'Groep', name);
    }
    const bodyHtml = `
      <div class="gb-detail-grid">
        <div class="gb-detail-row"><span class="gb-detail-label">Type</span><span>${liveEscapeHtml(d.groupType || '—')}</span></div>
        <div class="gb-detail-row"><span class="gb-detail-label">E-mail</span><span>${liveEscapeHtml(d.groupMail || '—')}</span></div>
        <div class="gb-detail-row"><span class="gb-detail-label">Leden</span><span>${liveEscapeHtml(d.memberCount || '0')}</span></div>
        <div class="gb-detail-row"><span class="gb-detail-label">Owners</span><span>${liveEscapeHtml(d.ownerCount || '0')}</span></div>
        <div class="gb-detail-row"><span class="gb-detail-label">Gasten</span><span>${liveEscapeHtml(d.guestCount || '0')}</span></div>
        <div class="gb-detail-row"><span class="gb-detail-label">Dynamisch</span><span>${d.isDynamic === 'true' ? 'Ja' : 'Nee'}</span></div>
        <div class="gb-detail-row"><span class="gb-detail-label">Aangemaakt</span><span>${liveEscapeHtml(formatDate(d.createdAt || ''))}</span></div>
        ${d.description ? `<div class="gb-detail-row" style="grid-column:1/-1"><span class="gb-detail-label">Omschrijving</span><span>${liveEscapeHtml(d.description)}</span></div>` : ''}
        <div class="gb-detail-row" style="grid-column:1/-1"><span class="gb-detail-label">Object ID</span><span style="font-family:var(--mono,monospace);font-size:0.72rem;word-break:break-all">${liveEscapeHtml(d.groupId || '—')}</span></div>
      </div>
    `;
    if (typeof window.updateSideRailDetail === 'function') {
      window.updateSideRailDetail(name, bodyHtml);
    }
  });

  window.loadLiveModuleSection = loadLiveModuleSection;
  window.switchLiveModuleTab = switchLiveModuleTab;
})();
