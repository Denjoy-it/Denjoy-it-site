// Results Viewer - opgebouwd in hetzelfde tab/pane patroon als de Kennisbank.

class ResultsViewer {
    constructor(containerId) {
        this.container = document.getElementById(containerId);
        this.currentTab = 'overview';
        this.currentPanel = 'viewer';
        this.currentPhaseFilter = 'overview';
        this.renderMode = 'legacy';
        this.jsonRunData = null;
        this.activeMetadata = null;
        this.phases = [];
        this.phaseSummaries = [];
        this.previousPhaseSummaries = [];
        this.findings = [];
        this.reportPath = null;
        this.portalNav = null;
        this.reportData = {
            tenantName: '-',
            tenantId: '-',
            assessmentDate: '-',
        };
        this.ctx = {
            latestRun: null,
            reportRuns: [],
            summary: {
                totalRuns: 0,
                completedRuns: 0,
                failedRuns: 0,
                reportRuns: 0,
            },
        };
    }

    async loadReport(reportPath, context = {}) {
        if (!this.container) return;
        this.reportPath = reportPath;
        this.currentPanel = 'viewer';
        this.currentPhaseFilter = 'overview';
        this.currentTab = 'overview';
        this.ctx = {
            ...this.ctx,
            ...context,
            summary: { ...this.ctx.summary, ...(context.summary || {}) },
            reportRuns: Array.isArray(context.reportRuns) ? context.reportRuns : [],
        };

        const runId = this.ctx?.latestRun?.id;
        if (runId) {
            const loadedFromJson = await this.loadRunJsonReport(runId);
            if (loadedFromJson) return;
        }

        if (!reportPath && this.ctx.tenantId) {
            await this.loadPortalAssessment(this.ctx.tenantId);
            return;
        }

        try {
            const response = await fetch(reportPath);
            if (!response.ok) throw new Error(`Failed to load report: ${response.statusText}`);

            const htmlContent = await response.text();
            const parser = new DOMParser();
            const doc = parser.parseFromString(htmlContent, 'text/html');
            const currentJsonMeta = await this.loadReportMetadata(reportPath, doc);
            this.extractPhases(doc, currentJsonMeta);

            const previousReportPath = this.ctx.reportRuns && this.ctx.reportRuns[1] ? this.ctx.reportRuns[1].report_path : '';
            if (previousReportPath) {
                try {
                    const previousResponse = await fetch(previousReportPath);
                    if (previousResponse.ok) {
                        const previousHtml = await previousResponse.text();
                        const previousDoc = parser.parseFromString(previousHtml, 'text/html');
                        const prevJsonMeta = await this.loadReportMetadata(previousReportPath, previousDoc);
                        this.previousPhaseSummaries = this.extractPhaseSummaries(previousDoc, prevJsonMeta);
                    }
                } catch (_) {
                    this.previousPhaseSummaries = [];
                }
            } else {
                this.previousPhaseSummaries = [];
            }

            this.render();
            this.renderComparisonPanel();
            this.renderAssessmentActionsPanel();
        } catch (error) {
            console.error('Error loading report:', error);
            if (this.ctx.tenantId) {
                await this.loadPortalAssessment(this.ctx.tenantId, error);
                return;
            }
            this.container.innerHTML = `
                <div class="nb-empty" style="display:block; color:#dc2626;">
                    Fout bij laden rapport: ${this.escapeHtml(error.message)}
                </div>
            `;
        }
    }

    async loadRunJsonReport(runId) {
        try {
            const response = await fetch(`/api/runs/${runId}/assessment-json`);
            const data = await response.json();
            if (!response.ok || !data?.ok || !Array.isArray(data?.phases) || !data.phases.length) {
                return false;
            }
            this.jsonRunData = data;
            this.portalSections = null;
            this.portalNav = null;
            this.renderMode = 'json';
            this.reportData = {
                tenantName: data.tenant_name || this.ctx?.latestRun?.tenant_name || '-',
                tenantId: data.tenant_id || this.ctx?.tenantId || '-',
                assessmentDate: data.generated_at || this.ctx?.latestRun?.completed_at || this.ctx?.latestRun?.started_at || '-',
            };
            this.phases = data.phases.map((phase, index) => ({
                id: String(phase.id || `phase${phase.number || (index + 1)}`),
                number: Number(phase.number || (index + 1)),
                title: String(phase.renderLabel || phase.navLabel || `Phase ${index + 1}`),
                navLabel: String(phase.navLabel || ''),
                summary: String(phase.summary || ''),
                critical: Number(phase.critical || 0),
                warning: Number(phase.warning || 0),
                info: Number(phase.info || 0),
                score: phase.score ?? null,
                payloads: Array.isArray(phase.payloads) ? phase.payloads : [],
                content: '',
            }));
            this.phaseSummaries = this.phases.map((phase) => ({
                number: phase.number,
                title: phase.title,
                score: phase.score,
                critical: phase.critical,
                warning: phase.warning,
                info: phase.info,
            }));
            this.findings = this.extractFindingsFromJsonPhases(this.phases);
            this.previousPhaseSummaries = [];
            this.render();
            this.renderComparisonPanel();
            this.renderAssessmentActionsPanel();
            return true;
        } catch (_) {
            return false;
        }
    }

    async loadPortalAssessment(tenantId, originalError = null) {
        try {
            const nav = await fetch(`/api/assessment/${tenantId}/nav`).then((r) => r.json());
            const navItems = Array.isArray(nav?.items) && nav.items.length
                ? nav.items
                : [{ key: 'summary', label: 'Overzicht', count: null }];
            const sectionResults = await Promise.all(
                navItems.map(async (item) => {
                    const key = item?.key || 'summary';
                    const data = await fetch(`/api/assessment/${tenantId}/section/${encodeURIComponent(key)}`).then((r) => r.json());
                    return [key, data];
                })
            );
            const sectionsByKey = Object.fromEntries(sectionResults);

            this.renderMode = 'metadata';
            this.portalNav = nav || null;
            this.reportData = {
                tenantName: nav?.tenant_name || '-',
                tenantId: nav?.tenant_id || tenantId || '-',
                assessmentDate: nav?.generated_at || '-',
            };
            this.phases = navItems.map((item, index) => {
                const section = sectionsByKey[item.key] || {};
                const rowCount = Array.isArray(section?.rows) ? section.rows.length : 0;
                const cardCount = Array.isArray(section?.cards) ? section.cards.length : 0;
                const barCount = Array.isArray(section?.bars) ? section.bars.length : 0;
                const computedInfo = rowCount || cardCount || barCount || Number(item.count || 0) || 0;
                return {
                    number: index + 1,
                    title: section?.title || item.label || item.key || `Onderdeel ${index + 1}`,
                    navLabel: item.label || section?.title || item.key || `Onderdeel ${index + 1}`,
                    summary: rowCount
                        ? `${rowCount} rij(en) beschikbaar in dit onderdeel.`
                        : (cardCount || barCount
                            ? `${cardCount} kaart(en) en ${barCount} grafiekonderdelen beschikbaar.`
                            : 'Assessmentonderdeel beschikbaar vanuit de laatste run.'),
                    score: item.key === 'summary' ? nav?.score ?? null : null,
                    critical: item.key === 'summary' ? nav?.critical_count || 0 : 0,
                    warning: item.key === 'summary' ? nav?.warning_count || 0 : 0,
                    info: computedInfo,
                    content: '',
                    assessmentKey: item.key,
                    coverage: item?.coverage || null,
                };
            });
            this.phaseSummaries = this.phases.map((phase) => ({
                number: phase.number,
                title: phase.title,
                score: phase.score,
                critical: phase.critical,
                warning: phase.warning,
                info: phase.info,
            }));
            this.findings = [
                ...(nav?.critical_count ? [{ phaseNumber: 1, phaseTitle: 'Assessment overzicht', severity: 'critical', text: `${nav.critical_count} kritieke finding(s) gedetecteerd.` }] : []),
                ...(nav?.warning_count ? [{ phaseNumber: 1, phaseTitle: 'Assessment overzicht', severity: 'warning', text: `${nav.warning_count} waarschuwing(en) gedetecteerd.` }] : []),
            ];
            this.portalSections = { byKey: sectionsByKey, originalError };
            this.render();
            this.renderComparisonPanel();
            this.renderAssessmentActionsPanel();
        } catch (err) {
            this.container.innerHTML = `
                <div class="nb-empty" style="display:block; color:#dc2626;">
                    Assessmentdata laden mislukt: ${this.escapeHtml(err.message || originalError?.message || 'Onbekende fout')}
                </div>
            `;
        }
    }

    async loadReportMetadata(reportPath, doc = null) {
        const sidecarPath = String(reportPath || '').replace(/\.html(\?.*)?$/i, '.metadata.json$1');
        if (sidecarPath && sidecarPath !== reportPath) {
            try {
                const metaResp = await fetch(sidecarPath);
                if (metaResp.ok) {
                    return await metaResp.json();
                }
            } catch (_) { }
        }

        if (doc) {
            const metaEl = doc.querySelector('#report-metadata');
            if (metaEl) {
                try { return JSON.parse(metaEl.textContent); } catch (_) { }
            }
        }
        return null;
    }

    extractPhases(doc, jsonMeta = null) {
        this.phases = [];
        this.findings = [];
        this.portalNav = null;

        // Fallback to embedded metadata if caller didn't provide JSON.
        if (!jsonMeta) {
            const metaEl = doc.querySelector('#report-metadata');
            if (metaEl) {
                try { jsonMeta = JSON.parse(metaEl.textContent); } catch (_) { }
            }
        }

        const tenantNameEl = doc.querySelector('.tenant-name');
        const tenantIdEl = doc.querySelector('.tenant-id');
        const assessmentDateEl = doc.querySelector('.assessment-date');
        const latest = this.ctx.latestRun || {};
        this.reportData = {
            tenantName: jsonMeta ? (jsonMeta.tenantName || '-')
                : (tenantNameEl ? tenantNameEl.textContent.replace('Tenant: ', '').trim() : (latest.tenant_name || latest.customer_name || '-')),
            tenantId: jsonMeta ? (jsonMeta.tenantId || '-')
                : (tenantIdEl ? tenantIdEl.textContent.replace('Tenant ID: ', '').trim() : '-'),
            assessmentDate: jsonMeta ? (jsonMeta.assessmentDate || '-')
                : (assessmentDateEl ? assessmentDateEl.textContent.replace('Assessment Date: ', '').trim() : (latest.completed_at || latest.started_at || '-')),
        };

        const hasJsonPhases = !!(jsonMeta && Array.isArray(jsonMeta.phases) && jsonMeta.phases.length);
        this.renderMode = hasJsonPhases ? 'metadata' : 'legacy';
        this.activeMetadata = hasJsonPhases ? jsonMeta : null;

        if (hasJsonPhases) {
            const sorted = [...jsonMeta.phases].sort((a, b) => Number(a.number || 0) - Number(b.number || 0));
            sorted.forEach((phase, index) => {
                const phaseNumber = Number(phase.number || (index + 1));
                this.phases.push({
                    id: String(phase.id || `phase${phaseNumber}`),
                    number: phaseNumber,
                    title: String(phase.renderLabel || phase.navLabel || `Phase ${phaseNumber}`),
                    navLabel: String(phase.navLabel || ''),
                    summary: String(phase.summary || ''),
                    critical: Number(phase.critical || 0),
                    warning: Number(phase.warning || 0),
                    info: Number(phase.info || 0),
                    score: phase.score ?? null,
                    content: '',
                });
            });
            this.findings = this.extractFindingsFromMetadata(sorted);
        } else {
            const phaseElements = doc.querySelectorAll('.phase-content');
            phaseElements.forEach((phaseEl, index) => {
                const phaseId = phaseEl.id || `phase${index + 1}`;
                const phaseTitle = this.extractPhaseTitle(phaseEl, index + 1);
                const phaseNumber = index + 1;
                const phaseContent = phaseEl.innerHTML;
                this.phases.push({ id: phaseId, number: phaseNumber, title: phaseTitle, content: phaseContent });
                this.findings.push(...this.extractFindingsFromPhase(phaseEl, phaseTitle, phaseNumber));
            });
        }

        this.phaseSummaries = this.extractPhaseSummaries(doc, jsonMeta);
    }

    extractPhaseSummaries(doc, jsonMeta = null) {
        if (jsonMeta && Array.isArray(jsonMeta.phases) && jsonMeta.phases.length > 0) {
            return jsonMeta.phases.map((p, index) => ({
                number: p.number || (index + 1),
                title: p.renderLabel || p.navLabel || `Phase ${index + 1}`,
                score: p.score ?? null,
                critical: p.critical || 0,
                warning: p.warning || 0,
                info: p.info || 0,
            }));
        }
        // Fallback: DOM scraping for older reports (pre-JSON)
        const phaseElements = Array.from(doc.querySelectorAll('.phase-content'));
        return phaseElements.map((phaseEl, index) => {
            const phaseTitle = this.extractPhaseTitle(phaseEl, index + 1);
            const readCount = (...selectors) => {
                for (const sel of selectors) {
                    const text = phaseEl.querySelector(sel)?.textContent || '';
                    const match = text.match(/(\d+)/);
                    if (match) return Number(match[1]);
                }
                return 0;
            };
            const scoreText = phaseEl.querySelector('.sev-score')?.textContent || '';
            const scoreMatch = scoreText.match(/(\d+)/);
            return {
                number: index + 1,
                title: phaseTitle,
                score: scoreMatch ? Number(scoreMatch[1]) : null,
                // generator uses sev-chip--danger/warn; older reports may use critical/warning
                critical: readCount('.sev-chip--danger', '.sev-chip--critical'),
                warning: readCount('.sev-chip--warn', '.sev-chip--warning'),
                info: readCount('.sev-chip--info'),
            };
        });
    }

    extractFindingsFromPhase(phaseEl, phaseTitle, phaseNumber) {
        const nodes = Array.from(phaseEl.querySelectorAll('.alert, .recommendation'));
        return nodes.map((node, index) => {
            const raw = (node.textContent || '').replace(/\s+/g, ' ').trim();
            let severity = 'warning';
            if (node.className.includes('critical')) severity = 'critical';
            else if (node.className.includes('info')) severity = 'info';
            else if (node.className.includes('warning')) severity = 'warning';
            const text = raw.replace(/^Aanbeveling:\s*/i, '');
            return {
                id: `phase-${phaseNumber}-finding-${index + 1}`,
                phaseNumber,
                phaseTitle,
                severity,
                text,
            };
        }).filter((item) => item.text);
    }

    extractFindingsFromMetadata(phases) {
        if (!Array.isArray(phases)) return [];
        return phases.map((phase, index) => {
            const text = String(phase.summary || '').trim();
            if (!text) return null;
            let severity = 'info';
            if (Number(phase.critical || 0) > 0) severity = 'critical';
            else if (Number(phase.warning || 0) > 0) severity = 'warning';
            return {
                id: `meta-phase-${phase.number || (index + 1)}-summary`,
                phaseNumber: Number(phase.number || (index + 1)),
                phaseTitle: String(phase.renderLabel || phase.navLabel || `Phase ${index + 1}`),
                severity,
                text,
            };
        }).filter(Boolean);
    }

    extractFindingsFromJsonPhases(phases) {
        if (!Array.isArray(phases)) return [];
        const findings = [];
        phases.forEach((phase, index) => {
            if (phase.summary) {
                findings.push({
                    id: `json-phase-${phase.number || (index + 1)}-summary`,
                    phaseNumber: Number(phase.number || (index + 1)),
                    phaseTitle: String(phase.title || `Phase ${index + 1}`),
                    severity: Number(phase.critical || 0) > 0 ? 'critical' : (Number(phase.warning || 0) > 0 ? 'warning' : 'info'),
                    text: String(phase.summary),
                });
            }
            (phase.payloads || []).forEach((payload) => {
                const notes = (payload?.meta?.notes || []).filter(Boolean);
                if (notes[0]) {
                    findings.push({
                        id: `json-phase-${phase.number}-${payload.section}-${payload.subsection}`,
                        phaseNumber: Number(phase.number || (index + 1)),
                        phaseTitle: String(phase.title || `Phase ${index + 1}`),
                        severity: 'info',
                        text: `${payload.label || `${payload.section}/${payload.subsection}`}: ${notes[0]}`,
                    });
                }
            });
        });
        return findings;
    }

    extractPhaseTitle(phaseEl, phaseNumber) {
        const h1 = phaseEl.querySelector('h1');
        if (h1) return h1.textContent.trim();
        const defaults = {
            1: 'Phase 1: Users, Licensing & Security Basics',
            2: 'Phase 2: Collaboration & Storage',
            3: 'Phase 3: Compliance & Security Policies',
            4: 'Phase 4: Advanced Security & Compliance',
            5: 'Phase 5: Intune Configuration',
            6: 'Phase 6: Azure Infrastructure',
        };
        return defaults[phaseNumber] || `Phase ${phaseNumber}`;
    }

    render() {
        if (!this.container) return;
        if (!this.phases.length) {
            this.container.innerHTML = '<p class="nb-empty" style="display:block;">Geen fasen gevonden in rapport.</p>';
            return;
        }

        this.container.innerHTML = `
            <div class="results-report-workspace">
                <div class="nb-content results-phase-content" id="resultsPhaseContent">
                    <div class="nb-pane active" data-rv-pane="overview">
                        ${this.renderOverviewPane()}
                    </div>
                    ${this.renderPhasePanes()}
                </div>
            </div>
        `;

        this.addStyles();
        this.enhanceEmbeddedContent();
        this.bindEvents();
        if (typeof window.updateSubnav === 'function' &&
            document.getElementById('resultsSection')?.classList.contains('active')) {
            window.updateSubnav('results', this.currentTab || 'overview');
        }
    }

    renderPhaseTabs() {
        return this.phases.map((phase) => `
            <a href="#" class="nb-tab" data-rv-tab="phase-${phase.number}">
                <span class="nb-tab-icon">${phase.number}</span>
                <span class="nb-tab-label">${this.escapeHtml(this.getShortTitle(phase.title))}</span>
            </a>
        `).join('');
    }

    renderOverviewPane() {
        const latest = this.ctx.latestRun || {};
        const summary = this.ctx.summary || {};
        const reportUrl = this.ctx.latestReportUrl || '';
        const csvUrl = this.ctx.latestCsvUrl || '';
        const fmtDate = this.ctx.formatDate || ((v) => (v || '-'));
        const fmtPhases = this.ctx.formatPhaseList || ((v) => (Array.isArray(v) ? v.join(', ') : '-'));
        const statusBadge = this.ctx.statusBadge || ((s) => this.escapeHtml(s || '-'));
        const esc = this.ctx.escapeHtml || this.escapeHtml.bind(this);
        const coveragePanel = this.renderCoverageOverview();

        return `
            <article class="results-overview-card results-overview-kpis">
                <div class="results-overview-kpi-grid">
                    <div class="results-overview-kpi-item">
                        <span>Totale score</span>
                        <strong>${esc(latest.score_overall ?? '—')}</strong>
                    </div>
                    <div class="results-overview-kpi-item results-overview-kpi-item--danger">
                        <span>Kritiek</span>
                        <strong>${esc(latest.critical_count ?? 0)}</strong>
                    </div>
                    <div class="results-overview-kpi-item results-overview-kpi-item--warn">
                        <span>Waarschuwingen</span>
                        <strong>${esc(latest.warning_count ?? 0)}</strong>
                    </div>
                    <div class="results-overview-kpi-item results-overview-kpi-item--info">
                        <span>Info</span>
                        <strong>${esc(latest.info_count ?? 0)}</strong>
                    </div>
                </div>
                <div class="results-overview-actions">
                    <button type="button" class="nb-btn nb-btn-primary nb-btn-sm" onclick="window.open('${esc(reportUrl)}','_blank')" ${reportUrl ? '' : 'disabled'}>↗ Open rapport</button>
                    <button type="button" class="nb-btn nb-btn-secondary nb-btn-sm" onclick="window.open('${esc(csvUrl)}','_blank')" ${csvUrl ? '' : 'disabled'}>↓ CSV</button>
                </div>
            </article>

            <div class="results-overview-grid">
                <article class="results-overview-card">
                    <h4>Run details</h4>
                    <div class="results-overview-rows">
                        <div class="results-overview-row"><span>Run ID</span><strong>${esc(latest.id || '-')}</strong></div>
                        <div class="results-overview-row"><span>Status</span><strong>${statusBadge(latest.status)}</strong></div>
                        <div class="results-overview-row"><span>Gestart</span><strong>${esc(fmtDate(latest.started_at))}</strong></div>
                        <div class="results-overview-row"><span>Voltooid</span><strong>${esc(fmtDate(latest.completed_at))}</strong></div>
                        <div class="results-overview-row"><span>Run mode</span><strong>${esc(latest.run_mode || '-')}</strong></div>
                        <div class="results-overview-row"><span>Fases</span><strong>${esc(fmtPhases(latest.phases))}</strong></div>
                    </div>
                </article>

                <article class="results-overview-card">
                    <h4>Tenant overzicht</h4>
                    <div class="results-overview-rows">
                        <div class="results-overview-row"><span>Tenant</span><strong>${esc(latest.tenant_name || this.reportData.tenantName || '-')}</strong></div>
                        <div class="results-overview-row"><span>Klant</span><strong>${esc(latest.customer_name || '-')}</strong></div>
                        <div class="results-overview-row"><span>Runs totaal</span><strong>${esc(summary.totalRuns || 0)}</strong></div>
                        <div class="results-overview-row"><span>Voltooid</span><strong>${esc(summary.completedRuns || 0)}</strong></div>
                        <div class="results-overview-row"><span>Mislukt</span><strong>${esc(summary.failedRuns || 0)}</strong></div>
                        <div class="results-overview-row"><span>Rapport-runs</span><strong>${esc(summary.reportRuns || 0)}</strong></div>
                    </div>
                </article>
            </div>

            <article class="results-overview-card results-overview-runs">
                <div class="results-overview-runs-head">
                    <h4>Recente rapport-runs</h4>
                    <span>${esc((this.ctx.reportRuns || []).length)} item(s)</span>
                </div>
                <div class="results-runs-table-wrap">
                    <table class="results-runs-table">
                        <thead>
                            <tr>
                                <th>Datum</th>
                                <th>Status</th>
                                <th>Mode</th>
                                <th>Fases</th>
                                <th>Score</th>
                                <th>K/W/I</th>
                                <th>Acties</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${this.renderRunsRows()}
                        </tbody>
                    </table>
                </div>
            </article>
            ${coveragePanel}
        `;
    }

    renderRunsRows() {
        const runs = this.ctx.reportRuns || [];
        const fmtDate = this.ctx.formatDate || ((v) => (v || '-'));
        const fmtPhases = this.ctx.formatPhaseList || ((v) => (Array.isArray(v) ? v.join(', ') : '-'));
        const statusBadge = this.ctx.statusBadge || ((s) => this.escapeHtml(s || '-'));
        const esc = this.ctx.escapeHtml || this.escapeHtml.bind(this);

        if (!runs.length) {
            return '<tr><td colspan="7" class="empty-state">Geen rapport-runs beschikbaar.</td></tr>';
        }

        return runs.map((r) => `
            <tr>
                <td>${esc(fmtDate(r.completed_at || r.started_at))}</td>
                <td>${statusBadge(r.status)}</td>
                <td>${esc(r.run_mode || '-')}</td>
                <td>${esc(fmtPhases(r.phases))}</td>
                <td>${esc(r.score_overall ?? '-')}</td>
                <td>${esc(r.critical_count ?? 0)} / ${esc(r.warning_count ?? 0)} / ${esc(r.info_count ?? 0)}</td>
                <td>
                    <div class="results-row-actions">
                        <button type="button" class="btn btn-secondary btn-sm" onclick="window.viewRunDetails && window.viewRunDetails('${esc(r.id)}')">Details</button>
                        ${r.report_path ? `<button type="button" class="btn btn-secondary btn-sm" onclick="window.open('${esc(r.report_path)}','_blank')">Rapport</button>` : ''}
                    </div>
                </td>
            </tr>
        `).join('');
    }

    renderPhasePanes() {
        if (this.portalSections) {
            return this.renderPortalAssessmentPanes();
        }
        return this.phases.map((phase) => `
            <div class="nb-pane" data-rv-pane="phase-${phase.number}">
                <article class="results-phase-card">
                    <header class="results-phase-card-head">
                        <h4>${this.escapeHtml(this.getPhaseNavLabel(phase.number, phase.title))}</h4>
                        <p>${this.escapeHtml(phase.title)}</p>
                    </header>
                    <section class="results-phase-body" id="phase-inline-${phase.number}">
                        ${this.renderMode === 'json'
                            ? this.renderJsonPhasePane(phase)
                            : (this.renderMode === 'metadata'
                                ? this.renderMetadataPhasePane(phase)
                                : (phase.content || '<p>Geen data gevonden voor deze fase.</p>'))}
                    </section>
                </article>
            </div>
        `).join('');
    }

    renderPortalAssessmentPanes() {
        return this.phases.map((phase) => {
            const section = this.portalSections?.byKey?.[phase.assessmentKey];
            return `
                <div class="nb-pane" data-rv-pane="phase-${phase.number}">
                    <article class="results-phase-card">
                        <header class="results-phase-card-head">
                            <h4>${this.escapeHtml(phase.navLabel || phase.title)}</h4>
                            <p>${this.escapeHtml(phase.title)}</p>
                        </header>
                        <section class="results-phase-body">
                            ${this.renderCoverageCallout(phase.coverage || section?.coverage || null)}
                            ${this.renderPortalAssessmentSection(section)}
                        </section>
                    </article>
                </div>
            `;
        }).join('');
    }

    renderPortalAssessmentSection(section) {
        if (!section) return '<p class="nb-empty" style="display:block;">Geen data beschikbaar.</p>';
        const cards = Array.isArray(section.cards) && section.cards.length
            ? `<div class="results-meta-phase-grid">${section.cards.map((card) => `<div class="results-meta-kpi"><span>${this.escapeHtml(card.label)}</span><strong>${this.escapeHtml(card.value)}</strong></div>`).join('')}</div>`
            : '';
        const bars = Array.isArray(section.bars) && section.bars.length
            ? `<article class="results-meta-summary-card"><h5>Onderdelen</h5>${section.bars.map((bar) => `<p>${this.escapeHtml(bar.label)}: <strong>${this.escapeHtml(bar.value)}</strong></p>`).join('')}</article>`
            : '';
        const table = Array.isArray(section.rows) && section.rows.length
            ? `
                <div class="results-runs-table-wrap">
                    <table class="results-runs-table">
                        <thead><tr>${(section.columns || []).map((col) => `<th>${this.escapeHtml(col)}</th>`).join('')}</tr></thead>
                        <tbody>
                            ${section.rows.map((row) => `<tr>${Object.values(row || {}).map((value) => `<td>${this.escapeHtml(value || '—')}</td>`).join('')}</tr>`).join('')}
                        </tbody>
                    </table>
                </div>
            `
            : '<p class="nb-empty" style="display:block;">Geen rijen beschikbaar voor dit onderdeel.</p>';
        return `${cards}${bars}${table}`;
    }

    renderCoverageOverview() {
        const items = Array.isArray(this.portalNav?.items)
            ? this.portalNav.items.filter((item) => item?.key && item.key !== 'summary')
            : [];
        if (!items.length) return '';
        const rows = items.map((item) => {
            const coverage = item?.coverage || {};
            const capability = coverage?.capability || null;
            const detail = capability?.status_label || coverage?.bucket_label || 'Onbekend';
            const source = capability?.supports_live
                ? 'Live + assessment'
                : (coverage?.bucket === 'report_only' ? 'Rapport' : 'Assessment');
            return `
                <tr>
                    <td><strong>${this.escapeHtml(item.label || item.key)}</strong><br><span>${this.escapeHtml(item.key)}</span></td>
                    <td>${this.renderCoverageBadge(coverage)}</td>
                    <td>${this.escapeHtml(detail)}</td>
                    <td>${this.escapeHtml(source)}</td>
                    <td>${this.renderCoverageAction(coverage)}</td>
                </tr>
            `;
        }).join('');
        return `
            <article class="results-overview-card results-overview-runs">
                <div class="results-overview-runs-head">
                    <h4>Portal dekking</h4>
                    <span>${this.escapeHtml(items.length)} onderdeel/onderdelen</span>
                </div>
                <div class="results-runs-table-wrap">
                    <table class="results-runs-table">
                        <thead>
                            <tr>
                                <th>Onderdeel</th>
                                <th>Status</th>
                                <th>Type</th>
                                <th>Bron</th>
                                <th>Actie</th>
                            </tr>
                        </thead>
                        <tbody>${rows}</tbody>
                    </table>
                </div>
            </article>
        `;
    }

    renderCoverageCallout(coverage) {
        if (!coverage) return '';
        const capability = coverage.capability || null;
        const detail = capability?.status_reason || coverage.detail || 'Geen aanvullende dekkinginformatie beschikbaar.';
        return `
            <article class="results-coverage-callout">
                <div class="results-coverage-copy">
                    ${this.renderCoverageBadge(coverage)}
                    <p>${this.escapeHtml(detail)}</p>
                </div>
                <div class="results-coverage-actions">
                    ${this.renderCoverageAction(coverage)}
                </div>
            </article>
        `;
    }

    renderCoverageBadge(coverage) {
        const bucket = String(coverage?.bucket || 'report_only');
        const label = coverage?.bucket_label || 'Alleen rapport';
        const classMap = {
            live_workspace: 'results-coverage-badge results-coverage-badge--live',
            snapshot_workspace: 'results-coverage-badge results-coverage-badge--snapshot',
            live_backend_only: 'results-coverage-badge results-coverage-badge--warn',
            snapshot_only: 'results-coverage-badge results-coverage-badge--snapshot',
            not_available: 'results-coverage-badge results-coverage-badge--stale',
            report_only: 'results-coverage-badge results-coverage-badge--neutral',
        };
        return `<span class="${classMap[bucket] || classMap.report_only}">${this.escapeHtml(label)}</span>`;
    }

    renderCoverageAction(coverage) {
        const target = coverage?.open_target || null;
        if (!target?.section) {
            return '<span class="results-coverage-hint">Nog geen directe portalroute</span>';
        }
        const attrs = [
            `data-open-section="${this.escapeHtml(target.section)}"`,
            target.tab_type ? `data-open-tab-type="${this.escapeHtml(target.tab_type)}"` : '',
            target.tab_key ? `data-open-tab-key="${this.escapeHtml(target.tab_key)}"` : '',
        ].filter(Boolean).join(' ');
        return `<button type="button" class="btn btn-secondary btn-sm" ${attrs}>Open in portal</button>`;
    }

    renderMetadataPhasePane(phase) {
        const score = phase.score != null ? String(phase.score) : '—';
        const summaryText = String(phase.summary || '').trim() || 'Geen extra samenvatting beschikbaar voor deze fase.';
        return `
            <div class="results-meta-phase-grid">
                <div class="results-meta-kpi"><span>Score</span><strong>${this.escapeHtml(score)}</strong></div>
                <div class="results-meta-kpi results-meta-kpi--critical"><span>Kritiek</span><strong>${this.escapeHtml(phase.critical || 0)}</strong></div>
                <div class="results-meta-kpi results-meta-kpi--warning"><span>Waarschuwing</span><strong>${this.escapeHtml(phase.warning || 0)}</strong></div>
                <div class="results-meta-kpi results-meta-kpi--info"><span>Info</span><strong>${this.escapeHtml(phase.info || 0)}</strong></div>
            </div>
            <article class="results-meta-summary-card">
                <h5>Samenvatting</h5>
                <p>${this.escapeHtml(summaryText)}</p>
            </article>
            <div class="results-meta-phase-hint">Metadata-only weergave actief: deze fase wordt opgebouwd vanuit <code>.metadata.json</code>.</div>
        `;
    }

    renderJsonPhasePane(phase) {
        const payloads = Array.isArray(phase.payloads) ? phase.payloads : [];
        const header = `
            <div class="results-meta-phase-grid">
                <div class="results-meta-kpi"><span>Score</span><strong>${this.escapeHtml(phase.score != null ? String(phase.score) : '—')}</strong></div>
                <div class="results-meta-kpi results-meta-kpi--critical"><span>Kritiek</span><strong>${this.escapeHtml(phase.critical || 0)}</strong></div>
                <div class="results-meta-kpi results-meta-kpi--warning"><span>Waarschuwing</span><strong>${this.escapeHtml(phase.warning || 0)}</strong></div>
                <div class="results-meta-kpi results-meta-kpi--info"><span>Items</span><strong>${this.escapeHtml(phase.info || 0)}</strong></div>
            </div>
            <article class="results-meta-summary-card">
                <h5>Samenvatting</h5>
                <p>${this.escapeHtml(phase.summary || 'Geen samenvatting beschikbaar.')}</p>
            </article>
        `;
        const sections = payloads.length
            ? payloads.map((payload) => this.renderJsonPayload(payload)).join('')
            : '<p class="nb-empty" style="display:block;">Geen JSON-payloads beschikbaar voor deze fase.</p>';
        return `${header}${sections}`;
    }

    renderJsonPayload(payload) {
        const summary = payload?.summary && typeof payload.summary === 'object' ? payload.summary : {};
        const items = Array.isArray(payload?.items) ? payload.items : [];
        const notes = Array.isArray(payload?.meta?.notes) ? payload.meta.notes.filter(Boolean) : [];
        const permissions = Array.isArray(payload?.meta?.permissions) ? payload.meta.permissions.filter(Boolean) : [];
        const summaryEntries = Object.entries(summary).filter(([, value]) => value !== null && value !== undefined && value !== '');
        const itemColumns = items.length ? Array.from(items.reduce((set, item) => {
            if (item && typeof item === 'object' && !Array.isArray(item)) {
                Object.keys(item).forEach((key) => set.add(key));
            }
            return set;
        }, new Set())) : [];
        const summaryHtml = summaryEntries.length
            ? `<div class="results-meta-phase-grid">${summaryEntries.slice(0, 8).map(([key, value]) => `<div class="results-meta-kpi"><span>${this.escapeHtml(key)}</span><strong>${this.escapeHtml(String(value))}</strong></div>`).join('')}</div>`
            : '';
        const notesHtml = notes.length
            ? `<article class="results-meta-summary-card"><h5>Notities</h5>${notes.map((note) => `<p>${this.escapeHtml(note)}</p>`).join('')}</article>`
            : '';
        const permissionsHtml = permissions.length
            ? `<div class="results-meta-phase-hint">Benodigde rechten: <code>${this.escapeHtml(permissions.join(', '))}</code></div>`
            : '';
        const tableHtml = itemColumns.length
            ? `
                <div class="results-runs-table-wrap">
                    <table class="results-runs-table">
                        <thead><tr>${itemColumns.map((column) => `<th>${this.escapeHtml(column)}</th>`).join('')}</tr></thead>
                        <tbody>
                            ${items.map((row) => `
                                <tr>${itemColumns.map((column) => `<td>${this.escapeHtml(this.formatJsonCell(row?.[column]))}</td>`).join('')}</tr>
                            `).join('')}
                        </tbody>
                    </table>
                </div>
            `
            : '<p class="nb-empty" style="display:block;">Geen rijdata beschikbaar voor dit onderdeel.</p>';
        return `
            <article class="results-meta-summary-card" style="margin-top:14px;">
                <h5>${this.escapeHtml(payload?.label || `${payload?.section || ''}/${payload?.subsection || ''}`)}</h5>
                ${summaryHtml}
                ${notesHtml}
                ${tableHtml}
                ${permissionsHtml}
            </article>
        `;
    }

    formatJsonCell(value) {
        if (value == null || value === '') return '—';
        if (Array.isArray(value)) {
            if (!value.length) return '—';
            return value.map((item) => {
                if (item && typeof item === 'object') {
                    const label = item.displayName || item.userPrincipalName || item.DisplayName || item.UserPrincipalName;
                    return label ? String(label) : JSON.stringify(item);
                }
                return String(item);
            }).join(', ');
        }
        if (typeof value === 'object') {
            const label = value.displayName || value.userPrincipalName || value.DisplayName || value.UserPrincipalName;
            return label ? String(label) : JSON.stringify(value);
        }
        return String(value);
    }

    bindEvents() {
        this.container.querySelectorAll('[data-rv-tab]').forEach((tab) => {
            tab.addEventListener('click', (e) => {
                e.preventDefault();
                const target = tab.dataset.rvTab || 'overview';
                this.switchTab(target);
            });
        });
        this.container.querySelectorAll('[data-open-section]').forEach((btn) => {
            btn.addEventListener('click', () => {
                if (typeof window.showSection !== 'function') return;
                const section = btn.dataset.openSection;
                const tabType = btn.dataset.openTabType;
                const tabKey = btn.dataset.openTabKey;
                const opts = {};
                if (tabType && tabKey) opts[tabType] = tabKey;
                window.showSection(section, opts);
            });
        });
    }

    renderComparisonPanel() {
        const el = document.getElementById('runDiffContainer');
        if (!el) return;

        if (!this.phaseSummaries.length) {
            el.className = 'nb-empty';
            el.innerHTML = 'Geen fasevergelijking beschikbaar.';
            return;
        }

        const activeSummary = this.getActivePhaseSummary();
        const filteredSummaries = activeSummary ? [activeSummary] : this.phaseSummaries;
        const previousMap = new Map((this.previousPhaseSummaries || []).map((item) => [item.number, item]));
        const rows = filteredSummaries.map((current) => {
            const prev = previousMap.get(current.number);
            const scoreDelta = prev && current.score != null && prev.score != null ? current.score - prev.score : null;
            const deltaLabel = scoreDelta == null ? 'Nieuw' : `${scoreDelta > 0 ? '+' : ''}${scoreDelta}`;
            return `
                <tr>
                    <td><strong>${this.escapeHtml(this.getPhaseNavLabel(current.number, current.title))}</strong><br><span>${this.escapeHtml(this.getShortTitle(current.title))}</span></td>
                    <td>${this.escapeHtml(current.score ?? '—')}</td>
                    <td>${prev ? this.escapeHtml(prev.score ?? '—') : '—'}</td>
                    <td>${this.renderDeltaBadge(deltaLabel, scoreDelta)}</td>
                    <td>${current.critical} / ${current.warning} / ${current.info}</td>
                    <td>${prev ? `${prev.critical} / ${prev.warning} / ${prev.info}` : '—'}</td>
                </tr>
            `;
        }).join('');
        const filterNote = activeSummary
            ? `Filter actief op ${this.escapeHtml(this.getPhaseNavLabel(activeSummary.number, activeSummary.title))}.`
            : 'Vergelijk alle assessment-fases met de vorige succesvolle run.';

        el.className = '';
        el.innerHTML = `
            <div class="results-workspace-card">
                <div class="results-workspace-card-head">
                    <h3>Fasevergelijking</h3>
                    <p>${filterNote}</p>
                </div>
                <div class="results-workspace-card-body">
                    <div class="results-runs-table-wrap">
                        <table class="results-runs-table">
                            <thead>
                                <tr>
                                    <th>Fase</th>
                                    <th>Huidig</th>
                                    <th>Vorige</th>
                                    <th>Delta</th>
                                    <th>Huidig K/W/I</th>
                                    <th>Vorige K/W/I</th>
                                </tr>
                            </thead>
                            <tbody>${rows}</tbody>
                        </table>
                    </div>
                </div>
            </div>
        `;
    }

    renderAssessmentActionsPanel() {
        const el = document.getElementById('resultsAssessmentActionsContainer');
        if (!el) return;

        if (!this.findings.length) {
            el.className = 'nb-empty';
            el.innerHTML = 'Geen aanbevelingen of bevindingen gevonden in de assessmentfasen.';
            return;
        }

        const activePhaseNumber = this.getActivePhaseNumber();
        const filteredFindings = activePhaseNumber
            ? this.findings.filter((item) => item.phaseNumber === activePhaseNumber)
            : this.findings;
        const counts = {
            critical: filteredFindings.filter((item) => item.severity === 'critical').length,
            warning: filteredFindings.filter((item) => item.severity === 'warning').length,
            info: filteredFindings.filter((item) => item.severity === 'info').length,
        };

        const PAGE_SIZE = 25;
        const currentPage = this._findingsPage || 0;
        const totalPages = Math.ceil(filteredFindings.length / PAGE_SIZE);
        const pageFindings = filteredFindings.slice(currentPage * PAGE_SIZE, (currentPage + 1) * PAGE_SIZE);

        const rows = pageFindings.map((item) => `
            <tr>
                <td><strong>${this.escapeHtml(this.getPhaseNavLabel(item.phaseNumber, item.phaseTitle))}</strong><br><span>${this.escapeHtml(this.getShortTitle(item.phaseTitle))}</span></td>
                <td>${this.renderSeverityBadge(item.severity)}</td>
                <td>${this.escapeHtml(item.text)}</td>
                <td><button type="button" class="btn btn-secondary btn-sm" onclick="window.showSection && window.showSection('results', { resultsPanel: 'viewer', resultsViewTab: 'phase-${item.phaseNumber}' });">Open fase</button></td>
            </tr>
        `).join('');

        const pagerHtml = totalPages > 1 ? `
            <div class="rv-pager">
                <button class="rv-pager-btn" ${currentPage === 0 ? 'disabled' : ''} onclick="window._rvInstance && (window._rvInstance._findingsPage = ${currentPage - 1}) && window._rvInstance.renderAssessmentActionsPanel()">← Vorige</button>
                <span>Pagina ${currentPage + 1} van ${totalPages} &nbsp;·&nbsp; ${filteredFindings.length} bevindingen</span>
                <button class="rv-pager-btn" ${currentPage >= totalPages - 1 ? 'disabled' : ''} onclick="window._rvInstance && (window._rvInstance._findingsPage = ${currentPage + 1}) && window._rvInstance.renderAssessmentActionsPanel()">Volgende →</button>
            </div>` : '';

        const activePhase = activePhaseNumber ? this.phases.find((phase) => phase.number === activePhaseNumber) : null;
        const filterNote = activePhase
            ? `Toont alleen aanbevelingen voor ${this.escapeHtml(this.getPhaseNavLabel(activePhase.number, activePhase.title))}.`
            : 'Alle aanbevelingen uit de assessmentfasen van het nieuwste rapport.';

        el.className = '';
        el.innerHTML = `
            <div class="results-workspace-card" style="margin-bottom:1rem;">
                <div class="results-workspace-card-head">
                    <h3>Samenvatting bevindingen</h3>
                    <p>${filterNote}</p>
                </div>
                <div class="results-workspace-card-body">
                    <div class="results-overview-kpi-grid">
                        <div class="results-overview-kpi-item results-overview-kpi-item--danger"><span>Kritiek</span><strong>${counts.critical}</strong></div>
                        <div class="results-overview-kpi-item results-overview-kpi-item--warn"><span>Waarschuwingen</span><strong>${counts.warning}</strong></div>
                        <div class="results-overview-kpi-item results-overview-kpi-item--info"><span>Info</span><strong>${counts.info}</strong></div>
                    </div>
                </div>
            </div>
            <div class="results-workspace-card">
                <div class="results-workspace-card-head">
                    <h3>Assessment acties</h3>
                    <p>Direct uit het rapport opgehaalde aanbevelingen per fase.</p>
                </div>
                <div class="results-workspace-card-body">
                    <div class="results-runs-table-wrap">
                        <table class="results-runs-table">
                            <thead>
                                <tr>
                                    <th>Fase</th>
                                    <th>Severity</th>
                                    <th>Bevinding</th>
                                    <th>Actie</th>
                                </tr>
                            </thead>
                            <tbody>${rows}</tbody>
                        </table>
                    </div>
                    ${pagerHtml}
                </div>
            </div>
        `;
    }

    renderSeverityBadge(severity) {
        const map = {
            critical: 'results-severity-badge results-severity-badge--critical',
            warning: 'results-severity-badge results-severity-badge--warning',
            info: 'results-severity-badge results-severity-badge--info',
        };
        const label = severity === 'critical' ? 'Critical' : severity === 'warning' ? 'Warning' : 'Info';
        return `<span class="${map[severity] || map.info}">${label}</span>`;
    }

    renderDeltaBadge(label, value) {
        let className = 'results-delta-badge';
        if (typeof value === 'number') {
            if (value > 0) className += ' results-delta-badge--up';
            if (value < 0) className += ' results-delta-badge--down';
        }
        return `<span class="${className}">${this.escapeHtml(label)}</span>`;
    }

    switchTab(tabKey) {
        this.currentTab = tabKey;
        this.currentPhaseFilter = tabKey;
        this.currentPanel = 'viewer';
        this.container.querySelectorAll('[data-rv-tab]').forEach((tab) => {
            tab.classList.toggle('active', tab.dataset.rvTab === tabKey);
        });
        this.container.querySelectorAll('[data-rv-pane]').forEach((pane) => {
            pane.classList.toggle('active', pane.dataset.rvPane === tabKey);
        });
        if (typeof window.setActiveSubnavItem === 'function') {
            window.setActiveSubnavItem(tabKey);
        }
    }

    setPanelContext(panelName, { viewerTab, phaseFilter } = {}) {
        this.currentPanel = panelName || this.currentPanel || 'viewer';
        if (phaseFilter !== undefined) {
            this.currentPhaseFilter = phaseFilter || 'overview';
        }

        if (this.currentPanel === 'viewer') {
            this.switchTab(viewerTab || this.currentPhaseFilter || 'overview');
            return;
        }

        if (typeof window.setActiveSubnavItem === 'function') {
            const activeKey = this.currentPhaseFilter && this.currentPhaseFilter !== 'overview'
                ? this.currentPhaseFilter
                : this.currentPanel;
            window.setActiveSubnavItem(activeKey);
        }

        if (this.currentPanel === 'diff') this.renderComparisonPanel();
        if (this.currentPanel === 'actions') this.renderAssessmentActionsPanel();
    }

    getPortalSubnavItems(allowedPanels = ['viewer', 'diff', 'actions']) {
        const items = [];
        if (allowedPanels.includes('viewer')) {
            items.push({ label: 'Overzicht', resultsViewTab: 'overview' });
            this.phases.forEach((phase) => {
                const summary = this.phaseSummaries.find((item) => item.number === phase.number);
                const badgeText = this.getPhaseBadgeText(summary);
                items.push({
                    label: this.getPhaseNavLabel(phase.number, phase.title),
                    resultsViewTab: `phase-${phase.number}`,
                    countText: badgeText,
                });
            });
        }
        if (allowedPanels.includes('diff')) {
            items.push({ label: 'Vergelijking', resultsPanel: 'diff' });
        }
        if (allowedPanels.includes('actions')) {
            items.push({ label: 'Acties', resultsPanel: 'actions' });
        }
        return items;
    }

    getActivePhaseNumber() {
        const match = String(this.currentPhaseFilter || '').match(/^phase-(\d+)$/);
        return match ? Number(match[1]) : null;
    }

    getActivePhaseSummary() {
        const phaseNumber = this.getActivePhaseNumber();
        if (!phaseNumber) return null;
        return this.phaseSummaries.find((item) => item.number === phaseNumber) || null;
    }

    getPhaseNavLabel(phaseNumber, fullTitle) {
        const explicitLabels = {
            1: 'Identiteit',
            2: 'Samenwerking',
            3: 'Compliance',
            4: 'Advanced Security',
            5: 'Intune',
            6: 'Azure',
        };
        if (explicitLabels[phaseNumber]) return explicitLabels[phaseNumber];

        const fallback = this.getShortTitle(fullTitle || `Fase ${phaseNumber}`);
        const normalized = fallback
            .toLowerCase()
            .replace(/[^a-z0-9]+/g, ' ')
            .trim();
        if (normalized.includes('hybrid identity') && normalized.includes('ad connect')) return 'Hybrid Identity';
        if (normalized.includes('hybrid identity')) return 'Hybrid Identity';
        if (normalized.includes('ad connect')) return 'AD Connect';
        if (normalized.includes('cis') && normalized.includes('foundation')) return 'CIS Benchmark';
        if (normalized.includes('compliance') && normalized.includes('matrix')) return 'Compliance Matrix';
        if (normalized.includes('secure score')) return 'Secure Score';
        if (normalized.includes('benchmark')) return 'Benchmark';
        if (normalized.includes('identity') && normalized.includes('access')) return 'Identity & Access';

        const compact = fallback.replace(/\s+/g, ' ').trim();
        return compact.length > 18 ? `${compact.slice(0, 17)}…` : compact;
    }

    getPhaseBadgeText(summary) {
        if (!summary) return '';
        if (summary.critical > 0) return `${summary.critical}K`;
        if (summary.warning > 0) return `${summary.warning}W`;
        if (summary.info > 0) return `${summary.info}I`;
        if (summary.score != null) return `S${summary.score}`;
        return '';
    }

    getShortTitle(fullTitle) {
        const match = String(fullTitle || '').match(/:\s*(.+)$/);
        return match ? match[1] : (fullTitle || 'Onbekende fase');
    }

    escapeHtml(value) {
        return String(value == null ? '' : value)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    enhanceEmbeddedContent() {
        if (this.renderMode === 'metadata') return;
        if (!this.container) return;
        const phaseBodies = this.container.querySelectorAll('.results-phase-body');
        phaseBodies.forEach((body) => {
            body.querySelectorAll('table').forEach((table) => {
                table.classList.add('results-embedded-table');
                const hasWrapper = table.parentElement && table.parentElement.classList.contains('results-table-scroll');
                if (hasWrapper) return;

                const bodyRows = table.querySelectorAll('tbody tr').length || Math.max(0, table.rows.length - (table.tHead?.rows.length || 0));
                if (bodyRows <= 10) return;

                const wrap = document.createElement('div');
                wrap.className = 'results-table-scroll';
                table.parentNode.insertBefore(wrap, table);
                wrap.appendChild(table);
            });
        });
    }

    addStyles() {
        if (document.getElementById('results-viewer-styles')) return;

        const style = document.createElement('style');
        style.id = 'results-viewer-styles';
        style.textContent = `
            .results-report-workspace,
            .results-overview-card,
            .results-phase-card,
            .results-workspace-card,
            .results-phase-body {
                font-family: var(--font, 'Outfit', 'Segoe UI', sans-serif);
            }

            .results-report-workspace {
                display: flex;
                flex-direction: column;
                gap: 12px;
            }

            .results-phase-content {
                min-height: 260px;
            }

            .results-overview-grid {
                display: grid;
                grid-template-columns: repeat(2, minmax(0, 1fr));
                gap: 12px;
                margin-bottom: 12px;
            }

            .results-overview-kpis {
                margin-bottom: 12px;
                padding: 12px;
            }

            .results-overview-kpi-grid {
                display: grid;
                grid-template-columns: repeat(4, minmax(0, 1fr));
                gap: 10px;
            }

            .results-overview-kpi-item {
                border: 1px solid var(--line, #e5e7eb);
                border-radius: 10px;
                padding: 10px 12px;
                background: var(--surface, #fff);
                display: flex;
                flex-direction: column;
                gap: 6px;
            }

            .results-overview-kpi-item span {
                font-size: 11px;
                color: var(--muted, #64748b);
                text-transform: uppercase;
                letter-spacing: 0.08em;
                font-family: var(--mono, 'JetBrains Mono', Consolas, monospace);
            }

            .results-overview-kpi-item strong {
                font-size: 28px;
                line-height: 1;
                color: var(--text, #0f172a);
            }

            .results-overview-kpi-item--danger strong { color: #dc2626; }
            .results-overview-kpi-item--warn strong { color: #d97706; }
            .results-overview-kpi-item--info strong { color: #2563eb; }

            .results-overview-actions {
                display: flex;
                justify-content: flex-end;
                gap: 8px;
                margin-top: 10px;
            }

            .results-severity-badge,
            .results-delta-badge {
                display: inline-flex;
                align-items: center;
                justify-content: center;
                padding: 4px 10px;
                border-radius: 999px;
                font-size: 11px;
                font-weight: 700;
                line-height: 1;
                background: #e5e7eb;
                color: #334155;
            }

            .results-severity-badge--critical,
            .results-delta-badge--down {
                background: #fee2e2;
                color: #b91c1c;
            }

            .results-severity-badge--warning {
                background: #fef3c7;
                color: #b45309;
            }

            .results-severity-badge--info,
            .results-delta-badge--up {
                background: #dbeafe;
                color: #1d4ed8;
            }

            .results-overview-card,
            .results-phase-card {
                border: 1px solid var(--line, #e5e7eb);
                border-radius: 12px;
                background: var(--surface, #fff);
                overflow: hidden;
            }

            .results-overview-card h4,
            .results-phase-card-head h4 {
                margin: 0;
                padding: 14px 16px;
                border-bottom: 1px solid var(--line, #e5e7eb);
                font-size: 15px;
                color: var(--text, #1f2937);
                background: rgba(247, 148, 29, 0.08);
            }

            .results-overview-rows {
                padding: 10px 16px;
            }

            .results-overview-row {
                display: flex;
                justify-content: space-between;
                gap: 10px;
                border-bottom: 1px solid var(--line, #eef2f7);
                padding: 8px 0;
                font-size: 13px;
            }

            .results-overview-row:last-child {
                border-bottom: 0;
            }

            .results-overview-runs {
                margin-top: 2px;
            }

            .results-coverage-callout {
                display: flex;
                align-items: center;
                justify-content: space-between;
                gap: 12px;
                border: 1px solid var(--line, #e5e7eb);
                border-radius: 12px;
                background: rgba(15, 23, 42, 0.03);
                padding: 12px 14px;
                margin-bottom: 14px;
            }

            .results-coverage-copy {
                display: flex;
                flex-direction: column;
                gap: 8px;
            }

            .results-coverage-copy p,
            .results-coverage-hint {
                margin: 0;
                color: var(--muted, #64748b);
                font-size: 13px;
            }

            .results-coverage-badge {
                display: inline-flex;
                align-items: center;
                justify-content: center;
                width: fit-content;
                padding: 6px 10px;
                border-radius: 999px;
                font-size: 11px;
                font-weight: 700;
                line-height: 1;
            }

            .results-coverage-badge--live {
                background: #dcfce7;
                color: #166534;
            }

            .results-coverage-badge--snapshot {
                background: #dbeafe;
                color: #1d4ed8;
            }

            .results-coverage-badge--warn {
                background: #fef3c7;
                color: #b45309;
            }

            .results-coverage-badge--stale {
                background: #fee2e2;
                color: #b91c1c;
            }

            .results-coverage-badge--neutral {
                background: #e5e7eb;
                color: #334155;
            }

            .results-overview-runs-head {
                display: flex;
                justify-content: space-between;
                align-items: center;
                padding: 14px 16px;
                border-bottom: 1px solid var(--line, #e5e7eb);
                background: rgba(15, 23, 42, 0.03);
            }

            .results-overview-runs-head h4 {
                padding: 0;
                margin: 0;
                border: 0;
                background: transparent;
            }

            .results-overview-runs-head span {
                font-size: 12px;
                color: var(--muted, #64748b);
            }

            .results-phase-card-head p {
                margin: 0;
                padding: 8px 16px 14px;
                font-size: 13px;
                color: var(--muted, #6b7280);
                background: rgba(247, 148, 29, 0.08);
            }

            .results-phase-body {
                padding: 16px;
                color: var(--text, #1f2937);
                overflow-x: auto;
                line-height: 1.45;
            }

            .results-phase-body > *:first-child {
                margin-top: 0 !important;
            }

            .results-phase-body h1,
            .results-phase-body h2,
            .results-phase-body h3,
            .results-phase-body h4 {
                color: var(--text, #0f172a);
                margin: 0 0 10px;
                line-height: 1.25;
            }

            .results-phase-body h1 { font-size: 34px; font-weight: 800; letter-spacing: -0.02em; }
            .results-phase-body h2 { font-size: 26px; font-weight: 750; letter-spacing: -0.01em; }
            .results-phase-body h3 { font-size: 20px; font-weight: 700; }
            .results-phase-body h4 { font-size: 16px; font-weight: 700; }

            .results-phase-body p,
            .results-phase-body li,
            .results-phase-body label,
            .results-phase-body span,
            .results-phase-body small {
                color: var(--text, #1f2937);
            }

            .results-phase-body ul,
            .results-phase-body ol {
                padding-left: 20px;
                margin: 8px 0 14px;
            }

            .results-phase-body .summary-cards,
            .results-phase-body .kpi-grid,
            .results-phase-body .score-grid {
                margin: 10px 0 14px;
            }

            .results-meta-phase-grid {
                display: grid;
                grid-template-columns: repeat(4, minmax(0, 1fr));
                gap: 10px;
                margin-bottom: 14px;
            }

            .results-meta-kpi {
                border: 1px solid var(--line, #e5e7eb);
                border-radius: 10px;
                padding: 10px 12px;
                background: var(--surface, #fff);
                display: flex;
                flex-direction: column;
                gap: 5px;
            }

            .results-meta-kpi span {
                font-size: 11px;
                color: var(--muted, #64748b);
                text-transform: uppercase;
                letter-spacing: 0.08em;
                font-family: var(--mono, 'JetBrains Mono', Consolas, monospace);
            }

            .results-meta-kpi strong {
                font-size: 20px;
                color: var(--text, #1f2937);
                line-height: 1;
            }

            .results-meta-kpi--critical strong { color: #dc2626; }
            .results-meta-kpi--warning strong { color: #d97706; }
            .results-meta-kpi--info strong { color: #2563eb; }

            .results-meta-summary-card {
                border: 1px solid var(--line, #e5e7eb);
                border-radius: 12px;
                background: var(--surface, #fff);
                padding: 14px 16px;
            }

            .results-meta-summary-card h5 {
                margin: 0 0 8px;
                font-size: 15px;
                color: var(--text, #0f172a);
            }

            .results-meta-summary-card p {
                margin: 0;
                font-size: 14px;
                line-height: 1.6;
                color: var(--text, #1f2937);
            }

            .results-meta-phase-hint {
                margin-top: 10px;
                font-size: 12px;
                color: var(--muted, #64748b);
            }

            .results-meta-phase-hint code {
                font-family: var(--mono, 'JetBrains Mono', Consolas, monospace);
                font-size: 11px;
                background: #f3f4f6;
                border: 1px solid #e5e7eb;
                border-radius: 6px;
                padding: 1px 6px;
                color: #334155;
            }

            .results-phase-body table {
                width: 100%;
                border-collapse: collapse;
                margin: 12px 0;
            }

            .results-table-scroll {
                max-height: 520px;
                overflow: auto;
                border: 1px solid var(--line, #e5e7eb);
                border-radius: 10px;
                background: var(--surface, #fff);
                margin: 12px 0;
            }

            .results-table-scroll > table {
                margin: 0;
                border: 0;
            }

            .results-table-scroll thead th {
                position: sticky;
                top: 0;
                z-index: 2;
                background: #f8fafc;
                box-shadow: 0 1px 0 var(--line, #e5e7eb);
            }

            .results-phase-body .results-embedded-table th {
                font-size: 11px;
                text-transform: uppercase;
                letter-spacing: 0.08em;
                color: var(--muted, #64748b);
                background: #f8fafc;
                font-family: var(--mono, 'JetBrains Mono', Consolas, monospace);
                font-weight: 600;
            }

            .results-phase-body .results-embedded-table td {
                font-size: 13px;
                color: var(--text, #1f2937);
                background: var(--surface, #fff);
            }

            .results-phase-body .results-embedded-table tbody tr:nth-child(even) td {
                background: #fcfdff;
            }

            .results-phase-body th,
            .results-phase-body td {
                border: 1px solid var(--line, #e5e7eb);
                padding: 8px;
                text-align: left;
                vertical-align: top;
            }

            [data-theme="dark"] .results-overview-card,
            [data-theme="dark"] .results-phase-card {
                background: #101f33;
                border-color: #1a2a41;
                color: #e6eef9;
            }

            [data-theme="dark"] .results-overview-kpi-item {
                background: #0f1a2b;
                border-color: #1a2a41;
            }

            [data-theme="dark"] .results-overview-kpi-item strong {
                color: #e6eef9;
            }

            [data-theme="dark"] .results-coverage-callout {
                background: rgba(16, 35, 58, 0.8);
                border-color: #1a2a41;
            }

            [data-theme="dark"] .results-coverage-copy p,
            [data-theme="dark"] .results-coverage-hint {
                color: #9fb6d1;
            }

            [data-theme="dark"] .results-coverage-badge--live {
                background: rgba(34, 197, 94, 0.18);
                color: #86efac;
            }

            [data-theme="dark"] .results-coverage-badge--snapshot {
                background: rgba(59, 130, 246, 0.2);
                color: #93c5fd;
            }

            [data-theme="dark"] .results-coverage-badge--warn {
                background: rgba(245, 158, 11, 0.2);
                color: #fcd34d;
            }

            [data-theme="dark"] .results-coverage-badge--stale {
                background: rgba(239, 68, 68, 0.2);
                color: #fca5a5;
            }

            [data-theme="dark"] .results-coverage-badge--neutral {
                background: #1f2937;
                color: #e5e7eb;
            }

            [data-theme="dark"] .results-severity-badge,
            [data-theme="dark"] .results-delta-badge {
                background: #1f2937;
                color: #e5e7eb;
            }

            [data-theme="dark"] .results-severity-badge--critical,
            [data-theme="dark"] .results-delta-badge--down {
                background: rgba(239, 68, 68, 0.2);
                color: #fca5a5;
            }

            [data-theme="dark"] .results-severity-badge--warning {
                background: rgba(245, 158, 11, 0.2);
                color: #fcd34d;
            }

            [data-theme="dark"] .results-severity-badge--info,
            [data-theme="dark"] .results-delta-badge--up {
                background: rgba(59, 130, 246, 0.2);
                color: #93c5fd;
            }

            [data-theme="dark"] .results-overview-kpi-item--danger strong { color: #f87171; }
            [data-theme="dark"] .results-overview-kpi-item--warn strong { color: #fbbf24; }
            [data-theme="dark"] .results-overview-kpi-item--info strong { color: #60a5fa; }

            [data-theme="dark"] .results-overview-card h4,
            [data-theme="dark"] .results-phase-card-head h4,
            [data-theme="dark"] .results-phase-card-head p,
            [data-theme="dark"] .results-overview-runs-head {
                background: rgba(247, 148, 29, 0.14);
                border-color: #1a2a41;
                color: #e6eef9;
            }

            [data-theme="dark"] .results-overview-row {
                border-color: #1a2a41;
            }

            [data-theme="dark"] .results-phase-body {
                color: #e6eef9;
            }

            [data-theme="dark"] .results-phase-body h1,
            [data-theme="dark"] .results-phase-body h2,
            [data-theme="dark"] .results-phase-body h3,
            [data-theme="dark"] .results-phase-body h4,
            [data-theme="dark"] .results-phase-body p,
            [data-theme="dark"] .results-phase-body li,
            [data-theme="dark"] .results-phase-body label,
            [data-theme="dark"] .results-phase-body span,
            [data-theme="dark"] .results-phase-body small {
                color: #e6eef9;
            }

            [data-theme="dark"] .results-table-scroll {
                border-color: #1a2a41;
                background: #0f1a2b;
            }

            [data-theme="dark"] .results-table-scroll thead th,
            [data-theme="dark"] .results-phase-body .results-embedded-table th {
                background: #10233a;
                color: #9fb6d1;
                box-shadow: 0 1px 0 #1a2a41;
            }

            [data-theme="dark"] .results-phase-body .results-embedded-table td {
                background: #0f1a2b;
                color: #e6eef9;
            }

            [data-theme="dark"] .results-phase-body .results-embedded-table tbody tr:nth-child(even) td {
                background: #0d1828;
            }

            [data-theme="dark"] .results-meta-kpi,
            [data-theme="dark"] .results-meta-summary-card {
                background: #0f1a2b;
                border-color: #1a2a41;
            }

            [data-theme="dark"] .results-meta-kpi strong,
            [data-theme="dark"] .results-meta-summary-card h5,
            [data-theme="dark"] .results-meta-summary-card p {
                color: #e6eef9;
            }

            [data-theme="dark"] .results-meta-phase-hint { color: #9fb6d1; }

            [data-theme="dark"] .results-meta-phase-hint code {
                background: #10233a;
                border-color: #1a2a41;
                color: #cfe0f3;
            }

            [data-theme="dark"] .results-phase-body th,
            [data-theme="dark"] .results-phase-body td {
                border-color: #1a2a41;
            }

            @media (max-width: 900px) {
                .results-overview-grid {
                    grid-template-columns: 1fr;
                }

                .results-overview-kpi-grid {
                    grid-template-columns: repeat(2, minmax(0, 1fr));
                }

                .results-meta-phase-grid {
                    grid-template-columns: repeat(2, minmax(0, 1fr));
                }

                .results-coverage-callout {
                    flex-direction: column;
                    align-items: flex-start;
                }

                .results-overview-actions {
                    justify-content: flex-start;
                    flex-wrap: wrap;
                }
            }
        `;
        document.head.appendChild(style);
    }
}

let resultsViewer = null;

async function initResultsViewer(containerId, reportPath, context = {}) {
    resultsViewer = new ResultsViewer(containerId);
    window.resultsViewer = resultsViewer;
    window._rvInstance = resultsViewer;
    await resultsViewer.loadReport(reportPath, context);
    return resultsViewer;
}
