// Results Viewer - Display HTML report phases in separate tabs

class ResultsViewer {
    constructor(containerId) {
        this.container = document.getElementById(containerId);
        this.currentPhase = 1;
        this.phases = [];
        this.reportData = null;
        this.reportPath = null;
        this.phaseIframeState = new Map();
        this.reportHeadHtml = '';
        this._themeListenerBound = false;
    }

    async loadReport(reportPath) {
        try {
            this.reportPath = reportPath;
            // Fetch the HTML report
            const response = await fetch(reportPath);
            if (!response.ok) {
                throw new Error(`Failed to load report: ${response.statusText}`);
            }

            const htmlContent = await response.text();
            
            // Parse the HTML
            const parser = new DOMParser();
            const doc = parser.parseFromString(htmlContent, 'text/html');

            // Reuse report CSS/links so phase rendering matches original report.
            const headNodes = Array.from(doc.head.querySelectorAll('style, link[rel="stylesheet"], meta[charset], meta[name="viewport"]'));
            this.reportHeadHtml = headNodes.map((n) => n.outerHTML).join('\n');
            
            // Extract phases
            this.extractPhases(doc);
            
            // Render the viewer
            this.render();
            
        } catch (error) {
            console.error('Error loading report:', error);
            this.container.innerHTML = `
                <div class="alert alert-danger">
                    <strong>Fout bij laden rapport:</strong> ${error.message}
                </div>
            `;
        }
    }

    extractPhases(doc) {
        this.phases = [];
        
        // Extract tenant info
        const tenantNameEl = doc.querySelector('.tenant-name');
        const tenantIdEl = doc.querySelector('.tenant-id');
        const assessmentDateEl = doc.querySelector('.assessment-date');
        
        this.reportData = {
            tenantName: tenantNameEl ? tenantNameEl.textContent.replace('Tenant: ', '') : 'Unknown',
            tenantId: tenantIdEl ? tenantIdEl.textContent.replace('Tenant ID: ', '') : 'Unknown',
            assessmentDate: assessmentDateEl ? assessmentDateEl.textContent.replace('Assessment Date: ', '') : 'Unknown'
        };
        
        // Extract each phase
        const phaseElements = doc.querySelectorAll('.phase-content');
        
        phaseElements.forEach((phaseEl, index) => {
            const phaseId = phaseEl.id || `phase${index + 1}`;
            const phaseTitle = this.extractPhaseTitle(phaseEl, index + 1);
            const phaseContent = phaseEl.innerHTML;
            
            this.phases.push({
                id: phaseId,
                number: index + 1,
                title: phaseTitle,
                content: phaseContent
            });
        });
    }

    extractPhaseTitle(phaseEl, phaseNumber) {
        // Try to get the h1 title
        const h1 = phaseEl.querySelector('h1');
        if (h1) {
            return h1.textContent.trim();
        }
        
        // Fallback titles
        const defaultTitles = {
            1: 'Phase 1: Users, Licensing & Security Basics',
            2: 'Phase 2: Collaboration & Storage',
            3: 'Phase 3: Compliance & Security Policies',
            4: 'Phase 4: Advanced Security & Compliance',
            5: 'Phase 5: Intune Configuration',
            6: 'Phase 6: Azure Infrastructure'
        };
        
        return defaultTitles[phaseNumber] || `Phase ${phaseNumber}`;
    }

    render() {
        if (this.phases.length === 0) {
            this.container.innerHTML = '<p class="empty-state">Geen fasen gevonden in rapport</p>';
            return;
        }

        // Create the viewer structure
        const viewerHTML = `
            <div class="results-viewer">
                <div class="viewer-header">
                    <h3>Assessment Rapport</h3>
                    <div class="report-meta">
                        <span><strong>Tenant:</strong> ${this.reportData.tenantName}</span>
                        <span><strong>Datum:</strong> ${this.reportData.assessmentDate}</span>
                    </div>
                </div>
                
                <div class="phase-tabs">
                    ${this.renderTabs()}
                </div>
                
                <div class="phase-panels">
                    ${this.renderPanels()}
                </div>
                
                <div class="viewer-footer">
                    <button class="btn btn-secondary" onclick="resultsViewer.previousPhase()" id="prevPhaseBtn">
                        ← Vorige Fase
                    </button>
                    <button class="btn btn-secondary" onclick="resultsViewer.nextPhase()" id="nextPhaseBtn">
                        Volgende Fase →
                    </button>
                </div>
            </div>
        `;

        this.container.innerHTML = viewerHTML;
        
        // Add styles if not already present
        this.addStyles();

        if (!this._themeListenerBound) {
            document.addEventListener('m365-theme-changed', () => {
                this.phases.forEach((p) => {
                    const iframe = document.getElementById(`phase-frame-${p.number}`);
                    if (iframe) this.syncIframeTheme(iframe);
                });
            });
            this._themeListenerBound = true;
        }
        
        // Show first phase
        this.showPhase(1);
    }

    renderTabs() {
        return this.phases.map(phase => `
            <button class="phase-tab ${phase.number === 1 ? 'active' : ''}" 
                    data-phase="${phase.number}"
                    onclick="resultsViewer.showPhase(${phase.number})">
                <span class="phase-number">Fase ${phase.number}</span>
                <span class="phase-title">${this.getShortTitle(phase.title)}</span>
            </button>
        `).join('');
    }

    renderPanels() {
        return this.phases.map(phase => `
            <div class="phase-panel ${phase.number === 1 ? 'active' : ''}" 
                 data-phase="${phase.number}"
                 id="panel-${phase.number}">
                <div class="phase-iframe-wrap">
                    <div class="phase-iframe-loading" id="phase-loader-${phase.number}">Fase laden...</div>
                    <iframe
                        class="phase-report-frame"
                        id="phase-frame-${phase.number}"
                        title="Rapport fase ${phase.number}"
                        loading="lazy"
                        referrerpolicy="no-referrer"
                    ></iframe>
                </div>
            </div>
        `).join('');
    }

    getShortTitle(fullTitle) {
        // Extract just the descriptive part after the colon
        const match = fullTitle.match(/:\s*(.+)$/);
        return match ? match[1] : fullTitle;
    }

    showPhase(phaseNumber) {
        // Update current phase
        this.currentPhase = phaseNumber;
        
        // Update tabs
        const tabs = document.querySelectorAll('.phase-tab');
        tabs.forEach(tab => {
            if (parseInt(tab.dataset.phase) === phaseNumber) {
                tab.classList.add('active');
            } else {
                tab.classList.remove('active');
            }
        });
        
        // Update panels
        const panels = document.querySelectorAll('.phase-panel');
        panels.forEach(panel => {
            if (parseInt(panel.dataset.phase) === phaseNumber) {
                panel.classList.add('active');
            } else {
                panel.classList.remove('active');
            }
        });

        this.ensurePhaseFrame(phaseNumber);
        
        // Update navigation buttons
        const prevBtn = document.getElementById('prevPhaseBtn');
        const nextBtn = document.getElementById('nextPhaseBtn');
        
        if (prevBtn) {
            prevBtn.disabled = phaseNumber === 1;
        }
        
        if (nextBtn) {
            nextBtn.disabled = phaseNumber === this.phases.length;
        }
        
        // Scroll to top
        this.container.scrollTop = 0;
    }

    ensurePhaseFrame(phaseNumber) {
        if (!this.reportPath) return;
        const phase = this.phases.find(p => p.number === phaseNumber);
        if (!phase) return;

        const iframe = document.getElementById(`phase-frame-${phaseNumber}`);
        const loader = document.getElementById(`phase-loader-${phaseNumber}`);
        if (!iframe) return;

        const state = this.phaseIframeState.get(phaseNumber);
        if (state && state.loaded) {
            this.syncIframeTheme(iframe);
            this.resizeIframe(iframe);
            return;
        }
        if (state && state.loading) return;

        this.phaseIframeState.set(phaseNumber, { loading: true, loaded: false });
        if (loader) loader.style.display = 'block';
        iframe.style.display = 'block';

        const finishLoad = (loaded) => {
            this.phaseIframeState.set(phaseNumber, { loading: false, loaded });
            if (loader) loader.style.display = 'none';
            iframe.style.display = 'block';
        };

        const loadTimeout = setTimeout(() => {
            // Fallback: some browsers/webviews do not reliably fire iframe onload.
            finishLoad(true);
        }, 6000);

        iframe.onload = () => {
            try {
                // Full report view per tab; jump to selected phase anchor.
                this.syncIframeTheme(iframe);
                this.resizeIframe(iframe);
                clearTimeout(loadTimeout);
                finishLoad(true);
            } catch (err) {
                console.error('Error configuring phase iframe:', err);
                clearTimeout(loadTimeout);
                if (loader) {
                    loader.textContent = `Fout bij laden fase: ${err.message}`;
                    loader.style.display = 'block';
                }
                this.phaseIframeState.set(phaseNumber, { loading: false, loaded: false });
            }
        };

        iframe.onerror = () => {
            clearTimeout(loadTimeout);
            if (loader) {
                loader.textContent = 'Fase kon niet worden geladen.';
                loader.style.display = 'block';
            }
            this.phaseIframeState.set(phaseNumber, { loading: false, loaded: false });
        };

        const safeTitle = this.escapeHtml(`${phase.title || `Fase ${phase.number}`}`);
        const theme = this.escapeHtml(document.documentElement.getAttribute('data-theme') || 'light');
        const phaseHtml = phase.content || '<p>Geen data gevonden voor deze fase.</p>';
        const docHtml = `<!doctype html>
<html lang="nl" data-theme="${theme}">
<head>
${this.reportHeadHtml || ''}
<style>
  html, body {
    margin: 0 !important;
    padding: 0 !important;
    min-height: 0 !important;
    background: transparent !important;
    overflow-x: hidden !important;
  }
  body {
    padding-top: 0 !important;
  }
  .header,
  .header-inner,
  .header-nav,
  .viewer-footer,
  .intro-header,
  .toc-title,
  .toc-list,
  .container > :not(.phase-content) {
    display: none !important;
  }
  .container {
    max-width: 100% !important;
    width: 100% !important;
    margin: 0 !important;
    padding: 0 !important;
    background: transparent !important;
    box-shadow: none !important;
    border: 0 !important;
  }
  .phase-content {
    max-width: 100% !important;
    width: 100% !important;
    margin: 0 !important;
    padding: 0 !important;
    border: 0 !important;
    box-shadow: none !important;
    background: transparent !important;
  }
  .phase-content > h1 {
    margin-top: 0 !important;
  }
  .phase-content > .section {
    margin-left: 0 !important;
    margin-right: 0 !important;
  }
  html[data-theme="dark"] body,
  html[data-theme="dark"] .container,
  html[data-theme="dark"] .phase-content {
    background: #0F1A2B !important;
    color: #E6EEF9 !important;
  }
  html[data-theme="dark"] h1,
  html[data-theme="dark"] h2,
  html[data-theme="dark"] h3,
  html[data-theme="dark"] h4,
  html[data-theme="dark"] h5,
  html[data-theme="dark"] h6,
  html[data-theme="dark"] p,
  html[data-theme="dark"] li,
  html[data-theme="dark"] strong,
  html[data-theme="dark"] span,
  html[data-theme="dark"] div {
    color: inherit;
  }
  html[data-theme="dark"] .section,
  html[data-theme="dark"] .section-body,
  html[data-theme="dark"] .card,
  html[data-theme="dark"] .recommendation,
  html[data-theme="dark"] .recommendation-card,
  html[data-theme="dark"] .info-box,
  html[data-theme="dark"] .warning-box,
  html[data-theme="dark"] .critical-box,
  html[data-theme="dark"] .stat-card,
  html[data-theme="dark"] .status-card {
    background: #101F33 !important;
    color: #E6EEF9 !important;
    border-color: #1A2A41 !important;
    box-shadow: none !important;
  }
  html[data-theme="dark"] table,
  html[data-theme="dark"] .table-container {
    background: #0F1A2B !important;
    color: #E6EEF9 !important;
    border-color: #1A2A41 !important;
  }
  html[data-theme="dark"] th {
    background: #101F33 !important;
    color: #C7D3E3 !important;
    border-color: #1A2A41 !important;
  }
  html[data-theme="dark"] td {
    background: #0F1A2B !important;
    color: #C7D3E3 !important;
    border-color: #1A2A41 !important;
  }
  html[data-theme="dark"] a {
    color: #7CC4FF !important;
  }
</style>
<title>${safeTitle}</title>
</head>
<body>
  <div class="phase-content" id="${this.escapeHtml(phase.id)}">
    ${phaseHtml}
  </div>
</body>
</html>`;
        iframe.srcdoc = docHtml;
    }

    syncIframeTheme(iframe) {
        try {
            const doc = iframe.contentDocument;
            if (!doc) return;
            const theme = document.documentElement.getAttribute('data-theme') || 'light';
            doc.documentElement.setAttribute('data-theme', theme);
        } catch (_) {}
    }

    resizeIframe(iframe) {
        try {
            const doc = iframe.contentDocument;
            if (!doc) return;
            const body = doc.body;
            const html = doc.documentElement;
            const h = Math.max(
                body ? body.scrollHeight : 0,
                html ? html.scrollHeight : 0,
                body ? body.offsetHeight : 0,
                html ? html.offsetHeight : 0
            );
            iframe.style.height = `${Math.max(320, h + 8)}px`;
        } catch (_) {}
    }

    escapeHtml(value) {
        return String(value == null ? '' : value)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    previousPhase() {
        if (this.currentPhase > 1) {
            this.showPhase(this.currentPhase - 1);
        }
    }

    nextPhase() {
        if (this.currentPhase < this.phases.length) {
            this.showPhase(this.currentPhase + 1);
        }
    }

    addStyles() {
        // Check if styles are already added
        if (document.getElementById('results-viewer-styles')) {
            return;
        }

        const style = document.createElement('style');
        style.id = 'results-viewer-styles';
        style.textContent = `
            .results-viewer {
                background: white;
                border-radius: 8px;
                box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
                overflow: hidden;
            }

            .viewer-header {
                padding: 20px 30px;
                background: linear-gradient(135deg, #F7941D 0%, #FF8C00 100%);
                color: white;
                border-bottom: 3px solid #E57C0D;
            }

            .viewer-header h3 {
                margin: 0 0 10px 0;
                font-size: 24px;
                font-weight: 600;
            }

            .report-meta {
                display: flex;
                gap: 30px;
                font-size: 14px;
                margin-top: 10px;
            }

            .report-meta span {
                display: flex;
                align-items: center;
                gap: 5px;
            }

            .phase-tabs {
                display: flex;
                background: #f8f9fa;
                border-bottom: 2px solid #dee2e6;
                overflow-x: auto;
                padding: 0 10px;
            }

            .phase-tab {
                flex: 1;
                min-width: 150px;
                padding: 15px 20px;
                border: none;
                background: transparent;
                cursor: pointer;
                transition: all 0.3s ease;
                border-bottom: 3px solid transparent;
                display: flex;
                flex-direction: column;
                align-items: center;
                text-align: center;
            }

            .phase-tab:hover {
                background: rgba(247, 148, 29, 0.1);
            }

            .phase-tab.active {
                background: white;
                border-bottom-color: #F7941D;
                font-weight: 600;
            }

            .phase-number {
                display: block;
                font-size: 12px;
                color: #6c757d;
                margin-bottom: 5px;
                text-transform: uppercase;
                letter-spacing: 0.5px;
            }

            .phase-tab.active .phase-number {
                color: #F7941D;
                font-weight: 700;
            }

            .phase-title {
                display: block;
                font-size: 14px;
                color: #2E2E2E;
            }

            .phase-panels {
                padding: 14px;
                min-height: 500px;
                max-height: none;
                overflow: visible;
                background: #f8fbff;
            }

            .phase-panel {
                display: none;
            }

            .phase-panel.active {
                display: block;
                animation: fadeIn 0.3s ease-in;
            }

            .phase-iframe-wrap {
                border: 1px solid #e6eef7;
                border-radius: 12px;
                background: #fff;
                overflow: hidden;
                box-shadow: 0 4px 14px rgba(15, 23, 42, 0.04);
            }

            .phase-iframe-loading {
                padding: 18px;
                color: #64748b;
                font-size: 13px;
                border-bottom: 1px solid #eef2f7;
                background: #fbfdff;
            }

            .phase-report-frame {
                width: 100%;
                min-height: 320px;
                border: 0;
                display: block;
                background: transparent;
            }

            @keyframes fadeIn {
                from {
                    opacity: 0;
                    transform: translateY(10px);
                }
                to {
                    opacity: 1;
                    transform: translateY(0);
                }
            }

            .viewer-footer {
                padding: 20px 30px;
                background: #f8f9fa;
                border-top: 1px solid #dee2e6;
                display: flex;
                justify-content: space-between;
                gap: 15px;
            }

            .viewer-footer button:disabled {
                opacity: 0.5;
                cursor: not-allowed;
            }

            html[data-theme="dark"] .results-viewer {
                background: #0F1A2B;
                border: 1px solid #1A2A41;
                box-shadow: 0 10px 28px rgba(0, 0, 0, 0.35);
            }

            html[data-theme="dark"] .phase-tabs {
                background: #101F33;
                border-bottom-color: #1A2A41;
            }

            html[data-theme="dark"] .phase-tab {
                color: #C7D3E3;
            }

            html[data-theme="dark"] .phase-tab:hover {
                background: rgba(255, 255, 255, 0.04);
            }

            html[data-theme="dark"] .phase-tab.active {
                background: #0F1A2B;
                border-bottom-color: #F7941D;
            }

            html[data-theme="dark"] .phase-number {
                color: #A9B7C9;
            }

            html[data-theme="dark"] .phase-title {
                color: #E6EEF9;
            }

            html[data-theme="dark"] .phase-panels {
                background: #0F1A2B;
            }

            html[data-theme="dark"] .phase-iframe-wrap {
                background: #0F1A2B;
                border-color: #1A2A41;
                box-shadow: none;
            }

            html[data-theme="dark"] .phase-iframe-loading {
                background: #101F33;
                color: #A9B7C9;
                border-bottom-color: #1A2A41;
            }

            html[data-theme="dark"] .viewer-footer {
                background: #101F33;
                border-top-color: #1A2A41;
            }

            /* Responsive design */
            @media (max-width: 768px) {
                .phase-tabs {
                    flex-wrap: wrap;
                }

                .phase-tab {
                    min-width: 120px;
                }

                .report-meta {
                    flex-direction: column;
                    gap: 10px;
                }
            }
        `;

        document.head.appendChild(style);
    }
}

// Global instance
let resultsViewer = null;

// Initialize viewer when needed
function initResultsViewer(containerId, reportPath) {
    resultsViewer = new ResultsViewer(containerId);
    resultsViewer.loadReport(reportPath);
}
