// Configuration for MSAL (Microsoft Authentication Library)
// This needs to be updated with actual app registration details after creation

const msalConfig = {
    auth: {
        clientId: localStorage.getItem('m365_clientId') || '', // Wordt ingesteld na app-registratie via bootstrap-config.json
        authority: 'https://login.microsoftonline.com/organizations', // Multi-tenant
        redirectUri: window.location.origin + '/dashboard.html'
    },
    cache: {
        cacheLocation: 'sessionStorage',
        storeAuthStateInCookie: false
    }
};

// Required API permissions (read-only)
const loginRequest = {
    scopes: [
        'User.Read.All',
        'Group.Read.All',
        'Directory.Read.All',
        'AuditLog.Read.All',
        'Policy.Read.All',
        'Sites.Read.All',
        'Team.ReadBasic.All',
        'Organization.Read.All',
        'Reports.Read.All',
        'ReportSettings.Read.All',
        'UserAuthenticationMethod.Read.All',
        'SecurityEvents.Read.All',
        'DelegatedAdminRelationship.Read.All',
        'DeviceManagementConfiguration.Read.All',
        'DeviceManagementManagedDevices.Read.All',
        'Policy.Read.ConditionalAccess',
        'Application.ReadWrite.All' // For app registration creation
    ]
};

// Token request for API calls
const tokenRequest = {
    scopes: ['https://graph.microsoft.com/.default']
};

let msalInstance = null;

// Initialize MSAL
function initializeMsal() {
    try {
        const clientId = localStorage.getItem('m365_clientId');

        if (clientId) {
            msalConfig.auth.clientId = clientId;
            msalInstance = new msal.PublicClientApplication(msalConfig);
            return msalInstance;
        }

        // If no clientId in localStorage, try to load bootstrap-config.json
        // and populate localStorage so the UI recognizes the app registration.
        return fetch(window.location.origin + '/web/bootstrap-config.json')
            .then(async (res) => {
                if (!res.ok) return null;
                const cfg = await res.json();
                if (cfg && cfg.appId) {
                    localStorage.setItem('m365_appId', cfg.objectId || cfg.appId);
                    localStorage.setItem('m365_clientId', cfg.appId);
                    localStorage.setItem('m365_appName', cfg.appName || 'M365-Baseline-Assessment');

                    // Update msalConfig with values from bootstrap file when present
                    msalConfig.auth.clientId = cfg.appId;
                    if (cfg.redirectUri) {
                        msalConfig.auth.redirectUri = cfg.redirectUri;
                    }

                    // Validate redirect URIs and store any mismatch info
                    try {
                        validateRedirectUris(cfg);
                    } catch (e) {
                        console.warn('Redirect validation error', e);
                    }

                    msalInstance = new msal.PublicClientApplication(msalConfig);
                    return msalInstance;
                }
                return null;
            })
            .catch((err) => {
                console.error('Error loading bootstrap-config.json:', err);
                return null;
            });
    } catch (error) {
        console.error('Error initializing MSAL:', error);
        return null;
    }
}

// Validate redirect URIs and store mismatch info in localStorage
function validateRedirectUris(cfg) {
    try {
        const configuredRedirect = cfg && cfg.redirectUri ? cfg.redirectUri : null;
        const msalRedirect = msalConfig.auth && msalConfig.auth.redirectUri ? msalConfig.auth.redirectUri : null;
        const expectedRedirect = window.location.origin + '/dashboard.html';

        const mismatches = [];

        if (configuredRedirect && configuredRedirect !== expectedRedirect) {
            mismatches.push({ source: 'bootstrap-config.json', value: configuredRedirect });
        }

        if (msalRedirect && msalRedirect !== expectedRedirect) {
            mismatches.push({ source: 'msalConfig', value: msalRedirect });
        }

        const mismatch = mismatches.length > 0;

        localStorage.setItem('m365_redirect_expected', expectedRedirect);
        localStorage.setItem('m365_redirect_configured', configuredRedirect || '');
        localStorage.setItem('m365_redirect_mismatch', mismatch ? '1' : '0');
        if (mismatch) {
            localStorage.setItem('m365_redirect_mismatch_details', JSON.stringify(mismatches));
        } else {
            localStorage.removeItem('m365_redirect_mismatch_details');
        }

        return { mismatch, expectedRedirect, configuredRedirect, msalRedirect, mismatches };
    } catch (e) {
        console.warn('Redirect URI validation failed', e);
        return { mismatch: false };
    }
}

// Sign in with redirect
async function signIn() {
    try {
        if (!msalInstance) {
            // For initial login before app registration, use a temporary approach
            // This will be replaced with proper app registration flow
            throw new Error('App registration niet gevonden. Maak eerst een app registratie aan.');
        }
        
        await msalInstance.loginRedirect(loginRequest);
    } catch (error) {
        console.error('Login error:', error);
        throw error;
    }
}

// Handle redirect response
async function handleRedirectResponse() {
    try {
        if (!msalInstance) return null;
        
        const response = await msalInstance.handleRedirectPromise();
        if (response) {
            return response;
        }
        
        // Check if user is already logged in
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length > 0) {
            return { account: accounts[0] };
        }
        
        return null;
    } catch (error) {
        console.error('Redirect handling error:', error);
        throw error;
    }
}

// Get access token
async function getAccessToken() {
    try {
        if (!msalInstance) {
            throw new Error('MSAL not initialized');
        }
        
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length === 0) {
            throw new Error('No accounts found');
        }
        
        const request = {
            ...tokenRequest,
            account: accounts[0]
        };
        
        try {
            const response = await msalInstance.acquireTokenSilent(request);
            return response.accessToken;
        } catch (error) {
            if (error instanceof msal.InteractionRequiredAuthError) {
                const response = await msalInstance.acquireTokenRedirect(request);
                return response.accessToken;
            }
            throw error;
        }
    } catch (error) {
        console.error('Token acquisition error:', error);
        throw error;
    }
}

// Sign out
async function signOut() {
    try {
        if (!msalInstance) return;
        
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length > 0) {
            await msalInstance.logoutRedirect({
                account: accounts[0]
            });
        }
    } catch (error) {
        console.error('Logout error:', error);
        throw error;
    }
}

// Get current user info
function getCurrentUser() {
    if (!msalInstance) return null;
    
    const accounts = msalInstance.getAllAccounts();
    return accounts.length > 0 ? accounts[0] : null;
}

// Create app registration via Microsoft Graph API
async function createAppRegistration(accessToken) {
    try {
        const appName = 'M365-Baseline-Assessment-' + Date.now();
        
        // Create application registration
        const appResponse = await fetch('https://graph.microsoft.com/v1.0/applications', {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                displayName: appName,
                signInAudience: 'AzureADMultipleOrgs',
                web: {
                    redirectUris: [
                        window.location.origin + '/dashboard.html',
                        window.location.origin + '/web/index.html'
                    ]
                },
                requiredResourceAccess: [
                    {
                        resourceAppId: '00000003-0000-0000-c000-000000000000', // Microsoft Graph
                        resourceAccess: [
                            { id: 'df021288-bdef-4463-88db-98f22de89214', type: 'Role' }, // User.Read.All
                            { id: '5b567255-7703-4780-807c-7be8301ae99b', type: 'Role' }, // Group.Read.All
                            { id: '7ab1d382-f21e-4acd-a863-ba3e13f7da61', type: 'Role' }, // Directory.Read.All
                            { id: 'b0afded3-3588-46d8-8b3d-9842eff778da', type: 'Role' }, // AuditLog.Read.All
                            { id: '246dd0d5-5bd0-4def-940b-0421030a5b68', type: 'Role' }, // Policy.Read.All
                            { id: '332a536c-c7ef-4017-ab91-336970924f0d', type: 'Role' }, // Sites.Read.All
                            { id: '660b7406-55f1-41ca-a0ed-0b035e182f3e', type: 'Role' }, // Team.ReadBasic.All
                            { id: '498476ce-e0fe-48b0-b801-37ba7e2685c6', type: 'Role' }, // Organization.Read.All
                            { id: '230c1aed-a721-4c5d-9cb4-a90514e508ef', type: 'Role' }, // Reports.Read.All
                            { id: '38d9df27-64da-44fd-b7c5-a6fbac20248f', type: 'Role' }, // UserAuthenticationMethod.Read.All
                            { id: 'bf394140-e372-4bf9-a898-299cfc7564e5', type: 'Role' }, // SecurityEvents.Read.All
                            { id: '0c0bf378-bf22-4481-8f81-9e89a9b4960a', type: 'Role' }, // DelegatedAdminRelationship.Read.All
                            { id: 'dc377aa6-52d8-4e23-b271-2a7ae04cedf3', type: 'Role' }, // DeviceManagementConfiguration.Read.All
                            { id: '2f51be20-0bb4-4fed-bf7b-db946066c75e', type: 'Role' }  // DeviceManagementManagedDevices.Read.All
                        ]
                    }
                ]
            })
        });
        
        if (!appResponse.ok) {
            const errorText = await appResponse.text();
            throw new Error(`Failed to create app: ${errorText}`);
        }
        
        const appData = await appResponse.json();
        
        // Store app details
        localStorage.setItem('m365_appId', appData.id);
        localStorage.setItem('m365_clientId', appData.appId);
        localStorage.setItem('m365_appName', appName);
        
        return {
            appId: appData.id,
            clientId: appData.appId,
            appName: appName
        };
    } catch (error) {
        console.error('Error creating app registration:', error);
        throw error;
    }
}

// Check if app registration exists
function hasAppRegistration() {
    return !!localStorage.getItem('m365_clientId');
}

// Get stored app registration details
function getAppRegistrationDetails() {
    return {
        appId: localStorage.getItem('m365_appId'),
        clientId: localStorage.getItem('m365_clientId'),
        appName: localStorage.getItem('m365_appName')
    };
}

// Clear app registration
function clearAppRegistration() {
    localStorage.removeItem('m365_appId');
    localStorage.removeItem('m365_clientId');
    localStorage.removeItem('m365_appName');
}

// ============================================================
// CSRF token helper - vereist bij POST-aanroepen naar de API
// ============================================================
let _csrfToken = null;

async function getCsrfToken() {
    if (_csrfToken) return _csrfToken;
    try {
        const res = await fetch('/api/health');
        if (!res.ok) throw new Error(`Health check failed: ${res.status}`);
        const data = await res.json();
        _csrfToken = data.csrfToken || null;
        return _csrfToken;
    } catch (err) {
        console.error('Kon CSRF-token niet ophalen:', err);
        return null;
    }
}

// Voert een POST uit naar de lokale API met CSRF-token in de header
async function apiPost(path, body) {
    const token = await getCsrfToken();
    const headers = { 'Content-Type': 'application/json' };
    if (token) headers['X-CSRF-Token'] = token;

    const res = await fetch(path, {
        method: 'POST',
        headers,
        body: JSON.stringify(body)
    });

    if (!res.ok) {
        const err = await res.text();
        throw new Error(`API POST ${path} mislukt (${res.status}): ${err}`);
    }
    return res.json();
}
