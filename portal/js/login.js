// Login page functionality

document.addEventListener('DOMContentLoaded', async () => {
    // Show redirect-URI warning if present
    try {
        const mismatch = localStorage.getItem('m365_redirect_mismatch') === '1';
        if (mismatch) {
            const expected = localStorage.getItem('m365_redirect_expected') || '';
            const configured = localStorage.getItem('m365_redirect_configured') || '';
            const details = localStorage.getItem('m365_redirect_mismatch_details') || '';

            const banner = document.createElement('div');
            banner.style.background = '#fff3cd';
            banner.style.border = '1px solid #ffeeba';
            banner.style.padding = '12px';
            banner.style.marginBottom = '12px';
            banner.style.borderRadius = '6px';
            banner.innerHTML = `
                <strong>Waarschuwing:</strong> Redirect-URI mismatch gedetecteerd.<br>
                Verwacht: <code>${expected}</code><br>
                Gekonfigureerd: <code>${configured}</code><br>
                Controleer dat de redirect-URI in de app-registratie exact overeenkomt met de verwachte waarde.
                <div style="margin-top:8px; display:flex; gap:8px; flex-wrap:wrap;">
                    <button id="openBootstrapConfig" class="btn btn-sm" style="background:#fff; border:1px solid #ccc; padding:6px 8px;">Open bootstrap-config.json</button>
                    <button id="copyExpectedRedirect" class="btn btn-sm" style="background:#fff; border:1px solid #ccc; padding:6px 8px;">Kopieer verwachte URI</button>
                    <button id="openAzurePortal" class="btn btn-sm" style="background:#fff; border:1px solid #ccc; padding:6px 8px;">Open in Azure Portal</button>
                    <a href="https://learn.microsoft.com/azure/active-directory/develop/reply-urls" target="_blank" style="align-self:center; color:#0b5ed7; text-decoration:underline;">Hoe fix je dit in Azure Portal?</a>
                </div>
            `;

            const container = document.querySelector('.login-card');
            if (container) container.prepend(banner);

            // Attach actions
            setTimeout(() => {
                const openBtn = document.getElementById('openBootstrapConfig');
                const copyBtn = document.getElementById('copyExpectedRedirect');
                const portalBtn = document.getElementById('openAzurePortal');
                if (openBtn) {
                    openBtn.addEventListener('click', () => {
                        window.open('/web/bootstrap-config.json', '_blank');
                    });
                }
                if (copyBtn) {
                    copyBtn.addEventListener('click', async () => {
                        try {
                            await navigator.clipboard.writeText(expected);
                            copyBtn.textContent = 'Gekopieerd!';
                            setTimeout(() => copyBtn.textContent = 'Kopieer verwachte URI', 2000);
                        } catch (e) {
                            console.warn('Clipboard write failed', e);
                            alert('Kopieer niet gelukt. Verwachte URI: ' + expected);
                        }
                    });
                }
                if (portalBtn) {
                    portalBtn.addEventListener('click', () => {
                        // Construct Azure Portal URL for the app registration
                        const clientId = localStorage.getItem('m365_clientId') || localStorage.getItem('m365_appId');
                        const tenantId = localStorage.getItem('m365_tenantId') || '';
                        if (!clientId) {
                            alert('Geen clientId gevonden in localStorage. Open eerst bootstrap-config.json.');
                            return;
                        }

                        // Try a common blade URL for app registrations
                        let portalUrl = `https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/Overview/appId/${clientId}`;
                        if (tenantId) {
                            // prefix with tenant context if available
                            portalUrl = `https://portal.azure.com/${tenantId}/#blade/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/Overview/appId/${clientId}`;
                        }

                        window.open(portalUrl, '_blank');
                    });
                }
            }, 50);
        }
    } catch (e) {
        console.warn('Could not show redirect warning', e);
    }
    const loginButton = document.getElementById('loginButton');
    const loginSection = document.getElementById('loginSection');
    const appRegistrationSection = document.getElementById('appRegistrationSection');
    const successSection = document.getElementById('successSection');
    const loginError = document.getElementById('loginError');

    // Check if user is already authenticated
    if (hasAppRegistration()) {
        // Try to initialize and login with existing credentials
        const msal = await initializeMsal();
        if (msal) {
            window.location.href = 'dashboard.html';
            return;
        }
    }

    loginButton.addEventListener('click', async () => {
        try {
            loginError.style.display = 'none';
            loginButton.disabled = true;
            loginButton.innerHTML = '<span class="spinner"></span> Aanmelden...';

            // Step 1: Initialize MSAL with stored client ID
            const clientId = localStorage.getItem('m365_clientId');
            
            if (!clientId) {
                throw new Error('Geen app registratie gevonden. Gebruik New-BootstrapApp.ps1 om een bootstrap app aan te maken.');
            }

            // Initialize MSAL
            const msal = await initializeMsal();
            if (!msal) {
                throw new Error('Kon MSAL niet initialiseren met Client ID: ' + clientId);
            }

            // Step 2: Perform interactive login
            loginSection.style.display = 'none';
            appRegistrationSection.style.display = 'block';
            appRegistrationSection.querySelector('p').textContent = 'Aanmelden bij Microsoft...';
            
            // Redirect to Microsoft login
            await signIn();

        } catch (error) {
            console.error('Login error:', error);
            loginError.textContent = 'Login mislukt: ' + error.message;
            loginError.style.display = 'block';
            loginButton.disabled = false;
            loginButton.innerHTML = '<span class="btn-icon">🔐</span> Login met Microsoft';
            loginSection.style.display = 'block';
            appRegistrationSection.style.display = 'none';
        }
    });
});

// Handle redirect response after login
document.addEventListener('DOMContentLoaded', async () => {
    try {
        const msal = await initializeMsal();
        if (msal) {
            const response = await handleRedirectResponse();
            if (response && response.account) {
                // Store tenant ID from the account (try multiple locations)
                try {
                    const acct = response.account;
                    let tenantId = null;
                    let tenantName = null;

                    if (acct.idTokenClaims && acct.idTokenClaims.tid) {
                        tenantId = acct.idTokenClaims.tid;
                    } else if (acct.tenantId) {
                        tenantId = acct.tenantId;
                    } else if (acct.homeAccountId && acct.homeAccountId.indexOf('.') !== -1) {
                        tenantId = acct.homeAccountId.split('.')[1];
                    }

                    tenantName = acct.name || acct.username || localStorage.getItem('m365_tenantName');

                    if (tenantId) {
                        localStorage.setItem('m365_tenantId', tenantId);
                    }

                    if (tenantName) {
                        localStorage.setItem('m365_tenantName', tenantName);
                    }
                } catch (e) {
                    console.warn('Could not parse tenant from account object', e);
                }

                // Show success and redirect
                const successSection = document.getElementById('successSection');
                const loginSection = document.getElementById('loginSection');
                const appRegistrationSection = document.getElementById('appRegistrationSection');

                if (successSection && loginSection && appRegistrationSection) {
                    loginSection.style.display = 'none';
                    appRegistrationSection.style.display = 'none';
                    successSection.style.display = 'block';

                    setTimeout(() => {
                        window.location.href = 'dashboard.html';
                    }, 1500);
                }
            }
        }
    } catch (error) {
        console.error('Redirect handling error:', error);
    }
});

// Note: Before using this interface:
// 1. Run the New-BootstrapApp.ps1 PowerShell script to create the bootstrap app
// 2. Grant admin consent in the browser when prompted
// 3. The script will save the Client ID to auth.js
// 4. Then you can use this web interface to login
