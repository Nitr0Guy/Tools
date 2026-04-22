/* ═══════════════════════════════════════════════════════
   auth.js — MSAL.js 2.x wrapper (redirect flow)
   ═══════════════════════════════════════════════════════ */

const Auth = (() => {
    const msalConfig = {
        auth: {
            clientId: '8ac4b833-754b-41ba-a4ea-ea8d626a2fb7',
            authority: 'https://login.microsoftonline.com/organizations',
            redirectUri: window.location.origin,
        },
        cache: {
            cacheLocation: 'localStorage',
            storeAuthStateInCookie: true,
        },
    };

    const scopes = [
        'DeviceManagementManagedDevices.ReadWrite.All',
        'DeviceManagementManagedDevices.PrivilegedOperations.All',
        'DeviceManagementServiceConfig.ReadWrite.All',
        'User.Read.All',
        'Organization.Read.All',
        'Device.Read.All',
    ];

    let msalInstance = null;
    let initError = null;

    async function initialize() {
        try {
            msalInstance = new msal.PublicClientApplication(msalConfig);
        } catch (err) {
            initError = 'MSAL create failed: ' + err.message;
            console.error(initError);
            throw new Error(initError);
        }
        try {
            const response = await msalInstance.handleRedirectPromise();
            if (response) {
                msalInstance.setActiveAccount(response.account);
            } else {
                const accounts = msalInstance.getAllAccounts();
                if (accounts.length > 0) {
                    msalInstance.setActiveAccount(accounts[0]);
                }
            }
        } catch (err) {
            console.warn('handleRedirectPromise failed:', err.message);
        }
    }

    function getAccount() {
        return msalInstance ? msalInstance.getActiveAccount() : null;
    }

    function isSignedIn() {
        return !!getAccount();
    }

    function signIn() {
        if (!msalInstance) throw new Error(initError || 'Auth not initialized — MSAL failed to load');
        msalInstance.loginRedirect({ scopes });
    }

    function signOut() {
        if (!msalInstance) return;
        const account = getAccount();
        msalInstance.logoutRedirect({
            account,
            postLogoutRedirectUri: window.location.origin,
        });
    }

    async function getAccessToken() {
        const account = getAccount();
        if (!account) throw new Error('Not signed in');

        try {
            const response = await msalInstance.acquireTokenSilent({ scopes, account });
            return response.accessToken;
        } catch (err) {
            if (err instanceof msal.InteractionRequiredAuthError) {
                msalInstance.acquireTokenRedirect({ scopes });
                return null;
            }
            throw err;
        }
    }

    return { initialize, getAccount, isSignedIn, signIn, signOut, getAccessToken };
})();
