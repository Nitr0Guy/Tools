/* ═══════════════════════════════════════════════════════
   graph.js — Microsoft Graph API HTTP client
   ═══════════════════════════════════════════════════════ */

const Graph = (() => {
    const BASE = 'https://graph.microsoft.com';
    const MAX_RETRIES = 2;

    /**
     * Make a Graph API request.
     * @param {'GET'|'POST'|'PATCH'|'DELETE'} method
     * @param {string} path  — e.g. "/v1.0/deviceManagement/managedDevices"
     * @param {object|null} body
     * @returns {Promise<object|null>}  parsed JSON or null for 204
     */
    async function request(method, path, body = null) {
        const token = await Auth.getAccessToken();
        if (!token) return null; // redirect in progress

        let lastErr = null;
        for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
            const opts = {
                method,
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/json',
                },
            };
            if (body) opts.body = JSON.stringify(body);

            let res;
            try {
                res = await fetch(`${BASE}${path}`, opts);
            } catch (networkErr) {
                lastErr = new Error('Network error. Check your connection.');
                if (attempt < MAX_RETRIES) continue;
                throw lastErr;
            }

            // No-content success (actions like wipe, sync return 204)
            if (res.status === 204) return null;

            // Success
            if (res.ok) return await res.json();

            // 401 — token expired, trigger re-auth
            if (res.status === 401) {
                Auth.signIn();
                return null;
            }

            // 403 — insufficient permissions (don't retry)
            if (res.status === 403) {
                throw new Error('You don\'t have permission for this action. Contact your Intune admin.');
            }

            // 404 — not found (don't retry)
            if (res.status === 404) {
                throw new Error('Resource not found.');
            }

            // 429 — throttled
            if (res.status === 429) {
                const retryAfter = parseInt(res.headers.get('Retry-After') || '5', 10);
                await sleep(retryAfter * 1000);
                continue;
            }

            // 5xx — server error, retry
            if (res.status >= 500) {
                lastErr = new Error(`Server error (${res.status}). Please try again.`);
                if (attempt < MAX_RETRIES) {
                    await sleep(1000 * (attempt + 1));
                    continue;
                }
                throw lastErr;
            }

            // Other 4xx — don't retry
            const errBody = await res.json().catch(() => ({}));
            throw new Error(errBody?.error?.message || `Request failed (${res.status})`);
        }
        throw lastErr || new Error('Request failed');
    }

    function sleep(ms) {
        return new Promise(r => setTimeout(r, ms));
    }

    return { request };
})();
