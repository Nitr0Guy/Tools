/* ═══════════════════════════════════════════════════════
   app.js — Main application (router, views, event wiring)
   ═══════════════════════════════════════════════════════ */

const App = (() => {
    // ── State ──
    let currentDevice = null;       // full device object from Graph
    let currentAutopilot = null;    // Autopilot identity (or null)

    // ── DOM refs ──
    const $ = (sel) => document.querySelector(sel);
    const $$ = (sel) => document.querySelectorAll(sel);

    // ── Initialization ──
    async function init() {
        // Bind events first so buttons always work even if auth init fails
        bindEvents();

        // Check if libraries loaded
        const statusEl = document.querySelector('.signin-subtitle');
        if (typeof msal === 'undefined') {
            if (statusEl) statusEl.textContent = 'Error: MSAL library failed to load. Check your connection.';
            return;
        }

        try {
            await Auth.initialize();
        } catch (err) {
            if (statusEl) statusEl.textContent = 'Auth error: ' + err.message;
        }

        if (Auth.isSignedIn()) {
            showView('scan');
            populateAccountInfo();
        } else {
            showView('signin');
        }
    }

    // ── Navigation / Views ──
    function showView(name) {
        $$('.view').forEach(v => v.classList.add('hidden'));
        const view = $(`#view-${name}`);
        if (view) view.classList.remove('hidden');

        // Stop scanner when leaving scan view
        if (name !== 'scan' && Scanner.isRunning()) {
            Scanner.stop();
            updateScannerButtons(false);
        }
    }

    // ── Event Binding ──
    function bindEvents() {
        // Sign in / out
        $('#btn-signin').addEventListener('click', () => {
            const statusEl = document.querySelector('.signin-subtitle');
            try {
                if (statusEl) statusEl.textContent = 'Redirecting to sign-in...';
                Auth.signIn();
            } catch (err) {
                if (statusEl) statusEl.textContent = 'Sign-in error: ' + err.message;
            }
        });
        $('#btn-signout').addEventListener('click', () => Auth.signOut());

        // Bottom nav
        $$('.nav-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                const view = btn.dataset.view;
                showView(view);
                // Update active state in the current bottom nav
                btn.closest('.bottom-nav').querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
                btn.classList.add('active');

                if (view === 'recent') renderRecent();
            });
        });

        // Mode toggle (Scan / Search)
        $('#mode-scan').addEventListener('click', () => setMode('scan'));
        $('#mode-search').addEventListener('click', () => setMode('search'));

        // Scanner controls
        $('#btn-start-scanner').addEventListener('click', startScanner);
        $('#btn-stop-scanner').addEventListener('click', stopScanner);

        // Search
        $('#btn-search').addEventListener('click', doSearch);
        $('#search-input').addEventListener('keydown', (e) => {
            if (e.key === 'Enter') doSearch();
        });

        // Back button
        $('#btn-back').addEventListener('click', () => {
            currentDevice = null;
            currentAutopilot = null;
            showView('scan');
        });

        // Device actions
        $('#btn-sync').addEventListener('click', handleSync);
        $('#btn-freshstart').addEventListener('click', handleFreshStart);
        $('#btn-wipe').addEventListener('click', handleWipe);
        $('#btn-change-grouptag').addEventListener('click', handleChangeGroupTag);
    }

    // ── Mode Toggle ──
    function setMode(mode) {
        $('#mode-scan').classList.toggle('active', mode === 'scan');
        $('#mode-search').classList.toggle('active', mode === 'search');
        $('#scanner-area').classList.toggle('hidden', mode !== 'scan');
        $('#search-area').classList.toggle('hidden', mode !== 'search');
        $('#scan-status').textContent = '';
        $('#search-results').classList.add('hidden');

        if (mode === 'search') {
            stopScanner();
            $('#search-input').focus();
        }
    }

    // ── Barcode Scanner ──
    async function startScanner() {
        try {
            Scanner.init('scanner-reader', onBarcodeScan);
            await Scanner.start();
            updateScannerButtons(true);
            $('#scanner-placeholder').classList.add('hidden');
            $('#scan-status').textContent = 'Point camera at a barcode...';
        } catch (err) {
            showToast(err.message, 'error');
        }
    }

    async function stopScanner() {
        await Scanner.stop();
        updateScannerButtons(false);
        $('#scan-status').textContent = '';
    }

    function updateScannerButtons(scanning) {
        $('#btn-start-scanner').classList.toggle('hidden', scanning);
        $('#btn-stop-scanner').classList.toggle('hidden', !scanning);
    }

    async function onBarcodeScan(deviceName) {
        await Scanner.stop();
        updateScannerButtons(false);
        $('#scan-status').textContent = `Scanned: ${deviceName}`;
        await lookupDevice(deviceName, true);
    }

    // ── Search ──
    async function doSearch() {
        const query = $('#search-input').value.trim();
        if (!query) return;

        $('#scan-status').textContent = 'Searching...';
        $('#search-results').classList.add('hidden');

        try {
            const devices = await Intune.searchDevices(query);
            if (devices.length === 0) {
                $('#scan-status').textContent = 'No devices found.';
            } else if (devices.length === 1) {
                $('#scan-status').textContent = '';
                navigateToDevice(devices[0].id, devices[0].deviceName);
            } else {
                $('#scan-status').textContent = `${devices.length} devices found`;
                renderSearchResults(devices);
            }
        } catch (err) {
            $('#scan-status').textContent = '';
            showToast(err.message, 'error');
        }
    }

    async function lookupDevice(name, exact) {
        try {
            const devices = exact
                ? await Intune.searchDeviceExact(name)
                : await Intune.searchDevices(name);

            if (devices.length === 0) {
                $('#scan-status').textContent = `No device found: "${name}"`;
            } else if (devices.length === 1) {
                navigateToDevice(devices[0].id, devices[0].deviceName);
            } else {
                $('#scan-status').textContent = `${devices.length} devices found`;
                renderSearchResults(devices);
            }
        } catch (err) {
            showToast(err.message, 'error');
        }
    }

    function renderSearchResults(devices) {
        const list = $('#results-list');
        list.innerHTML = '';
        devices.forEach(d => {
            const li = document.createElement('li');
            const serial = d.serialNumber ? ` · S/N: ${esc(d.serialNumber)}` : '';
            li.innerHTML = `
                <div class="result-name">${esc(d.deviceName)}</div>
                <div class="result-sub">${esc(d.operatingSystem || '')} · ${esc(d.userDisplayName || 'No user')}${serial}</div>
            `;
            li.addEventListener('click', () => navigateToDevice(d.id, d.deviceName));
            list.appendChild(li);
        });
        $('#search-results').classList.remove('hidden');
    }

    // ── Device Detail ──
    async function navigateToDevice(id, name) {
        currentDevice = null;
        currentAutopilot = null;
        $('#search-results').classList.add('hidden');

        showView('device');
        $('#device-title').textContent = name || 'Device';
        $('#device-loading').classList.remove('hidden');
        $('#device-details').classList.add('hidden');

        addRecentLookup(name);

        try {
            currentDevice = await Intune.getDevice(id);
            renderDeviceDetails(currentDevice);
            $('#device-loading').classList.add('hidden');
            $('#device-details').classList.remove('hidden');

            // Load Autopilot info in background
            loadAutopilotInfo(currentDevice.serialNumber);
        } catch (err) {
            $('#device-loading').innerHTML = `<p style="color:var(--danger);">${esc(err.message)}</p>`;
        }
    }

    function renderDeviceDetails(d) {
        // Title
        $('#device-title').textContent = d.deviceName || 'Unknown';

        // Health card
        setHealth('compliance', d.complianceState === 'compliant', d.complianceState || '—');
        setHealth('encryption', d.isEncrypted, d.isEncrypted ? 'Encrypted' : 'Not encrypted');
        setHealthSync(d.lastSyncDateTime);
        setHealth('mgmt', d.managementState === 'managed', d.managementState || '—');

        // Properties
        $('#prop-name').textContent = d.deviceName || '—';
        $('#prop-os').textContent = d.operatingSystem || '—';
        $('#prop-osver').textContent = d.osVersion || '—';
        $('#prop-serial').textContent = d.serialNumber || '—';
        $('#prop-manufacturer').textContent = d.manufacturer || '—';
        $('#prop-model').textContent = d.model || '—';
        $('#prop-user').textContent = d.userDisplayName
            ? `${d.userDisplayName} (${d.userPrincipalName || ''})`
            : '—';
        $('#prop-enrolled').textContent = d.enrolledDateTime ? formatDate(d.enrolledDateTime) : '—';
    }

    function setHealth(key, good, text) {
        const dot = $(`#health-${key}-dot`);
        const val = $(`#health-${key}`);
        dot.className = 'health-dot ' + (good ? 'good' : 'bad');
        val.textContent = text;
    }

    function setHealthSync(syncDate) {
        const dot = $('#health-sync-dot');
        const val = $('#health-sync');
        if (!syncDate) {
            dot.className = 'health-dot';
            val.textContent = '—';
            return;
        }
        const daysSince = (Date.now() - new Date(syncDate).getTime()) / 86400000;
        const good = daysSince <= 7;
        dot.className = 'health-dot ' + (good ? 'good' : daysSince <= 14 ? 'warn' : 'bad');
        val.textContent = formatDate(syncDate);
    }

    // ── Autopilot ──
    async function loadAutopilotInfo(serialNumber) {
        const loadEl = $('#autopilot-loading');
        const detEl = $('#autopilot-details');
        const noneEl = $('#autopilot-none');

        loadEl.classList.remove('hidden');
        detEl.classList.add('hidden');
        noneEl.classList.add('hidden');

        try {
            currentAutopilot = await Intune.getAutopilotIdentity(serialNumber);
            loadEl.classList.add('hidden');

            if (!currentAutopilot) {
                noneEl.classList.remove('hidden');
                $('#btn-change-grouptag').disabled = true;
                return;
            }

            $('#prop-grouptag').textContent = currentAutopilot.groupTag || '(none)';
            $('#btn-change-grouptag').disabled = false;

            // Fetch profile
            try {
                const apDetail = await Intune.getAutopilotProfile(currentAutopilot.id);
                $('#prop-enrollprofile').textContent =
                    apDetail?.deploymentProfile?.displayName || '—';
                $('#prop-profilestatus').textContent =
                    formatProfileStatus(apDetail?.deploymentProfileAssignmentStatus);
            } catch (_) {
                $('#prop-enrollprofile').textContent = '—';
                $('#prop-profilestatus').textContent = '—';
            }

            detEl.classList.remove('hidden');
        } catch (err) {
            loadEl.classList.add('hidden');
            noneEl.classList.remove('hidden');
        }
    }

    function formatProfileStatus(status) {
        const map = {
            'assignedInSync': 'Assigned (In Sync)',
            'assignedOutOfSync': 'Assigned (Out of Sync)',
            'assigned': 'Assigned',
            'notAssigned': 'Not Assigned',
            'pending': 'Pending',
            'failed': 'Failed',
        };
        return map[status] || status || '—';
    }

    // ── Actions ──
    async function handleSync() {
        if (!currentDevice) return;
        showToast('Syncing device...');
        try {
            await Intune.syncDevice(currentDevice.id);
            showToast('Sync initiated successfully', 'success');
        } catch (err) {
            showToast(err.message, 'error');
        }
    }

    function handleFreshStart() {
        if (!currentDevice) return;
        showModal(
            'Fresh Start',
            `<p>This will reinstall Windows on <strong>${esc(currentDevice.deviceName)}</strong>. The device stays enrolled but OEM apps will be removed.</p>
             <label><input type="radio" name="freshstart-data" value="true" checked> Keep user data</label>
             <label><input type="radio" name="freshstart-data" value="false"> Remove user data</label>`,
            [
                { text: 'Cancel', style: 'secondary', action: closeModal },
                {
                    text: 'Confirm Fresh Start', style: 'danger', action: async () => {
                        const keep = document.querySelector('input[name="freshstart-data"]:checked').value === 'true';
                        closeModal();
                        showToast('Initiating Fresh Start...');
                        try {
                            await Intune.freshStart(currentDevice.id, keep);
                            showToast('Fresh Start initiated successfully', 'success');
                        } catch (err) {
                            showToast(err.message, 'error');
                        }
                    }
                },
            ]
        );
    }

    function handleWipe() {
        if (!currentDevice) return;
        // First confirmation
        showModal(
            'Wipe Device',
            `<p><strong>Warning:</strong> This will factory reset <strong>${esc(currentDevice.deviceName)}</strong>. All data will be permanently deleted.</p>
             <p style="margin-top:12px;">Type the device name to confirm:</p>
             <input type="text" id="wipe-confirm-input" placeholder="${esc(currentDevice.deviceName)}" autocomplete="off">`,
            [
                { text: 'Cancel', style: 'secondary', action: closeModal },
                {
                    text: 'Wipe Device', style: 'danger', action: async () => {
                        const input = document.getElementById('wipe-confirm-input');
                        if (input.value.trim() !== currentDevice.deviceName) {
                            input.style.borderColor = 'var(--danger)';
                            input.placeholder = 'Name does not match!';
                            return; // don't close modal
                        }
                        closeModal();
                        showToast('Initiating device wipe...');
                        try {
                            await Intune.wipeDevice(currentDevice.id);
                            showToast('Device wipe initiated successfully', 'success');
                        } catch (err) {
                            showToast(err.message, 'error');
                        }
                    }
                },
            ]
        );
    }

    function handleChangeGroupTag() {
        if (!currentAutopilot) return;
        const currentTag = currentAutopilot.groupTag || '';

        showModal(
            'Change Group Tag',
            `<p>Update the Autopilot group tag for <strong>${esc(currentDevice.deviceName)}</strong>.</p>
             <input type="text" id="grouptag-input" value="${esc(currentTag)}" placeholder="Enter new group tag">`,
            [
                { text: 'Cancel', style: 'secondary', action: closeModal },
                {
                    text: 'Save', style: 'primary', action: async () => {
                        const newTag = document.getElementById('grouptag-input').value.trim();
                        closeModal();
                        showToast('Updating group tag...');
                        try {
                            await Intune.changeGroupTag(currentAutopilot.id, newTag);
                            currentAutopilot.groupTag = newTag;
                            $('#prop-grouptag').textContent = newTag || '(none)';
                            showToast('Group tag updated (pending sync)', 'success');
                        } catch (err) {
                            showToast(err.message, 'error');
                        }
                    }
                },
            ]
        );
    }

    // ── Modal ──
    function showModal(title, bodyHtml, buttons) {
        $('#modal-title').textContent = title;
        $('#modal-body').innerHTML = bodyHtml;

        const actionsEl = $('#modal-actions');
        actionsEl.innerHTML = '';
        buttons.forEach(b => {
            const btn = document.createElement('button');
            btn.textContent = b.text;
            btn.className = b.style === 'danger' ? 'btn-danger'
                          : b.style === 'primary' ? 'btn-primary'
                          : 'btn-secondary';
            btn.addEventListener('click', b.action);
            actionsEl.appendChild(btn);
        });

        $('#modal-overlay').classList.remove('hidden');
    }

    function closeModal() {
        $('#modal-overlay').classList.add('hidden');
    }

    // ── Toast ──
    let toastTimeout = null;
    function showToast(msg, type = '') {
        const el = $('#toast');
        el.textContent = msg;
        el.className = 'toast' + (type ? ` toast-${type}` : '');
        el.classList.remove('hidden');
        clearTimeout(toastTimeout);
        toastTimeout = setTimeout(() => el.classList.add('hidden'), 3500);
    }

    // ── Recent Lookups ──
    function getRecent() {
        try {
            return JSON.parse(localStorage.getItem('recentLookups') || '[]');
        } catch { return []; }
    }

    function addRecentLookup(name) {
        if (!name) return;
        let recent = getRecent().filter(r => r !== name);
        recent.unshift(name);
        if (recent.length > 20) recent = recent.slice(0, 20);
        localStorage.setItem('recentLookups', JSON.stringify(recent));
    }

    function renderRecent() {
        const list = $('#recent-list');
        const empty = $('#recent-empty');
        const recent = getRecent();

        list.innerHTML = '';
        if (recent.length === 0) {
            empty.classList.remove('hidden');
            return;
        }
        empty.classList.add('hidden');
        recent.forEach(name => {
            const li = document.createElement('li');
            li.innerHTML = `<div class="result-name">${esc(name)}</div>`;
            li.addEventListener('click', () => lookupDevice(name, true));
            list.appendChild(li);
        });
    }

    // ── Account Info ──
    function populateAccountInfo() {
        const account = Auth.getAccount();
        if (!account) return;
        $('#header-user').textContent = account.name || '';
        $('#account-name').textContent = account.name || '—';
        $('#account-upn').textContent = account.username || '—';
        $('#account-tenant').textContent = `Tenant: ${account.tenantId || '—'}`;
    }

    // ── Helpers ──
    function esc(str) {
        const el = document.createElement('span');
        el.textContent = str || '';
        return el.innerHTML;
    }

    function formatDate(iso) {
        if (!iso) return '—';
        const d = new Date(iso);
        return d.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' }) +
               ' ' + d.toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' });
    }

    return { init };
})();

// ── Boot ──
document.addEventListener('DOMContentLoaded', () => App.init());
