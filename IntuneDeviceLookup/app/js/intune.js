/* ═══════════════════════════════════════════════════════
   intune.js — Intune-specific Microsoft Graph operations
   ═══════════════════════════════════════════════════════ */

const Intune = (() => {

    // ── Device Search ──

    /** Exact match by device name (used after barcode scan). */
    async function searchDeviceExact(name) {
        const safe = encodeURIComponent(name.replace(/'/g, "''"));
        const data = await Graph.request(
            'GET',
            `/v1.0/deviceManagement/managedDevices?$filter=deviceName eq '${safe}'&$select=id,deviceName,operatingSystem,osVersion,complianceState,lastSyncDateTime,isEncrypted,serialNumber,manufacturer,model,enrolledDateTime,userDisplayName,userPrincipalName,azureADDeviceId,managementState`
        );
        return data?.value || [];
    }

    /**
     * Partial match by device name or serial number (used from search field).
     * Paginates all managed devices and filters client-side for reliable
     * substring matching on both deviceName and serialNumber.
     */
    async function searchDevices(query) {
        const select = '$select=id,deviceName,operatingSystem,osVersion,complianceState,lastSyncDateTime,serialNumber,manufacturer,model,userDisplayName';
        const lower = query.trim().toLowerCase();
        if (!lower) return [];

        let allMatches = [];
        let url = `/v1.0/deviceManagement/managedDevices?$top=500&${select}`;
        let pageCount = 0;
        const MAX_PAGES = 20;

        while (url && allMatches.length < 25 && pageCount < MAX_PAGES) {
            pageCount++;
            const data = await Graph.request('GET', url);
            if (!data || !data.value) break;

            const matches = data.value.filter(d =>
                (d.deviceName && d.deviceName.toLowerCase().includes(lower)) ||
                (d.serialNumber && d.serialNumber.toLowerCase().includes(lower))
            );
            allMatches.push(...matches);

            if (allMatches.length >= 25) break;
            url = data['@odata.nextLink']
                ? data['@odata.nextLink'].replace('https://graph.microsoft.com', '')
                : null;
        }

        return allMatches.slice(0, 25);
    }

    // ── Device Details ──

    /** Get full managed device details. */
    async function getDevice(id) {
        return await Graph.request('GET', `/v1.0/deviceManagement/managedDevices/${encodeURIComponent(id)}`);
    }

    /** Get hardware info (BIOS version etc.) via beta endpoint. */
    async function getHardwareInfo(id) {
        return await Graph.request(
            'GET',
            `/beta/deviceManagement/managedDevices/${encodeURIComponent(id)}?$select=hardwareInformation`
        );
    }

    // ── Autopilot ──

    /**
     * Find the Autopilot identity for a device by serial number.
     * Paginates through all identities and matches client-side.
     */
    async function getAutopilotIdentity(serialNumber) {
        if (!serialNumber) return null;
        const serial = serialNumber.trim();

        let url = '/v1.0/deviceManagement/windowsAutopilotDeviceIdentities?$top=100';
        while (url) {
            const data = await Graph.request('GET', url);
            if (!data) return null;

            const match = (data.value || []).find(
                ap => ap.serialNumber && ap.serialNumber.trim() === serial
            );
            if (match) return match;

            url = data['@odata.nextLink']
                ? data['@odata.nextLink'].replace('https://graph.microsoft.com', '')
                : null;
        }
        return null;
    }

    /** Get the assigned deployment profile for an Autopilot identity. */
    async function getAutopilotProfile(autopilotId) {
        return await Graph.request(
            'GET',
            `/beta/deviceManagement/windowsAutopilotDeviceIdentities/${encodeURIComponent(autopilotId)}?$expand=deploymentProfile($select=displayName)`
        );
    }

    // ── Device Actions ──

    /** Sync a managed device. */
    async function syncDevice(deviceId) {
        await Graph.request(
            'POST',
            `/v1.0/deviceManagement/managedDevices/${encodeURIComponent(deviceId)}/syncDevice`,
            {}
        );
    }

    /** Fresh Start (cleanWindowsDevice). */
    async function freshStart(deviceId, keepUserData) {
        await Graph.request(
            'POST',
            `/beta/deviceManagement/managedDevices/${encodeURIComponent(deviceId)}/cleanWindowsDevice`,
            { keepUserData: !!keepUserData }
        );
    }

    /** Full device wipe. */
    async function wipeDevice(deviceId) {
        await Graph.request(
            'POST',
            `/v1.0/deviceManagement/managedDevices/${encodeURIComponent(deviceId)}/wipe`,
            {}
        );
    }

    /** Change Autopilot group tag. */
    async function changeGroupTag(autopilotId, newTag) {
        await Graph.request(
            'POST',
            `/v1.0/deviceManagement/windowsAutopilotDeviceIdentities/${encodeURIComponent(autopilotId)}/updateDeviceProperties`,
            { groupTag: newTag }
        );
    }

    return {
        searchDeviceExact,
        searchDevices,
        getDevice,
        getHardwareInfo,
        getAutopilotIdentity,
        getAutopilotProfile,
        syncDevice,
        freshStart,
        wipeDevice,
        changeGroupTag,
    };
})();
