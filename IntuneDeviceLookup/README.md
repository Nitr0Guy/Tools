# Intune Device Lookup

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)](https://github.com/PowerShell/PowerShell)
[![Graph API](https://img.shields.io/badge/Microsoft%20Graph-v1.0%20%7C%20beta-0078D4.svg)](https://learn.microsoft.com/en-us/graph/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

A PowerShell WPF GUI tool that authenticates to **Microsoft Graph** and provides a rich, single-pane view of **Intune-managed devices** — including compliance status, security keys, installed apps, Autopilot details, and remote actions.

---

## Features

- **Sign In / Sign Out** — Interactive browser-based delegated auth via Microsoft Graph; MSAL cache cleared on every sign-in to avoid stale tokens
- **Check Permissions** — Dedicated dialog showing all required Graph scopes with Granted / Missing badges
- **User & Device Search** — Search by display name, UPN, email address, or device name
- **Device List** — All Intune-managed devices for a selected user, sorted by last sync date with color-coded compliance dots
- **Device Health Card** — At-a-glance health score (x/5) with indicators for Compliance, Encryption, Last Sync freshness, Management state, and Defender/AV status
- **Dark / Light Theme** — Toggle at runtime; persists for the session
- **Export PDF** — Generates a styled report and converts it to PDF via Edge/Chrome headless; saves to the Desktop
- **Remote Actions** — Sync and Fresh Start with confirmation dialogs and structured error feedback
- **Change Autopilot Group Tag** — Fetch all tenant group tags, select one, and apply it to the device
- **Module auto-install** — Prompts to install `Microsoft.Graph.Authentication` if absent

---

## Screenshots

<img width="1618" height="1092" alt="image" src="https://github.com/user-attachments/assets/5aca26e0-6b86-4ede-994a-ce9fadcde285" />

---

## Prerequisites

| Requirement | Details |
|---|---|
| **OS** | Windows (WPF requires Windows) |
| **PowerShell** | 5.1 or later |
| **Module** | `Microsoft.Graph.Authentication` (auto-install prompt on first run) |
| **Permissions** | See [Required Graph Permissions](#required-graph-permissions) |

---

## Installation

Download `IntuneDeviceLookup.ps1` (or the compiled `IntuneDeviceLookup.exe`) directly, or clone the repository:

```powershell
git clone https://github.com/<your-username>/IntuneDeviceLookup.git
cd IntuneDeviceLookup
```

---

## Usage

```powershell
.\IntuneDeviceLookup.ps1
```

1. Click **Sign In** — a browser window opens for Microsoft authentication
2. Once connected, the tenant name appears in the header
3. Type a user's name, email, UPN, or device name in the search box and press **Enter** or click **Search**
4. Select a user from the results — their devices load automatically
5. Select a device to view full details in the right panel
6. Switch between the **Details**, **Applications**, **Compliance**, **Security**, and **Groups** tabs

---

## Tabs

### Details
Core device properties and Autopilot information:

| Field | Notes |
|---|---|
| Device Name | — |
| OS Version | Friendly label (e.g. "Windows 11") |
| OS Build | Raw `osVersion` value |
| Last Logged-On User | Resolved display name + timestamp (beta endpoint) |
| Primary User | Display name + UPN |
| All Logged-In Users | Full resolved list |
| Installed Enrollment Profile | Profile name recorded at enrollment time |
| Assigned Enrollment Profile | Current Autopilot deployment profile |
| Autopilot Group Tag | Tag value + assignment status badge with per-state tooltips |
| Last Enrolled | Enrollment date/time |
| Last Password Change | Primary user's Azure AD password last-changed date |
| BIOS / UEFI Version | Windows only, from `hardwareInformation` |
| Defender / AV | Detailed status, last scan, and signature date |

### Applications
All detected apps on the device. Built-in / system packages are hidden by default (toggle to reveal). Live search filters by name or publisher. Shows count of visible vs. hidden apps.

### Compliance
- **Overview** — Device-level compliance state, last sync timestamp, grace period expiry, and mismatch warning cards
- **Policies** — Per-policy compliance cards with individual setting rows translated to friendly names (Windows, iOS, macOS, Android); "X of Y policies compliant" summary

### Security
- **LAPS Password** — Retrieve with Reveal / Hide toggle and Copy button
- **BitLocker Recovery Keys** — Lists all keys (ID, volume type, creation date) with per-key Reveal / Hide and Copy; full key value fetched on demand

### Groups
- **Device Groups** — All Entra ID groups the device is a member of
- **User Groups** — All Entra ID groups the primary user is a member of
- Live search filter in both lists

---

## Remote Actions

| Action | API | Notes |
|---|---|---|
| **Refresh** | Re-loads all device data | Clears lazy-load flags and re-fetches everything |
| **Sync** | `POST .../managedDevices/{id}/syncDevice` | Sends an MDM sync signal |
| **Fresh Start** | `POST .../managedDevices/{id}/cleanWindowsDevice` | Confirmation dialog with Keep / Remove user files choice |
| **Change Group Tag** | `POST .../windowsAutopilotDeviceIdentities/{id}/updateDeviceProperties` | Fetches all tenant tags, presents selection dialog, confirms before applying |

> Remote actions require an **Intune RBAC role** with *Remote tasks* permission in addition to the Graph scopes below.

---

## Required Graph Permissions

All permissions are **delegated** (signed-in user):

| Permission | Purpose |
|---|---|
| `User.Read.All` | Search and resolve Entra ID users |
| `Device.Read.All` | Resolve Azure AD device object ID for group lookups |
| `Organization.Read.All` | Display tenant name in the header |
| `GroupMember.Read.All` | Device and user group memberships |
| `DeviceManagementManagedDevices.ReadWrite.All` | Read Intune device data, apps, compliance |
| `DeviceManagementManagedDevices.PrivilegedOperations.All` | Trigger remote actions (Sync, Fresh Start) |
| `DeviceManagementConfiguration.Read.All` | Read enrollment profiles |
| `DeviceManagementServiceConfig.ReadWrite.All` | Autopilot identities, group tags |
| `DeviceLocalCredential.Read.All` | Retrieve LAPS passwords |
| `BitlockerKey.ReadBasic.All` | List BitLocker recovery key IDs |
| `BitlockerKey.Read.All` | Reveal full BitLocker recovery key values |

Use the **Check Permissions** button in the tool to verify which scopes your current token has consented.

---

## How It Works

1. **Authentication** — `Connect-MgGraph` runs in a background runspace so the WPF UI thread stays responsive. The MSAL token cache is cleared before each sign-in. After sign-in, consented scopes are validated against the required list.
2. **User Search** — Queries `/v1.0/users` with `$filter` (startsWith on displayName, UPN, mail) plus a parallel device-name search via `/v1.0/deviceManagement/managedDevices` to resolve the owning user.
3. **Device Details** — Core properties from `/v1.0/deviceManagement/managedDevices/{id}`. Logged-on users and hardware info from the `/beta` endpoint.
4. **Autopilot** — Pages through `windowsAutopilotDeviceIdentities` and matches by serial number. The assigned deployment profile is resolved via `$expand=deploymentProfile` in a single call.
5. **Apps** — Fetches `/beta/deviceManagement/managedDevices/{id}/detectedApps`; built-in packages filtered client-side.
6. **Compliance** — Fetches policy states and per-setting states; setting identifiers translated to human-readable names.
7. **Security** — LAPS via `/beta/directory/deviceLocalCredentials/{aadDeviceId}`; BitLocker key list via `/v1.0/informationProtection/bitlocker/recoveryKeys` with per-key on-demand reveal.
8. **Groups** — Device membership via `/v1.0/devices/{objId}/memberOf`; user membership via `/v1.0/users/{upn}/memberOf`.
9. **Lazy loading** — Applications, Compliance, Security, and Groups data are fetched only on first tab activation per device, then cached.
10. **Export PDF** — Builds a styled HTML report and converts it using Edge or Chrome headless (`--headless=new --print-to-pdf`); falls back to opening the HTML file if no supported browser is found.

---

## Troubleshooting

**"The 'Microsoft.Graph.Authentication' module is required"**
Click **Yes** when prompted to auto-install, or run manually:
```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
```

**Sign-in fails or scopes are missing**
Use the **Check Permissions** button after signing in. If scopes show as missing, sign out, sign back in, and re-consent. If scopes are present but actions still fail with 403, an Intune RBAC role with the relevant permission is required.

**Fresh Start / Sync returns Forbidden**
Verify `DeviceManagementManagedDevices.PrivilegedOperations.All` is consented (Check Permissions) and that your account has an Intune RBAC role with *Remote tasks* enabled.

**Assigned Enrollment Profile shows "Not assigned"**
The device may not be registered in Autopilot, or no deployment profile has been assigned to it yet.

**LAPS button not shown**
The device must be Azure AD joined with LAPS configured. The button is hidden if no Azure AD device ID is found for the device.

**BitLocker keys missing**
Requires `BitlockerKey.Read.All`. Keys are only stored in Azure AD for Azure AD joined / hybrid joined devices with key escrow enabled.

**Apps tab shows 0 applications**
`DeviceManagementManagedDevices.Read.All` is required, or the device may not have reported detected apps to Intune yet.

**PDF export opens HTML instead of PDF**
Microsoft Edge or Google Chrome must be installed. The tool falls back to opening the HTML file when neither browser is detected.

---

**If you find this tool useful, please give it a ⭐!**


