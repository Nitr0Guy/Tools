# Intune Device Lookup

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)](https://github.com/PowerShell/PowerShell)
[![Graph API](https://img.shields.io/badge/Microsoft%20Graph-v1.0%20%7C%20beta-0078D4.svg)](https://learn.microsoft.com/en-us/graph/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

A PowerShell WPF GUI tool that authenticates to **Microsoft Graph**, searches **Entra ID** users, and provides a rich view of their **Intune-managed devices** — including compliance status, installed apps, LAPS passwords, and more.

---

## Features

- **Sign In / Sign Out** — Interactive browser-based sign-in via Microsoft Graph (device code / delegated auth)
- **User Search** — Search by display name, UPN, email address, or device name across your tenant
- **Device List** — Lists all Intune-managed devices for a selected user, sorted by last sync date
- **Device Details** — View device name, OS version, OS build, manufacturer, model, and serial number
- **Last Logged-On User** — Shows last logon user and timestamp via the Graph beta endpoint
- **All Enrolled Users** — Lists all users associated with a device
- **Enrollment Profile & Date** — Displays enrollment profile name and enrollment timestamp
- **LAPS Password** — Retrieve and reveal/hide the Local Administrator Password Solution (LAPS) password
- **Installed Apps** — Browse all detected apps on the device with a live search/filter and auto-hides system/built-in packages
- **Compliance Policies** — View all assigned compliance policies and their per-setting states (compliant/non-compliant/error/grace period)
- **Friendly Setting Names** — Compliance setting identifiers are translated to human-readable names (Windows, iOS, macOS, Android)

---

## Screenshots

> *(Add screenshots of the running GUI here)*

---

## Prerequisites

| Requirement | Details |
|---|---|
| **OS** | Windows (WPF requires Windows) |
| **PowerShell** | 5.1 or later |
| **Module** | `Microsoft.Graph.Authentication` |
| **Permissions** | See [Required Graph Permissions](#required-graph-permissions) |

The script will prompt to auto-install `Microsoft.Graph.Authentication` if it is not found.

---

## Installation

```powershell
# Clone the repository
git clone https://github.com/<your-username>/CIS-Edge-Benchmark.git

# Navigate to the script
cd CIS-Edge-Benchmark
```

Or download `IntuneDeviceLookup.ps1` directly.

---

## Usage

```powershell
# Run the tool
.\IntuneDeviceLookup.ps1
```

1. Click **Sign In** — a browser window opens for Microsoft authentication
2. Once connected, the tenant name appears in the header
3. Type a user's name, email, UPN, or device name in the search box and click **Search** (or press Enter)
4. Select a user from the results list — their devices load automatically
5. Select a device to view full details in the right panel
6. Switch between the **Details**, **Apps**, and **Compliance** tabs

---

## Required Graph Permissions

The signed-in account needs the following **delegated** Microsoft Graph permissions:

| Permission | Purpose |
|---|---|
| `User.Read.All` | Search and read Entra ID users |
| `DeviceManagementManagedDevices.Read.All` | Read Intune device data, compliance states, detected apps |
| `DeviceManagementConfiguration.Read.All` | Read enrollment profiles |
| `DeviceLocalCredential.Read.All` | Read LAPS passwords |
| `Organization.Read.All` | Display tenant name in the header |

> **Note:** LAPS password retrieval requires the `DeviceLocalCredential.Read.All` permission and the device must be Azure AD joined with LAPS configured.

---

## Tabs

### Details
Core device info: device name, OS version & build, last logon timestamp, primary user, all enrolled users, enrollment profile, and enrollment date.

### Apps
All detected applications on the device. System/built-in packages (e.g. `Microsoft.BingSearch`, `MicrosoftWindows.*`, `Windows.*`, MSIX packages) are automatically hidden. A live search box lets you filter by name.

### Compliance
All assigned compliance policies with per-policy state badges (✓ Compliant, ✗ Non-Compliant, ~ Grace Period, — Not Applicable). Each policy card expands to show individual setting results with friendly names.

---

## How It Works

1. **Authentication** — Uses `Connect-MgGraph` (delegated) to authenticate interactively via the default browser. The token is managed by `Microsoft.Graph.Authentication`.
2. **User Search** — Queries `/v1.0/users` with `$filter` (startsWith on displayName, UPN, mail) plus a parallel device-name search via `/v1.0/deviceManagement/managedDevices` to resolve the owning user.
3. **Device Details** — Core properties from `/v1.0/deviceManagement/managedDevices/{id}`. Last-logon data from the `/beta` endpoint (graceful failure if unavailable).
4. **Apps** — Fetches `/beta/deviceManagement/managedDevices/{id}/detectedApps` and client-side filters system packages.
5. **Compliance** — Fetches `/v1.0/deviceManagement/managedDevices/{id}/deviceCompliancePolicyStates` and renders per-policy result cards with individual setting rows.
6. **LAPS** — Retrieves the local administrator password from the Microsoft Graph LAPS API using the device's Entra (Azure AD) device ID.

---

## Troubleshooting

**"The 'Microsoft.Graph.Authentication' module is required"**
Click **Yes** when prompted to auto-install, or run manually:
```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
```

**"Sign-in failed" error**
Ensure your account has the required Graph permissions and that your tenant allows the required delegated scopes.

**LAPS button not shown**
The device must be Azure AD joined and have LAPS configured. The button is hidden if no Azure AD device ID is found.

**Apps tab shows 0 applications**
The `DeviceManagementManagedDevices.Read.All` permission is required, or the device may not have reported detected apps to Intune yet.

---

## License

This project is licensed under the MIT License — see the [LICENSE](LICENSE) file for details.

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

---

**If you find this tool useful, please give it a ⭐!**
