# Set-RestrictedSiteAccess

These scripts can be used to enable and disable Restricted Site Access on SPO Sites

# Restricted Site Access Script

## Synopsis

Finds every site under `/sites/` or `/teams/` across all site templates, including Microsoft 365 Group-connected and Teams-connected sites. Checks each site's `SharingCapability` and, where external sharing is `Disabled`, turns on Restricted Site Access (RSA) and adds up to three groups to it.

Groups already in RSA and a group-connected site's own Microsoft 365 group are always kept; the configured groups are added to them.

## Description

- Uses app-only certificate authentication for both Microsoft Graph and SharePoint Online so the script can run unattended (Task Scheduler, Azure Automation, etc.).
- Reports each step and the progress of the run to the PowerShell session.
- Honors SharePoint Online and Microsoft Graph throttling (HTTP `429` and `503`): waits for the `Retry-After` value or uses exponential backoff with jitter, then retries.
- Writes a CSV log of every site processed.

## Notes

### Module Responsibilities

| Module | Purpose |
| --- | --- |
| Microsoft Graph PowerShell | App-only sign-in and group lookup. Runs in a separate process to avoid the MSAL DLL conflict with PnP.PowerShell. |
| PnP.PowerShell | Reads site `SharingCapability` and sets RSA. Microsoft Graph cannot perform either operation. |

### Required Modules

- `Microsoft.Graph.Authentication`
- `Microsoft.Graph.Groups`
- `PnP.PowerShell`

### Application Permissions

| Service | Permission | Permission Type | Consent |
| --- | --- | --- | --- |
| Microsoft Graph | `Group.Read.All` | Application | Admin consent |
| SharePoint | `Sites.FullControl.All` | Application | Admin consent |

### Tenant Setting

Run once to enable restricted access control at the tenant level:

```powershell
Set-SPOTenant -EnableRestrictedAccessControl $true
```

### Licensing

Requires SharePoint Advanced Management or Microsoft 365 Copilot.

## Examples

### Preview Only

Preview changes and list excluded sites:

```powershell
.\Set-RestrictedSiteAccess.ps1 -WhatIf -Verbose
```

### Apply Changes

```powershell
.\Set-RestrictedSiteAccess.ps1
```
