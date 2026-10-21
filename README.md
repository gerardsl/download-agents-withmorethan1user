# Agent sharing report (Agent 365)

`Get-SharedAgents.ps1` gets every package in your tenant's Microsoft Agent 365 / Copilot agent catalog and sorts the agents built in your organization into three lists:

1. **Not shared with anyone**: only the creator (owner) has access.
2. **Shared with more than one person**: 2 or more unique people. Groups are expanded and the owner isn't counted.
3. **Shared with the whole organization**: published to everyone in the organization.

## How it works
1. Signs in to Microsoft Graph with your account (delegated, interactive).
2. `GET https://graph.microsoft.com/v1.0/copilot/admin/catalog/packages`: lists all packages, following `@odata.nextLink` until there are no more pages.
3. `GET .../copilot/admin/catalog/packages/{id}`: gets each package's details (`sharedWithUsersAndGroups`, `allowedUsersAndGroups`, `availableTo`, `deployedTo`, `ownerId`).
4. Expands groups to their users, including nested groups (`/groups/{id}/transitiveMembers/microsoft.graph.user`). It then removes duplicates and the owner to count **unique people**.
5. Assigns each package a `SharingCategory`:

| SharingCategory | Rule | Listed as |
|---|---|---|
| `NotShared` | Org-built agent (`type` = `shared`, `lob` or `custom`) shared with 0 people besides the owner | List 1 |
| `SharedWithMultiple` | Org-built agent shared with `MinSharedPeople` or more unique people (default 2) | List 2 |
| `SharedWithOrganization` | Published by your org (`type` = `lob`/`custom`) with `availableTo = allowedForAll` or `deployedTo = acquiredForAll` | List 3 |
| `SharedBelowThreshold` | Shared with fewer than `MinSharedPeople` people (that is, exactly 1 by default) | Inventory only |
| `MicrosoftOrExternal` | Microsoft or partner apps (`firstParty`, `thirdParty`, `microsoft`, `external`) | Inventory only |
| `DuplicateEntry` | Another registry entry for an agent that's already listed (see below) | Inventory only |

**One row per agent:** the registry can hold several entries for the same agent, for example a **Shared by creator** entry and a **Published by your org** entry. Entries are linked when they share a `manifestId`, `agentIdentityId` or `assetId`, or have the same name and platform while at least one of them has no `agentIdentityId`. Each linked group is listed once, in its widest category (organization > more than one person > one person > not shared). The other entries get the category `DuplicateEntry`, and the `LinkedEntries` column lists the related entries.

> **Why "whole organization" uses `lob`/`custom`:** the API has no org-wide sharing flag for agents shared by their creator (`type = shared`), and their `availableTo` is always `allowedForAll`. When a creator shares an agent with everyone in the organization and an admin approves it, the registry creates a separate **Published by your org** entry. That entry's `availableTo`/`deployedTo` shows whether everyone can use it.

For agents published by your org, the admin availability list (`allowedUsersAndGroups`) is treated as the sharing audience.

Retries throttled (429) and transient (5xx) responses. User and group lookups are cached.

## Prerequisites
- A **Microsoft Agent 365** license in the tenant (required by the Package Management API).
- A signed-in account with the **AI Administrator** or **Global Administrator** role.
- PowerShell 7+ (recommended) or Windows PowerShell 5.1.
- The Graph authentication module:
  ```powershell
  Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
  ```
- Consent to these delegated permissions, requested at sign-in:
  `CopilotPackages.Read.All`, `User.ReadBasic.All`, `GroupMember.Read.All`

## Usage
```powershell
# Default: three lists, "more than one person" = 2+ unique people, reports written to .\output
.\Get-SharedAgents.ps1

# Specific tenant, higher threshold, plus a CSV of the full inventory
.\Get-SharedAgents.ps1 -TenantId contoso.onmicrosoft.com -MinSharedPeople 10 -ExportAllAgents

# Also count users/groups that the admin made the agent available to
.\Get-SharedAgents.ps1 -IncludeAllowedUsersAndGroups
```

| Parameter | Description |
|---|---|
| `-MinSharedPeople` | Minimum number of unique people for list 2. Default `2`. |
| `-OutputFolder` | Folder for the reports. Default `.\output`. |
| `-TenantId` | Tenant ID or domain to sign in to. Optional. |
| `-IncludeAllowedUsersAndGroups` | Also count `allowedUsersAndGroups`. |
| `-ExportAllAgents` | Also write `AllAgents_<timestamp>.csv` with every package. |
| `-UseDeviceCode` | Sign in with a device code. Use this in embedded or headless terminals if sign-in fails with "A window handle must be configured". |
| `-ClientId` | App registration ID for app-only sign-in. Requires `-TenantId` and either `-CertificateThumbprint` or `-ClientSecret`. |
| `-CertificateThumbprint` | Certificate thumbprint for app-only sign-in (the certificate must be in `Cert:\CurrentUser\My`). |
| `-ClientSecret` | Client secret (SecureString) for app-only sign-in. |
| `-MaxRetries` | Retries on 429/5xx. Default `5`. |

## Output
- **Console**: the three lists as tables, plus a count of every package by category.
- **`AgentsNotShared_<timestamp>.csv`**, **`AgentsSharedWithMultiple_<timestamp>.csv`**, **`AgentsSharedWithOrganization_<timestamp>.csv`**: one file per list.
- **`AgentSharingReport_<timestamp>.html`**: a self-contained report with summary cards and one section per list. Each section has a filter box and sortable columns.
- **`AllAgents_<timestamp>.csv`** (with `-ExportAllAgents`): every package, including its `SharingCategory`.

Main columns: `DisplayName`, `Id`, `SharingCategory`, `Owner`, `Type`, `Platform`, `Publisher`, `AvailableTo`, `DeployedTo`, `IsBlocked`, `UniquePeopleCount`, `DirectUserShares`, `GroupShares`, `SharedWithUsers`, `SharedWithGroups`, `ActiveUsers`, `LastUsedDateTime`, `CreatedDateTime`, `LastModifiedDateTime`.

## App-only sign-in (when Conditional Access blocks interactive sign-in)
If sign-in shows **"You don't have access to this"** (Conditional Access, error 53003), use an app registration instead. User-based Conditional Access policies don't apply to app-only tokens.

**Quick setup:** run `New-SharedAgentsApp.ps1` from **Azure Cloud Shell** (PowerShell) in the target tenant, or from any device the tenant allows. It creates the app, adds the Graph application permissions, grants admin consent, creates a client secret (30 days by default), and prints the command to run:
```powershell
./New-SharedAgentsApp.ps1 -TenantId <tenantId>
```

**Manual setup:**
1. In Entra admin center, go to **App registrations > New registration** (single tenant).
2. Under **API permissions > Add > Microsoft Graph > Application permissions**, add `CopilotPackages.Read.All`, `User.Read.All` and `GroupMember.Read.All`. Then select **Grant admin consent**.
3. Under **Certificates & secrets**, upload a certificate (recommended) or create a client secret.
4. Run:
   ```powershell
   # Certificate
   .\Get-SharedAgents.ps1 -TenantId <tenant> -ClientId <appId> -CertificateThumbprint <thumbprint>
   # Client secret
   .\Get-SharedAgents.ps1 -TenantId <tenant> -ClientId <appId> -ClientSecret (Read-Host -AsSecureString 'Secret')
   ```

## Notes
- Users that can't be resolved (for example, deleted users) still count and are shown as `(unresolved) <id>`.
- Group members that aren't users (devices, service principals) aren't counted.
- If a package's details can't be retrieved, a warning is shown and the package is treated as shared with 0 people.
- Some Microsoft or partner apps, and some agents published by your org, may also be preinstalled or available to everyone. Lists 1 to 3 only cover agents built in your organization. Use `-ExportAllAgents` to see every package.
- HTTP 401/403 when listing packages usually means the license, admin role or permission consent is missing.
