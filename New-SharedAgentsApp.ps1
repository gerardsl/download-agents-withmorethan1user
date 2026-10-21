<#
.SYNOPSIS
    Creates the app registration used by Get-SharedAgents.ps1 for app-only sign-in.

.DESCRIPTION
    Uses the Azure CLI (works as-is in Azure Cloud Shell, where az is already signed in):
      1. Creates a single-tenant app registration.
      2. Adds Microsoft Graph application permissions: CopilotPackages.Read.All,
         User.Read.All, GroupMember.Read.All.
      3. Creates the service principal and grants admin consent (requires Global
         Administrator or Privileged Role Administrator, or Cloud Application Administrator
         plus permission to consent to Graph app roles).
      4. Creates a client secret and prints the values needed to run Get-SharedAgents.ps1.

    The secret is shown once. Store it securely (e.g. a password manager or Key Vault).

.EXAMPLE
    ./New-SharedAgentsApp.ps1 -TenantId <your-tenant-id>
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$TenantId,
    [string]$DisplayName = 'Agent365 Shared Agents Report',
    [ValidateRange(1, 730)]
    [int]$SecretValidityDays = 30,
    [switch]$SkipAdminConsent
)

$ErrorActionPreference = 'Stop'
$graphAppId = '00000003-0000-0000-c000-000000000000'
$permissions = 'CopilotPackages.Read.All', 'User.Read.All', 'GroupMember.Read.All'

function Invoke-Az {
    $out = & az @args 2>&1
    if ($LASTEXITCODE -ne 0) { throw "az $($args -join ' ') failed: $out" }
    return ($out | Where-Object { $_ -isnot [System.Management.Automation.ErrorRecord] }) -join "`n"
}

& az account get-access-token --tenant $TenantId --resource-type ms-graph --query expiresOn -o tsv 2>$null | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-Host "Signing in to tenant $TenantId..." -ForegroundColor Cyan
    Invoke-Az login --tenant $TenantId --allow-no-subscriptions --only-show-errors | Out-Null
}

Write-Host 'Resolving Microsoft Graph app role IDs...' -ForegroundColor Cyan
$roles = Invoke-Az ad sp show --id $graphAppId --query "appRoles[].{id:id, value:value}" -o json | ConvertFrom-Json
$roleIds = foreach ($p in $permissions) {
    $r = $roles | Where-Object value -EQ $p
    if (-not $r) { throw "Graph application permission '$p' was not found in this tenant." }
    "$($r.id)=Role"
}

$appId = Invoke-Az ad app list --display-name $DisplayName --query "[0].appId" -o tsv
if ($appId) {
    Write-Host "Reusing existing app '$DisplayName' ($appId)" -ForegroundColor Yellow
}
else {
    Write-Host "Creating app registration '$DisplayName'..." -ForegroundColor Cyan
    $appId = Invoke-Az ad app create --display-name $DisplayName --sign-in-audience AzureADMyOrg --query appId -o tsv
}

Invoke-Az ad app permission add --id $appId --api $graphAppId --api-permissions @roleIds --only-show-errors | Out-Null

$spId = Invoke-Az ad sp list --filter "appId eq '$appId'" --query "[0].id" -o tsv
if (-not $spId) {
    Invoke-Az ad sp create --id $appId --only-show-errors | Out-Null
}

if (-not $SkipAdminConsent) {
    Write-Host 'Granting admin consent...' -ForegroundColor Cyan
    # Newly created service principals can take a few seconds to replicate.
    for ($i = 1; $i -le 6; $i++) {
        try { Invoke-Az ad app permission admin-consent --id $appId | Out-Null; break }
        catch { if ($i -eq 6) { throw }; Start-Sleep -Seconds 10 }
    }
}

Write-Host 'Creating client secret...' -ForegroundColor Cyan
$end = (Get-Date).ToUniversalTime().AddDays($SecretValidityDays).ToString('yyyy-MM-ddTHH:mm:ssZ')
$secret = Invoke-Az ad app credential reset --id $appId --append --display-name 'Get-SharedAgents' --end-date $end --query password -o tsv

Write-Host ''
Write-Host 'Done. Values for Get-SharedAgents.ps1:' -ForegroundColor Green
Write-Host "  TenantId : $TenantId"
Write-Host "  ClientId : $appId"
Write-Host "  Secret   : $secret   (expires $end; store it securely, it won't be shown again)"
Write-Host ''
Write-Host 'Run:' -ForegroundColor Green
Write-Host "  .\Get-SharedAgents.ps1 -TenantId $TenantId -ClientId $appId -ClientSecret (Read-Host -AsSecureString 'Secret') -ExportAllAgents"
if ($SkipAdminConsent) {
    Write-Host 'Admin consent was skipped: a Global Administrator must grant consent in Entra > App registrations > API permissions.' -ForegroundColor Yellow
}
