<#
.SYNOPSIS
    Classifies Agent 365 / Microsoft 365 Copilot agents by how widely they are shared.

.DESCRIPTION
    Uses the Microsoft Graph Agent 365 Package Management API (v1.0):
      GET /copilot/admin/catalog/packages        -> all packages in the tenant catalog
      GET /copilot/admin/catalog/packages/{id}   -> details incl. sharedWithUsersAndGroups

    Agents built in the organization (type shared / lob / custom) are classified into
    three lists, reported to the console, CSV files and one HTML report:
      1. NotShared              - nobody but the owner/creator has access.
      2. SharedWithMultiple     - shared with >= MinSharedPeople unique people. Group shares
                                  are expanded to their (transitive) user members and the
                                  owner is excluded.
      3. SharedWithOrganization - published by the org (lob/custom) and available to or
                                  installed for all users (availableTo = allowedForAll or
                                  deployedTo = acquiredForAll).
    Agents shared with fewer than MinSharedPeople people are 'SharedBelowThreshold', and
    Microsoft / partner packages are 'MicrosoftOrExternal'. When the registry holds several
    entries for one agent (e.g. 'Shared by creator' + 'Published by your org'), only the
    entry with the widest sharing is listed; the others become 'DuplicateEntry'. These
    categories appear only in the -ExportAllAgents inventory.

    Requirements: Microsoft Agent 365 license, signed-in account with the AI Administrator
    or Global Administrator role, and the Microsoft.Graph.Authentication module.

.PARAMETER MinSharedPeople
    Minimum number of unique people for the "shared with more than one person" list. Default 2.

.PARAMETER OutputFolder
    Folder where CSV/HTML reports are written. Created if missing. Default .\output

.PARAMETER TenantId
    Optional tenant ID or domain to sign in to.

.PARAMETER IncludeAllowedUsersAndGroups
    Also count users/groups from allowedUsersAndGroups (admin availability) in addition to
    sharedWithUsersAndGroups.

.PARAMETER ExportAllAgents
    Also write a CSV with every package in the catalog (inventory), not only the shared ones.

.PARAMETER UseDeviceCode
    Sign in with the device code flow (useful in embedded/headless terminals where WAM or
    browser sign-in fails with "A window handle must be configured").

.PARAMETER ClientId
    App registration (client) ID for app-only sign-in. Use with -CertificateThumbprint or
    -ClientSecret and -TenantId. The app needs the application permissions
    CopilotPackages.Read.All, User.Read.All and GroupMember.Read.All (admin consented).
    App-only sign-in is not subject to user Conditional Access policies.

.PARAMETER CertificateThumbprint
    Thumbprint of a certificate in Cert:\CurrentUser\My (or LocalMachine\My) for app-only sign-in.

.PARAMETER ClientSecret
    Client secret (SecureString) for app-only sign-in, e.g. (Read-Host -AsSecureString).

.PARAMETER MaxRetries
    Retries for throttled (429) or transient (5xx) Graph responses. Default 5.

.EXAMPLE
    .\Get-SharedAgents.ps1

.EXAMPLE
    .\Get-SharedAgents.ps1 -MinSharedPeople 10 -ExportAllAgents -OutputFolder C:\Reports
#>
[CmdletBinding()]
param(
    [ValidateRange(1, [int]::MaxValue)]
    [int]$MinSharedPeople = 2,
    [string]$OutputFolder = (Join-Path -Path (Get-Location) -ChildPath 'output'),
    [string]$TenantId,
    [switch]$IncludeAllowedUsersAndGroups,
    [switch]$ExportAllAgents,
    [switch]$UseDeviceCode,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [securestring]$ClientSecret,
    [ValidateRange(0, 20)]
    [int]$MaxRetries = 5
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:GraphBase = 'https://graph.microsoft.com/v1.0'
$script:RequiredScopes = @('CopilotPackages.Read.All', 'User.ReadBasic.All', 'GroupMember.Read.All')
# Package types built inside the tenant ('shared' = shared by creator, 'lob'/'custom' = published by your org).
$script:OrgBuiltTypes = @('shared', 'lob', 'custom')
$script:OrgPublishedTypes = @('lob', 'custom')
$script:EmptyGuid = '00000000-0000-0000-0000-000000000000'
# Wider sharing ranks higher; used to pick one entry per agent when the registry has several.
$script:CategoryRank = @{ SharedWithOrganization = 4; SharedWithMultiple = 3; SharedBelowThreshold = 2; NotShared = 1 }
$script:UserCache = @{}
$script:GroupCache = @{}

#region Graph helpers

function Get-GraphResponse {
    param($ErrorRecord)
    $p = $ErrorRecord.Exception.PSObject.Properties['Response']
    if ($p) { return $p.Value } else { return $null }
}

function Get-GraphStatusCode {
    param($ErrorRecord)
    $resp = Get-GraphResponse $ErrorRecord
    if ($resp -and $resp.PSObject.Properties['StatusCode']) { return [int]$resp.StatusCode }
    if ($ErrorRecord.Exception.Message -match '\b(4\d\d|5\d\d)\b') { return [int]$Matches[1] }
    return 0
}

function Invoke-GraphWithRetry {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [string]$Method = 'GET'
    )
    $attempt = 0
    while ($true) {
        try {
            return Invoke-MgGraphRequest -Method $Method -Uri $Uri -OutputType HashTable
        }
        catch {
            $status = Get-GraphStatusCode $_
            $attempt++
            if (($status -eq 429 -or $status -ge 500) -and $attempt -le $MaxRetries) {
                $delay = [math]::Min(60, [math]::Pow(2, $attempt))
                $resp = Get-GraphResponse $_
                if ($resp -and $resp.PSObject.Properties['Headers'] -and $resp.Headers.PSObject.Properties['RetryAfter'] -and
                    $resp.Headers.RetryAfter -and $resp.Headers.RetryAfter.Delta) {
                    $delay = [math]::Ceiling($resp.Headers.RetryAfter.Delta.TotalSeconds)
                }
                Write-Verbose "Graph returned $status for $Uri. Retry $attempt/$MaxRetries in $delay s."
                Start-Sleep -Seconds $delay
                continue
            }
            throw
        }
    }
}

function Invoke-GraphPaged {
    param([Parameter(Mandatory)][string]$Uri)
    $items = [System.Collections.Generic.List[object]]::new()
    $next = $Uri
    while ($next) {
        $page = Invoke-GraphWithRetry -Uri $next
        if ($page.ContainsKey('value') -and $page['value']) {
            foreach ($v in $page['value']) { $items.Add($v) }
        }
        $next = if ($page.ContainsKey('@odata.nextLink')) { $page['@odata.nextLink'] } else { $null }
    }
    return , $items
}

function Get-Prop {
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) { return $Object[$Name] } else { return $null }
    }
    $p = $Object.PSObject.Properties[$Name]
    if ($p) { return $p.Value } else { return $null }
}

#endregion

#region Directory resolution (cached)

function Resolve-User {
    param([Parameter(Mandatory)][string]$UserId)
    if (-not $script:UserCache.ContainsKey($UserId)) {
        try {
            $u = Invoke-GraphWithRetry -Uri "$script:GraphBase/users/$($UserId)?`$select=id,displayName,userPrincipalName"
            $script:UserCache[$UserId] = [pscustomobject]@{
                Id = $UserId; DisplayName = $u['displayName']; UserPrincipalName = $u['userPrincipalName']
            }
        }
        catch {
            Write-Verbose "Could not resolve user $UserId : $($_.Exception.Message)"
            $script:UserCache[$UserId] = [pscustomobject]@{
                Id = $UserId; DisplayName = "(unresolved) $UserId"; UserPrincipalName = $null
            }
        }
    }
    return $script:UserCache[$UserId]
}

function Resolve-Group {
    param([Parameter(Mandatory)][string]$GroupId)
    if (-not $script:GroupCache.ContainsKey($GroupId)) {
        $name = "(unresolved) $GroupId"
        $members = @()
        try {
            $g = Invoke-GraphWithRetry -Uri "$script:GraphBase/groups/$($GroupId)?`$select=id,displayName"
            $name = $g['displayName']
            $raw = Invoke-GraphPaged -Uri "$script:GraphBase/groups/$GroupId/transitiveMembers/microsoft.graph.user?`$select=id,displayName,userPrincipalName&`$top=999"
            $members = @(foreach ($m in $raw) {
                    $id = Get-Prop $m 'id'
                    if (-not $script:UserCache.ContainsKey($id)) {
                        $script:UserCache[$id] = [pscustomobject]@{
                            Id = $id; DisplayName = (Get-Prop $m 'displayName'); UserPrincipalName = (Get-Prop $m 'userPrincipalName')
                        }
                    }
                    $id
                })
        }
        catch {
            Write-Warning "Could not expand group $GroupId : $($_.Exception.Message)"
        }
        $script:GroupCache[$GroupId] = [pscustomobject]@{ Id = $GroupId; DisplayName = $name; MemberIds = $members }
    }
    return $script:GroupCache[$GroupId]
}

#endregion

#region Core logic

function Get-AgentShareSummary {
    <# Builds a report row (incl. SharingCategory) for one copilotPackageDetail object. #>
    param(
        [Parameter(Mandatory)]$Detail,
        [switch]$IncludeAllowed,
        [int]$MinPeople = 2
    )
    $pkgType = "$(Get-Prop $Detail 'type')".ToLowerInvariant()
    $isOrgPublished = $pkgType -in $script:OrgPublishedTypes

    $entries = @()
    $shared = Get-Prop $Detail 'sharedWithUsersAndGroups'
    if ($shared) { $entries += @($shared) }
    # For agents published by the org, the admin availability list is the sharing audience.
    if ($IncludeAllowed -or $isOrgPublished) {
        $allowed = Get-Prop $Detail 'allowedUsersAndGroups'
        if ($allowed) { $entries += @($allowed) }
    }

    $ownerId = Get-Prop $Detail 'ownerId'
    if ($ownerId -eq $script:EmptyGuid) { $ownerId = $null }
    $people = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $directUsers = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $groups = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $other = 0

    foreach ($e in $entries) {
        $id = Get-Prop $e 'resourceId'
        $type = "$(Get-Prop $e 'resourceType')".ToLowerInvariant()
        if ([string]::IsNullOrWhiteSpace($id)) { continue }
        switch ($type) {
            'user' { if ($id -ne $ownerId) { [void]$directUsers.Add($id); [void]$people.Add($id) } }
            'group' {
                if ($groups.Add($id)) {
                    foreach ($m in (Resolve-Group -GroupId $id).MemberIds) { [void]$people.Add($m) }
                }
            }
            default { $other++ }
        }
    }
    # The owner/creator always has access; don't count them as someone the agent is shared with.
    if ($ownerId) { [void]$people.Remove($ownerId) }

    $userNames = @($directUsers | ForEach-Object {
            $u = Resolve-User -UserId $_
            if ($u.UserPrincipalName) { "$($u.DisplayName) <$($u.UserPrincipalName)>" } else { $u.DisplayName }
        } | Sort-Object)
    $groupNames = @($groups | ForEach-Object {
            $g = Resolve-Group -GroupId $_
            "$($g.DisplayName) ($(@($g.MemberIds).Count) members)"
        } | Sort-Object)

    $join = { param($v) if ($null -eq $v) { '' } else { (@($v) -join ', ') } }

    $owner = ''
    if ($ownerId) {
        $o = Resolve-User -UserId $ownerId
        $owner = if ($o.UserPrincipalName) { "$($o.DisplayName) <$($o.UserPrincipalName)>" } else { $o.DisplayName }
    }

    $availableTo = Get-Prop $Detail 'availableTo'
    $deployedTo = Get-Prop $Detail 'deployedTo'
    $category =
    if ($pkgType -notin $script:OrgBuiltTypes) { 'MicrosoftOrExternal' }
    elseif ($isOrgPublished -and ($availableTo -eq 'allowedForAll' -or $deployedTo -eq 'acquiredForAll')) { 'SharedWithOrganization' }
    elseif ($people.Count -ge $MinPeople) { 'SharedWithMultiple' }
    elseif ($people.Count -gt 0) { 'SharedBelowThreshold' }
    else { 'NotShared' }

    [pscustomobject][ordered]@{
        DisplayName          = Get-Prop $Detail 'displayName'
        Id                   = Get-Prop $Detail 'id'
        SharingCategory      = $category
        Owner                = $owner
        Type                 = Get-Prop $Detail 'type'
        Platform             = Get-Prop $Detail 'platform'
        Publisher            = Get-Prop $Detail 'publisher'
        ElementTypes         = & $join (Get-Prop $Detail 'elementTypes')
        AvailableTo          = $availableTo
        DeployedTo           = $deployedTo
        IsBlocked            = Get-Prop $Detail 'isBlocked'
        UniquePeopleCount    = $people.Count
        DirectUserShares     = $directUsers.Count
        GroupShares          = $groups.Count
        OtherShares          = $other
        SharedWithUsers      = $userNames -join '; '
        SharedWithGroups     = $groupNames -join '; '
        ActiveUsers          = Get-Prop $Detail 'activeUsers'
        LastUsedDateTime     = Get-Prop $Detail 'lastUsedDateTime'
        CreatedDateTime      = Get-Prop $Detail 'createdDateTime'
        LastModifiedDateTime = Get-Prop $Detail 'lastModifiedDateTime'
        AppId                = Get-Prop $Detail 'appId'
        ManifestId           = Get-Prop $Detail 'manifestId'
        AssetId              = Get-Prop $Detail 'assetId'
        AgentIdentityId      = Get-Prop $Detail 'agentIdentityId'
        LinkedEntries        = ''
    }
}

function Merge-LinkedEntries {
    <#
        The registry can hold several entries for one agent (e.g. 'Shared by creator' and
        'Published by your org'). Entries are linked when they share a manifestId,
        agentIdentityId or assetId, or have the same name and platform while at least one of
        them has no agentIdentityId. Each linked group keeps the entry with the widest sharing
        category; the others become 'DuplicateEntry'. Mutates the rows in place.
    #>
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows)
    $orgRows = @($Rows | Where-Object { $script:CategoryRank.ContainsKey($_.SharingCategory) })
    if ($orgRows.Count -lt 2) { return }

    $parent = @{}
    function Find-Root([int]$i) { while ($parent[$i] -ne $i) { $parent[$i] = $parent[$parent[$i]]; $i = $parent[$i] }; $i }
    for ($i = 0; $i -lt $orgRows.Count; $i++) { $parent[$i] = $i }
    $byKey = @{}
    for ($i = 0; $i -lt $orgRows.Count; $i++) {
        $r = $orgRows[$i]
        $keys = @($r.ManifestId, $r.AgentIdentityId, $r.AssetId | Where-Object { $_ -and $_ -ne $script:EmptyGuid } | ForEach-Object { "id:$($_.ToLowerInvariant())" })
        if (-not $r.AgentIdentityId -or $r.AgentIdentityId -eq $script:EmptyGuid) {
            $keys += "nameNoIdentity:$("$($r.DisplayName)|$($r.Platform)".ToLowerInvariant())"
        }
        foreach ($k in $keys) {
            if ($byKey.ContainsKey($k)) { $parent[(Find-Root $i)] = Find-Root $byKey[$k] } else { $byKey[$k] = $i }
        }
    }
    # Name fallback: an entry without agentIdentityId also links to same-name/platform entries that have one.
    for ($i = 0; $i -lt $orgRows.Count; $i++) {
        $r = $orgRows[$i]
        if ($r.AgentIdentityId -and $r.AgentIdentityId -ne $script:EmptyGuid) {
            $k = "nameNoIdentity:$("$($r.DisplayName)|$($r.Platform)".ToLowerInvariant())"
            if ($byKey.ContainsKey($k)) { $parent[(Find-Root $i)] = Find-Root $byKey[$k] }
        }
    }

    $groups = @{}
    for ($i = 0; $i -lt $orgRows.Count; $i++) {
        $root = Find-Root $i
        if (-not $groups.ContainsKey($root)) { $groups[$root] = [System.Collections.Generic.List[object]]::new() }
        $groups[$root].Add($orgRows[$i])
    }
    foreach ($g in $groups.Values) {
        if ($g.Count -lt 2) { continue }
        $keep = $g | Sort-Object -Property @{Expression = { $script:CategoryRank[$_.SharingCategory] }; Descending = $true },
        @{Expression = { $_.Type -in $script:OrgPublishedTypes }; Descending = $true }, @{Expression = { [bool]$_.Owner }; Descending = $true } |
            Select-Object -First 1
        foreach ($r in $g) {
            $r.LinkedEntries = (@($g | Where-Object { $_ -ne $r } | ForEach-Object { "$($_.DisplayName) [$($_.Type)] $($_.Id)" }) -join '; ')
            if ($r -ne $keep) { $r.SharingCategory = 'DuplicateEntry' }
        }
    }
}

function Get-AllAgentDetails {
    $packages = Invoke-GraphPaged -Uri "$script:GraphBase/copilot/admin/catalog/packages"
    Write-Host "Found $($packages.Count) packages in the catalog. Retrieving details..." -ForegroundColor Cyan
    $details = [System.Collections.Generic.List[object]]::new()
    $i = 0
    foreach ($p in $packages) {
        $i++
        $id = Get-Prop $p 'id'
        Write-Progress -Activity 'Retrieving agent details' -Status "$i / $($packages.Count): $(Get-Prop $p 'displayName')" -PercentComplete ([int](100 * $i / [math]::Max(1, $packages.Count)))
        try {
            $detail = Invoke-GraphWithRetry -Uri "$script:GraphBase/copilot/admin/catalog/packages/$([uri]::EscapeDataString($id))"
            # The details response can omit fields the list response has (e.g. manifestId/assetId).
            if ($detail -is [System.Collections.IDictionary] -and $p -is [System.Collections.IDictionary]) {
                foreach ($k in @($p.Keys)) {
                    if (-not $detail.Contains($k) -or $null -eq $detail[$k] -or "$($detail[$k])" -eq '') { $detail[$k] = $p[$k] }
                }
            }
            $details.Add($detail)
        }
        catch {
            $status = Get-GraphStatusCode $_
            Write-Warning "Could not get details for '$(Get-Prop $p 'displayName')' ($id) [HTTP $status]: $($_.Exception.Message)"
            $details.Add($p)  # fall back to summary object (no sharing info)
        }
    }
    Write-Progress -Activity 'Retrieving agent details' -Completed
    return , $details
}

#endregion

#region Output

$script:ReportSections = @(
    [pscustomobject]@{
        Key     = 'NotShared'; FilePrefix = 'AgentsNotShared'
        Title   = 'Agents not shared with anyone (only the creator has access)'
        Columns = 'DisplayName', 'Owner', 'Type', 'Platform', 'ActiveUsers', 'LastUsedDateTime', 'CreatedDateTime', 'IsBlocked', 'Id'
    },
    [pscustomobject]@{
        Key     = 'SharedWithMultiple'; FilePrefix = 'AgentsSharedWithMultiple'
        Title   = 'Agents shared with more than one person'
        Columns = 'DisplayName', 'UniquePeopleCount', 'DirectUserShares', 'GroupShares', 'Owner', 'Type', 'Platform',
        'SharedWithUsers', 'SharedWithGroups', 'ActiveUsers', 'LastUsedDateTime', 'IsBlocked', 'Id'
    },
    [pscustomobject]@{
        Key     = 'SharedWithOrganization'; FilePrefix = 'AgentsSharedWithOrganization'
        Title   = 'Agents shared with the whole organization'
        Columns = 'DisplayName', 'Owner', 'Type', 'Platform', 'Publisher', 'AvailableTo', 'DeployedTo', 'ActiveUsers',
        'LastUsedDateTime', 'IsBlocked', 'Id'
    }
)

function Get-ReportLists {
    <# Splits rows into the three report lists, keyed by section Key. #>
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$AllRows)
    $lists = [ordered]@{}
    foreach ($s in $script:ReportSections) {
        $rows = @($AllRows | Where-Object { $_.SharingCategory -eq $s.Key })
        $sorted = if ($s.Key -eq 'SharedWithMultiple') {
            $rows | Sort-Object -Property @{Expression = 'UniquePeopleCount'; Descending = $true }, DisplayName
        }
        else { $rows | Sort-Object DisplayName }
        $lists[$s.Key] = @($sorted | Where-Object { $null -ne $_ })
    }
    return $lists
}

function ConvertTo-HtmlReport {
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary]$Lists,
        [int]$TotalAgents,
        [int]$MinPeople,
        [string]$Tenant
    )
    $enc = { param($v) [System.Net.WebUtility]::HtmlEncode("$v") }
    $numeric = 'UniquePeopleCount', 'DirectUserShares', 'GroupShares', 'ActiveUsers'

    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine('<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><title>Agent sharing report</title>')
    [void]$sb.AppendLine(@'
<style>
body{font-family:Segoe UI,Arial,sans-serif;margin:24px;color:#222}
h1{font-size:22px;margin-bottom:4px}h2{font-size:18px;margin:28px 0 8px;color:#0f6cbd}.meta{color:#555;margin-bottom:16px}
.cards{display:flex;gap:12px;margin-bottom:16px;flex-wrap:wrap}.card{background:#f3f6fb;border:1px solid #d6e0f0;border-radius:6px;padding:10px 16px}
.card b{display:block;font-size:22px;color:#0f6cbd}.card a{color:inherit;text-decoration:none}
input{padding:6px;width:320px;margin-bottom:10px}.empty{color:#777;font-style:italic}
table{border-collapse:collapse;width:100%;font-size:13px}th,td{border:1px solid #ddd;padding:6px;vertical-align:top;text-align:left}
th{background:#0f6cbd;color:#fff;cursor:pointer;position:sticky;top:0}tr:nth-child(even){background:#fafafa}
td.list{max-width:420px;white-space:pre-line}td.num{text-align:right}
</style></head><body>
'@)
    [void]$sb.AppendLine('<h1>Agent sharing report</h1>')
    [void]$sb.AppendLine("<div class='meta'>Generated $(& $enc (Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz')) &middot; Tenant: $(& $enc $Tenant) &middot; 'More than one person' = &ge; $MinPeople unique people (groups expanded, owner excluded)</div>")
    [void]$sb.Append("<div class='cards'><div class='card'><b>$TotalAgents</b>Packages in catalog</div>")
    foreach ($s in $script:ReportSections) {
        [void]$sb.Append("<div class='card'><a href='#$($s.Key)'><b>$(@($Lists[$s.Key]).Count)</b>$(& $enc $s.Title)</a></div>")
    }
    [void]$sb.AppendLine('</div>')

    foreach ($s in $script:ReportSections) {
        $rows = @($Lists[$s.Key])
        $tid = "t_$($s.Key)"
        [void]$sb.AppendLine("<h2 id='$($s.Key)'>$(& $enc $s.Title) ($($rows.Count))</h2>")
        if ($rows.Count -eq 0) { [void]$sb.AppendLine("<div class='empty'>No agents in this category.</div>"); continue }
        [void]$sb.AppendLine("<input placeholder='Filter...' onkeyup=""f(this,'$tid')"">")
        [void]$sb.Append("<table id='$tid'><thead><tr>")
        $ci = 0
        foreach ($c in $s.Columns) { [void]$sb.Append("<th onclick=""s('$tid',$ci,$(if($c -in $numeric){1}else{0}))"">$(& $enc $c)</th>"); $ci++ }
        [void]$sb.AppendLine('</tr></thead><tbody>')
        foreach ($r in $rows) {
            [void]$sb.Append('<tr>')
            foreach ($c in $s.Columns) {
                $v = $r.$c
                $cls = if ($c -in $numeric) { " class='num'" } elseif ($c -like 'SharedWith*') { " class='list'" } else { '' }
                if ($c -like 'SharedWith*') { $v = ("$v" -split '; ') -join "`n" }
                [void]$sb.Append("<td$cls>$(& $enc $v)</td>")
            }
            [void]$sb.AppendLine('</tr>')
        }
        [void]$sb.AppendLine('</tbody></table>')
    }
    [void]$sb.AppendLine(@'
<script>
function f(inp,id){var q=inp.value.toLowerCase();document.querySelectorAll('#'+id+' tbody tr').forEach(function(r){r.style.display=r.innerText.toLowerCase().indexOf(q)>-1?'':'none';});}
var dir={};function s(id,i,n){var b=document.querySelector('#'+id+' tbody');var rows=Array.from(b.rows);var k=id+i;dir[k]=!dir[k];
rows.sort(function(a,c){var x=a.cells[i].innerText,y=c.cells[i].innerText;if(n){x=parseFloat(x)||0;y=parseFloat(y)||0;return dir[k]?x-y:y-x;}return dir[k]?x.localeCompare(y):y.localeCompare(x);});
rows.forEach(function(r){b.appendChild(r);});}
</script></body></html>
'@)
    return $sb.ToString()
}

function Export-CsvWithHeader {
    <# Export-Csv that still writes a header line when there are no rows. #>
    param([AllowEmptyCollection()][AllowNull()][object[]]$Rows, [string]$Path, [string[]]$Headers)
    $Rows = @($Rows | Where-Object { $null -ne $_ })
    if ($Rows.Count -gt 0) { $Rows | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding utf8 }
    else { (($Headers | ForEach-Object { '"' + $_ + '"' }) -join ',') | Set-Content -LiteralPath $Path -Encoding utf8 }
}

function Write-Reports {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$AllRows,
        [Parameter(Mandatory)][int]$MinPeople,
        [Parameter(Mandatory)][string]$Folder,
        [string]$Tenant,
        [switch]$ExportAll
    )
    $lists = Get-ReportLists -AllRows $AllRows

    if (-not (Test-Path -LiteralPath $Folder)) { New-Item -ItemType Directory -Path $Folder -Force | Out-Null }
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $headers = if ($AllRows.Count -gt 0) { @($AllRows[0].PSObject.Properties.Name) } else { @('DisplayName', 'Id', 'SharingCategory') }

    $csvs = [ordered]@{}
    foreach ($s in $script:ReportSections) {
        $csvs[$s.Key] = Join-Path $Folder "$($s.FilePrefix)_$stamp.csv"
        Export-CsvWithHeader -Rows $lists[$s.Key] -Path $csvs[$s.Key] -Headers $headers
    }

    $html = Join-Path $Folder "AgentSharingReport_$stamp.html"
    ConvertTo-HtmlReport -Lists $lists -TotalAgents $AllRows.Count -MinPeople $MinPeople -Tenant $Tenant |
        Set-Content -LiteralPath $html -Encoding utf8

    $allCsv = $null
    if ($ExportAll) {
        $allCsv = Join-Path $Folder "AllAgents_$stamp.csv"
        Export-CsvWithHeader -Rows @($AllRows | Sort-Object DisplayName) -Path $allCsv -Headers $headers
    }

    [pscustomobject]@{ Lists = $lists; Csvs = $csvs; Html = $html; AllCsv = $allCsv }
}

#endregion

function Invoke-Main {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        throw "Module 'Microsoft.Graph.Authentication' is not installed. Run: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser"
    }
    Import-Module Microsoft.Graph.Authentication

    $connectArgs = @{ NoWelcome = $true }
    if ($TenantId) { $connectArgs['TenantId'] = $TenantId }
    if ($ClientId) {
        if (-not $TenantId) { throw '-TenantId is required for app-only sign-in (-ClientId).' }
        if ($CertificateThumbprint) {
            $connectArgs['ClientId'] = $ClientId
            $connectArgs['CertificateThumbprint'] = $CertificateThumbprint
        }
        elseif ($ClientSecret) {
            $connectArgs['ClientSecretCredential'] = [pscredential]::new($ClientId, $ClientSecret)
        }
        else { throw 'App-only sign-in requires -CertificateThumbprint or -ClientSecret.' }
    }
    else {
        $connectArgs['Scopes'] = $script:RequiredScopes
        if ($UseDeviceCode) { $connectArgs['UseDeviceCode'] = $true }
    }
    Connect-MgGraph @connectArgs
    $ctx = Get-MgContext
    $who = if ($ctx.Account) { $ctx.Account } else { "app $($ctx.AppName) ($($ctx.ClientId))" }
    Write-Host "Connected as $who (tenant $($ctx.TenantId))" -ForegroundColor Green

    try {
        $details = Get-AllAgentDetails
    }
    catch {
        $status = Get-GraphStatusCode $_
        if ($status -eq 403 -or $status -eq 401) {
            throw "Access denied listing packages (HTTP $status). Ensure the tenant has a Microsoft Agent 365 license, the account has the AI Administrator or Global Administrator role, and CopilotPackages.Read.All was consented. Details: $($_.Exception.Message)"
        }
        throw
    }

    Write-Host 'Expanding shares to unique people...' -ForegroundColor Cyan
    $rows = @(foreach ($d in $details) { Get-AgentShareSummary -Detail $d -IncludeAllowed:$IncludeAllowedUsersAndGroups -MinPeople $MinSharedPeople })
    Merge-LinkedEntries -Rows $rows

    $result = Write-Reports -AllRows $rows -MinPeople $MinSharedPeople -Folder $OutputFolder -Tenant $ctx.TenantId -ExportAll:$ExportAllAgents
    Write-ConsoleSummary -AllRows $rows -Result $result
}

function Write-ConsoleSummary {
    param([object[]]$AllRows, $Result)
    $consoleCols = @{
        NotShared              = 'DisplayName', 'Owner', 'Type', 'Platform', 'ActiveUsers', 'IsBlocked'
        SharedWithMultiple     = 'DisplayName', 'UniquePeopleCount', 'DirectUserShares', 'GroupShares', 'Owner', 'Platform', 'IsBlocked'
        SharedWithOrganization = 'DisplayName', 'Owner', 'Type', 'Platform', 'AvailableTo', 'DeployedTo', 'IsBlocked'
    }
    $i = 0
    foreach ($s in $script:ReportSections) {
        $i++
        $list = @($Result.Lists[$s.Key])
        Write-Host ''
        Write-Host "[$i] $($s.Title): $($list.Count)" -ForegroundColor Green
        if ($list.Count -gt 0) { $list | Format-Table -Property $consoleCols[$s.Key] -AutoSize | Out-Host }
    }
    Write-Host ''
    Write-Host 'All packages by category:' -ForegroundColor Cyan
    $AllRows | Group-Object SharingCategory | Sort-Object Count -Descending |
        Format-Table @{n = 'SharingCategory'; e = { $_.Name } }, Count -AutoSize | Out-Host
    foreach ($s in $script:ReportSections) { Write-Host ("CSV  [{0}]: {1}" -f $s.Key, $Result.Csvs[$s.Key]) }
    Write-Host "HTML: $($Result.Html)"
    if ($Result.AllCsv) { Write-Host "All : $($Result.AllCsv)" }
}

# Run only when executed, not when dot-sourced (allows offline testing of the functions).
if ($MyInvocation.InvocationName -ne '.') {
    Invoke-Main
}
