<#
.DISCLAIMER
 The sample scripts are provided AS IS without warranty of any kind. 
    Microsoft further disclaims all implied warranties including, without limitation, 
    any implied warranties of merchantability or of fitness for a particular purpose. 
    The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. 
    In no event shall Microsoft, its authors, or anyone else involved in the creation, 
    production, or delivery of the scripts be liable for any damages whatsoever 
    (including, without limitation, damages for loss of business profits, business interruption, 
    loss of business information, or other pecuniary loss) arising out of the use of or inability 
    to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.
    
.SYNOPSIS
    Finds every site under /sites/ or /teams/ - all site templates, including Microsoft 365
    Group-connected (and Teams-connected) sites - checks each site's SharingCapability, and where
    external sharing is Disabled turns on Restricted Site Access (RSA) and adds up to three groups
    to it. Groups already in RSA, and a group-connected site's own Microsoft 365 group, are always
    kept; the configured groups are added to them.

.DESCRIPTION
    - App-only (certificate) authentication for both Microsoft Graph and SharePoint Online,
      so the script can run unattended (Task Scheduler, Azure Automation, etc.).
    - Reports each step and the progress of the run to the PowerShell session.
    - Honors SharePoint Online / Graph throttling (HTTP 429 and 503): waits for the
      Retry-After value, or uses exponential back-off with jitter, then retries.
    - Writes a CSV log of every site processed.

.NOTES
    Microsoft Graph PowerShell : app-only sign-in and group lookup (runs in a separate
                                 process to avoid the MSAL .dll conflict with PnP.PowerShell)
    PnP.PowerShell             : reads site SharingCapability and sets RSA
                                 (Microsoft Graph cannot do either)

    Required modules : Microsoft.Graph.Authentication, Microsoft.Graph.Groups, PnP.PowerShell
    App permissions  : Microsoft Graph  - Group.Read.All          (Application, admin consent)
                       SharePoint       - Sites.FullControl.All   (Application, admin consent)
    Tenant setting   : Set-SPOTenant -EnableRestrictedAccessControl $true  (one time)
    Licensing        : SharePoint Advanced Management or Microsoft 365 Copilot

.EXAMPLE
    .\Set-RestrictedSiteAccess.ps1 -WhatIf -Verbose   # preview only, lists excluded sites
.EXAMPLE
    .\Set-RestrictedSiteAccess.ps1                    # apply changes
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param()

#region ---------- Configuration ----------
$TenantID              = "00000000-0000-0000-0000-000000000000"   # Entra tenant ID (GUID)
$AppID                 = "00000000-0000-0000-0000-000000000000"   # App registration (client) ID
$CertThumbprint = "0000000000000000000000000000000000000000"
$TenantName            = "contoso"                                 # e.g. contoso -> contoso-admin.sharepoint.com
$AdminUrl              = "https://$TenantName-admin.sharepoint.com"



# Group display name or object ID (GUID). Leave a value empty ("") to skip it.
$Group1 = 'Guid_of_Security_Group_1'
$Group2 = 'Guid_of_Security_Group_2'
#$Group3 = 'Add additional security group GUID here'

# Throttling / retry settings
$MaxRetries       = 8        # attempts per call before giving up
$BaseDelaySeconds = 5        # first back-off when no Retry-After header is returned
$MaxDelaySeconds  = 300      # cap on any single wait
$PacingMs         = 500      # pause after each site update to stay under limits

# Only sites whose server-relative URL starts with one of these managed paths are processed
$IncludedPaths   = @('sites', 'teams')


# Site templates to process. '*' = ALL templates (team, communication, classic, group-connected,
# channel sites, etc.). Every site must ALSO have SharingCapability = Disabled to be changed.
# To limit the run, list templates instead, e.g. @('STS#3', 'GROUP#0', 'SITEPAGEPUBLISHING#0').
# A Microsoft 365 Group-connected site (GROUP#0) always keeps its own group in Restricted Site Access.
$IncludedTemplates = @('*')
$PathFilterRegex = '^/(' + (($IncludedPaths | ForEach-Object { [regex]::Escape($_) }) -join '|') + ')/'

$AdminUrl = "https://$TenantName-admin.sharepoint.com"
$LogFile  = Join-Path $PSScriptRoot ("RSA-Log_{0:yyyyMMdd_HHmmss}.csv" -f (Get-Date))
#endregion

#region ---------- Status / progress helpers ----------
$script:TotalSteps    = 5
$script:ThrottleCount = 0
$script:StartTime     = Get-Date

function Write-Step {
    param([int]$Number, [string]$Title)
    Write-Host ""
    Write-Host ("=" * 78) -ForegroundColor DarkCyan
    Write-Host (" STEP {0}/{1}: {2}" -f $Number, $script:TotalSteps, $Title) -ForegroundColor Cyan
    Write-Host ("=" * 78) -ForegroundColor DarkCyan
}

function Write-Status {
    param(
        [string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR','THROTTLE','WHATIF')][string]$Level = 'INFO'
    )
    $color = @{ INFO='Gray'; OK='Green'; WARN='Yellow'; ERROR='Red'; THROTTLE='Magenta'; WHATIF='DarkYellow' }[$Level]
    Write-Host ("[{0:HH:mm:ss}] [{1,-8}] {2}" -f (Get-Date), $Level, $Message) -ForegroundColor $color
    Write-Verbose $Message
}
#endregion

#region ---------- Throttling: back-off and retry ----------
function Get-ThrottleInfo {
    <# Walks the exception chain looking for HTTP 429/503 and a Retry-After value. #>
    param([System.Exception]$Exception)

    $ex = $Exception
    while ($ex) {
        $status = $null; $retryAfter = $null

        # CSOM / System.Net.WebException
        if ($ex -is [System.Net.WebException] -and $ex.Response) {
            $status = [int]$ex.Response.StatusCode
            $ra = $ex.Response.Headers['Retry-After']
            if ($ra) { $retryAfter = $ra }
        }
        # Microsoft Graph SDK / Kiota (ResponseStatusCode + ResponseHeaders)
        elseif ($ex.PSObject.Properties['ResponseStatusCode']) {
            $status = [int]$ex.ResponseStatusCode
            if ($ex.ResponseHeaders -and $ex.ResponseHeaders['Retry-After']) {
                $retryAfter = @($ex.ResponseHeaders['Retry-After'])[0]
            }
        }
        # HttpResponseException (Invoke-RestMethod / PnP)
        elseif ($ex.PSObject.Properties['Response'] -and $ex.Response -and $ex.Response.PSObject.Properties['StatusCode']) {
            $status = [int]$ex.Response.StatusCode
            if ($ex.Response.Headers.RetryAfter -and $ex.Response.Headers.RetryAfter.Delta) {
                $retryAfter = $ex.Response.Headers.RetryAfter.Delta.TotalSeconds
            }
        }

        if ($status -in 429, 503) {
            return [pscustomobject]@{ IsThrottled = $true; StatusCode = $status; RetryAfter = $retryAfter }
        }
        $ex = $ex.InnerException
    }

    # Fallback: message text (some wrappers drop the response object)
    if ($Exception.ToString() -match '\b(429|503)\b|Too Many Requests|Server Unavailable|throttl') {
        $code = if ($Matches[1]) { $Matches[1] } else { '429/503' }
        return [pscustomobject]@{ IsThrottled = $true; StatusCode = $code; RetryAfter = $null }
    }
    return [pscustomobject]@{ IsThrottled = $false }
}

function Invoke-WithRetry {
    <# Runs a script block; on HTTP 429/503 waits Retry-After (or exponential back-off) and retries. #>
    param(
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [string]$Operation = "SharePoint request"
    )
    $attempt = 0
    while ($true) {
        $attempt++
        try {
            return & $ScriptBlock
        }
        catch {
            $info = Get-ThrottleInfo -Exception $_.Exception
            if (-not $info.IsThrottled -or $attempt -ge $MaxRetries) {
                if ($info.IsThrottled) { Write-Status "$Operation still throttled after $attempt attempts; giving up." -Level ERROR }
                throw
            }

            $script:ThrottleCount++
            [double]$delay = 0
            if ($info.RetryAfter -and [double]::TryParse([string]$info.RetryAfter, [ref]$delay) -and $delay -gt 0) {
                $source = "Retry-After header"
            } else {
                # Exponential back-off with jitter: 5, 10, 20, 40 ... seconds
                $delay  = $BaseDelaySeconds * [math]::Pow(2, $attempt - 1) + (Get-Random -Minimum 0 -Maximum 3)
                $source = "exponential back-off"
            }
            $delay = [int][math]::Min([math]::Ceiling($delay), $MaxDelaySeconds)

            Write-Status ("{0} throttled (HTTP {1}). Waiting {2}s ({3}), retry {4}/{5}..." -f `
                $Operation, $info.StatusCode, $delay, $source, $attempt, ($MaxRetries - 1)) -Level THROTTLE

            for ($s = $delay; $s -gt 0; $s--) {
                Write-Progress -Id 2 -ParentId 1 -Activity "Throttled by SharePoint Online (HTTP $($info.StatusCode))" `
                    -Status "Retrying '$Operation' in $s s" -PercentComplete ((($delay - $s) / $delay) * 100)
                Start-Sleep -Seconds 1
            }
            Write-Progress -Id 2 -Activity "Throttled" -Completed
        }
    }
}

function Invoke-CsomQuery {
    param([Microsoft.SharePoint.Client.ClientContext]$Context, [string]$Operation)
    Invoke-WithRetry -Operation $Operation -ScriptBlock { $Context.ExecuteQuery() }
}
#endregion


$results = [System.Collections.Generic.List[object]]::new()
$counts  = [ordered]@{ Scanned = 0; PathExcluded = 0; TemplateExcluded = 0; Candidates = 0; SharingNotDisabled = 0; SharingDisabled = 0;
             GroupConnected = 0; NotGroupConnected = 0; Updated = 0; AlreadyCompliant = 0; WhatIf = 0; Failed = 0 }

try {
    if ($WhatIfPreference) { Write-Status "Running in WhatIf mode: no changes will be made." -Level WHATIF }

    # ---- STEP 1 + 2: Graph sign-in and group lookup in a SEPARATE PROCESS ----
    # Microsoft.Graph and PnP.PowerShell ship different versions of Microsoft.Identity.Client (MSAL).
    # .NET can only load one version per process, so whichever module loads second fails with
    # "Could not load type 'Microsoft.Identity.Client.IMsalSFHttpClientFactory'".
    # Running the Graph work in a background job (its own process) keeps the two apart.
    Write-Step 1 "Connecting to Microsoft Graph (app-only, certificate) in an isolated process"
    Write-Status "Tenant: $TenantID | App: $AppID | Cert: $CertThumbprint"
    Write-Status "Starting a separate PowerShell process so Graph and PnP MSAL versions don't collide..."

    $graphJob = Start-Job -Name 'RSA-GraphLookup' -ArgumentList `
        $TenantID, $AppID, $CertThumbprint, @($Group1, $Group2, $Group3), $MaxRetries, $BaseDelaySeconds, $MaxDelaySeconds `
        -ScriptBlock {
        param($TenantID, $AppID, $CertThumbprint, $Groups, $MaxRetries, $BaseDelaySeconds, $MaxDelaySeconds)
        $ErrorActionPreference = 'Stop'

        function Say([string]$m, [string]$c = 'Gray') { Write-Host ("[{0:HH:mm:ss}] [GRAPH   ] {1}" -f (Get-Date), $m) -ForegroundColor $c }

        function Invoke-GraphRetry([scriptblock]$Sb, [string]$Op) {
            for ($attempt = 1; ; $attempt++) {
                try { return & $Sb }
                catch {
                    $ex = $_.Exception; $status = $null; $ra = $null
                    while ($ex) {
                        if ($ex.PSObject.Properties['ResponseStatusCode']) {
                            $status = [int]$ex.ResponseStatusCode
                            if ($ex.ResponseHeaders -and $ex.ResponseHeaders['Retry-After']) { $ra = @($ex.ResponseHeaders['Retry-After'])[0] }
                            break
                        }
                        $ex = $ex.InnerException
                    }
                    if (-not $status -and $_.Exception.ToString() -match '\b(429|503)\b|Too Many Requests') { $status = 429 }
                    if ($status -notin 429, 503 -or $attempt -ge $MaxRetries) { throw }
                    [double]$d = 0
                    if (-not ($ra -and [double]::TryParse([string]$ra, [ref]$d) -and $d -gt 0)) {
                        $d = $BaseDelaySeconds * [math]::Pow(2, $attempt - 1) + (Get-Random -Minimum 0 -Maximum 3)
                    }
                    $d = [int][math]::Min([math]::Ceiling($d), $MaxDelaySeconds)
                    Say "$Op throttled (HTTP $status). Waiting ${d}s, retry $attempt/$($MaxRetries - 1)..." 'Magenta'
                    Start-Sleep -Seconds $d
                }
            }
        }

        Import-Module Microsoft.Graph.Authentication, Microsoft.Graph.Groups
        Connect-MgGraph -TenantId $TenantID -ClientId $AppID -CertificateThumbprint $CertThumbprint -NoWelcome
        Say "Connected to Microsoft Graph as app '$((Get-MgContext).AppName)'." 'Green'

        try {
            foreach ($Group in $Groups | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) {
                $guid = [guid]::Empty
                if ([guid]::TryParse($Group, [ref]$guid)) {
                    Say "Looking up group by object ID: $Group"
                    $g = Invoke-GraphRetry { Get-MgGroup -GroupId $Group } "Get group $Group"
                } else {
                    Say "Looking up group by display name: '$Group'"
                    $escaped = $Group.Replace("'", "''")
                    $g = @(Invoke-GraphRetry { Get-MgGroup -Filter "displayName eq '$escaped'" } "Find group '$Group'")
                    if ($g.Count -ne 1) { throw "Group '$Group' returned $($g.Count) matches; use the object ID instead." }
                    $g = $g[0]
                }
                Say "Found '$($g.DisplayName)' -> $($g.Id)" 'Green'
                [pscustomobject]@{ Input = $Group; Id = $g.Id; DisplayName = $g.DisplayName }
            }
        }
        finally {
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        }
    }

    Write-Step 2 "Finding the Restricted Site Access groups"
    Write-Status "Waiting for the Graph lookup process to finish..."
    $null = Wait-Job -Job $graphJob
    try {
        $groupInfo = @(Receive-Job -Job $graphJob -ErrorAction Stop)
        if ($graphJob.State -eq 'Failed') { throw $graphJob.ChildJobs[0].JobStateInfo.Reason }
    }
    finally {
        Remove-Job -Job $graphJob -Force -ErrorAction SilentlyContinue
    }

    $GroupIds = @($groupInfo | ForEach-Object { [guid]$_.Id } | Select-Object -Unique)
    if ($GroupIds.Count -eq 0) { throw "No groups were resolved from `$Group1, `$Group2 or `$Group3." }
    Write-Status "$($GroupIds.Count) unique group(s) ready: $($GroupIds -join ', ')" -Level OK

    # ---- STEP 3: SharePoint admin app-only sign-in ----
    Write-Step 3 "Connecting to SharePoint Online admin center (app-only, certificate)"
    Write-Status "Admin URL: $AdminUrl"
    Invoke-WithRetry -Operation "Connect to $AdminUrl" -ScriptBlock {
        Connect-PnPOnline -Url $AdminUrl -ClientId $AppID -Tenant "$TenantName.onmicrosoft.com" -Thumbprint $CertThumbprint -ErrorAction Stop
    }
    $ctx = Get-PnPContext

    # NOTE: no ExecutingWebRequest event handler here. CSOM raises that event on a .NET
    # worker thread with no PowerShell runspace, which causes "There is no Runspace available
    # to run scripts in this thread". PnP.PowerShell already tags every request with its own
    # "NONISV|SharePointPnP|PnPPS/<version>" User-Agent, which meets Microsoft's guidance.
    $tenant = New-Object Microsoft.Online.SharePoint.TenantAdministration.Tenant($ctx)
    Write-Status "Connected to SharePoint admin center." -Level OK

    # ---- STEP 4: Enumerate and filter sites ----
    Write-Step 4 "Listing all sites under $(($IncludedPaths | ForEach-Object { "/$_/" }) -join ' and ')"
    Write-Status "Getting the site list (this can take several minutes in large tenants)..."
    $sites = @(Invoke-WithRetry -Operation "Get-PnPTenantSite" -ScriptBlock {
        Get-PnPTenantSite -IncludeOneDriveSites:$false -ErrorAction Stop
    })
    $counts.Scanned = $sites.Count
    Write-Status "$($counts.Scanned) sites returned by the tenant." -Level OK

    # Filter 1: managed path must be /sites/ or /teams/
    Write-Status "Applying path filter: $(($IncludedPaths | ForEach-Object { "/$_/" }) -join ', ')"
    $inScope = [System.Collections.Generic.List[object]]::new()
    foreach ($s in $sites) {
        $relativePath = ([uri]$s.Url).AbsolutePath          # e.g. /sites/Finance or /teams/HR
        if ($relativePath -match $PathFilterRegex) { $inScope.Add($s) }
        else { Write-Verbose "Excluded (path): $($s.Url)" }
    }
    $counts.PathExcluded = $counts.Scanned - $inScope.Count
    Write-Status "$($inScope.Count) site(s) in scope; $($counts.PathExcluded) excluded by path (root, /personal/, /portals/, etc.)." -Level OK

    # Filter 2: site template ('*' = every template, including Microsoft 365 Group-connected sites)
    $allTemplates = $IncludedTemplates -contains '*'
    $candidates   = [System.Collections.Generic.List[object]]::new()
    if ($allTemplates) {
        Write-Status "Template filter: ALL site templates are included."
        foreach ($s in $inScope) { $candidates.Add($s) }
    } else {
        Write-Status "Applying template filter: $($IncludedTemplates -join ', ')"
        foreach ($s in $inScope) {
            if ([string]$s.Template -in $IncludedTemplates) { $candidates.Add($s) }
            else { Write-Verbose "Excluded (template $($s.Template)): $($s.Url)" }
        }
    }
    $counts.TemplateExcluded = $inScope.Count - $candidates.Count
    $counts.Candidates       = $candidates.Count
    Write-Status "$($counts.Candidates) site(s) to check; $($counts.TemplateExcluded) excluded by template." -Level OK

    # Show the template mix so you can see what is in scope
    $candidates | Group-Object { [string]$_.Template } | Sort-Object Count -Descending | ForEach-Object {
        Write-Status ("   {0,-28} {1,6} site(s)" -f $(if ($_.Name) { $_.Name } else { '(unknown)' }), $_.Count)
    }
    Write-Status "Each site's SharingCapability is read from its full site properties in Step 5; only sites with external sharing Disabled are changed."

    # ---- STEP 5: Check sharing and apply Restricted Site Access ----
    Write-Step 5 "Checking sharing and applying Restricted Site Access on $($counts.Candidates) site(s)"
    $i = 0
    foreach ($site in $candidates) {
        $i++
        $pct = [int](($i / [math]::Max($candidates.Count, 1)) * 100)
        Write-Progress -Id 1 -Activity "Checking sharing and applying Restricted Site Access" `
            -Status ("Site {0} of {1} | Updated {2} | Skipped (external sharing on) {3} | Failed {4} | Throttled {5}x" -f `
                $i, $candidates.Count, $counts.Updated, $counts.SharingNotDisabled, $counts.Failed, $script:ThrottleCount) `
            -CurrentOperation $site.Url -PercentComplete $pct

        Write-Status "($i/$($candidates.Count)) Processing $($site.Url)  [template $($site.Template)]"
        $entry = [ordered]@{ Url = $site.Url; Template = $site.Template; SiteType = ''; M365GroupId = ''; TeamsConnected = '';
                             SharingCapability = ''; Action = ''; Status = ''; Detail = '' }

        try {
            # Read the full site properties: authoritative SharingCapability, group and RSA state
            Write-Status "   Reading full site properties..."
            $props = $tenant.GetSitePropertiesByUrl($site.Url, $true)
            $ctx.Load($props)
            Invoke-CsomQuery -Context $ctx -Operation "Read properties of $($site.Url)"

            $template      = [string]$props.Template
            $groupId       = [guid]::Empty
            if ($props.GroupId) { $groupId = [guid]$props.GroupId }
            # Channel sites (TEAMCHANNEL#*) carry the PARENT team's group ID but are not group-connected
            $isChannelSite = $template -like 'TEAMCHANNEL#*'
            $isGroupSite   = (-not $isChannelSite) -and (($groupId -ne [guid]::Empty) -or ($template -like 'GROUP#*'))
            $isTeams       = [bool]$props.IsTeamsConnected
            $sharing       = [string]$props.SharingCapability

            $entry.SiteType = if ($isGroupSite) { 'M365 Group-connected' } elseif ($isChannelSite) { 'Teams channel site' } else { 'Not group-connected' }
            $entry.M365GroupId       = if ($isGroupSite) { [string]$groupId } else { '' }
            $entry.TeamsConnected    = $isTeams
            $entry.SharingCapability = $sharing
            Write-Status ("   Type: {0}{1} | SharingCapability: {2}" -f $entry.SiteType, $(if ($isTeams) { ' (Teams-connected)' } else { '' }), $sharing)

            # Sharing check applies to EVERY site type, group-connected or not
            if ($sharing -ne 'Disabled') {
                Write-Status "   External sharing is '$sharing', not Disabled; no change." -Level INFO
                $entry.Action = 'None'; $entry.Status = 'SkippedSharingNotDisabled'
                $counts.SharingNotDisabled++
                $results.Add([pscustomobject]$entry)
                continue
            }
            $counts.SharingDisabled++
            if ($isGroupSite) { $counts.GroupConnected++ } else { $counts.NotGroupConnected++ }

            if ($isGroupSite -and $groupId -eq [guid]::Empty) {
                throw "Site is group-connected but its Microsoft 365 group ID could not be read; skipped so the group is not removed from RSA."
            }

            $rsaOn    = [bool]$props.RestrictedAccessControl
            $existing = @($props.RestrictedAccessControlGroups | Where-Object { $_ } | ForEach-Object { [guid]$_ })
            Write-Status ("   RSA currently {0}; existing groups: {1}" -f `
                ($(if ($rsaOn) { 'ON' } else { 'OFF' })), ($(if ($existing) { $existing -join ', ' } else { 'none' })))

            # Groups that MUST stay in RSA. For a group-connected site this is its own Microsoft 365
            # group: SharePoint adds it by default when RSA is turned on, but writing the group list
            # replaces the whole list, so it has to be included explicitly or it is removed.
            $required = @()
            if ($isGroupSite) {
                $required = @($groupId)
                Write-Status "   Its Microsoft 365 group $groupId will be kept in RSA."
            }

            # Final list = existing + site's own M365 group + configured groups (nothing removed)
            $merged  = @(@($existing) + $required + $GroupIds | Select-Object -Unique)
            $missing = @($merged | Where-Object { $_ -notin $existing })

            if ($rsaOn -and $missing.Count -eq 0) {
                Write-Status "   Already compliant; skipping." -Level OK
                $entry.Action = 'None'; $entry.Status = 'AlreadyCompliant'
                $counts.AlreadyCompliant++
            }
            elseif ($merged.Count -gt 10) {
                throw "The site would have $($merged.Count) RSA groups (including any Microsoft 365 group); the limit is 10."
            }
            elseif ($PSCmdlet.ShouldProcess($site.Url, "Turn on Restricted Site Access and add $($missing.Count) group(s)")) {
                Write-Status "   Turning on RSA and adding $($missing.Count) group(s): $($missing -join ', ')"
                $props.RestrictedAccessControl       = $true
                $props.RestrictedAccessControlGroups = [guid[]]$merged
                $op = $props.Update()
                $ctx.Load($op)
                Invoke-CsomQuery -Context $ctx -Operation "Update RSA on $($site.Url)"

                # Verify: re-read the site and confirm every required group is present
                Write-Status "   Verifying RSA groups..."
                $check = $tenant.GetSitePropertiesByUrl($site.Url, $true)
                $ctx.Load($check)
                Invoke-CsomQuery -Context $ctx -Operation "Verify RSA on $($site.Url)"
                $after    = @($check.RestrictedAccessControlGroups | Where-Object { $_ } | ForEach-Object { [guid]$_ })
                $notThere = @($merged | Where-Object { $_ -notin $after })
                if ($notThere.Count -gt 0) {
                    Write-Status "   Missing after update: $($notThere -join ', '). Re-applying full group list..." -Level WARN
                    $check.RestrictedAccessControl       = $true
                    $check.RestrictedAccessControlGroups = [guid[]]@(@($after) + $merged | Select-Object -Unique)
                    $op2 = $check.Update(); $ctx.Load($op2)
                    Invoke-CsomQuery -Context $ctx -Operation "Repair RSA on $($site.Url)"
                }

                Write-Status "   RSA enabled. Groups now: $($merged -join ', ')" -Level OK
                $entry.Action = 'RSA enabled'; $entry.Status = 'Success'; $entry.Detail = "Groups: $($merged -join ';')"
                $counts.Updated++
                Start-Sleep -Milliseconds $PacingMs
            }
            else {
                Write-Status "   Would turn on RSA; final groups would be: $($merged -join ', ')" -Level WHATIF
                $entry.Action = 'WhatIf'; $entry.Status = 'Skipped'; $entry.Detail = "Groups: $($merged -join ';')"
                $counts.WhatIf++
            }
        }
        catch {
            $entry.Status = 'Failed'; $entry.Detail = $_.Exception.Message
            $counts.Failed++
            Write-Status "   FAILED: $($_.Exception.Message)" -Level ERROR
        }
        $results.Add([pscustomobject]$entry)
    }
    Write-Progress -Id 1 -Activity "Checking sharing and applying Restricted Site Access" -Completed
}
catch {
    Write-Status "Fatal: $($_.Exception.Message)" -Level ERROR
}
finally {
    Write-Host ""
    Write-Host ("=" * 78) -ForegroundColor DarkCyan
    Write-Host " SUMMARY" -ForegroundColor Cyan
    Write-Host ("=" * 78) -ForegroundColor DarkCyan
    $counts.GetEnumerator() | ForEach-Object { Write-Host ("  {0,-20}: {1}" -f $_.Key, $_.Value) }
    Write-Host ("  {0,-20}: {1}" -f 'Throttle events', $script:ThrottleCount)
    Write-Host ("  {0,-20}: {1:hh\:mm\:ss}" -f 'Elapsed', ((Get-Date) - $script:StartTime))

    if ($results.Count) {
        # -WhatIf:$false so the log is still written during a preview run
        $results | Export-Csv -Path $LogFile -NoTypeInformation -WhatIf:$false
        Write-Status "Log saved to $LogFile" -Level OK
    }
    Write-Status "Disconnecting SharePoint session..."
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
    Write-Status "Done." -Level OK
}
