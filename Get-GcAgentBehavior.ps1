<#
.SYNOPSIS
    Genesys Cloud agent behavior audit tool - Constrained Language Mode (CLM) safe.

.DESCRIPTION
    Pulls a per-agent behavior picture from the Genesys Cloud Platform API:

      1. Call / interaction activity   (analytics conversation aggregates, grouped by agent)
      2. Presence & routing status     (analytics user aggregates - time in each status)
      3. WFM real-time adherence       (workforce management adherence state per agent)
      4. Knowledge Base feedback       (how often each agent submits KB article feedback)
      5. Copilot / AI summary activity (summaries per conversation, feedback + edit signals)
      6. Edit an AI Copilot note       (-EditCopilotNote action: update a conversation summary)

    Everything is exported as CSV into an output folder, plus a per-agent roll-up
    (AgentBehaviorSummary.csv) that joins all data sets on the agent.

    CONSTRAINED LANGUAGE MODE COMPLIANCE (Windows PowerShell 5.1, AppLocker/WDAC):
      - No .NET static method calls ([Convert]::, [Text.Encoding]::, [uri]::, [Guid]::, [math]::)
      - Base64 for the OAuth Basic header is implemented in pure PowerShell using
        bit operators (-shl, -shr, -band, -bor) over [int][char] values
      - No ::new() constructors - plain arrays with += only
      - No [pscustomobject]@{} casts - New-Object PSObject -Property @{} instead,
        with explicit Select-Object column ordering before every Export-Csv
      - URL encoding done with minimal string .Replace() only
      - HTTP status codes read via [int]$_.Exception.Response.StatusCode inside
        try/catch, with fallback string matching on $_.Exception.Message
      - Every API response validated explicitly before use

.PARAMETER ClientId
    OAuth Client Credentials grant client id. Optional - falls back to the
    EMBEDDED CREDENTIALS block near the top of the script.

.PARAMETER ClientSecret
    OAuth Client Credentials grant client secret. Optional - falls back to the
    EMBEDDED CREDENTIALS block near the top of the script.

.PARAMETER Region
    Genesys Cloud region API domain. Defaults to the embedded region -
    mypurecloud.com.au (Australia / Sydney). Other examples: mypurecloud.com,
    mypurecloud.ie, usw2.pure.cloud, euw2.pure.cloud, cac1.pure.cloud ...

.PARAMETER DaysBack
    Reporting window: now minus N days, in UTC. Default 7.

.PARAMETER OutputDir
    Folder for CSV output. Created if missing. Default: .\GcAgentBehavior_<yyyyMMdd_HHmmss>

.PARAMETER AgentEmailFilter
    Optional wildcard filter on agent email (e.g. "*@contoso.com" or "j.smith*").

.PARAMETER MaxConversationsForCopilot
    Cap on conversations scanned for Copilot summaries (1 API call each). Default 200.

.PARAMETER MaxKbDocuments
    Cap on KB documents scanned for feedback (1 API call each). Default 500.

.PARAMETER SkipWfm / SkipKb / SkipCopilot
    Skip the corresponding (heavier / permission-sensitive) sections.

.PARAMETER EditCopilotNote
    Action switch: instead of reporting, edit one AI Copilot note (conversation
    summary). Requires -ConversationId, -SummaryId and -NewNoteText.

.EXAMPLE
    # Uses the embedded credentials and Australia region - no arguments needed
    .\Get-GcAgentBehavior.ps1 -DaysBack 14

.EXAMPLE
    .\Get-GcAgentBehavior.ps1 -EditCopilotNote -ConversationId 'abc-123' `
        -SummaryId 'def-456' -NewNoteText 'Corrected summary text'

.EXAMPLE
    # Command-line credentials still override the embedded block
    .\Get-GcAgentBehavior.ps1 -ClientId $id -ClientSecret $secret -Region 'mypurecloud.ie'

.NOTES
    Required OAuth scopes / permissions (grant only what you use):
      analytics:conversationAggregate:view, analytics:userAggregate:view,
      analytics:conversationDetail:view, directory:user:view,
      wfm:realtimeAdherence:view, knowledge:knowledgebase:view,
      knowledge:document:view, conversation:summary:view (+ edit for -EditCopilotNote)

    The Copilot summary feedback / edit endpoints are the newest surface in this
    script and Genesys still iterates on them. If your org gets 404s there, open
    the API Explorer (https://developer.genesys.cloud/devapps/api-explorer),
    search "summaries", and adjust the two $script:Copilot*PathTemplate values
    near the top of the script - nothing else needs to change.
#>

[CmdletBinding()]
param(
    # ClientId / ClientSecret / Region fall back to the embedded values in the
    # EMBEDDED CREDENTIALS block below when not supplied on the command line.
    [Parameter(Mandatory = $false)] [string]$ClientId = '',
    [Parameter(Mandatory = $false)] [string]$ClientSecret = '',
    [Parameter(Mandatory = $false)] [string]$Region = '',
    [Parameter(Mandatory = $false)] [int]$DaysBack = 7,
    [Parameter(Mandatory = $false)] [string]$OutputDir = '',
    [Parameter(Mandatory = $false)] [string]$AgentEmailFilter = '*',
    [Parameter(Mandatory = $false)] [int]$MaxConversationsForCopilot = 200,
    [Parameter(Mandatory = $false)] [int]$MaxKbDocuments = 500,
    [Parameter(Mandatory = $false)] [switch]$SkipWfm,
    [Parameter(Mandatory = $false)] [switch]$SkipKb,
    [Parameter(Mandatory = $false)] [switch]$SkipCopilot,
    [Parameter(Mandatory = $false)] [switch]$EditCopilotNote,
    [Parameter(Mandatory = $false)] [string]$ConversationId = '',
    [Parameter(Mandatory = $false)] [string]$SummaryId = '',
    [Parameter(Mandatory = $false)] [string]$NewNoteText = ''
)

$ErrorActionPreference = 'Stop'

# ===========================================================================
# EMBEDDED CREDENTIALS - Australia (APSE2) region
# ---------------------------------------------------------------------------
# Paste your OAuth Client Credentials here on YOUR LOCAL COPY ONLY.
# NEVER commit real credentials to source control - anyone with repo access
# (including via pull requests and git history) can read them. If a secret
# does get committed, treat it as compromised: delete/rotate the OAuth client
# in Genesys Cloud Admin immediately.
# Command-line -ClientId / -ClientSecret / -Region still override these.
# ===========================================================================
$script:EmbeddedClientId     = 'PASTE-YOUR-CLIENT-ID-HERE'
$script:EmbeddedClientSecret = 'PASTE-YOUR-CLIENT-SECRET-HERE'
$script:EmbeddedRegion       = 'mypurecloud.com.au'   # Australia (Sydney / APSE2)

if ($ClientId -eq '')     { $ClientId     = $script:EmbeddedClientId }
if ($ClientSecret -eq '') { $ClientSecret = $script:EmbeddedClientSecret }
if ($Region -eq '')       { $Region       = $script:EmbeddedRegion }

if ($ClientId -like 'PASTE-YOUR-*' -or $ClientSecret -like 'PASTE-YOUR-*') {
    throw 'No credentials configured. Edit the EMBEDDED CREDENTIALS block near the top of this script (local copy only), or pass -ClientId and -ClientSecret on the command line.'
}

# ---------------------------------------------------------------------------
# Endpoint templates most likely to differ between orgs / API releases.
# {0} = conversationId, {1} = summaryId. Verify in the API Explorer if needed.
# ---------------------------------------------------------------------------
$script:CopilotSummariesPathTemplate = '/api/v2/conversations/{0}/summaries'
$script:CopilotEditPathTemplate      = '/api/v2/conversations/{0}/summaries/{1}'

$script:ApiBase   = 'https://api.' + $Region
$script:LoginBase = 'https://login.' + $Region
$script:Token     = ''

# ===========================================================================
# region CLM-safe primitives
# ===========================================================================

function ConvertTo-GcBase64 {
    <#  Pure-PowerShell Base64 encoder (no [Convert]:: / [Text.Encoding]::).
        Operates on [int][char] values, so it is correct for ASCII input -
        which OAuth clientId:clientSecret always is.  #>
    param([Parameter(Mandatory = $true)][string]$Text)

    $alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/'
    $bytes = @()
    foreach ($ch in $Text.ToCharArray()) { $bytes += [int][char]$ch }

    $out = ''
    $i = 0
    while ($i -lt $bytes.Count) {
        $b0 = $bytes[$i]
        $b1 = -1
        $b2 = -1
        if (($i + 1) -lt $bytes.Count) { $b1 = $bytes[$i + 1] }
        if (($i + 2) -lt $bytes.Count) { $b2 = $bytes[$i + 2] }

        $out += [string]$alphabet[(($b0 -shr 2) -band 63)]
        if ($b1 -ge 0) {
            $out += [string]$alphabet[((($b0 -shl 4) -bor ($b1 -shr 4)) -band 63)]
            if ($b2 -ge 0) {
                $out += [string]$alphabet[((($b1 -shl 2) -bor ($b2 -shr 6)) -band 63)]
                $out += [string]$alphabet[($b2 -band 63)]
            } else {
                $out += [string]$alphabet[(($b1 -shl 2) -band 63)]
                $out += '='
            }
        } else {
            $out += [string]$alphabet[(($b0 -shl 4) -band 63)]
            $out += '=='
        }
        $i += 3
    }
    return $out
}

function ConvertTo-GcUrlEncoded {
    <#  Minimal percent-encoding via string .Replace() only (no [uri]::).
        Covers the characters that actually occur in ids / query values.  #>
    param([Parameter(Mandatory = $true)][string]$Value)
    $v = $Value
    $v = $v.Replace('%', '%25')   # must be first
    $v = $v.Replace('&', '%26')
    $v = $v.Replace('+', '%2B')
    $v = $v.Replace('/', '%2F')
    $v = $v.Replace('=', '%3D')
    $v = $v.Replace('?', '%3F')
    $v = $v.Replace('#', '%23')
    $v = $v.Replace(' ', '%20')
    return $v
}

function Get-GcHttpStatusFromError {
    <#  Extract an HTTP status code from an ErrorRecord the CLM-safe way:
        [int] cast on the Response.StatusCode first, string matching fallback.  #>
    param([Parameter(Mandatory = $true)]$ErrorRecord)

    $code = 0
    try {
        if ($ErrorRecord.Exception.Response) {
            $code = [int]$ErrorRecord.Exception.Response.StatusCode
        }
    } catch { $code = 0 }

    if ($code -eq 0) {
        $msg = ''
        try { $msg = [string]$ErrorRecord.Exception.Message } catch { $msg = '' }
        if     ($msg -like '*429*') { $code = 429 }
        elseif ($msg -like '*401*') { $code = 401 }
        elseif ($msg -like '*403*') { $code = 403 }
        elseif ($msg -like '*404*') { $code = 404 }
        elseif ($msg -like '*400*') { $code = 400 }
    }
    return $code
}

function Format-GcHours {
    <#  Milliseconds -> hours with 2 decimals, without [math]:: statics.  #>
    param([Parameter(Mandatory = $true)]$Milliseconds)
    $ms = 0
    try { $ms = [double]$Milliseconds } catch { $ms = 0 }
    return ('{0:N2}' -f ($ms / 3600000))
}

# endregion

# ===========================================================================
# region Auth + API wrapper
# ===========================================================================

function Connect-GcCloud {
    Write-Host ('Authenticating against ' + $script:LoginBase + ' ...') -ForegroundColor Cyan

    $basic = ConvertTo-GcBase64 -Text ($ClientId + ':' + $ClientSecret)
    $headers = @{ Authorization = ('Basic ' + $basic) }
    $body = @{ grant_type = 'client_credentials' }

    $resp = $null
    try {
        $resp = Invoke-RestMethod -Method Post -Uri ($script:LoginBase + '/oauth/token') `
            -Headers $headers -Body $body -ContentType 'application/x-www-form-urlencoded'
    } catch {
        $status = Get-GcHttpStatusFromError -ErrorRecord $_
        throw ('OAuth token request failed (HTTP ' + $status + '). Check ClientId/ClientSecret/Region. Detail: ' + $_.Exception.Message)
    }

    # Explicit validation - never assume success.
    if ($null -eq $resp -or $null -eq $resp.access_token -or ([string]$resp.access_token).Length -lt 10) {
        throw 'OAuth response did not contain a usable access_token. Aborting.'
    }

    $script:Token = [string]$resp.access_token
    Write-Host ('Token acquired (expires in ' + [string]$resp.expires_in + 's).') -ForegroundColor Green
}

function Invoke-GcApi {
    <#  Central API wrapper: bearer auth, JSON handling, 429 backoff, and
        graceful 403/404 handling (returns $null instead of dying) so one
        missing permission never kills the whole report.  #>
    param(
        [Parameter(Mandatory = $false)][string]$Method = 'GET',
        [Parameter(Mandatory = $true)] [string]$Path,
        [Parameter(Mandatory = $false)]$Body = $null,
        [Parameter(Mandatory = $false)][switch]$SoftFail
    )

    $uri = $script:ApiBase + $Path
    $headers = @{
        Authorization  = ('Bearer ' + $script:Token)
        'Content-Type' = 'application/json'
    }

    $jsonBody = $null
    if ($null -ne $Body) { $jsonBody = ConvertTo-Json -InputObject $Body -Depth 12 }

    $attempt = 0
    $maxAttempts = 5
    while ($attempt -lt $maxAttempts) {
        $attempt += 1
        try {
            if ($null -ne $jsonBody) {
                return Invoke-RestMethod -Method $Method -Uri $uri -Headers $headers -Body $jsonBody
            }
            return Invoke-RestMethod -Method $Method -Uri $uri -Headers $headers
        } catch {
            $status = Get-GcHttpStatusFromError -ErrorRecord $_

            if ($status -eq 429) {
                $wait = 3 * $attempt
                Write-Host ('  Rate limited (429) on ' + $Path + ' - waiting ' + $wait + 's (attempt ' + $attempt + '/' + $maxAttempts + ')') -ForegroundColor Yellow
                Start-Sleep -Seconds $wait
                continue
            }
            if ($status -eq 401) {
                # Token may have expired mid-run: refresh once and retry.
                Write-Host '  401 received - refreshing token and retrying...' -ForegroundColor Yellow
                Connect-GcCloud
                $headers['Authorization'] = ('Bearer ' + $script:Token)
                continue
            }
            if ($status -eq 403) {
                Write-Warning ('Permission denied (403) for ' + $Method + ' ' + $Path + ' - section skipped. Grant the matching permission to the OAuth client to enable it.')
                return $null
            }
            if ($status -eq 404) {
                Write-Verbose ('404 for ' + $Method + ' ' + $Path)
                return $null
            }

            if ($SoftFail) {
                Write-Warning ('API call failed (HTTP ' + $status + '): ' + $Method + ' ' + $Path + ' :: ' + $_.Exception.Message)
                return $null
            }
            throw ('API call failed (HTTP ' + $status + '): ' + $Method + ' ' + $Path + ' :: ' + $_.Exception.Message)
        }
    }

    Write-Warning ('Gave up on ' + $Method + ' ' + $Path + ' after ' + $maxAttempts + ' attempts (persistent rate limiting).')
    return $null
}

# endregion

# ===========================================================================
# region Data collectors
# ===========================================================================

function Get-GcInterval {
    <#  ISO-8601 interval string "start/end" in UTC for the analytics APIs.  #>
    $endUtc   = (Get-Date).ToUniversalTime()
    $startUtc = $endUtc.AddDays(0 - $DaysBack)
    $startStr = Get-Date -Date $startUtc -Format 'yyyy-MM-ddTHH:mm:ss'
    $endStr   = Get-Date -Date $endUtc   -Format 'yyyy-MM-ddTHH:mm:ss'
    return ($startStr + '.000Z/' + $endStr + '.000Z')
}

function Get-GcUsers {
    Write-Host 'Loading users (agents)...' -ForegroundColor Cyan
    $users = @()
    $page = 1
    while ($true) {
        $resp = Invoke-GcApi -Method 'GET' -Path ('/api/v2/users?pageSize=200&state=active&pageNumber=' + $page)
        if ($null -eq $resp -or $null -eq $resp.entities) { break }
        foreach ($u in $resp.entities) {
            $email = ''
            if ($null -ne $u.email) { $email = [string]$u.email }
            if ($email -like $AgentEmailFilter -or [string]$u.name -like $AgentEmailFilter) {
                $users += New-Object PSObject -Property @{
                    UserId     = [string]$u.id
                    Name       = [string]$u.name
                    Email      = $email
                    Department = [string]$u.department
                    Title      = [string]$u.title
                }
            }
        }
        if ($null -eq $resp.pageCount -or $page -ge [int]$resp.pageCount) { break }
        $page += 1
    }
    Write-Host ('  ' + $users.Count + ' users matched filter "' + $AgentEmailFilter + '".') -ForegroundColor Green
    return ,$users
}

function New-GcUserIdFilter {
    <#  Builds the analytics "or userId matches" predicate filter for a chunk of ids.  #>
    param([Parameter(Mandatory = $true)]$UserIds)
    $predicates = @()
    foreach ($id in $UserIds) {
        $predicates += @{ type = 'dimension'; dimension = 'userId'; operator = 'matches'; value = [string]$id }
    }
    return @{ type = 'or'; predicates = $predicates }
}

function Get-GcCallActivity {
    <#  Section 1: conversation aggregates grouped by agent - offered/answered
        counts plus talk / hold / ACW / handle time.  #>
    param($Users, [string]$Interval)

    Write-Host 'Querying call/interaction activity (conversation aggregates)...' -ForegroundColor Cyan
    $rows = @()
    $chunkSize = 100
    $index = 0

    while ($index -lt $Users.Count) {
        $chunk = @()
        $j = $index
        while ($j -lt $Users.Count -and $j -lt ($index + $chunkSize)) { $chunk += $Users[$j].UserId; $j += 1 }
        $index += $chunkSize

        $body = @{
            interval = $Interval
            groupBy  = @('userId')
            metrics  = @('nOffered', 'tAnswered', 'tTalk', 'tHeld', 'tAcw', 'tHandle', 'tNotResponding')
            filter   = New-GcUserIdFilter -UserIds $chunk
        }
        $resp = Invoke-GcApi -Method 'POST' -Path '/api/v2/analytics/conversations/aggregates/query' -Body $body -SoftFail
        if ($null -eq $resp -or $null -eq $resp.results) { continue }

        foreach ($result in $resp.results) {
            $uid = ''
            if ($null -ne $result.group -and $null -ne $result.group.userId) { $uid = [string]$result.group.userId }
            if ($uid -eq '') { continue }

            $offered = 0; $answered = 0; $talkMs = 0; $heldMs = 0; $acwMs = 0; $handleMs = 0; $noRespond = 0
            foreach ($d in $result.data) {
                foreach ($m in $d.metrics) {
                    $sum = 0; $count = 0
                    if ($null -ne $m.stats) {
                        if ($null -ne $m.stats.sum)   { $sum   = [double]$m.stats.sum }
                        if ($null -ne $m.stats.count) { $count = [int]$m.stats.count }
                    }
                    if     ($m.metric -eq 'nOffered')       { $offered   += $count }
                    elseif ($m.metric -eq 'tAnswered')      { $answered  += $count }
                    elseif ($m.metric -eq 'tTalk')          { $talkMs    += $sum }
                    elseif ($m.metric -eq 'tHeld')          { $heldMs    += $sum }
                    elseif ($m.metric -eq 'tAcw')           { $acwMs     += $sum }
                    elseif ($m.metric -eq 'tHandle')        { $handleMs  += $sum }
                    elseif ($m.metric -eq 'tNotResponding') { $noRespond += $count }
                }
            }

            $avgHandleSec = 0
            if ($answered -gt 0) { $avgHandleSec = [int](($handleMs / $answered) / 1000) }

            $rows += New-Object PSObject -Property @{
                UserId            = $uid
                Offered           = $offered
                Answered          = $answered
                TalkHours         = Format-GcHours -Milliseconds $talkMs
                HoldHours         = Format-GcHours -Milliseconds $heldMs
                AcwHours          = Format-GcHours -Milliseconds $acwMs
                HandleHours       = Format-GcHours -Milliseconds $handleMs
                AvgHandleSeconds  = $avgHandleSec
                NotRespondingCnt  = $noRespond
            }
        }
    }
    Write-Host ('  Call activity rows: ' + $rows.Count) -ForegroundColor Green
    return ,$rows
}

function Get-GcStatusTime {
    <#  Section 2: user aggregates - time spent in each organization presence
        (Available, Busy, Away, Meal, Meeting, custom...) and each routing
        status (IDLE, INTERACTING, NOT_RESPONDING, OFF_QUEUE...).  #>
    param($Users, [string]$Interval)

    Write-Host 'Querying presence & routing status time (user aggregates)...' -ForegroundColor Cyan
    $rows = @()
    $chunkSize = 100
    $index = 0

    while ($index -lt $Users.Count) {
        $chunk = @()
        $j = $index
        while ($j -lt $Users.Count -and $j -lt ($index + $chunkSize)) { $chunk += $Users[$j].UserId; $j += 1 }
        $index += $chunkSize

        $body = @{
            interval = $Interval
            groupBy  = @('userId')
            metrics  = @('tOrganizationPresence', 'tSystemPresence', 'tAgentRoutingStatus')
            filter   = New-GcUserIdFilter -UserIds $chunk
        }
        $resp = Invoke-GcApi -Method 'POST' -Path '/api/v2/analytics/users/aggregates/query' -Body $body -SoftFail
        if ($null -eq $resp -or $null -eq $resp.results) { continue }

        foreach ($result in $resp.results) {
            $uid = ''
            if ($null -ne $result.group -and $null -ne $result.group.userId) { $uid = [string]$result.group.userId }
            if ($uid -eq '') { continue }

            foreach ($d in $result.data) {
                foreach ($m in $d.metrics) {
                    $sum = 0
                    if ($null -ne $m.stats -and $null -ne $m.stats.sum) { $sum = [double]$m.stats.sum }
                    $qualifier = ''
                    if ($null -ne $m.qualifier) { $qualifier = [string]$m.qualifier }

                    $metricType = 'Other'
                    if     ($m.metric -eq 'tOrganizationPresence') { $metricType = 'Presence' }
                    elseif ($m.metric -eq 'tSystemPresence')       { $metricType = 'SystemPresence' }
                    elseif ($m.metric -eq 'tAgentRoutingStatus')   { $metricType = 'RoutingStatus' }

                    $rows += New-Object PSObject -Property @{
                        UserId     = $uid
                        MetricType = $metricType
                        Status     = $qualifier
                        Hours      = Format-GcHours -Milliseconds $sum
                        RawMs      = [long]$sum
                    }
                }
            }
        }
    }
    Write-Host ('  Status rows: ' + $rows.Count) -ForegroundColor Green
    return ,$rows
}

function Get-GcWfmAdherence {
    <#  Section 3: WFM real-time adherence per agent - current adherence state,
        scheduled vs actual activity category, and impact.  #>
    param($Users)

    Write-Host 'Querying WFM real-time adherence...' -ForegroundColor Cyan
    $rows = @()
    $chunkSize = 50
    $index = 0

    while ($index -lt $Users.Count) {
        $params = @()
        $j = $index
        while ($j -lt $Users.Count -and $j -lt ($index + $chunkSize)) {
            $params += ('userId=' + (ConvertTo-GcUrlEncoded -Value $Users[$j].UserId))
            $j += 1
        }
        $index += $chunkSize

        $resp = Invoke-GcApi -Method 'GET' -Path ('/api/v2/workforcemanagement/adherence?' + ($params -join '&')) -SoftFail
        if ($null -eq $resp) { continue }

        # This endpoint returns a bare array.
        foreach ($a in $resp) {
            $uid = ''
            if ($null -ne $a.user -and $null -ne $a.user.id) { $uid = [string]$a.user.id }
            elseif ($null -ne $a.userId) { $uid = [string]$a.userId }
            if ($uid -eq '') { continue }

            $rows += New-Object PSObject -Property @{
                UserId                    = $uid
                AdherenceState            = [string]$a.adherenceState
                ScheduledActivityCategory = [string]$a.scheduledActivityCategory
                ActualActivityCategory    = [string]$a.actualActivityCategory
                Impact                    = [string]$a.impact
                IsOutOfOffice             = [string]$a.isOutOfOffice
                TimeOfAdherenceChange     = [string]$a.timeOfAdherenceChange
            }
        }
    }
    Write-Host ('  WFM adherence rows: ' + $rows.Count) -ForegroundColor Green
    return ,$rows
}

function Get-GcKbFeedback {
    <#  Section 4: Knowledge Base feedback frequency per agent.
        Walks knowledge bases -> documents -> per-document feedback records and
        counts submissions per user. Document scanning is capped by
        -MaxKbDocuments because each document costs one API call.  #>

    Write-Host ('Querying Knowledge Base feedback (this walks KB documents - capped at ' + $MaxKbDocuments + ')...') -ForegroundColor Cyan
    $rows = @()
    $docsScanned = 0

    $kbResp = Invoke-GcApi -Method 'GET' -Path '/api/v2/knowledge/knowledgebases?pageSize=100' -SoftFail
    if ($null -eq $kbResp -or $null -eq $kbResp.entities) {
        Write-Warning '  No knowledge bases visible (missing permission or none exist).'
        return ,$rows
    }

    foreach ($kb in $kbResp.entities) {
        $kbId = [string]$kb.id
        $kbName = [string]$kb.name
        Write-Host ('  Scanning KB "' + $kbName + '"...') -ForegroundColor DarkCyan

        # Knowledge v2 document listing uses cursor paging via nextUri/after.
        $docPath = '/api/v2/knowledge/knowledgebases/' + $kbId + '/documents?pageSize=100'
        while ($null -ne $docPath -and $docPath -ne '' -and $docsScanned -lt $MaxKbDocuments) {
            $docResp = Invoke-GcApi -Method 'GET' -Path $docPath -SoftFail
            if ($null -eq $docResp -or $null -eq $docResp.entities) { break }

            foreach ($doc in $docResp.entities) {
                if ($docsScanned -ge $MaxKbDocuments) { break }
                $docsScanned += 1
                $docId = [string]$doc.id

                $fbResp = Invoke-GcApi -Method 'GET' -Path ('/api/v2/knowledge/knowledgebases/' + $kbId + '/documents/' + $docId + '/feedback?pageSize=100') -SoftFail
                if ($null -eq $fbResp -or $null -eq $fbResp.entities) { continue }

                foreach ($fb in $fbResp.entities) {
                    # Feedback attribution field has varied across releases - try each.
                    $fbUserId = ''
                    if ($null -ne $fb.userId) { $fbUserId = [string]$fb.userId }
                    elseif ($null -ne $fb.createdBy -and $null -ne $fb.createdBy.id) { $fbUserId = [string]$fb.createdBy.id }
                    elseif ($null -ne $fb.agent -and $null -ne $fb.agent.id) { $fbUserId = [string]$fb.agent.id }

                    $rating = ''
                    if ($null -ne $fb.rating) { $rating = [string]$fb.rating }
                    $reason = ''
                    if ($null -ne $fb.reason) { $reason = [string]$fb.reason }

                    $rows += New-Object PSObject -Property @{
                        UserId        = $fbUserId
                        KnowledgeBase = $kbName
                        DocumentId    = $docId
                        DocumentTitle = [string]$doc.title
                        Rating        = $rating
                        Reason        = $reason
                        Comment       = [string]$fb.comment
                        DateCreated   = [string]$fb.dateCreated
                    }
                }
            }

            # Follow cursor if present; otherwise stop.
            $docPath = ''
            if ($null -ne $docResp.nextUri -and [string]$docResp.nextUri -ne '') {
                $docPath = [string]$docResp.nextUri
            }
        }
        if ($docsScanned -ge $MaxKbDocuments) {
            Write-Warning ('  Hit MaxKbDocuments cap (' + $MaxKbDocuments + '). Raise -MaxKbDocuments for full coverage.')
            break
        }
    }

    Write-Host ('  KB feedback records: ' + $rows.Count + ' (from ' + $docsScanned + ' documents)') -ForegroundColor Green
    return ,$rows
}

function Get-GcCopilotActivity {
    <#  Section 5: Copilot / AI summary behavior. Pulls recent conversation ids
        via analytics detail query, then fetches each conversation's summaries
        and records feedback + edit signals where the payload exposes them.
        Capped by -MaxConversationsForCopilot (one API call per conversation).  #>
    param([string]$Interval)

    Write-Host ('Querying Copilot / AI summaries (capped at ' + $MaxConversationsForCopilot + ' conversations)...') -ForegroundColor Cyan
    $rows = @()

    # 1) Recent conversation ids from analytics details.
    $convIds = @()
    $page = 1
    while ($convIds.Count -lt $MaxConversationsForCopilot) {
        $body = @{
            interval = $Interval
            order    = 'desc'
            orderBy  = 'conversationStart'
            paging   = @{ pageSize = 100; pageNumber = $page }
        }
        $resp = Invoke-GcApi -Method 'POST' -Path '/api/v2/analytics/conversations/details/query' -Body $body -SoftFail
        if ($null -eq $resp -or $null -eq $resp.conversations -or $resp.conversations.Count -eq 0) { break }
        foreach ($c in $resp.conversations) {
            if ($convIds.Count -ge $MaxConversationsForCopilot) { break }
            $convIds += [string]$c.conversationId
        }
        if ($resp.conversations.Count -lt 100) { break }
        $page += 1
    }
    Write-Host ('  Scanning ' + $convIds.Count + ' conversations for summaries...') -ForegroundColor DarkCyan

    # 2) Summaries per conversation.
    $notFoundCount = 0
    foreach ($cid in $convIds) {
        $path = $script:CopilotSummariesPathTemplate.Replace('{0}', $cid)
        $sumResp = Invoke-GcApi -Method 'GET' -Path $path -SoftFail
        if ($null -eq $sumResp) { $notFoundCount += 1; continue }

        $summaries = @()
        if ($null -ne $sumResp.entities) { $summaries = $sumResp.entities }
        elseif ($null -ne $sumResp.summaries) { $summaries = $sumResp.summaries }
        elseif ($null -ne $sumResp.id) { $summaries = @($sumResp) }

        foreach ($s in $summaries) {
            $agentId = ''
            if ($null -ne $s.agent -and $null -ne $s.agent.id) { $agentId = [string]$s.agent.id }
            elseif ($null -ne $s.userId) { $agentId = [string]$s.userId }
            elseif ($null -ne $s.modifiedBy -and $null -ne $s.modifiedBy.id) { $agentId = [string]$s.modifiedBy.id }

            $summaryText = ''
            if ($null -ne $s.summary) {
                if ($null -ne $s.summary.text) { $summaryText = [string]$s.summary.text }
                else { $summaryText = [string]$s.summary }
            } elseif ($null -ne $s.text) { $summaryText = [string]$s.text }

            $feedbackValue = ''
            if ($null -ne $s.feedback) {
                if ($null -ne $s.feedback.rating) { $feedbackValue = [string]$s.feedback.rating }
                else { $feedbackValue = [string]$s.feedback }
            }

            # Edit signal: modified date differing from created date implies the
            # agent touched the AI-generated note.
            $wasEdited = 'Unknown'
            $created = ''; $modified = ''
            if ($null -ne $s.dateCreated)  { $created  = [string]$s.dateCreated }
            if ($null -ne $s.dateModified) { $modified = [string]$s.dateModified }
            if ($created -ne '' -and $modified -ne '') {
                if ($created -eq $modified) { $wasEdited = 'No' } else { $wasEdited = 'Yes' }
            }

            $rows += New-Object PSObject -Property @{
                ConversationId = $cid
                SummaryId      = [string]$s.id
                UserId         = $agentId
                FeedbackGiven  = $feedbackValue
                WasEdited      = $wasEdited
                DateCreated    = $created
                DateModified   = $modified
                SummarySnippet = $summaryText
            }
        }
    }

    if ($notFoundCount -eq $convIds.Count -and $convIds.Count -gt 0) {
        Write-Warning ('  No summaries returned for any conversation. Either Copilot summarization is not enabled, or your org uses a different endpoint - verify "' + $script:CopilotSummariesPathTemplate + '" in the API Explorer and adjust the template at the top of this script.')
    }
    Write-Host ('  Copilot summary rows: ' + $rows.Count) -ForegroundColor Green
    return ,$rows
}

function Set-GcCopilotNote {
    <#  Section 6 (action): edit an AI Copilot note / conversation summary.
        Tries PATCH first, then PUT, against the configurable edit template.  #>
    param(
        [Parameter(Mandatory = $true)][string]$TargetConversationId,
        [Parameter(Mandatory = $true)][string]$TargetSummaryId,
        [Parameter(Mandatory = $true)][string]$Text
    )

    $path = $script:CopilotEditPathTemplate.Replace('{0}', $TargetConversationId).Replace('{1}', $TargetSummaryId)
    $body = @{ summary = @{ text = $Text } }

    Write-Host ('Updating Copilot note ' + $TargetSummaryId + ' on conversation ' + $TargetConversationId + '...') -ForegroundColor Cyan

    $resp = Invoke-GcApi -Method 'PATCH' -Path $path -Body $body -SoftFail
    if ($null -eq $resp) {
        Write-Host '  PATCH failed or unsupported - retrying with PUT...' -ForegroundColor Yellow
        $resp = Invoke-GcApi -Method 'PUT' -Path $path -Body $body -SoftFail
    }

    if ($null -eq $resp) {
        Write-Warning ('Could not update the summary. Verify the edit endpoint ("' + $script:CopilotEditPathTemplate + '") in the API Explorer and that the OAuth client has summary edit permission.')
        return $false
    }

    Write-Host '  Copilot note updated successfully.' -ForegroundColor Green
    return $true
}

# endregion

# ===========================================================================
# region Main
# ===========================================================================

Connect-GcCloud

# --- Action mode: edit a Copilot note, then exit -------------------------
if ($EditCopilotNote) {
    if ($ConversationId -eq '' -or $SummaryId -eq '' -or $NewNoteText -eq '') {
        throw 'EditCopilotNote requires -ConversationId, -SummaryId and -NewNoteText.'
    }
    $ok = Set-GcCopilotNote -TargetConversationId $ConversationId -TargetSummaryId $SummaryId -Text $NewNoteText
    if ($ok) { exit 0 } else { exit 1 }
}

# --- Report mode ----------------------------------------------------------
if ($OutputDir -eq '') {
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $OutputDir = '.\GcAgentBehavior_' + $stamp
}
if (-not (Test-Path -Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
}
Write-Host ('Output folder: ' + $OutputDir) -ForegroundColor Cyan

$interval = Get-GcInterval
Write-Host ('Reporting interval (UTC): ' + $interval) -ForegroundColor Cyan

$users = Get-GcUsers
if ($users.Count -eq 0) { throw 'No users matched - check -AgentEmailFilter and the directory:user:view permission.' }

# Fast lookup: userId -> user object
$userMap = @{}
foreach ($u in $users) { $userMap[$u.UserId] = $u }

function Get-GcUserName { param([string]$Id)
    if ($userMap.ContainsKey($Id)) { return $userMap[$Id].Name } else { return '' } }
function Get-GcUserEmail { param([string]$Id)
    if ($userMap.ContainsKey($Id)) { return $userMap[$Id].Email } else { return '' } }

# ---- Collect ----
$callRows    = Get-GcCallActivity -Users $users -Interval $interval
$statusRows  = Get-GcStatusTime  -Users $users -Interval $interval
$wfmRows     = @()
if (-not $SkipWfm) { $wfmRows = Get-GcWfmAdherence -Users $users }
$kbRows      = @()
if (-not $SkipKb) { $kbRows = Get-GcKbFeedback }
$copilotRows = @()
if (-not $SkipCopilot) { $copilotRows = Get-GcCopilotActivity -Interval $interval }

# ---- Export raw sections (explicit column order via Select-Object) ----
Write-Host 'Exporting CSVs...' -ForegroundColor Cyan

$users |
    Select-Object UserId, Name, Email, Department, Title |
    Export-Csv -Path (Join-Path $OutputDir 'Agents.csv') -NoTypeInformation

$callExport = @()
foreach ($r in $callRows) {
    $callExport += New-Object PSObject -Property @{
        AgentName = Get-GcUserName -Id $r.UserId; AgentEmail = Get-GcUserEmail -Id $r.UserId
        UserId = $r.UserId; Offered = $r.Offered; Answered = $r.Answered
        TalkHours = $r.TalkHours; HoldHours = $r.HoldHours; AcwHours = $r.AcwHours
        HandleHours = $r.HandleHours; AvgHandleSeconds = $r.AvgHandleSeconds
        NotRespondingCnt = $r.NotRespondingCnt
    }
}
$callExport |
    Select-Object AgentName, AgentEmail, UserId, Offered, Answered, TalkHours, HoldHours, AcwHours, HandleHours, AvgHandleSeconds, NotRespondingCnt |
    Sort-Object AgentName |
    Export-Csv -Path (Join-Path $OutputDir 'CallActivity.csv') -NoTypeInformation

$statusExport = @()
foreach ($r in $statusRows) {
    $statusExport += New-Object PSObject -Property @{
        AgentName = Get-GcUserName -Id $r.UserId; AgentEmail = Get-GcUserEmail -Id $r.UserId
        UserId = $r.UserId; MetricType = $r.MetricType; Status = $r.Status; Hours = $r.Hours
    }
}
$statusExport |
    Select-Object AgentName, AgentEmail, UserId, MetricType, Status, Hours |
    Sort-Object AgentName, MetricType, Status |
    Export-Csv -Path (Join-Path $OutputDir 'StatusTime.csv') -NoTypeInformation

if ($wfmRows.Count -gt 0) {
    $wfmExport = @()
    foreach ($r in $wfmRows) {
        $wfmExport += New-Object PSObject -Property @{
            AgentName = Get-GcUserName -Id $r.UserId; AgentEmail = Get-GcUserEmail -Id $r.UserId
            UserId = $r.UserId; AdherenceState = $r.AdherenceState
            ScheduledActivityCategory = $r.ScheduledActivityCategory
            ActualActivityCategory = $r.ActualActivityCategory
            Impact = $r.Impact; IsOutOfOffice = $r.IsOutOfOffice
            TimeOfAdherenceChange = $r.TimeOfAdherenceChange
        }
    }
    $wfmExport |
        Select-Object AgentName, AgentEmail, UserId, AdherenceState, ScheduledActivityCategory, ActualActivityCategory, Impact, IsOutOfOffice, TimeOfAdherenceChange |
        Sort-Object AgentName |
        Export-Csv -Path (Join-Path $OutputDir 'WfmAdherence.csv') -NoTypeInformation
}

if ($kbRows.Count -gt 0) {
    $kbExport = @()
    foreach ($r in $kbRows) {
        $kbExport += New-Object PSObject -Property @{
            AgentName = Get-GcUserName -Id $r.UserId; AgentEmail = Get-GcUserEmail -Id $r.UserId
            UserId = $r.UserId; KnowledgeBase = $r.KnowledgeBase; DocumentTitle = $r.DocumentTitle
            DocumentId = $r.DocumentId; Rating = $r.Rating; Reason = $r.Reason
            Comment = $r.Comment; DateCreated = $r.DateCreated
        }
    }
    $kbExport |
        Select-Object AgentName, AgentEmail, UserId, KnowledgeBase, DocumentTitle, DocumentId, Rating, Reason, Comment, DateCreated |
        Sort-Object AgentName, DateCreated |
        Export-Csv -Path (Join-Path $OutputDir 'KbFeedback.csv') -NoTypeInformation
}

if ($copilotRows.Count -gt 0) {
    $cpExport = @()
    foreach ($r in $copilotRows) {
        $cpExport += New-Object PSObject -Property @{
            AgentName = Get-GcUserName -Id $r.UserId; AgentEmail = Get-GcUserEmail -Id $r.UserId
            UserId = $r.UserId; ConversationId = $r.ConversationId; SummaryId = $r.SummaryId
            FeedbackGiven = $r.FeedbackGiven; WasEdited = $r.WasEdited
            DateCreated = $r.DateCreated; DateModified = $r.DateModified
            SummarySnippet = $r.SummarySnippet
        }
    }
    $cpExport |
        Select-Object AgentName, AgentEmail, UserId, ConversationId, SummaryId, FeedbackGiven, WasEdited, DateCreated, DateModified, SummarySnippet |
        Sort-Object AgentName, DateCreated |
        Export-Csv -Path (Join-Path $OutputDir 'CopilotSummaries.csv') -NoTypeInformation
}

# ---- Per-agent roll-up ----
Write-Host 'Building per-agent behavior roll-up...' -ForegroundColor Cyan
$summaryRows = @()
foreach ($u in $users) {
    $uid = $u.UserId

    $offered = 0; $answered = 0; $handleHours = '0.00'
    foreach ($r in $callRows) {
        if ($r.UserId -eq $uid) { $offered = $r.Offered; $answered = $r.Answered; $handleHours = $r.HandleHours }
    }

    $availableHours = 0; $awayHours = 0; $onQueueMs = 0
    foreach ($r in $statusRows) {
        if ($r.UserId -ne $uid) { continue }
        if ($r.MetricType -eq 'SystemPresence') {
            if ($r.Status -like '*AVAILABLE*') { $availableHours += [double]$r.Hours }
            if ($r.Status -like '*AWAY*' -or $r.Status -like '*BREAK*') { $awayHours += [double]$r.Hours }
            if ($r.Status -like '*ON_QUEUE*') { $onQueueMs += [long]$r.RawMs }
        }
    }

    $adherence = ''
    foreach ($r in $wfmRows) { if ($r.UserId -eq $uid) { $adherence = $r.AdherenceState } }

    $kbCount = 0
    foreach ($r in $kbRows) { if ($r.UserId -eq $uid) { $kbCount += 1 } }

    $cpSummaries = 0; $cpFeedback = 0; $cpEdited = 0
    foreach ($r in $copilotRows) {
        if ($r.UserId -ne $uid) { continue }
        $cpSummaries += 1
        if ($r.FeedbackGiven -ne '') { $cpFeedback += 1 }
        if ($r.WasEdited -eq 'Yes') { $cpEdited += 1 }
    }

    $summaryRows += New-Object PSObject -Property @{
        AgentName            = $u.Name
        AgentEmail           = $u.Email
        UserId               = $uid
        Offered              = $offered
        Answered             = $answered
        HandleHours          = $handleHours
        OnQueueHours         = Format-GcHours -Milliseconds $onQueueMs
        AvailableHours       = ('{0:N2}' -f $availableHours)
        AwayBreakHours       = ('{0:N2}' -f $awayHours)
        WfmAdherenceState    = $adherence
        KbFeedbackCount      = $kbCount
        CopilotSummariesSeen = $cpSummaries
        CopilotFeedbackCount = $cpFeedback
        CopilotNotesEdited   = $cpEdited
    }
}

$summaryRows |
    Select-Object AgentName, AgentEmail, UserId, Offered, Answered, HandleHours, OnQueueHours, AvailableHours, AwayBreakHours, WfmAdherenceState, KbFeedbackCount, CopilotSummariesSeen, CopilotFeedbackCount, CopilotNotesEdited |
    Sort-Object AgentName |
    Export-Csv -Path (Join-Path $OutputDir 'AgentBehaviorSummary.csv') -NoTypeInformation

Write-Host ''
Write-Host ('DONE. Reports written to ' + $OutputDir) -ForegroundColor Green
Write-Host '  Agents.csv               - directory of matched agents'
Write-Host '  CallActivity.csv         - offered/answered, talk/hold/ACW/handle time'
Write-Host '  StatusTime.csv           - hours per presence + routing status'
Write-Host '  WfmAdherence.csv         - real-time WFM adherence state'
Write-Host '  KbFeedback.csv           - individual KB feedback submissions'
Write-Host '  CopilotSummaries.csv     - AI summaries, feedback + edit signals'
Write-Host '  AgentBehaviorSummary.csv - one row per agent, everything joined'

# endregion
