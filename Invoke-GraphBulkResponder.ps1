#Requires -Version 5.1
<#
.SYNOPSIS
    Bulk reply to Microsoft 365 mailbox messages using the Microsoft Graph API.

.DESCRIPTION
    Authenticates to Microsoft 365 via OAuth 2.0 device code flow — no EWS DLL
    required. The script displays a short code and a URL; the user opens the URL
    in any browser, enters the code, and signs in. Once authenticated, the script
    searches a specified mailbox folder for messages matching a subject string and
    date range, then sends replies in configurable batches.

    Every run appends to a persistent log file (bulk-responder.log) and creates a
    per-run CSV tracker (bulk-responder-YYYYmmddHHmm.csv) recording each message's
    submission status. Use -ResumeFromCSV to continue a session that was paused or
    interrupted.

.PARAMETER TenantId
    Azure Active Directory tenant ID (GUID), or 'common' for multi-tenant apps.
    Found at: Azure Portal > Azure Active Directory > Overview > Tenant ID.

.PARAMETER ClientId
    Application (client) ID of the Azure AD app registration that has
    Mail.Read and Mail.Send delegated permissions with device code flow enabled.
    Found at: Azure Portal > App registrations > your app > Overview.

.PARAMETER MailboxEmail
    SMTP address of the mailbox to access. The signed-in account must have
    access to this mailbox. If omitted you will be prompted interactively.

.PARAMETER FolderPath
    Mailbox folder to search. Separate subfolder levels with a backslash.
    Well-known names (Inbox, Drafts, "Sent Items", "Deleted Items", "Junk Email",
    Archive) are resolved automatically; any other name is searched under the
    mailbox root. Example: "Inbox" or "Inbox\Projects\2024"
    Required for a new search. Ignored when -ResumeFromCSV is used.

.PARAMETER Subject
    Text to match against message subjects (case-insensitive, substring match).
    Required for a new search. Ignored when -ResumeFromCSV is used.

.PARAMETER StartDate
    Earliest received date to include (inclusive).
    Required for a new search. Ignored when -ResumeFromCSV is used.

.PARAMETER EndDate
    Latest received date to include (inclusive, covers through end of day UTC).
    Required for a new search. Ignored when -ResumeFromCSV is used.

.PARAMETER ReplyBody
    Plain-text reply body supplied on the command line.
    Mutually exclusive with -ReplyHtmlFile.

.PARAMETER ReplyHtmlFile
    Path to an HTML file whose contents become the reply body (sent as HTML).
    Mutually exclusive with -ReplyBody.

.PARAMETER EstimateOnly
    Count matching messages and display a summary. No reply is sent and no CSV is
    created. Works in resume mode too (shows pending count from the tracker CSV).

.PARAMETER ReplyAll
    Use Reply All instead of Reply when sending responses.

.PARAMETER BatchSize
    Number of replies to send per batch before pausing. When 0 (default) the script
    prompts interactively. Enter 0 at the prompt to send all messages without pausing.

.PARAMETER ResumeFromCSV
    Path to a tracker CSV produced by a previous run. The script re-authenticates,
    skips messages already marked ResponseSubmitted = True, and resumes sending the
    remaining messages. All search parameters are ignored.

.PARAMETER USGovDoD
    Target the Microsoft 365 GCC High / DoD sovereign cloud endpoints instead of
    the commercial cloud:
      Login : https://login.microsoftonline.us
      Graph : https://dod-graph.microsoft.us/v1.0
    The app registration and mailbox must both exist in the GCC High / DoD tenant.
    Combine with -TenantId set to the DoD tenant GUID.

.EXAMPLE
    # Preview match count only
    .\Invoke-GraphBulkResponder.ps1 `
        -TenantId     "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -ClientId     "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
        -MailboxEmail "you@contoso.com" `
        -FolderPath   "Inbox" `
        -Subject      "Project Alpha Status" `
        -StartDate    "2024-03-01" `
        -EndDate      "2024-03-31" `
        -EstimateOnly

.EXAMPLE
    # Send plain-text replies in batches of 10 (body entered interactively)
    .\Invoke-GraphBulkResponder.ps1 `
        -TenantId     "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -ClientId     "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
        -MailboxEmail "support@contoso.com" `
        -FolderPath   "Inbox\Client Requests" `
        -Subject      "Support Ticket" `
        -StartDate    "2024-04-01" `
        -EndDate      "2024-04-30" `
        -BatchSize    10

.EXAMPLE
    # Send HTML replies from a file, reply all
    .\Invoke-GraphBulkResponder.ps1 `
        -TenantId      "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -ClientId      "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
        -MailboxEmail  "billing@contoso.com" `
        -FolderPath    "Inbox" `
        -Subject       "Invoice" `
        -StartDate     "2024-01-01" `
        -EndDate       "2024-01-31" `
        -ReplyHtmlFile ".\auto-reply.html" `
        -ReplyAll

.EXAMPLE
    # Resume a previous session from its saved tracker CSV
    .\Invoke-GraphBulkResponder.ps1 `
        -TenantId      "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -ClientId      "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
        -MailboxEmail  "you@contoso.com" `
        -ResumeFromCSV ".\bulk-responder-202403151030.csv" `
        -ReplyHtmlFile ".\auto-reply.html" `
        -BatchSize     10
#>

[CmdletBinding(SupportsShouldProcess)]
param (
    # --- Authentication ---
    [string]$TenantId,
    [string]$ClientId,
    [string]$MailboxEmail,

    # --- New-search parameters (required unless -ResumeFromCSV is provided) ---
    [string]$FolderPath,
    [string]$Subject,
    [datetime]$StartDate,
    [datetime]$EndDate,

    # --- Reply body (one or neither; interactive prompt if both omitted) ---
    [string]$ReplyBody,
    [string]$ReplyHtmlFile,

    # --- Behaviour ---
    [switch]$EstimateOnly,
    [switch]$ReplyAll,
    [int]$BatchSize = 0,
    [string]$ResumeFromCSV,

    # --- Sovereign cloud ---
    [switch]$USGovDoD
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Script-level state
# ---------------------------------------------------------------------------

$script:RunTimestamp = Get-Date -Format 'yyyyMMddHHmm'
$script:ScriptDir    = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }
$script:LogFilePath  = Join-Path $script:ScriptDir 'bulk-responder.log'

$script:CsvFilePath = if ($ResumeFromCSV) {
    (Resolve-Path -Path $ResumeFromCSV -ErrorAction Stop).ProviderPath
} else {
    Join-Path $script:ScriptDir "bulk-responder-$($script:RunTimestamp).csv"
}

$script:CsvData = [System.Collections.Generic.List[psobject]]::new()

# OAuth token state
$script:AccessToken  = [string]::Empty
$script:RefreshToken = [string]::Empty
$script:TokenExpiry  = [datetime]::MinValue

# Endpoint constants – switched to DoD sovereign cloud when -USGovDoD is set
if ($USGovDoD) {
    $script:LoginBase  = 'https://login.microsoftonline.us'
    $script:GraphBase  = 'https://dod-graph.microsoft.us/v1.0'
    $script:OAuthScope = 'https://dod-graph.microsoft.us/Mail.Read https://dod-graph.microsoft.us/Mail.Send offline_access'
} else {
    $script:LoginBase  = 'https://login.microsoftonline.com'
    $script:GraphBase  = 'https://graph.microsoft.com/v1.0'
    $script:OAuthScope = 'https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Mail.Send offline_access'
}

# ---------------------------------------------------------------------------
# Helper – append a timestamped entry to the persistent log file
# ---------------------------------------------------------------------------

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')][string]$Level = 'INFO'
    )
    $ts   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$ts] [$Level] [GRAPH] $Message"
    Add-Content -Path $script:LogFilePath -Value $line -Encoding UTF8
}

# ---------------------------------------------------------------------------
# Helper – flush in-memory CSV state to disk
# ---------------------------------------------------------------------------

function Save-CsvTracker {
    $script:CsvData | Export-Csv -Path $script:CsvFilePath -NoTypeInformation -Encoding UTF8
}

# ---------------------------------------------------------------------------
# Helper – update a single CSV row by MessageId and flush
# ---------------------------------------------------------------------------

function Update-CsvRow {
    param(
        [string]$MessageId,
        [string]$ResponseSubmitted = ''
    )
    $row = $script:CsvData | Where-Object { $_.MessageId -eq $MessageId } | Select-Object -First 1
    if ($row -and $ResponseSubmitted -ne '') { $row.ResponseSubmitted = $ResponseSubmitted }
    Save-CsvTracker
}

# ---------------------------------------------------------------------------
# Helper – build the resume command string shown when the user exits mid-batch
# ---------------------------------------------------------------------------

function Get-ResumeCommand {
    $bt   = [char]96
    $nl   = [Environment]::NewLine
    $cont = " $bt$nl"

    $parts = [System.Collections.Generic.List[string]]::new()
    $parts.Add('.\Invoke-GraphBulkResponder.ps1')
    if ($TenantId)     { $parts.Add("    -TenantId `"$TenantId`"") }
    if ($ClientId)     { $parts.Add("    -ClientId `"$ClientId`"") }
    if ($MailboxEmail) { $parts.Add("    -MailboxEmail `"$MailboxEmail`"") }
    $parts.Add("    -ResumeFromCSV `"$script:CsvFilePath`"")
    if ($ReplyHtmlFile) {
        $parts.Add("    -ReplyHtmlFile `"$ReplyHtmlFile`"")
    } elseif ($ReplyBody) {
        $parts.Add('    -ReplyBody "(paste your reply body here)"')
    }
    if ($BatchSize -gt 0)      { $parts.Add("    -BatchSize $BatchSize") }
    if ($ReplyAll.IsPresent)   { $parts.Add('    -ReplyAll') }
    if ($USGovDoD.IsPresent)   { $parts.Add('    -USGovDoD') }

    return ($parts -join $cont)
}

# ---------------------------------------------------------------------------
# Authentication – OAuth 2.0 device code flow
# ---------------------------------------------------------------------------

function Invoke-DeviceCodeAuth {
    $deviceCodeUri = "$script:LoginBase/$TenantId/oauth2/v2.0/devicecode"
    $tokenUri      = "$script:LoginBase/$TenantId/oauth2/v2.0/token"

    Write-Host ""
    Write-Host "Requesting device code ..." -ForegroundColor Cyan

    $dcResponse = Invoke-RestMethod -Method Post -Uri $deviceCodeUri `
        -Body @{ client_id = $ClientId; scope = $script:OAuthScope } `
        -ErrorAction Stop

    Write-Host ""
    Write-Host ("=" * 62) -ForegroundColor Yellow
    Write-Host " AUTHENTICATION REQUIRED" -ForegroundColor Yellow
    Write-Host ("=" * 62) -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  1. Open a browser and navigate to:" -ForegroundColor White
    Write-Host "     $($dcResponse.verification_uri)" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  2. Enter this code when prompted:" -ForegroundColor White
    Write-Host "     $($dcResponse.user_code)" -ForegroundColor Green
    Write-Host ""
    Write-Host "  Waiting for sign-in (expires in $($dcResponse.expires_in)s) ..." -ForegroundColor DarkGray
    Write-Host ("=" * 62) -ForegroundColor Yellow
    Write-Host ""

    $deadline = (Get-Date).AddSeconds($dcResponse.expires_in)
    $interval = [int]$dcResponse.interval

    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds $interval

        try {
            $tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenUri -Body @{
                grant_type  = 'urn:ietf:params:oauth:grant-type:device_code'
                client_id   = $ClientId
                device_code = $dcResponse.device_code
            } -ErrorAction Stop

            $script:AccessToken  = $tokenResponse.access_token
            $script:RefreshToken = $tokenResponse.refresh_token
            $script:TokenExpiry  = (Get-Date).AddSeconds([int]$tokenResponse.expires_in)

            Write-Host "Authentication successful." -ForegroundColor Green
            Write-Log "AUTH SUCCESS | TokenExpiry: $($script:TokenExpiry.ToString('yyyy-MM-dd HH:mm:ss'))"
            return

        } catch {
            $errBody = $null
            try { $errBody = $_.ErrorDetails.Message | ConvertFrom-Json } catch {}

            switch ($errBody.error) {
                'authorization_pending' { continue }
                'authorization_declined' { throw "Authentication was declined by the user." }
                'expired_token'          { throw "Device code expired. Please run the script again." }
                'bad_verification_code'  { throw "Invalid device code. Please run the script again." }
                default                  { throw "Authentication error: $($_.Exception.Message)" }
            }
        }
    }

    throw "Authentication timed out. Please run the script again."
}

# ---------------------------------------------------------------------------
# Token management – silently refresh when the access token is near expiry
# ---------------------------------------------------------------------------

function Get-ValidToken {
    if ((Get-Date) -ge $script:TokenExpiry.AddMinutes(-5)) {
        if ($script:RefreshToken) {
            try {
                $tokenResponse = Invoke-RestMethod -Method Post `
                    -Uri "$script:LoginBase/$TenantId/oauth2/v2.0/token" `
                    -Body @{
                        grant_type    = 'refresh_token'
                        client_id     = $ClientId
                        refresh_token = $script:RefreshToken
                        scope         = $script:OAuthScope
                    } -ErrorAction Stop

                $script:AccessToken  = $tokenResponse.access_token
                $script:RefreshToken = if ($tokenResponse.refresh_token) { $tokenResponse.refresh_token } else { $script:RefreshToken }
                $script:TokenExpiry  = (Get-Date).AddSeconds([int]$tokenResponse.expires_in)
                Write-Log "TOKEN REFRESHED | Expiry: $($script:TokenExpiry.ToString('yyyy-MM-dd HH:mm:ss'))"

            } catch {
                Write-Log "TOKEN REFRESH FAILED — re-authenticating | Error: $_" -Level WARN
                Invoke-DeviceCodeAuth
            }
        } else {
            Invoke-DeviceCodeAuth
        }
    }
    return $script:AccessToken
}

# ---------------------------------------------------------------------------
# Helper – make an authenticated Microsoft Graph REST call
# ---------------------------------------------------------------------------

function Invoke-GraphRequest {
    param(
        [string]$Method = 'GET',
        [Parameter(Mandatory)][string]$Uri,
        [hashtable]$ExtraHeaders = @{},
        [object]$Body = $null
    )

    $token   = Get-ValidToken
    $headers = @{ Authorization = "Bearer $token" }
    foreach ($k in $ExtraHeaders.Keys) { $headers[$k] = $ExtraHeaders[$k] }

    $params = @{
        Method      = $Method
        Uri         = $Uri
        Headers     = $headers
        ErrorAction = 'Stop'
    }

    if ($Body) {
        $params['Body']        = ($Body | ConvertTo-Json -Depth 10 -Compress)
        $params['ContentType'] = 'application/json'
    }

    return Invoke-RestMethod @params
}

# ---------------------------------------------------------------------------
# Helper – resolve a backslash-delimited folder path to a Graph mailFolder object
# ---------------------------------------------------------------------------

function Get-GraphFolder {
    param([string]$Email, [string]$Path)

    $wellKnown = @{
        'Inbox'         = 'inbox'
        'Drafts'        = 'drafts'
        'SentItems'     = 'sentitems'
        'Sent Items'    = 'sentitems'
        'DeletedItems'  = 'deleteditems'
        'Deleted Items' = 'deleteditems'
        'JunkEmail'     = 'junkemail'
        'Junk Email'    = 'junkemail'
        'Archive'       = 'archive'
    }

    $segments = $Path -split '\\'
    $rootName = $segments[0]

    if ($wellKnown.ContainsKey($rootName)) {
        $folder = Invoke-GraphRequest -Uri "$script:GraphBase/users/$Email/mailFolders/$($wellKnown[$rootName])"
    } else {
        $encodedFilter = [Uri]::EscapeDataString("displayName eq '$($rootName -replace "'", "''")'")
        $result = Invoke-GraphRequest -Uri "$script:GraphBase/users/$Email/mailFolders?`$filter=$encodedFilter"
        if ($result.value.Count -eq 0) {
            throw "Root folder '$rootName' not found in mailbox."
        }
        $folder = $result.value[0]
    }

    foreach ($seg in $segments[1..($segments.Count - 1)]) {
        $folderId      = $folder.id
        $encodedFilter = [Uri]::EscapeDataString("displayName eq '$($seg -replace "'", "''")'")
        $result = Invoke-GraphRequest `
            -Uri "$script:GraphBase/users/$Email/mailFolders/$folderId/childFolders?`$filter=$encodedFilter"
        if ($result.value.Count -eq 0) {
            throw "Subfolder '$seg' not found under '$($folder.displayName)'."
        }
        $folder = $result.value[0]
    }

    return $folder
}

# ---------------------------------------------------------------------------
# Helper – search a folder for matching messages (paginated via nextLink)
# ---------------------------------------------------------------------------

function Find-GraphMessages {
    param(
        [string]$Email,
        [string]$FolderId,
        [string]$SubjectText,
        [datetime]$RangeStart,
        [datetime]$RangeEnd
    )

    $escapedSubject = $SubjectText -replace "'", "''"
    $startStr = $RangeStart.Date.ToUniversalTime().ToString("yyyy-MM-dd'T'00:00:00'Z'")
    $endStr   = $RangeEnd.Date.AddDays(1).ToUniversalTime().ToString("yyyy-MM-dd'T'00:00:00'Z'")

    $filter = "contains(subject, '$escapedSubject') and receivedDateTime ge $startStr and receivedDateTime lt $endStr"
    $select = 'id,subject,from,receivedDateTime'

    $encodedFilter = [Uri]::EscapeDataString($filter)
    $uri = "$script:GraphBase/users/$Email/mailFolders/$FolderId/messages" +
           "?`$filter=$encodedFilter&`$select=$select&`$top=100&`$count=true"

    $allMessages = [System.Collections.Generic.List[psobject]]::new()

    do {
        $result = Invoke-GraphRequest -Uri $uri -ExtraHeaders @{ 'ConsistencyLevel' = 'eventual' }
        foreach ($msg in $result.value) { $allMessages.Add($msg) }
        $uri = $result.'@odata.nextLink'
    } while ($uri)

    return $allMessages
}

# ===========================================================================
# MAIN
# ===========================================================================

# ---------------------------------------------------------------------------
# 1. Collect required connection parameters
# ---------------------------------------------------------------------------

if (-not $TenantId)     { $TenantId     = Read-Host 'Enter Azure AD Tenant ID' }
if (-not $ClientId)     { $ClientId     = Read-Host 'Enter App Registration Client ID' }
if (-not $MailboxEmail) { $MailboxEmail = Read-Host 'Enter mailbox email address' }

# ---------------------------------------------------------------------------
# 2. Authenticate via device code flow
# ---------------------------------------------------------------------------

Invoke-DeviceCodeAuth

# ---------------------------------------------------------------------------
# 3. Log script start
# ---------------------------------------------------------------------------

$runMode  = if ($ResumeFromCSV) { 'RESUME' } else { 'NEW SEARCH' }
$cloudEnv = if ($USGovDoD) { 'USGovDoD' } else { 'Commercial' }
Write-Log ("SCRIPT START | Mailbox: $MailboxEmail | Mode: $runMode | Cloud: $cloudEnv | TenantId: $TenantId")

# ---------------------------------------------------------------------------
# 4. Collect pending message IDs and display metadata
# ---------------------------------------------------------------------------

$pendingIds   = [System.Collections.Generic.List[string]]::new()
$displayItems = [System.Collections.Generic.List[psobject]]::new()

if ($ResumeFromCSV) {
    # -----------------------------------------------------------------------
    # Resume mode – load tracker CSV, skip already-submitted rows
    # -----------------------------------------------------------------------

    if (-not (Test-Path $ResumeFromCSV)) {
        Write-Error "CSV tracker file not found: $ResumeFromCSV"
        exit 1
    }

    $loadedRows  = @(Import-Csv -Path $script:CsvFilePath -Encoding UTF8)
    foreach ($row in $loadedRows) { $script:CsvData.Add($row) }

    $pendingRows = @($loadedRows | Where-Object { $_.ResponseSubmitted -ne 'True' })
    $doneCount   = $loadedRows.Count - $pendingRows.Count

    Write-Host ""
    Write-Host ("Loaded tracker: $($loadedRows.Count) total | $doneCount already submitted | $($pendingRows.Count) pending") `
        -ForegroundColor Cyan
    Write-Log ("RESUME | CSV: $script:CsvFilePath | Total: $($loadedRows.Count) | " +
        "AlreadySubmitted: $doneCount | Pending: $($pendingRows.Count)")

    foreach ($row in $pendingRows) {
        $pendingIds.Add($row.MessageId)
        $displayItems.Add([PSCustomObject]@{
            Subject     = $row.Subject
            FromAddress = $row.FromAddress
            Received    = ''
        })
    }

    if ($EstimateOnly) {
        Write-Host ""
        Write-Host ("=" * 50) -ForegroundColor Yellow
        Write-Host " RESUME ESTIMATE" -ForegroundColor Yellow
        Write-Host ("=" * 50) -ForegroundColor Yellow
        Write-Host "  Pending replies   : $($pendingIds.Count)" -ForegroundColor Green
        Write-Host "  Already submitted : $doneCount"
        Write-Host "  Tracker CSV       : $script:CsvFilePath"
        Write-Host ("=" * 50) -ForegroundColor Yellow
        Write-Host ""
        exit 0
    }

} else {
    # -----------------------------------------------------------------------
    # New search mode – validate params, search, create CSV tracker
    # -----------------------------------------------------------------------

    $missing = [System.Collections.Generic.List[string]]::new()
    if ([string]::IsNullOrWhiteSpace($FolderPath))         { $missing.Add('-FolderPath') }
    if ([string]::IsNullOrWhiteSpace($Subject))            { $missing.Add('-Subject') }
    if (-not $PSBoundParameters.ContainsKey('StartDate'))  { $missing.Add('-StartDate') }
    if (-not $PSBoundParameters.ContainsKey('EndDate'))    { $missing.Add('-EndDate') }
    if ($missing.Count -gt 0) {
        Write-Error ("Missing required parameters for a new search: $($missing -join ', ')." +
            "`nTo resume an existing session use -ResumeFromCSV.")
        exit 1
    }

    Write-Host ""
    Write-Host "Resolving folder: $FolderPath" -ForegroundColor Cyan
    $targetFolder = Get-GraphFolder -Email $MailboxEmail -Path $FolderPath
    Write-Host "Folder resolved : $($targetFolder.displayName)" -ForegroundColor Green

    Write-Host "Searching for messages ..." -ForegroundColor Cyan
    $foundMessages = Find-GraphMessages `
        -Email       $MailboxEmail `
        -FolderId    $targetFolder.id `
        -SubjectText $Subject `
        -RangeStart  $StartDate `
        -RangeEnd    $EndDate

    $matchCount = @($foundMessages).Count

    Write-Log ("SEARCH | Folder: $FolderPath | Subject: `"$Subject`" | " +
        "DateRange: $($StartDate.ToString('yyyy-MM-dd')) to $($EndDate.ToString('yyyy-MM-dd')) | " +
        "Matches: $matchCount")

    if ($EstimateOnly) {
        Write-Host ""
        Write-Host ("=" * 50) -ForegroundColor Yellow
        Write-Host " ESTIMATE RESULTS" -ForegroundColor Yellow
        Write-Host ("=" * 50) -ForegroundColor Yellow
        Write-Host "  Folder      : $($targetFolder.displayName)"
        Write-Host "  Subject     : $Subject"
        Write-Host "  Date range  : $($StartDate.ToString('yyyy-MM-dd'))  to  $($EndDate.ToString('yyyy-MM-dd'))"
        Write-Host "  Matches     : $matchCount" -ForegroundColor Green
        Write-Host ("=" * 50) -ForegroundColor Yellow
        Write-Host ""
        exit 0
    }

    if ($matchCount -eq 0) {
        Write-Host "No messages matched the given criteria. Nothing to do." -ForegroundColor Yellow
        Write-Log "SEARCH | No matches found. Exiting."
        exit 0
    }

    foreach ($msg in $foundMessages) {
        $fromAddr = $msg.from.emailAddress.address
        $fromName = $msg.from.emailAddress.name
        $received = if ($msg.receivedDateTime) {
            [datetime]::Parse($msg.receivedDateTime).ToLocalTime().ToString('yyyy-MM-dd HH:mm')
        } else { '' }

        $script:CsvData.Add([PSCustomObject]@{
            MessageId         = $msg.id
            FromAddress       = $fromAddr
            Subject           = $msg.subject
            ResponseSubmitted = 'False'
        })

        $pendingIds.Add($msg.id)
        $displayItems.Add([PSCustomObject]@{
            Subject     = $msg.subject
            FromAddress = $fromAddr
            FromName    = $fromName
            Received    = $received
        })
    }

    Save-CsvTracker
    Write-Host "CSV tracker  : $($script:CsvFilePath)" -ForegroundColor Cyan
    Write-Log "CSV CREATED | Path: $script:CsvFilePath | Rows: $matchCount"
}

# ---------------------------------------------------------------------------
# 5. Display pending messages summary table
# ---------------------------------------------------------------------------

$totalPending = $pendingIds.Count

if ($totalPending -eq 0) {
    Write-Host "No pending messages. All messages in the tracker have been submitted." -ForegroundColor Green
    exit 0
}

Write-Host ""
Write-Host ("=" * 88) -ForegroundColor Yellow
Write-Host (" PENDING MESSAGES ($totalPending)") -ForegroundColor Yellow
Write-Host ("=" * 88) -ForegroundColor Yellow
Write-Host ("{0,-5}  {1,-36}  {2,-26}  {3}" -f "#", "Subject", "From", "Received") -ForegroundColor Cyan
Write-Host ("-" * 88) -ForegroundColor DarkGray

$idx = 1
foreach ($di in $displayItems) {
    $subj = $di.Subject
    if ($subj.Length -gt 33) { $subj = $subj.Substring(0, 30) + '...' }
    $from = if ($di.PSObject.Properties['FromName'] -and $di.FromName) { $di.FromName } else { $di.FromAddress }
    if ($from.Length -gt 23) { $from = $from.Substring(0, 20) + '...' }
    Write-Host ("{0,-5}  {1,-36}  {2,-26}  {3}" -f $idx, $subj, $from, $di.Received)
    $idx++
}

Write-Host ("-" * 88) -ForegroundColor DarkGray
Write-Host ""

# ---------------------------------------------------------------------------
# 6. Gather reply body
# ---------------------------------------------------------------------------

$replyIsHtml  = $false
$replyContent = ''

if ($ReplyHtmlFile) {
    $htmlPath = Resolve-Path -Path $ReplyHtmlFile -ErrorAction SilentlyContinue
    if (-not $htmlPath) { Write-Error "HTML file not found: $ReplyHtmlFile"; exit 1 }
    $replyContent = Get-Content -Path $htmlPath.ProviderPath -Raw -Encoding UTF8
    $replyIsHtml  = $true
    Write-Host "Reply body loaded from: $($htmlPath.ProviderPath)  [HTML]" -ForegroundColor Cyan
} elseif ($ReplyBody) {
    $replyContent = $ReplyBody
} else {
    Write-Host "Enter the reply message body (type END on a new line to finish):" -ForegroundColor Cyan
    Write-Host "  HTML is detected if your input begins with an HTML tag (e.g. <p>)." -ForegroundColor DarkGray
    Write-Host ""
    $lines = [System.Collections.Generic.List[string]]::new()
    do {
        $line = Read-Host '>'
        if ($line -ne 'END') { $lines.Add($line) }
    } while ($line -ne 'END')
    $replyContent = $lines -join "`r`n"
    if ($replyContent.TrimStart() -match '^<[a-zA-Z]') {
        $replyIsHtml = $true
        Write-Host "Detected HTML content." -ForegroundColor DarkGray
    }
}

if ([string]::IsNullOrWhiteSpace($replyContent)) {
    Write-Error "Reply body is empty. Aborting."
    exit 1
}

# ---------------------------------------------------------------------------
# 7. Determine batch size
# ---------------------------------------------------------------------------

$effectiveBatch = $BatchSize

if ($effectiveBatch -le 0) {
    Write-Host ""
    Write-Host "Enter batch size (messages to send before pausing)." -ForegroundColor Cyan
    Write-Host "  Enter 0 to send all $totalPending message(s) without pausing." -ForegroundColor DarkGray
    $bsRaw = Read-Host "Batch size"
    $effectiveBatch = [int]$bsRaw
}

if ($effectiveBatch -le 0) { $effectiveBatch = $totalPending }

Write-Host ""
Write-Host "Batch size : $effectiveBatch" -ForegroundColor DarkGray

# ---------------------------------------------------------------------------
# 8. Confirm before the first batch
# ---------------------------------------------------------------------------

$action = if ($ReplyAll) { 'REPLY ALL' } else { 'REPLY' }
Write-Host ("Ready to {0} {1} message(s) in batches of {2}." -f $action, $totalPending, $effectiveBatch) `
    -ForegroundColor Yellow

$confirm = Read-Host "Proceed? (yes / no)"
if ($confirm -notmatch '^(yes|y)$') {
    Write-Host "Cancelled. No replies were sent." -ForegroundColor Yellow
    Write-Log "USER CANCELLED | No replies sent."
    exit 0
}

# ---------------------------------------------------------------------------
# 9. Batch send loop
# ---------------------------------------------------------------------------

$contentType   = if ($replyIsHtml) { 'html' } else { 'text' }
$replyAction   = if ($ReplyAll) { 'replyAll' } else { 'reply' }

$allIds      = @($pendingIds)
$totalSent   = 0
$totalFailed = 0
$batchNum    = 0
$offset      = 0

while ($offset -lt $allIds.Count) {
    $batchNum++
    $batchEnd     = [Math]::Min($offset + $effectiveBatch - 1, $allIds.Count - 1)
    $currentBatch = $allIds[$offset..$batchEnd]

    Write-Host ""
    Write-Host ("--- Batch $batchNum | Messages $($offset + 1)-$($batchEnd + 1) of $totalPending ---") `
        -ForegroundColor Cyan
    Write-Log "BATCH START | Batch: $batchNum | Range: $($offset + 1)-$($batchEnd + 1) of $totalPending"

    $batchSent   = 0
    $batchFailed = 0

    foreach ($msgId in $currentBatch) {
        $csvRow      = $script:CsvData | Where-Object { $_.MessageId -eq $msgId } | Select-Object -First 1
        $displayAddr = if ($csvRow) { $csvRow.FromAddress } else { '(unknown)' }

        try {
            # Mark submitted before the network call so an interrupted run is visible in the CSV
            Update-CsvRow -MessageId $msgId -ResponseSubmitted 'True'
            Write-Log "REPLY SUBMITTED | MessageId: $msgId | Endpoint: Graph /$replyAction | From: $displayAddr"

            $uri  = "$script:GraphBase/users/$MailboxEmail/messages/$msgId/$replyAction"
            $body = @{
                message = @{
                    body = @{
                        contentType = $contentType
                        content     = $replyContent
                    }
                }
            }

            Invoke-GraphRequest -Method 'POST' -Uri $uri -Body $body

            Write-Log "REPLY SENT | MessageId: $msgId | From: $displayAddr" -Level SUCCESS
            $batchSent++
            $totalSent++
            Write-Host ("  [SENT]   -> {0}" -f $displayAddr) -ForegroundColor Green

        } catch {
            Write-Log "REPLY FAILED | MessageId: $msgId | From: $displayAddr | Error: $_" -Level ERROR
            $batchFailed++
            $totalFailed++
            Write-Host ("  [FAIL]   $displayAddr — $_") -ForegroundColor Red
        }
    }

    Write-Host ""
    Write-Host ("Batch $batchNum complete. Sent: $batchSent | Failed: $batchFailed") -ForegroundColor Cyan
    Write-Log "BATCH COMPLETE | Batch: $batchNum | Sent: $batchSent | Failed: $batchFailed"

    $offset += $effectiveBatch

    if ($offset -lt $allIds.Count) {
        $remaining = $allIds.Count - $offset
        Write-Host ""
        Write-Host ("$remaining message(s) still pending.") -ForegroundColor Yellow
        $choice = Read-Host "[C] Continue next batch   [E] Exit and save progress"

        if ($choice -notmatch '^[Cc]') {
            Write-Host ""
            Write-Host "Progress saved." -ForegroundColor Green
            Write-Host "  CSV tracker : $script:CsvFilePath"
            Write-Host "  Log file    : $script:LogFilePath"
            Write-Host ""
            Write-Host "To resume this session later, run:" -ForegroundColor Yellow
            Write-Host ""
            Write-Host (Get-ResumeCommand) -ForegroundColor Cyan
            Write-Host ""
            Write-Log ("USER EXIT | After batch $batchNum | TotalSent: $totalSent | " +
                "TotalFailed: $totalFailed | Remaining: $remaining | CSV: $script:CsvFilePath")
            exit 0
        }
    }
}

# ---------------------------------------------------------------------------
# 10. Final summary
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host ("=" * 55) -ForegroundColor Yellow
Write-Host " ALL BATCHES COMPLETE" -ForegroundColor Yellow
Write-Host ("=" * 55) -ForegroundColor Yellow
Write-Host ("  Total sent   : $totalSent") -ForegroundColor Green
if ($totalFailed -gt 0) {
    Write-Host ("  Total failed : $totalFailed") -ForegroundColor Red
} else {
    Write-Host ("  Total failed : $totalFailed")
}
Write-Host ("  CSV tracker  : $script:CsvFilePath")
Write-Host ("  Log file     : $script:LogFilePath")
Write-Host ("=" * 55) -ForegroundColor Yellow
Write-Host ""

Write-Log ("SCRIPT COMPLETE | TotalSent: $totalSent | TotalFailed: $totalFailed | " +
    "CSV: $script:CsvFilePath")
