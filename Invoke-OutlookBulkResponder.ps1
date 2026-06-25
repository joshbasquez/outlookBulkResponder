#Requires -Version 5.1
<#
.SYNOPSIS
    Bulk reply to Exchange mailbox messages matching a subject filter and date range via EWS.

.DESCRIPTION
    Loads the EWS Managed API, connects to an Exchange mailbox, and searches a specified
    folder for email messages whose subject contains the given string within the given date
    range. Replies are sent in configurable batches; after each batch the script pauses
    and lets you continue or save progress and exit for later.

    Every run appends to a persistent log file (bulk-responder.log) and creates a CSV
    tracker (bulk-responder-YYYYmmddHHmm.csv) that records each message's send status.

    Use -EstimateOnly to preview match counts without sending.
    Use -ResumeFromCSV to continue a previous run from its saved tracker CSV.

.PARAMETER FolderPath
    Mailbox folder to search. Separate subfolder levels with a backslash.
    Well-known names (Inbox, Drafts, "Sent Items", "Deleted Items", "Junk Email")
    are resolved automatically; any other name is searched under the mailbox root.
    Example: "Inbox" or "Inbox\Projects\2024"
    Required for a new search. Ignored when -ResumeFromCSV is used.

.PARAMETER Subject
    Text to match against message subjects (case-insensitive, partial match).
    Required for a new search. Ignored when -ResumeFromCSV is used.

.PARAMETER StartDate
    Earliest received date to include (inclusive).
    Required for a new search. Ignored when -ResumeFromCSV is used.

.PARAMETER EndDate
    Latest received date to include (inclusive, covers through end of day).
    Required for a new search. Ignored when -ResumeFromCSV is used.

.PARAMETER MailboxEmail
    SMTP address of the mailbox to open. Used for AutoDiscover and credential prompt.
    If omitted you will be prompted interactively.

.PARAMETER EwsDllPath
    Path to Microsoft.Exchange.WebServices.dll.
    Defaults to .\Microsoft.Exchange.WebServices.dll (same directory as the script).

.PARAMETER EwsUrl
    Override the EWS endpoint URL. When omitted, AutoDiscover is used.
    Example: https://mail.contoso.com/EWS/Exchange.asmx

.PARAMETER ReplyBody
    Plain-text reply body supplied on the command line.
    Mutually exclusive with -ReplyHtmlFile.

.PARAMETER ReplyHtmlFile
    Path to an HTML file whose contents become the reply body (sent as HTML).
    Mutually exclusive with -ReplyBody.

.PARAMETER EstimateOnly
    Count matching messages and display a summary. No reply is sent and no CSV is created.

.PARAMETER ReplyAll
    Use Reply All instead of Reply when sending responses.

.PARAMETER UseCurrentCredentials
    Authenticate with the Windows identity of the calling process (NTLM/Kerberos).
    Suitable for on-premises Exchange on a domain-joined machine.

.PARAMETER BatchSize
    Number of replies to send per batch before pausing. When 0 (default) the script
    prompts interactively. Pass 0 at the prompt to send all messages without pausing.

.PARAMETER ResumeFromCSV
    Path to a tracker CSV produced by a previous run. The script reconnects to Exchange,
    skips messages already marked ResponseSentSuccessfully = True, and resumes sending
    the remaining messages. All other search parameters are ignored.

.EXAMPLE
    # Preview match count only
    .\Invoke-OutlookBulkResponder.ps1 `
        -FolderPath "Inbox" `
        -Subject "Project Alpha Status" `
        -StartDate "2024-03-01" `
        -EndDate "2024-03-31" `
        -MailboxEmail "you@contoso.com" `
        -EstimateOnly

.EXAMPLE
    # Send replies in batches of 10 (body entered interactively)
    .\Invoke-OutlookBulkResponder.ps1 `
        -FolderPath "Inbox\Client Requests" `
        -Subject "Support Ticket" `
        -StartDate "2024-04-01" `
        -EndDate "2024-04-30" `
        -MailboxEmail "support@contoso.com" `
        -BatchSize 10

.EXAMPLE
    # Send HTML replies from a file, reply all, current Windows credentials
    .\Invoke-OutlookBulkResponder.ps1 `
        -FolderPath "Inbox" `
        -Subject "Invoice" `
        -StartDate "2024-01-01" `
        -EndDate "2024-01-31" `
        -ReplyHtmlFile ".\auto-reply.html" `
        -ReplyAll `
        -UseCurrentCredentials

.EXAMPLE
    # Resume a previous session from its saved tracker CSV
    .\Invoke-OutlookBulkResponder.ps1 `
        -ResumeFromCSV ".\bulk-responder-202403151030.csv" `
        -MailboxEmail "you@contoso.com" `
        -ReplyHtmlFile ".\auto-reply.html" `
        -BatchSize 10
#>

[CmdletBinding(SupportsShouldProcess)]
param (
    # --- New-search parameters (required unless -ResumeFromCSV is provided) ---
    [string]$FolderPath,
    [string]$Subject,
    [datetime]$StartDate,
    [datetime]$EndDate,

    # --- Connection parameters ---
    [string]$MailboxEmail,
    [string]$EwsDllPath = ".\Microsoft.Exchange.WebServices.dll",
    [string]$EwsUrl,

    # --- Reply body (one or neither; interactive prompt if both omitted) ---
    [string]$ReplyBody,
    [string]$ReplyHtmlFile,

    # --- Behaviour switches ---
    [switch]$EstimateOnly,
    [switch]$ReplyAll,
    [switch]$UseCurrentCredentials,

    # --- Batch and resume controls ---
    [int]$BatchSize = 0,
    [string]$ResumeFromCSV
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Script-level state
# ---------------------------------------------------------------------------

$script:RunTimestamp = Get-Date -Format 'yyyyMMddHHmm'
$script:ScriptDir    = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }
$script:LogFilePath  = Join-Path $script:ScriptDir 'bulk-responder.log'

$script:CsvFilePath  = if ($ResumeFromCSV) {
    (Resolve-Path -Path $ResumeFromCSV -ErrorAction Stop).ProviderPath
} else {
    Join-Path $script:ScriptDir "bulk-responder-$($script:RunTimestamp).csv"
}

# In-memory CSV rows (PSCustomObjects). Written to disk after every state change.
$script:CsvData = [System.Collections.Generic.List[psobject]]::new()

# ---------------------------------------------------------------------------
# Helper – append a timestamped entry to the persistent log file
# ---------------------------------------------------------------------------

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')][string]$Level = 'INFO'
    )
    $ts   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$ts] [$Level] $Message"
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
    $bt   = [char]96  # backtick – PowerShell line-continuation character
    $nl   = [Environment]::NewLine
    $cont = " $bt$nl"

    $parts = [System.Collections.Generic.List[string]]::new()
    $parts.Add(".\Invoke-OutlookBulkResponder.ps1")
    $parts.Add("    -ResumeFromCSV `"$script:CsvFilePath`"")

    if ($MailboxEmail)  { $parts.Add("    -MailboxEmail `"$MailboxEmail`"") }
    if ($EwsUrl)        { $parts.Add("    -EwsUrl `"$EwsUrl`"") }

    if ($EwsDllPath -and $EwsDllPath -ne '.\Microsoft.Exchange.WebServices.dll') {
        $parts.Add("    -EwsDllPath `"$EwsDllPath`"")
    }

    if ($ReplyHtmlFile) {
        $parts.Add("    -ReplyHtmlFile `"$ReplyHtmlFile`"")
    } elseif ($ReplyBody) {
        $parts.Add('    -ReplyBody "(paste your reply body here)"')
    }

    if ($BatchSize -gt 0)               { $parts.Add("    -BatchSize $BatchSize") }
    if ($ReplyAll.IsPresent)            { $parts.Add('    -ReplyAll') }
    if ($UseCurrentCredentials.IsPresent) { $parts.Add('    -UseCurrentCredentials') }

    return ($parts -join $cont)
}

# ---------------------------------------------------------------------------
# Helper – resolve a backslash-delimited folder path to an EWS Folder object
# ---------------------------------------------------------------------------

function Get-EwsFolder {
    param (
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service,
        [string]$Path
    )

    $wellKnown = @{
        'Inbox'            = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox
        'Drafts'           = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Drafts
        'SentItems'        = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems
        'Sent Items'       = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems
        'DeletedItems'     = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::DeletedItems
        'Deleted Items'    = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::DeletedItems
        'JunkEmail'        = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::JunkEmail
        'Junk Email'       = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::JunkEmail
        'Archive'          = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot
        'RecoverableItems' = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::RecoverableItemsRoot
    }

    $segments = $Path -split '\\'
    $rootName = $segments[0]

    if ($wellKnown.ContainsKey($rootName)) {
        $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, $wellKnown[$rootName])
    } else {
        $msgRoot = [Microsoft.Exchange.WebServices.Data.Folder]::Bind(
            $Service,
            [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot
        )
        $view   = New-Object Microsoft.Exchange.WebServices.Data.FolderView(20)
        $filter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo(
            [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $rootName
        )
        $found = $msgRoot.FindFolders($filter, $view)
        if ($found.Folders.Count -eq 0) {
            throw "Root folder '$rootName' not found in mailbox."
        }
        $folder = $found.Folders[0]
    }

    foreach ($seg in $segments[1..($segments.Count - 1)]) {
        $view   = New-Object Microsoft.Exchange.WebServices.Data.FolderView(20)
        $filter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo(
            [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $seg
        )
        $found = $folder.FindFolders($filter, $view)
        if ($found.Folders.Count -eq 0) {
            throw "Subfolder '$seg' not found under '$($folder.DisplayName)'."
        }
        $folder = $found.Folders[0]
    }

    return $folder
}

# ---------------------------------------------------------------------------
# Helper – search a folder for matching messages (paginated, lightweight props)
# ---------------------------------------------------------------------------

function Find-MatchingMessages {
    param (
        [Microsoft.Exchange.WebServices.Data.Folder]$Folder,
        [string]$SubjectText,
        [datetime]$RangeStart,
        [datetime]$RangeEnd
    )

    $clauses = [System.Collections.Generic.List[Microsoft.Exchange.WebServices.Data.SearchFilter]]::new()

    $clauses.Add((New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo(
        [Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass, 'IPM.Note'
    )))

    $clauses.Add((New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring(
        [Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject,
        $SubjectText,
        [Microsoft.Exchange.WebServices.Data.ContainmentMode]::Substring,
        [Microsoft.Exchange.WebServices.Data.ComparisonMode]::IgnoreCase
    )))

    $clauses.Add((New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThanOrEqualTo(
        [Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived,
        $RangeStart.Date
    )))

    # End date is inclusive through end of day
    $clauses.Add((New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThan(
        [Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived,
        $RangeEnd.Date.AddDays(1)
    )))

    $andFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection(
        [Microsoft.Exchange.WebServices.Data.LogicalOperator]::And,
        $clauses
    )

    $propSet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet(
        [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly
    )
    $propSet.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject)
    $propSet.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived)
    $propSet.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::From)
    $propSet.Add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Sender)

    $pageSize      = 100
    $offset        = 0
    $allItems      = [System.Collections.Generic.List[Microsoft.Exchange.WebServices.Data.Item]]::new()
    $moreAvailable = $true

    do {
        $view = New-Object Microsoft.Exchange.WebServices.Data.ItemView($pageSize, $offset)
        $view.PropertySet = $propSet

        $page = $Folder.FindItems($andFilter, $view)
        foreach ($item in $page.Items) { $allItems.Add($item) }

        $moreAvailable = $page.MoreAvailable
        $offset       += $pageSize
    } while ($moreAvailable)

    return $allItems
}

# ===========================================================================
# MAIN
# ===========================================================================

# ---------------------------------------------------------------------------
# 1. Load EWS Managed API DLL
# ---------------------------------------------------------------------------

$resolvedDll = Resolve-Path -Path $EwsDllPath -ErrorAction SilentlyContinue
if (-not $resolvedDll) {
    Write-Error ("EWS DLL not found at '{0}'.`nDownload Microsoft.Exchange.WebServices " +
        "and place the DLL in the script folder, or pass -EwsDllPath.") -f $EwsDllPath
    exit 1
}

try {
    Add-Type -Path $resolvedDll.ProviderPath
} catch {
    Write-Error "Failed to load EWS DLL: $_"
    exit 1
}

# ---------------------------------------------------------------------------
# 2. Create and configure ExchangeService
# ---------------------------------------------------------------------------

$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService(
    [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
)

if ($UseCurrentCredentials) {
    $service.UseDefaultCredentials = $true
} else {
    if (-not $MailboxEmail) {
        $MailboxEmail = Read-Host 'Enter mailbox email address'
    }
    $cred = Get-Credential -Message "Credentials for $MailboxEmail" -UserName $MailboxEmail
    $service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials(
        $cred.UserName,
        $cred.GetNetworkCredential().Password
    )
}

if (-not $MailboxEmail) {
    $MailboxEmail = Read-Host 'Enter mailbox email address'
}

if ($EwsUrl) {
    $service.Url = [Uri]$EwsUrl
} else {
    Write-Host "Running AutoDiscover for $MailboxEmail ..." -ForegroundColor Cyan
    $service.AutodiscoverUrl($MailboxEmail, { param($redirectionUrl) $true })
}

# ---------------------------------------------------------------------------
# 3. Log script start
# ---------------------------------------------------------------------------

$runMode = if ($ResumeFromCSV) { 'RESUME' } else { 'NEW SEARCH' }
Write-Log ("SCRIPT START | User: $($env:USERDOMAIN)\$($env:USERNAME) | " +
    "Mailbox: $MailboxEmail | Mode: $runMode")

# ---------------------------------------------------------------------------
# 4. Collect pending message IDs and display metadata
# ---------------------------------------------------------------------------

# $pendingIds   – ordered list of EWS UniqueId strings still needing a reply
# $displayItems – parallel list of PSCustomObjects for the summary table
$pendingIds   = [System.Collections.Generic.List[string]]::new()
$displayItems = [System.Collections.Generic.List[psobject]]::new()

if ($ResumeFromCSV) {
    # -----------------------------------------------------------------------
    # Resume mode – load tracker CSV, skip already-sent rows
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
    Write-Host "Loaded tracker: $($loadedRows.Count) total | $doneCount already submitted | $($pendingRows.Count) pending" `
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
        Write-Host "  Pending replies    : $($pendingIds.Count)" -ForegroundColor Green
        Write-Host "  Already submitted  : $doneCount"
        Write-Host "  Tracker CSV     : $script:CsvFilePath"
        Write-Host ("=" * 50) -ForegroundColor Yellow
        Write-Host ""
        exit 0
    }

} else {
    # -----------------------------------------------------------------------
    # New search mode – validate params, search, create CSV tracker
    # -----------------------------------------------------------------------

    $missing = [System.Collections.Generic.List[string]]::new()
    if ([string]::IsNullOrWhiteSpace($FolderPath))                    { $missing.Add('-FolderPath') }
    if ([string]::IsNullOrWhiteSpace($Subject))                       { $missing.Add('-Subject') }
    if (-not $PSBoundParameters.ContainsKey('StartDate'))              { $missing.Add('-StartDate') }
    if (-not $PSBoundParameters.ContainsKey('EndDate'))                { $missing.Add('-EndDate') }
    if ($missing.Count -gt 0) {
        Write-Error ("Missing required parameters for a new search: $($missing -join ', ')." +
            "`nTo resume an existing session use -ResumeFromCSV.")
        exit 1
    }

    Write-Host ""
    Write-Host "Connecting to folder: $FolderPath" -ForegroundColor Cyan
    $targetFolder = Get-EwsFolder -Service $service -Path $FolderPath
    Write-Host "Folder resolved : $($targetFolder.DisplayName)" -ForegroundColor Green

    Write-Host "Searching for messages ..." -ForegroundColor Cyan
    $foundItems = Find-MatchingMessages `
        -Folder      $targetFolder `
        -SubjectText $Subject `
        -RangeStart  $StartDate `
        -RangeEnd    $EndDate

    $matchCount = @($foundItems).Count

    Write-Log ("SEARCH | Folder: $FolderPath | Subject: `"$Subject`" | " +
        "DateRange: $($StartDate.ToString('yyyy-MM-dd')) to $($EndDate.ToString('yyyy-MM-dd')) | " +
        "Matches: $matchCount")

    # --- EstimateOnly: show count and exit without creating any files ---
    if ($EstimateOnly) {
        Write-Host ""
        Write-Host ("=" * 50) -ForegroundColor Yellow
        Write-Host " ESTIMATE RESULTS" -ForegroundColor Yellow
        Write-Host ("=" * 50) -ForegroundColor Yellow
        Write-Host "  Folder      : $($targetFolder.DisplayName)"
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

    # Build in-memory CSV rows and collect IDs / display data
    foreach ($item in $foundItems) {
        $emailItem  = $item -as [Microsoft.Exchange.WebServices.Data.EmailMessage]
        $fromAddr   = if ($emailItem -and $emailItem.From)   { $emailItem.From.Address }   `
                      elseif ($emailItem -and $emailItem.Sender) { $emailItem.Sender.Address } else { '' }
        $fromName   = if ($emailItem -and $emailItem.From)   { $emailItem.From.Name }      `
                      elseif ($emailItem -and $emailItem.Sender) { $emailItem.Sender.Name }   else { '(unknown)' }
        $received   = if ($emailItem) { $emailItem.DateTimeReceived.ToString('yyyy-MM-dd HH:mm') } else { '' }

        $script:CsvData.Add([PSCustomObject]@{
            MessageId         = $item.Id.UniqueId
            FromAddress       = $fromAddr
            Subject           = $item.Subject
            ResponseSubmitted = 'False'
        })

        $pendingIds.Add($item.Id.UniqueId)
        $displayItems.Add([PSCustomObject]@{
            Subject     = $item.Subject
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

$fullPropSet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet(
    [Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties
)
$bodyType = if ($replyIsHtml) {
    [Microsoft.Exchange.WebServices.Data.BodyType]::HTML
} else {
    [Microsoft.Exchange.WebServices.Data.BodyType]::Text
}

$allIds      = @($pendingIds)   # snapshot – unchanged throughout the loop
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

    foreach ($uniqueId in $currentBatch) {
        try {
            # Mark as submitted before the network call so a crash mid-send is visible in the CSV
            Update-CsvRow -MessageId $uniqueId -ResponseSubmitted 'True'
            Write-Log "REPLY SUBMITTED | MessageId: $uniqueId"

            $itemId  = New-Object Microsoft.Exchange.WebServices.Data.ItemId($uniqueId)
            $fullMsg = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind(
                $service, $itemId, $fullPropSet
            )

            $msgBody = New-Object Microsoft.Exchange.WebServices.Data.MessageBody(
                $bodyType, $replyContent
            )
            $fullMsg.Reply($msgBody, $ReplyAll.IsPresent)

            $fromAddr = if ($fullMsg.From) { $fullMsg.From.Address } else { '?' }
            Write-Log "REPLY SENT | MessageId: $uniqueId | From: $fromAddr | Subject: $($fullMsg.Subject)" `
                -Level SUCCESS

            $batchSent++
            $totalSent++
            Write-Host ("  [SENT]   -> {0}" -f $fromAddr) -ForegroundColor Green

        } catch {
            Write-Log "REPLY FAILED | MessageId: $uniqueId | Error: $_" -Level ERROR

            $batchFailed++
            $totalFailed++
            Write-Host ("  [FAIL]   $uniqueId — $_") -ForegroundColor Red
        }
    }

    Write-Host ""
    Write-Host ("Batch $batchNum complete. Sent: $batchSent | Failed: $batchFailed") -ForegroundColor Cyan
    Write-Log "BATCH COMPLETE | Batch: $batchNum | Sent: $batchSent | Failed: $batchFailed"

    $offset += $effectiveBatch

    # Pause between batches and let the user decide whether to continue
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
