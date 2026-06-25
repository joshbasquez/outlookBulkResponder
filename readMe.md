# Outlook Bulk Responder

Two PowerShell scripts for bulk-replying to mailbox messages. Choose based on your environment:

| Script | Target environment | Authentication |
|--------|--------------------|----------------|
| [`Invoke-GraphBulkResponder.ps1`](#invoke-graphbulkresponderps1--microsoft-365--graph-api) | Microsoft 365 / Exchange Online | OAuth 2.0 device code — browser sign-in, no DLL |
| [`Invoke-OutlookBulkResponder.ps1`](#invoke-outlookbulkresponderps1--on-premises-exchange--ews) | On-premises Exchange | EWS Managed API DLL + credentials |

Both scripts share the same batch-send, CSV tracker, log file, and `-ResumeFromCSV` workflow. The CSV files from one script are **not** interchangeable with the other.

---

## `Invoke-GraphBulkResponder.ps1` — Microsoft 365 / Graph API

Connects to a Microsoft 365 mailbox using the **Microsoft Graph API** and **OAuth 2.0 device code flow**. No EWS DLL is required — all HTTP calls use PowerShell's built-in `Invoke-RestMethod`. The script prints a short code and a URL; the user opens the URL in any browser, enters the code, and signs in to their Microsoft 365 account.

### Required Files (Graph)

| File | Purpose |
|------|---------|
| `Invoke-GraphBulkResponder.ps1` | The script |

No external DLLs. Requires PowerShell 5.1+ and outbound HTTPS access to `login.microsoftonline.com` and `graph.microsoft.com`.

---

### Azure AD App Registration

The script requires an Azure AD app registration with delegated Mail permissions and device code flow enabled. Follow these steps once per tenant.

#### Step 1 — Create the app registration

1. Sign in to the [Azure Portal](https://portal.azure.com) as a Global Administrator or Application Administrator.
2. Go to **Azure Active Directory** → **App registrations** → **New registration**.
3. Fill in the form:
   - **Name**: `Outlook Bulk Responder` (or any name you prefer)
   - **Supported account types**: *Accounts in this organizational directory only* (single-tenant) — use *any Azure AD directory* only if you need multi-tenant access
   - **Redirect URI**: leave blank (device code flow does not use a redirect URI)
4. Click **Register**.

#### Step 2 — Enable public client / device code flow

1. In the app registration, go to **Authentication**.
2. Under **Advanced settings**, set **Allow public client flows** to **Yes**.
3. Click **Save**.

#### Step 3 — Grant API permissions

1. In the app registration, go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**.
2. Search for and add:
   - `Mail.Read`
   - `Mail.Send`
3. If your tenant requires admin consent for delegated permissions, click **Grant admin consent for [your tenant]** and confirm.

#### Step 4 — Collect your IDs

From the app registration **Overview** page, copy:

| Value | Where to find it | Script parameter |
|-------|-----------------|-----------------|
| **Tenant ID** | Azure Active Directory → Overview → *Tenant ID* | `-TenantId` |
| **Client ID** | App registrations → your app → *Application (client) ID* | `-ClientId` |

---

### Authentication Flow (device code)

When the script starts it contacts `login.microsoftonline.com` and prints:

```
==============================================================
 AUTHENTICATION REQUIRED
==============================================================

  1. Open a browser and navigate to:
     https://microsoft.com/devicelogin

  2. Enter this code when prompted:
     WDJB-MJHT

  Waiting for sign-in (expires in 900s) ...
==============================================================
```

Open the URL in any browser (on any device), enter the code, and sign in with the account that has access to the target mailbox. The script polls silently and continues automatically once sign-in is complete.

The script acquires a **refresh token** alongside the access token and uses it to silently renew the access token during long batch runs, so you will not be prompted again mid-session unless the refresh token itself expires (typically 90 days for most tenants).

---

### Parameters (Graph)

#### Authentication

| Parameter | Type | Description |
|-----------|------|-------------|
| `-TenantId` | String | Azure AD tenant ID (GUID) or `common`. Prompted if omitted. |
| `-ClientId` | String | App registration client ID (GUID). Prompted if omitted. |
| `-MailboxEmail` | String | SMTP address of the mailbox to access. The signed-in account must have access to this mailbox. Prompted if omitted. |

#### New-search parameters *(required unless `-ResumeFromCSV` is used)*

| Parameter | Type | Description |
|-----------|------|-------------|
| `-FolderPath` | String | Folder to search. Well-known names: `Inbox`, `Drafts`, `Sent Items`, `Deleted Items`, `Junk Email`, `Archive`. Use `\` for subfolders. |
| `-Subject` | String | Substring to match against message subjects (case-insensitive). |
| `-StartDate` | DateTime | Inclusive start of the received-date range. |
| `-EndDate` | DateTime | Inclusive end of the received-date range (covers through midnight UTC). |

#### Reply body *(one or neither — prompted interactively if both omitted)*

| Parameter | Type | Description |
|-----------|------|-------------|
| `-ReplyBody` | String | Plain-text reply body on the command line. |
| `-ReplyHtmlFile` | String | Path to an HTML file to use as the reply body. |

#### Behaviour

| Parameter | Type | Description |
|-----------|------|-------------|
| `-EstimateOnly` | Switch | Show match count only. No reply sent, no CSV created. Also works with `-ResumeFromCSV`. |
| `-ReplyAll` | Switch | Use Reply All instead of Reply. |
| `-BatchSize` | Int | Replies per batch. `0` (default) prompts interactively; enter `0` at the prompt to send all at once. |
| `-ResumeFromCSV` | String | Path to a tracker CSV from a prior run. Skips rows where `ResponseSubmitted = True`. |
| `-USGovDoD` | Switch | Route all traffic to the Microsoft 365 GCC High / DoD sovereign cloud. See [USGovDoD endpoints](#usgov-dod-sovereign-cloud) below. |

---

### Example Usages (Graph)

#### Estimate match count

```powershell
.\Invoke-GraphBulkResponder.ps1 `
    -TenantId     "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -ClientId     "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
    -MailboxEmail "you@contoso.com" `
    -FolderPath   "Inbox" `
    -Subject      "Project Alpha Status" `
    -StartDate    "2024-03-01" `
    -EndDate      "2024-03-31" `
    -EstimateOnly
```

#### Send plain-text replies in batches of 10

```powershell
.\Invoke-GraphBulkResponder.ps1 `
    -TenantId     "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -ClientId     "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
    -MailboxEmail "support@contoso.com" `
    -FolderPath   "Inbox\Client Requests" `
    -Subject      "Support Ticket" `
    -StartDate    "2024-04-01" `
    -EndDate      "2024-04-30" `
    -BatchSize    10
```

After authentication, the script displays matched messages, prompts for the reply body (type `END` to finish), and asks for confirmation before each batch.

#### Send HTML replies from a file, Reply All

```powershell
.\Invoke-GraphBulkResponder.ps1 `
    -TenantId      "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -ClientId      "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
    -MailboxEmail  "billing@contoso.com" `
    -FolderPath    "Inbox" `
    -Subject       "Invoice" `
    -StartDate     "2024-01-01" `
    -EndDate       "2024-01-31" `
    -ReplyHtmlFile ".\auto-reply.html" `
    -ReplyAll `
    -BatchSize     25
```

#### Resume a paused session

```powershell
.\Invoke-GraphBulkResponder.ps1 `
    -TenantId      "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -ClientId      "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
    -MailboxEmail  "you@contoso.com" `
    -ResumeFromCSV ".\bulk-responder-202403151030.csv" `
    -ReplyHtmlFile ".\auto-reply.html" `
    -BatchSize     10
```

#### Estimate pending count from a saved CSV

```powershell
.\Invoke-GraphBulkResponder.ps1 `
    -TenantId      "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -ClientId      "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
    -MailboxEmail  "you@contoso.com" `
    -ResumeFromCSV ".\bulk-responder-202403151030.csv" `
    -EstimateOnly
```

---

### USGov DoD Sovereign Cloud

Add `-USGovDoD` to any invocation to route all authentication and API traffic away from the commercial cloud endpoints and to the DoD sovereign cloud endpoints:

| Endpoint | Commercial | USGovDoD |
|----------|-----------|---------|
| Login / token | `login.microsoftonline.com` | `login.microsoftonline.us` |
| Graph API | `graph.microsoft.com/v1.0` | `dod-graph.microsoft.us/v1.0` |
| OAuth scope resource | `https://graph.microsoft.com` | `https://dod-graph.microsoft.us` |

#### Prerequisites for DoD

- The app registration must be created in the **GCC High / DoD Azure AD tenant** (portal: `portal.azure.us`), not the commercial portal.
- The target mailbox must be hosted in the same DoD tenant.
- The app registration steps are identical to the commercial steps above, performed at `portal.azure.us` instead of `portal.azure.com`.
- The device code sign-in URL displayed by the script (`https://microsoft.com/devicelogin`) is the same for both clouds; the token endpoint (`login.microsoftonline.us`) is what changes.

#### Example — DoD bulk reply in batches of 20

```powershell
.\Invoke-GraphBulkResponder.ps1 `
    -TenantId     "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -ClientId     "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
    -MailboxEmail "you@mail.mil" `
    -FolderPath   "Inbox" `
    -Subject      "TASKORD" `
    -StartDate    "2024-06-01" `
    -EndDate      "2024-06-30" `
    -ReplyHtmlFile ".\reply.html" `
    -BatchSize    20 `
    -USGovDoD
```

#### Example — DoD estimate only

```powershell
.\Invoke-GraphBulkResponder.ps1 `
    -TenantId     "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -ClientId     "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
    -MailboxEmail "you@mail.mil" `
    -FolderPath   "Inbox" `
    -Subject      "TASKORD" `
    -StartDate    "2024-06-01" `
    -EndDate      "2024-06-30" `
    -EstimateOnly `
    -USGovDoD
```

#### Example — DoD resume from a saved CSV

```powershell
.\Invoke-GraphBulkResponder.ps1 `
    -TenantId      "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -ClientId      "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
    -MailboxEmail  "you@mail.mil" `
    -ResumeFromCSV ".\bulk-responder-202406011430.csv" `
    -ReplyHtmlFile ".\reply.html" `
    -BatchSize     20 `
    -USGovDoD
```

When the user exits mid-batch, the resume command printed to the console automatically includes `-USGovDoD` so the correct endpoints are used on the next run.

---

### Notes (Graph)

- **Body quoting** — The script sets `message.body` on the reply, which replaces the full body. The original quoted message is not automatically appended. To include quoted content, add it manually to your HTML file.
- **Graph throttling** — The Graph API enforces per-user and per-app rate limits. If you encounter `429 Too Many Requests` errors, reduce `-BatchSize` and add pauses between batches.
- **Admin consent** — Some tenants enforce admin consent for all delegated permissions. If the device code sign-in fails with a consent error, have a Global Administrator grant consent via the Azure Portal step above.
- **Shared / delegated mailboxes** — The signed-in account must have `Full Access` (or at minimum `Send As`) permission on the target mailbox in Exchange Online for the reply to succeed.
- **Log entries** — All log lines from this script are tagged `[GRAPH]` so they are distinguishable from EWS script entries in the shared `bulk-responder.log`.

---

## `Invoke-OutlookBulkResponder.ps1` — On-premises Exchange (EWS)

A PowerShell script that connects to an Exchange mailbox via the **EWS Managed API**, searches a specified folder for email messages matching a subject string and date range, and sends replies in configurable batches. A persistent log file records every run and every sent reply. A per-run CSV tracker file records the status of every matched message and can be used to resume a session that was paused or interrupted.

---

## Required Files

| File | Purpose |
|------|---------|
| `Invoke-OutlookBulkResponder.ps1` | The script |
| `Microsoft.Exchange.WebServices.dll` | EWS Managed API assembly |
| `Microsoft.Exchange.WebServices.Auth.dll` | Companion auth assembly (ships alongside the main DLL) |

### Downloading the EWS Managed API DLL

The EWS Managed API is no longer distributed through the Microsoft Download Center. The current canonical source is the NuGet package **Microsoft.Exchange.WebServices** (version 2.2).

**Option A – NuGet CLI**

```powershell
# Install the NuGet CLI if not already present, then:
nuget install Microsoft.Exchange.WebServices -Version 2.2 -OutputDirectory .\packages
```

The DLL will be at:
```
.\packages\Microsoft.Exchange.WebServices.2.2\lib\40\Microsoft.Exchange.WebServices.dll
```

Copy that file (and the Auth DLL in the same directory) into the folder containing the script.

**Option B – Download via PowerShell (no NuGet CLI needed)**

```powershell
$url  = "https://www.nuget.org/api/v2/package/Microsoft.Exchange.WebServices/2.2"
$dest = ".\ews.zip"
Invoke-WebRequest -Uri $url -OutFile $dest
Expand-Archive -Path $dest -DestinationPath .\ews-package
# DLL is under .\ews-package\lib\40\
```

**Option C – GitHub source**

The open-sourced version is available at `OfficeDev/ews-managed-api` on GitHub. Building it from source requires the .NET 4.x SDK.

---

## Prerequisites

### Exchange / Authentication

| Scenario | What to configure |
|----------|------------------|
| **On-premises Exchange (domain-joined machine)** | Use `-UseCurrentCredentials`. No additional setup needed — the script authenticates as the current Windows user via NTLM/Kerberos. |
| **On-premises Exchange (explicit credentials)** | Omit `-UseCurrentCredentials`. The script will call `Get-Credential` and pass the result to EWS as `WebCredentials`. |
| **Exchange Online / Microsoft 365 (basic auth)** | Basic auth is **disabled by default** on M365 as of October 2022. This script uses `WebCredentials` which relies on basic auth. It will not work against M365 unless basic auth has been re-enabled for EWS in your tenant (not recommended). |
| **Exchange Online with OAuth** | OAuth support requires acquiring a token via MSAL and assigning it to `$service.Credentials` as an `OAuthCredentials` object. This script does not implement OAuth; for M365 production use, consider the [Microsoft Graph API](https://learn.microsoft.com/en-us/graph/overview) instead. |

### AutoDiscover

By default the script calls `AutodiscoverUrl` to locate the EWS endpoint. AutoDiscover must be resolvable from the machine running the script. If it is not, supply the full EWS URL with `-EwsUrl`:

```
-EwsUrl "https://mail.contoso.com/EWS/Exchange.asmx"
```

### PowerShell Execution Policy

The script requires at minimum `RemoteSigned`. If your machine blocks unsigned scripts, run:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

### .NET Framework

The EWS Managed API DLL targets .NET 4.0. PowerShell 5.1 (which ships with Windows) runs on .NET 4.x and is fully compatible. PowerShell 7+ runs on .NET Core/5+ and can load the DLL via `Add-Type -Path`, but some edge cases around assembly binding may differ.

---

## Output Files

Every run produces or appends to two files in the script directory.

### Log File — `bulk-responder.log`

A single, persistent log file that all runs append to. Each line is timestamped and tagged with a severity level.

```
[2024-03-15 10:30:00] [INFO]    SCRIPT START | User: CORP\jsmith | Mailbox: jsmith@contoso.com | Mode: NEW SEARCH
[2024-03-15 10:30:01] [INFO]    SEARCH | Folder: Inbox | Subject: "Invoice" | DateRange: 2024-01-01 to 2024-01-31 | Matches: 47
[2024-03-15 10:30:02] [INFO]    CSV CREATED | Path: C:\scripts\bulk-responder-202403151030.csv | Rows: 47
[2024-03-15 10:30:03] [INFO]    BATCH START | Batch: 1 | Range: 1-10 of 47
[2024-03-15 10:30:04] [INFO]    REPLY SUBMITTED | MessageId: AAMkAGI2...
[2024-03-15 10:30:04] [SUCCESS] REPLY SENT | MessageId: AAMkAGI2... | From: vendor@example.com | Subject: Invoice #1021
[2024-03-15 10:30:09] [INFO]    BATCH COMPLETE | Batch: 1 | Sent: 10 | Failed: 0
[2024-03-15 10:31:00] [INFO]    USER EXIT | After batch 1 | TotalSent: 10 | TotalFailed: 0 | Remaining: 37
[2024-03-15 14:00:00] [INFO]    SCRIPT START | User: CORP\jsmith | Mailbox: jsmith@contoso.com | Mode: RESUME
[2024-03-15 14:00:01] [INFO]    RESUME | CSV: bulk-responder-202403151030.csv | Total: 47 | AlreadySent: 10 | Pending: 37
```

Logged events:

| Event | When |
|-------|------|
| `SCRIPT START` | Beginning of every run |
| `SEARCH` | After search query executes (new search only) |
| `CSV CREATED` | When the tracker CSV is first written |
| `RESUME` | When loading a tracker CSV to continue |
| `BATCH START / COMPLETE` | Around each batch |
| `REPLY SUBMITTED` | Before each `Reply()` call |
| `REPLY SENT` | After `Reply()` returns without error |
| `REPLY FAILED` | When `Reply()` throws an exception |
| `USER EXIT` | When the user chooses to save progress and exit |
| `USER CANCELLED` | When the user declines the send confirmation |
| `SCRIPT COMPLETE` | End of a run that processed all pending messages |

### CSV Tracker — `bulk-responder-YYYYmmddHHmm.csv`

Created at the start of each new search run. The filename timestamp (`YYYYmmddHHmm`) marks when that run began. When resuming, you pass this file back to the script with `-ResumeFromCSV`.

| Column | Description |
|--------|-------------|
| `MessageId` | EWS UniqueId of the matched message. Used to re-bind the item when resuming. |
| `FromAddress` | Sender's SMTP address. |
| `Subject` | Message subject line. |
| `ResponseSubmitted` | `True` once `EmailMessage.Reply()` is called for this message. |

The CSV is written to disk after every individual message is processed, so progress is never lost if the script is killed.

**Example CSV content:**

```csv
"MessageId","FromAddress","Subject","ResponseSubmitted"
"AAMkAGI2...","vendor1@example.com","Invoice #1021","True"
"AAMkAGI3...","vendor2@example.com","Invoice #1022","True"
"AAMkAGI4...","vendor3@example.com","Invoice #1023","False"
```

When resuming, the script skips any row where `ResponseSubmitted = True` and only sends replies for the remaining rows.

---

## Parameters

### New Search Parameters
*(Required unless `-ResumeFromCSV` is used; ignored when resuming)*

| Parameter | Type | Description |
|-----------|------|-------------|
| `-FolderPath` | String | Mailbox folder to search. Use `\` for subfolders. E.g. `"Inbox"` or `"Inbox\Projects\2024"`. Well-known names: `Inbox`, `Drafts`, `Sent Items`, `Deleted Items`, `Junk Email`, `Archive`. |
| `-Subject` | String | Text to match against message subjects (case-insensitive, substring match). |
| `-StartDate` | DateTime | Inclusive start of the received-date range. |
| `-EndDate` | DateTime | Inclusive end of the received-date range (covers through 23:59:59 of that day). |

### Connection Parameters

| Parameter | Type | Description |
|-----------|------|-------------|
| `-MailboxEmail` | String | SMTP address of the mailbox. Used for AutoDiscover and credential prompt. Prompted interactively if omitted. |
| `-EwsDllPath` | String | Path to `Microsoft.Exchange.WebServices.dll`. Defaults to `.\Microsoft.Exchange.WebServices.dll`. |
| `-EwsUrl` | String | Manual EWS endpoint URL. Bypasses AutoDiscover. |
| `-UseCurrentCredentials` | Switch | Authenticate as the current Windows user (NTLM/Kerberos). |

### Reply Body Parameters
*(Both optional — if neither is provided the script prompts interactively)*

| Parameter | Type | Description |
|-----------|------|-------------|
| `-ReplyBody` | String | Plain-text reply body passed on the command line. |
| `-ReplyHtmlFile` | String | Path to an HTML file used as the reply body (sent as HTML). |

### Behaviour Parameters

| Parameter | Type | Description |
|-----------|------|-------------|
| `-EstimateOnly` | Switch | Show match count only. No reply is sent and no CSV is created. Works in resume mode too (shows pending count). |
| `-ReplyAll` | Switch | Use Reply All instead of Reply. |
| `-BatchSize` | Int | Number of replies to send before pausing. `0` (default) prompts interactively; enter `0` at the prompt to send all without pausing. |
| `-ResumeFromCSV` | String | Path to a tracker CSV from a previous run. Skips messages already marked `ResponseSentSuccessfully = True`. |

---

## Batch Processing

After confirmation, the script sends `-BatchSize` replies, then pauses:

```
--- Batch 1 | Messages 1-10 of 47 ---
  [SENT]   -> vendor1@example.com
  [SENT]   -> vendor2@example.com
  ...

Batch 1 complete. Sent: 10 | Failed: 0

37 message(s) still pending.
[C] Continue next batch   [E] Exit and save progress
```

Pressing **E** saves the CSV, writes a log entry, and prints the exact command needed to resume:

```
Progress saved.
  CSV tracker : C:\scripts\bulk-responder-202403151030.csv
  Log file    : C:\scripts\bulk-responder.log

To resume this session later, run:

.\Invoke-OutlookBulkResponder.ps1 `
    -ResumeFromCSV "C:\scripts\bulk-responder-202403151030.csv" `
    -MailboxEmail "you@contoso.com" `
    -ReplyHtmlFile ".\auto-reply.html" `
    -BatchSize 10
```

---

## Example Usages

### Estimate only – count matching messages without sending

```powershell
.\Invoke-OutlookBulkResponder.ps1 `
    -FolderPath   "Inbox" `
    -Subject      "Project Alpha Status" `
    -StartDate    "2024-03-01" `
    -EndDate      "2024-03-31" `
    -MailboxEmail "you@contoso.com" `
    -EstimateOnly
```

Sample output:
```
==================================================
 ESTIMATE RESULTS
==================================================
  Folder      : Inbox
  Subject     : Project Alpha Status
  Date range  : 2024-03-01  to  2024-03-31
  Matches     : 14
==================================================
```

---

### Send plain-text replies in batches of 10 – body entered interactively

```powershell
.\Invoke-OutlookBulkResponder.ps1 `
    -FolderPath   "Inbox\Client Requests" `
    -Subject      "Support Ticket" `
    -StartDate    "2024-04-01" `
    -EndDate      "2024-04-30" `
    -MailboxEmail "support@contoso.com" `
    -BatchSize    10
```

The script will:
1. Connect and search the folder.
2. Create `bulk-responder-YYYYmmddHHmm.csv` and display the matched messages.
3. Prompt you to type the reply body (end with `END` on its own line).
4. Ask for confirmation, then send the first batch of 10.
5. Pause after each batch and ask whether to continue or exit.

---

### Send HTML replies from a file

```powershell
.\Invoke-OutlookBulkResponder.ps1 `
    -FolderPath    "Inbox" `
    -Subject       "Invoice" `
    -StartDate     "2024-01-01" `
    -EndDate       "2024-01-31" `
    -MailboxEmail  "billing@contoso.com" `
    -ReplyHtmlFile ".\auto-reply.html" `
    -BatchSize     25
```

`auto-reply.html` example:
```html
<p>Thank you for your message. Your invoice has been received and is being processed.</p>
<p>If you have questions, please contact <a href="mailto:billing@contoso.com">billing@contoso.com</a>.</p>
```

---

### Reply All, HTML body, current Windows credentials, manual EWS URL

```powershell
.\Invoke-OutlookBulkResponder.ps1 `
    -FolderPath          "Inbox\Announcements" `
    -Subject             "All Hands Meeting" `
    -StartDate           "2024-06-01" `
    -EndDate             "2024-06-15" `
    -ReplyHtmlFile       ".\meeting-reply.html" `
    -ReplyAll `
    -UseCurrentCredentials `
    -EwsUrl              "https://mail.corp.local/EWS/Exchange.asmx" `
    -BatchSize           20
```

---

### Supply the reply body as a command-line argument (fully non-interactive)

```powershell
.\Invoke-OutlookBulkResponder.ps1 `
    -FolderPath   "Inbox" `
    -Subject      "Survey" `
    -StartDate    "2024-05-01" `
    -EndDate      "2024-05-31" `
    -MailboxEmail "hr@contoso.com" `
    -ReplyBody    "Thank you for completing the survey. Results will be shared next week." `
    -BatchSize    0
```

Passing `-BatchSize 0` sends all messages in a single uninterrupted run.

---

### Resume a previous session

After choosing to exit mid-batch, the script prints the exact resume command. Copy and run it:

```powershell
.\Invoke-OutlookBulkResponder.ps1 `
    -ResumeFromCSV ".\bulk-responder-202403151030.csv" `
    -MailboxEmail  "billing@contoso.com" `
    -ReplyHtmlFile ".\auto-reply.html" `
    -BatchSize     25
```

On resume the script:
1. Loads the CSV and counts how many messages still have `ResponseSentSuccessfully = False`.
2. Displays the pending-messages table.
3. Prompts for confirmation and proceeds with batching exactly as in a new run.
4. Retries any message where `ResponseSubmitted = True` but `ResponseSentSuccessfully = False` (send was attempted but failed).

### Estimate pending count from a saved CSV

```powershell
.\Invoke-OutlookBulkResponder.ps1 `
    -ResumeFromCSV ".\bulk-responder-202403151030.csv" `
    -MailboxEmail  "billing@contoso.com" `
    -EstimateOnly
```

Sample output:
```
==================================================
 RESUME ESTIMATE
==================================================
  Pending replies : 37
  Already sent    : 10
  Tracker CSV     : C:\scripts\bulk-responder-202403151030.csv
==================================================
```

---

## Notes

- **Rate limiting** – Exchange throttling policies may slow or block bulk operations. Use `-BatchSize` to spread sends over multiple sessions and avoid hitting server limits.
- **Sent Items** – Replies use `EmailMessage.Reply()`, which sends immediately. A copy is not guaranteed to appear in Sent Items via this method; whether it does depends on your Exchange server's configuration.
- **Resume skips submitted messages** – `ResponseSubmitted` is written to the CSV as `True` when `Reply()` is called. On resume, any row already marked `True` is skipped. If a send call was made but the reply was not actually delivered, correct that message manually before resuming to avoid missing it.
- **Test first** – Always run with `-EstimateOnly` before a bulk send to verify the filter matches the expected messages.
- **EWS deprecation on M365** – Microsoft has indicated that EWS access for M365 tenants will be progressively restricted. For new M365 integrations, the Microsoft Graph API (`/v1.0/me/messages`) is the recommended path.
