# NPSBox — Migrate Box collaboration permissions to OneDrive for Business

This repository contains a PowerShell script that **translates Box collaboration/export data** into **OneDrive for Business item-level permissions** using the **Microsoft Graph PowerShell SDK**.

> **Note:** This is sample/illustration code. See [`DISCLAIMER.MD`](DISCLAIMER.MD) for terms.
> For a full audit of assumptions and constraints, see [`assumptions.md`](assumptions.md).

---

## Why Microsoft Graph instead of PnP PowerShell

The project originally used `PnP.PowerShell` to resolve OneDrive items via SharePoint CSOM and set permissions through SharePoint role definitions. That approach was replaced with a pure Microsoft Graph implementation for the following reasons:

| Concern | PnP PowerShell (old) | Microsoft Graph (current) |
|---|---|---|
| **Assembly conflict** | PnP.PowerShell bundles `Microsoft.Graph.Core` 1.x, which fatally conflicts with the Graph SDK v2 (3.x) in the same PowerShell session. Running both required separate processes. | Single SDK — no assembly conflicts. |
| **Item resolution** | Resolved items via SharePoint list item IDs, which is sensitive to SharePoint URL encoding rules (`#`, `%`, `&`). Deep nested paths were fragile. | Resolves items by OneDrive path via [`driveItem: get`](https://learn.microsoft.com/graph/api/driveitem-get?view=graph-rest-1.0#access-a-driveitem-by-path). Handles special characters via per-segment URL encoding. |
| **Permission model** | Mapped Box roles → SharePoint role definitions (`Contribute`, `Read`). Required site collection admin access and user profile property lookups. | Maps Box roles → Graph sharing roles (`write`, `read`). Uses the [`driveItem: invite`](https://learn.microsoft.com/graph/api/driveitem-invite?view=graph-rest-1.0) API with app-only permissions. |
| **Authentication** | Multiple auth modes (`Interactive`, `ManagedIdentity`, `Certificate`, etc.) added complexity. Required SPO admin URL and MySite host URL parameters. | Certificate-based app-only auth via [`Connect-MgGraph`](https://learn.microsoft.com/powershell/module/microsoft.graph.authentication/connect-mggraph). Three parameters: `TenantId`, `ClientId`, `CertificateThumbprint`. |
| **File upload** | Not supported — permissions only. | Uploads local files to OneDrive. Files ≤ 4 MB use the [simple upload API](https://learn.microsoft.com/graph/api/driveitem-put-content?view=graph-rest-1.0); larger files use a chunked [resumable upload session](https://learn.microsoft.com/graph/api/driveitem-createuploadsession?view=graph-rest-1.0). |
| **Retry / throttling** | No built-in retry. PnP cmdlets surfaced raw 429 errors. | `Invoke-WithGraphRetry` with exponential backoff, `Retry-After` header parsing, and configurable max attempts. |
| **Testability** | CSOM calls were difficult to mock in Pester. | All Graph calls go through `Invoke-MgGraphRequest`, which is straightforward to mock. Full Pester 5 test suite (87 tests for `Update-UserFile.ps1`). |

The legacy PnP scripts (`Set-BoxToOneDriveItemPermission.ps1`, `Connect-SharePointOnlineAdmin.ps1`, `Add-SCA.ps1`, `Add-AppCertificateCredential.ps1`, `Get-UserLocalFiles.ps1`) have been removed from the active project. Archived sample data and the standalone `Test-AzureAppRegistration.ps1` validator (now folded into `-Test` mode of the main script) are in `holding/`.

---

## Table of contents

- [Repository layout](#repository-layout)
- [Input CSV format](#input-csv-format)
- [Update-UserFile.ps1](#update-userfileps1)
  - [What it does](#what-it-does)
  - [Prerequisites](#prerequisites)
  - [Parameters](#parameters)
  - [Parameter sets and operating modes](#parameter-sets-and-operating-modes)
  - [Configuration file (config.json)](#configuration-file-configjson)
  - [Box role mapping](#box-role-mapping)
  - [Authentication](#authentication)
  - [Usage examples](#usage-examples)
  - [File upload](#file-upload)
  - [Output objects](#output-objects)
  - [Logging](#logging)
  - [Retry and throttle handling](#retry-and-throttle-handling)
  - [Known limitations](#known-limitations)
- [Set-NPSBoxEnv.ps1](#set-npsboxenvps1)
- [GitHub Copilot customization (.github/)](#github-copilot-customization-github)
- [Disclaimer](#disclaimer)
- [Assumptions & constraints](#assumptions--constraints)

---

## Repository layout

| Path | Purpose |
|---|---|
| `Update-UserFile.ps1` | Main script — resolves OneDrive items, uploads files, and grants permissions via Microsoft Graph. Also runs as a pre-flight validator with `-Test`. |
| `config.json` | Optional JSON config providing defaults for `InputFile`, `TenantId`, `ClientId`, `CertificateThumbprint`, `LogFolder`, and `AllFilesDirectory`. Loaded automatically from the script directory when present. |
| `Set-NPSBoxEnv.ps1` | Helper to persist `NPSBOX_TENANT_ID`, `NPSBOX_CLIENT_ID`, and `NPSBOX_CERT_THUMBPRINT` as user-scope environment variables (consumed as parameter defaults by `Update-UserFile.ps1`). |
| `UserInfo.csv` | Input CSV — Box collaboration data to process |
| `tests/Update-UserFile.Tests.ps1` | Pester 5 test suite for `Update-UserFile.ps1` (87 tests; all Graph/file-system calls are mocked) |
| `tests/Test-AzureAppRegistration.Tests.ps1` | Pester 5 test suite that exercises the archived `holding/Test-AzureAppRegistration.ps1` validator (11 tests) |
| `tests/results/` | Pester test result output (NUnit XML), named by version |
| `LocalFiles/` | Per-user subfolders (named by UPN) containing files to upload with `-UploadFiles` |
| `logs/` | Timestamped log output (git-ignored) |
| `en-US/about_Update-UserFile.help.txt` | About-topic help text shown when the script is invoked with no parameters |
| `holding/` | Archived files: `Test-AzureAppRegistration.ps1` (validator, now folded into `-Test` mode), `Box_Collaboration_Sample_Data.csv`, `Top10.MD`, `setupcert.txt` |
| `.github/` | GitHub Copilot customization files (instructions, agents, prompts) |
| [`DISCLAIMER.MD`](DISCLAIMER.MD) | Disclaimer text |
| [`assumptions.md`](assumptions.md) | Full assumptions and constraints audit |

---

## Input CSV format

The script consumes a CSV with these required columns:

| Column | Description | Example |
|---|---|---|
| `Owner Login` | UPN of the file owner | `user@contoso.com` |
| `Path` | Box path to the item | `All Files/Documents` |
| `Item Name` | File or folder name | `Report.pdf` |
| `Collaborator Login` | UPN of the person to share with | `colleague@contoso.com` |
| `Collaborator Permission` | Box role | `Editor`, `Viewer`, `Co-owner` |

Column names must match exactly (case-sensitive). See `UserInfo.csv` for a working example.

---

## Update-UserFile.ps1

### What it does

For a given Box owner (user), the script:

1. **Authenticates** to Microsoft Graph using certificate-based app-only auth.
2. **Reads the CSV** and identifies unique owners to process.
   If `-UserToProcess` is specified, only that user is processed;
   otherwise all unique owners in the CSV are processed.
3. **Looks up** the user's OneDrive drive via the [Graph API](https://learn.microsoft.com/graph/overview).
4. **Optionally uploads** local files and folders to the user's OneDrive (when `-UploadFiles` is specified).
5. **Validates each owner account** with `Get-MgUser` (exists, enabled) and verifies the OneDrive drive is provisioned and reachable before processing rows.
6. **Deduplicates** CSV rows for each owner (same Path + Item Name + Collaborator Login + Permission) to reduce redundant API calls.
7. **Resolves each item** by path in OneDrive using the [`driveItem: get` API](https://learn.microsoft.com/graph/api/driveitem-get?view=graph-rest-1.0#access-a-driveitem-by-path).
8. **Validates** the collaborator's email format and domain against `-AllowedDomains` to prevent unintended external sharing.
9. **Grants permissions** to each collaborator using the [`driveItem: invite` API](https://learn.microsoft.com/graph/api/driveitem-invite?view=graph-rest-1.0) with `sendInvitation = false` (no email notifications) and `requireSignIn = true`.
10. **Outputs** a structured `[pscustomobject]` result per CSV row to the pipeline for reporting (see [Output objects](#output-objects)).

When invoked with **no parameters**, the script displays its about-topic help (`en-US/about_Update-UserFile.help.txt`) and exits without authenticating.

### Prerequisites

- **PowerShell 7.0 or later** — [Install PowerShell](https://learn.microsoft.com/powershell/scripting/install/installing-powershell)
- **Microsoft Graph PowerShell SDK** modules:

```powershell
Install-Module Microsoft.Graph.Authentication  -Scope CurrentUser
Install-Module Microsoft.Graph.Users           -Scope CurrentUser
Install-Module Microsoft.Graph.Files           -Scope CurrentUser
Install-Module Microsoft.Graph.Applications    -Scope CurrentUser
```

- [Microsoft Graph PowerShell SDK installation guide](https://learn.microsoft.com/powershell/microsoftgraph/installation)
- **Azure AD App Registration** with the following **Application** permissions
  granted with admin consent:
  - `Files.ReadWrite.All`
  - `User.Read.All`
  - `Application.Read.All` *(optional — enables startup permission validation via `Assert-GraphPermissions` and detailed per-permission reporting in `-Test` mode; without it the check is reported as `Skipped`)*
  - [Register an application](https://learn.microsoft.com/entra/identity-platform/quickstart-register-app)
- A certificate uploaded to the app registration — [Certificate credentials](https://learn.microsoft.com/entra/identity-platform/certificate-credentials)
- Run `Update-UserFile.ps1 -Test ...` to verify modules, assembly compatibility, certificate auth, and Graph permissions before your first migration.

> **Important:** Do not load `PnP.PowerShell` in the same session. PnP.PowerShell bundles `Microsoft.Graph.Core` 1.x, which conflicts with the Graph SDK v2. Start a fresh `pwsh` session.

### Parameters

| Parameter | Type | Required | Default | Description |
|---|---|---|---|---|
| `-ConfigFile` | `string` | No | `<script>\config.json` | Path to a JSON config file providing defaults for `InputFile`, `TenantId`, `ClientId`, `CertificateThumbprint`, `LogFolder`, and `AllFilesDirectory`. Missing file is OK. See [Configuration file (config.json)](#configuration-file-configjson). |
| `-InputFile` | `FileInfo` | No | `C:\Repos\NPSBox\UserInfo.csv` | Path to the input CSV file |
| `-UserToProcess` | `string` | No | — | Owner UPN to filter on. When omitted, all unique owners in the CSV are processed. Accepts pipeline input. Aliases: `Owner Login`, `User`, `UPN`, `Account` |
| `-TenantId` | `string` | No | config.json &rarr; `$env:NPSBOX_TENANT_ID` | Microsoft 365 tenant ID (GUID). [Find your tenant ID](https://learn.microsoft.com/entra/fundamentals/how-to-find-tenant) |
| `-ClientId` | `string` | No | config.json &rarr; `$env:NPSBOX_CLIENT_ID` | App registration client ID (GUID) |
| `-CertificateThumbprint` | `string` | No | config.json &rarr; `$env:NPSBOX_CERT_THUMBPRINT` | SHA-1 thumbprint of a certificate in `Cert:\CurrentUser\My` |
| `-LogFolder` | `string` | No | `C:\Repos\NPSBox\Logs` | Folder for timestamped log files |
| `-AllFilesDirectory` | `string` | No | `C:\Repos\NPSBox\LocalFiles` | Root folder containing per-user subfolders of local files to upload |
| `-UploadFiles` | `switch` | No | — | When present, uploads local files to OneDrive before applying permissions (parameter set `Run` only) |
| `-AllowedDomains` | `string[]` | No | — | List of allowed email domains for collaborators. Blocks external sharing to unlisted domains. (parameter set `Run` only) |
| `-Test` | `switch` | **Yes (`Test` set)** | — | Runs validation only (modules, assembly compatibility, certificate auth, Graph permissions) and exits without processing CSV rows. |

The credential parameters (`TenantId`, `ClientId`, `CertificateThumbprint`) are optional on the command line — values are resolved in this order (highest first):

1. Explicit `-TenantId` / `-ClientId` / `-CertificateThumbprint` on the command line.
2. Values from `config.json` (or whatever `-ConfigFile` points at).
3. Environment variables `NPSBOX_TENANT_ID`, `NPSBOX_CLIENT_ID`, `NPSBOX_CERT_THUMBPRINT` (see [`Set-NPSBoxEnv.ps1`](#set-npsboxenvps1)).

If none of the three sources provides a value, `ValidateNotNullOrEmpty()` rejects the empty string and the script aborts.

Common parameters `-WhatIf`, `-Confirm`, and `-Verbose` are fully supported.

### Parameter sets and operating modes

The script declares three parameter sets that select how it runs:

| Parameter set | Trigger | Behavior |
|---|---|---|
| `Help` (default) | No parameters supplied | Displays `en-US/about_Update-UserFile.help.txt` (or `Get-Help -Detailed` as fallback) and exits without authenticating. |
| `Test` | `-Test` plus credentials | Runs the validation checklist (assembly compatibility, required modules, certificate auth, permission validation, and per-permission detail) and emits a structured `[pscustomobject]` per step. No CSV processing, no OneDrive writes. |
| `Run` | Credentials supplied without `-Test` | Normal migration: reads CSV, optionally uploads files, grants permissions per row. |

`-UploadFiles` and `-AllowedDomains` are only valid in the `Run` set.

### Configuration file (config.json)

When `Update-UserFile.ps1` starts (in `Run` or `Test` mode) it looks for a JSON file at the path supplied via `-ConfigFile` (default: `config.json` next to the script). Any keys present in that file become **defaults** for the matching parameters — but only for parameters the caller did not pass explicitly on the command line.

**Recognised keys:**

```json
{
    "InputFile":             "C:\\Repos\\NPSBox\\UserInfo.csv",
    "TenantId":              "<tenant-guid>",
    "ClientId":              "<app-guid>",
    "CertificateThumbprint": "<sha1-thumbprint>",
    "LogFolder":             "C:\\Repos\\NPSBox\\Logs",
    "AllFilesDirectory":     "C:\\Repos\\NPSBox\\LocalFiles"
}
```

**Behaviour:**

- **Precedence:** explicit parameter &gt; `config.json` &gt; environment variable / hard-coded default.
- **Missing file:** silently ignored — the script falls back to env vars and built-in defaults. A verbose message records the path that was checked.
- **Empty or whitespace values:** ignored (so you can keep blank placeholders in a checked-in template).
- **Unrecognised keys:** ignored.
- **Malformed JSON:** a warning is written and the script continues with env vars / defaults.

> **Security note:** `config.json` contains tenant/client/thumbprint identifiers. Treat it the same way you would treat `Set-NPSBoxEnv.ps1` inputs — add it to `.gitignore` (or commit only a sanitised template) before checking changes in.

### Box role mapping

The script maps Box permissions to [Microsoft Graph sharing roles](https://learn.microsoft.com/graph/api/resources/permission?view=graph-rest-1.0#roles-property-values):

| Box Permission | Graph Role | Effect |
|---|---|---|
| `Co-owner` | `write` | View + edit access |
| `Editor` | `write` | View + edit access |
| `Viewer Uploader` | `read` | View-only access (upload aspect not supported in Graph) |
| `Viewer` | `read` | View-only access |
| `Previewer Uploader` | *(skipped)* | No Graph equivalent |
| `Previewer` | *(skipped)* | No Graph equivalent |
| `Uploader` | *(skipped)* | No Graph equivalent |

Rows with unmapped roles are skipped with a log entry.

### Authentication

`Update-UserFile.ps1` uses certificate-based app-only authentication via `Connect-GraphCertAuth` (an internal helper function). The app registration must have **Application** permissions `Files.ReadWrite.All` and `User.Read.All` granted with admin consent.

> **Important — `Connect-Graph` alias collision (fixed in v1.2.0.17):**
> The `Microsoft.Graph.Authentication` module exports `Connect-Graph` as an **alias** for `Connect-MgGraph`. PowerShell resolves aliases before functions, so any script function named `Connect-Graph` is silently shadowed after the module is imported. The bare alias invokes `Connect-MgGraph` with no parameters, which falls back to **delegated** auth using the default Microsoft Graph Command Line Tools app (`14d82eec-...`) instead of certificate-based app-only auth. This caused all Graph API calls to fail with `Authorization_RequestDenied` (insufficient privileges). The fix was renaming the script's function to `Connect-GraphCertAuth` to avoid the collision.

The script reuses an existing Graph session if the `TenantId` and `AuthType` match, avoiding redundant authentication.

- [Connect-MgGraph reference](https://learn.microsoft.com/powershell/module/microsoft.graph.authentication/connect-mggraph)
- [Graph authentication overview](https://learn.microsoft.com/powershell/microsoftgraph/authentication-commands)

### Usage examples

#### 1) Dry-run (recommended first)

```powershell
$params = @{
    InputFile             = '.\UserInfo.csv'
    UserToProcess         = 'user@contoso.com'
    TenantId              = '<tenant-guid>'
    ClientId              = '<app-guid>'
    CertificateThumbprint = '<thumbprint>'
    Verbose               = $true
    WhatIf                = $true
}
.\Update-UserFile.ps1 @params
```

#### 2) Apply permissions for a specific user

```powershell
$params = @{
    InputFile             = '.\UserInfo.csv'
    UserToProcess         = 'user@contoso.com'
    TenantId              = '<tenant-guid>'
    ClientId              = '<app-guid>'
    CertificateThumbprint = '<thumbprint>'
    Verbose               = $true
}
.\Update-UserFile.ps1 @params
```

#### 3) Upload files and apply permissions

```powershell
$params = @{
    UserToProcess         = 'user@contoso.com'
    TenantId              = '<tenant-guid>'
    ClientId              = '<app-guid>'
    CertificateThumbprint = '<thumbprint>'
    UploadFiles           = $true
    Verbose               = $true
}
.\Update-UserFile.ps1 @params
```

#### 4) Export results to CSV

```powershell
$params = @{
    InputFile             = '.\UserInfo.csv'
    UserToProcess         = 'user@contoso.com'
    TenantId              = '<tenant-guid>'
    ClientId              = '<app-guid>'
    CertificateThumbprint = '<thumbprint>'
    Verbose               = $true
}
.\Update-UserFile.ps1 @params | Export-Csv -Path '.\permission_results.csv' -NoTypeInformation
```

#### 5) Process multiple users via pipeline

```powershell
$params = @{
    InputFile             = '.\UserInfo.csv'
    TenantId              = '<tenant-guid>'
    ClientId              = '<app-guid>'
    CertificateThumbprint = '<thumbprint>'
    Verbose               = $true
}
'user1@contoso.com', 'user2@contoso.com' | .\Update-UserFile.ps1 @params
```

#### 6) Process all users in the CSV

```powershell
$params = @{
    InputFile             = '.\UserInfo.csv'
    TenantId              = '<tenant-guid>'
    ClientId              = '<app-guid>'
    CertificateThumbprint = '<thumbprint>'
    Verbose               = $true
}
.\Update-UserFile.ps1 @params
```

### File upload

When `-UploadFiles` is specified, the script uploads local files from `<AllFilesDirectory>\<UserToProcess>\` to the user's OneDrive root before applying permissions.

- Folders are created first (parent before child, sorted by path depth) via `POST /drives/{driveId}/root:/{parent}:/children` (or `/root/children` at the root) using the [documented `driveItem: post children` API](https://learn.microsoft.com/graph/api/driveitem-post-children?view=graph-rest-1.0) with `@microsoft.graph.conflictBehavior = fail`. A `409 nameAlreadyExists` response is treated as success so reruns are safe.
- **Small files (≤ 4 MB)** are uploaded via `PUT /drives/{driveId}/root:/{path}:/content` ([simple upload API](https://learn.microsoft.com/graph/api/driveitem-put-content?view=graph-rest-1.0)).
- **Large files (&gt; 4 MB)** are uploaded via a [resumable upload session](https://learn.microsoft.com/graph/api/driveitem-createuploadsession?view=graph-rest-1.0):
  1. `POST .../createUploadSession` returns a short-lived, pre-authenticated `uploadUrl`.
  2. The file is streamed to that URL in 10 MiB chunks (a multiple of the required 320 KiB block size) with `Content-Range: bytes <start>-<end>/<total>` headers.
  3. Chunk PUTs use `Invoke-WebRequest` (not `Invoke-MgGraphRequest`) so they do **not** add a Graph bearer token to the pre-authenticated upload URL.
  4. Each chunk goes through the same `Invoke-WithGraphRetry` exponential-backoff path used for other Graph calls.
- `@microsoft.graph.conflictBehavior` is set to `replace` on the upload session so reruns overwrite the existing file.
- Use `-WhatIf` to preview what would be uploaded without making changes.

**Directory structure example:**

```
LocalFiles/
  user@contoso.com/
    Documents/
      Report.pdf
    Photos/
      vacation.jpg
```

### Output objects

The script emits one `[pscustomobject]` per CSV row with the following properties:

| Property | Description |
|---|---|
| `OwnerLogin` | Owner UPN from the CSV |
| `ItemName`, `Path`, `NormalizedPath` | Original CSV path/name and the OneDrive-relative path after stripping `All Files/` and re-encoding |
| `CollaboratorLogin`, `CollaboratorPermission` | Source values from the CSV |
| `GraphRole` | Mapped Graph role (`read` / `write`) or `$null` when the Box role has no equivalent |
| `DriveId`, `OneDriveWebUrl` | Resolved drive identifier and the owner's OneDrive web URL |
| `IsValidAccount` | `$true` when `Get-MgUser` confirmed the owner exists and is enabled |
| `OneDriveProvisioned` | `$true` when the owner's OneDrive drive root was reachable |
| `ExistsInOneDrive`, `DriveItemId` | Whether the target item was found by path, and its driveItem ID |
| `Action` | `Granted`, `WhatIf`, `Skipped`, etc. |
| `Status` | `Applied`, `WhatIf`, `Skipped`, or `Failed` |
| `Error` | Error message when `Status = 'Failed'`; otherwise `$null` |

When `-UploadFiles` is used, additional objects are emitted by `Invoke-OneDriveUpload` describing each folder/file operation (`OwnerLogin`, `LocalPath`, `OneDrivePath`, `ItemType`, `SizeBytes`, `Action`, `Status`, `DurationSec`, `RateMBps`, `Error`). For `UploadFile` results, `Status` is one of `Completed`, `Incomplete`, or `WhatIf`; `DurationSec` is the elapsed upload time in seconds (one decimal) and `RateMBps` is the throughput in MB/sec (one decimal). Both are `0.0` for `WhatIf` previews.

When `-Test` is used, the output is a per-step validation report (`Step`, `Status`, `Detail`) plus per-permission rows from `Get-AppPermissionDetail` and a final `Overall` summary row.

### Logging

Each run creates a timestamped log file in `-LogFolder`:

```
Update-UserFile_20260429_143022_123.log
```

Log entries include ISO 8601 timestamps, severity levels (`INFO`, `WARN`, `ERROR`), and contextual messages. Log files are written even during `-WhatIf` runs to ensure an audit trail. Use `-Verbose` to also see log lines in the console.

### Retry and throttle handling

The script wraps all Graph API calls in `Invoke-WithGraphRetry` with automatic retry and exponential backoff:

- **Retried errors:** HTTP 429 (throttling), 500, 502, 503, 504, timeouts, and transient failures.
- **Not retried:** HTTP 401, 403, 404, and other permanent errors.
- **`Retry-After` header honored:** When Graph returns a 429 with a `Retry-After` value, the script waits that exact duration before retrying.
- **Defaults:** up to 6 attempts, starting at a 2-second delay, doubling each retry up to a 60-second maximum.

See [Microsoft Graph throttling guidance](https://learn.microsoft.com/graph/throttling) for details.

### Known limitations

- **Simple vs. resumable upload:** Files ≤ 4 MB use the single-request `PUT /content` endpoint; larger files automatically switch to a resumable upload session (10 MiB chunks). The chunk size is hard-coded at the moment — expose it as a parameter if you need to tune it for very slow or unreliable links.
- **Path translation:** Box paths prefixed with `All Files/` are stripped automatically (case-insensitive). If your migration tool renamed items or altered folder hierarchy, paths may not resolve correctly.
- **Box path encoding:** Special characters (spaces, parentheses, etc.) are URL-encoded per segment for the Graph API. Highly unusual characters may still cause issues.
- **Unmapped roles:** Permissions with no Graph equivalent (`Previewer`, `Uploader`, `Previewer Uploader`) and any unrecognized role string are skipped with a log entry.
- **No email notifications:** All permission grants use `sendInvitation = false`.
- **No idempotency check on permissions:** Re-running re-invokes the invite API for every row without checking existing roles. Granting a different role to an existing collaborator may create a second permission entry. (Folder creation *is* idempotent — a `409 nameAlreadyExists` is treated as success.)
- **No Graph batching:** Each permission grant is a separate HTTP request. Graph supports batching up to 20 requests per call (future enhancement).
- **OneDrive must be pre-provisioned:** If the owner's OneDrive site has not been created, the script fails for that owner with guidance to run `Request-SPOPersonalSite` or have the user sign in at https://portal.office.com. It does not provision OneDrive automatically.
- **Email format validation is heuristic:** `Test-EmailFormat` uses a basic regex (`local@domain.tld`); it is not a full RFC 5322 implementation.
- **Permission auto-validation requires `Application.Read.All`:** Without this permission, `Assert-GraphPermissions` reports `Skipped` and the script proceeds; you lose the pre-flight safety check.
- **Hardcoded default paths:** `-InputFile`, `-LogFolder`, and `-AllFilesDirectory` default to absolute paths under `C:\Repos\NPSBox\`. Override them explicitly when running from a different working directory.
- **PnP.PowerShell conflict:** Do not import `PnP.PowerShell` in the same session. It loads `Microsoft.Graph.Core` 1.x, which is incompatible with Graph SDK v2. The script detects this at startup via `Assert-GraphAssemblyCompatibility` and refuses to run.
- **`Connect-Graph` alias collision:** The `Microsoft.Graph.Authentication` module exports `Connect-Graph` as an alias. Do not name custom functions `Connect-Graph` — use `Connect-GraphCertAuth` or a similar unique name to avoid silent delegation fallback.
- Always run with `-Test` (and then `-WhatIf`) first in a new tenant or against new datasets.

---

## Set-NPSBoxEnv.ps1

A small helper that persists three user-scope environment variables consumed as parameter defaults by `Update-UserFile.ps1`:

| Variable | Used by parameter |
|---|---|
| `NPSBOX_TENANT_ID` | `-TenantId` |
| `NPSBOX_CLIENT_ID` | `-ClientId` |
| `NPSBOX_CERT_THUMBPRINT` | `-CertificateThumbprint` |

```powershell
.\Set-NPSBoxEnv.ps1 -TenantId '<tenant-guid>' -ClientId '<app-guid>' -CertificateThumbprint '<thumbprint>'
```

Values are written via `[Environment]::SetEnvironmentVariable(..., 'User')` so they survive across PowerShell sessions, and are also set in the current session so they take effect immediately. To clear a value, set it to `$null` at the `User` scope (see the script's `.NOTES` block).

> **Security note:** These variables are stored in plaintext in the user's environment block. Do not use this helper on shared workstations.

---

## Pre-flight validation (`-Test`)

Run the main script with `-Test` to verify the environment without touching any user data:

```powershell
.\Update-UserFile.ps1 -Test -TenantId '<tenant-guid>' -ClientId '<app-guid>' -CertificateThumbprint '<thumbprint>' -Verbose
```

The script reports a structured result for each step:

1. **Assembly compatibility** — confirms `Microsoft.Graph.Core` v3+ and that `PnP.PowerShell` is not loaded.
2. **Required modules** — `Microsoft.Graph.Authentication`, `.Users`, `.Files`, `.Applications` are present and importable.
3. **Certificate authentication** — `Connect-MgGraph` succeeds with the supplied thumbprint and tenant.
4. **Permission validation** — required Graph application permissions are admin-consented (or `Skipped` if `Application.Read.All` is missing).
5. **Permission detail** — per-permission `IsGranted` rows (same data as the archived `Test-AzureAppRegistration.ps1` validator).
6. **Overall** — summary row with `Passed` or `Passed with warnings`.

---

## GitHub Copilot customization (`.github/`)

The `.github/` directory contains [VS Code / GitHub Copilot customization files](https://code.visualstudio.com/docs/copilot/copilot-customization) that tailor AI-assisted development to this project's conventions.

| Path | Purpose |
|---|---|
| `copilot-instructions.md` | **Global Copilot instructions** — loaded into every Copilot conversation. Defines the project role, key files, architecture notes, versioning convention, coding standards, API reference requirements, and validation workflow. |
| `instructions/csv-data.instructions.md` | **CSV data schema** — applied automatically when editing `*.csv` files. Documents column schemas, Box-role-to-Graph-permission mapping, and path normalization conventions. |
| `instructions/pester-tests.instructions.md` | **Pester test conventions** — applied automatically when editing files under `tests/`. Covers Pester 5 test structure, `New-CsvRow` helper, mocking conventions, and assertion patterns. |
| `instructions/powershell.instructions.md` | **PowerShell coding standards** — applied automatically when editing `*.ps1` files. Enforces Allman-style braces, `SupportsShouldProcess`, `Write-LogLine` logging, pipeline output conventions, and error-handling patterns. |
| `agents/graph-api.agent.md` | **Graph API research agent** — read-only specialist for researching Microsoft Graph REST endpoints, SDK cmdlets, permissions, throttling, and error codes. |
| `prompts/add-graph-operation.prompt.md` | **Reusable prompt** — scaffolds a new Graph API operation with retry logic, URL encoding, `ShouldProcess` guards, logging, and matching Pester tests. |
| `prompts/learn-project.prompt.md` | **Reusable prompt** — onboarding walkthrough that summarizes the project layout, conventions, and key files. |
| `prompts/run-tests.prompt.md` | **Reusable prompt** — runs the Pester 5 test suite and reports a pass/fail summary with failure details. |
| `prompts/version-bump.prompt.md` | **Reusable prompt** — bumps the script version (revision, minor, or major) and date in the `.SYNOPSIS` block. |

---

## Disclaimer

This repository includes [`DISCLAIMER.MD`](DISCLAIMER.MD). Please review it before using or modifying the scripts.

---

## Assumptions & constraints

See [`assumptions.md`](assumptions.md) for a comprehensive audit of security, performance, throttling, data handling, infrastructure, file upload, logging, and permission mapping assumptions.

---

## Changelog

| Version | Date | Summary |
|---|---|---|
| **1.2.2.4** | 2026-05-07 | README rewritten to match current functionality: documented `-Test` mode and Help mode (no parameters), added `Set-NPSBoxEnv.ps1` section, corrected parameter table (TenantId/ClientId are Mandatory; defaults come from `NPSBOX_*` env vars; absolute default paths), added Output objects section, updated File upload to reflect documented `POST /children` with `conflictBehavior=fail`, refreshed Known limitations, updated test count to 87 active tests for `Update-UserFile.ps1`. |
| 1.2.1.0 | 2026-05-05 | Minor version bump. Added 11 new Pester tests covering: Viewer Uploader/Previewer Uploader/Uploader/unknown permission mapping, path edge cases (empty after prefix removal, whitespace-only, case-variant `ALL FILES`), and transient error classification (403 not retried, 404 not retried, 500 retried, timeout retried). |
| 1.2.0.17 | 2026-05-05 | **Critical fix:** Renamed `Connect-Graph` → `Connect-GraphCertAuth` to avoid alias collision with `Microsoft.Graph.Authentication` module. The module exports `Connect-Graph` as an alias for `Connect-MgGraph`; PowerShell resolves aliases before functions, causing the script's auth function to be silently bypassed. This resulted in delegated auth (wrong ClientId, insufficient permissions) instead of certificate-based app-only auth. Uploads and permission grants failed with `Authorization_RequestDenied`. |
| 1.2.0.16 | 2026-04-29 | Previous release |
