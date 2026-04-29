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
| **File upload** | Not supported — permissions only. | Supports uploading local files to OneDrive via the [simple upload API](https://learn.microsoft.com/graph/api/driveitem-put-content?view=graph-rest-1.0) (≤ 4 MB). |
| **Retry / throttling** | No built-in retry. PnP cmdlets surfaced raw 429 errors. | `Invoke-WithGraphRetry` with exponential backoff, `Retry-After` header parsing, and configurable max attempts. |
| **Testability** | CSOM calls were difficult to mock in Pester. | All Graph calls go through `Invoke-MgGraphRequest`, which is straightforward to mock. Full Pester 5 test suite (57 tests). |

The legacy PnP scripts (`Set-BoxToOneDriveItemPermission.ps1`, `Connect-SharePointOnlineAdmin.ps1`, `Add-SCA.ps1`, `Add-AppCertificateCredential.ps1`, `Get-UserLocalFiles.ps1`) have been removed from the active project. Archived sample data is in `holding/`.

---

## Table of contents

- [Repository layout](#repository-layout)
- [Input CSV format](#input-csv-format)
- [Update-UserFile.ps1](#update-userfileps1)
  - [What it does](#what-it-does)
  - [Prerequisites](#prerequisites)
  - [Parameters](#parameters)
  - [Box role mapping](#box-role-mapping)
  - [Authentication](#authentication)
  - [Usage examples](#usage-examples)
  - [File upload](#file-upload)
  - [Logging](#logging)
  - [Retry and throttle handling](#retry-and-throttle-handling)
  - [Known limitations](#known-limitations)
- [Test-AzureAppRegistration.ps1](#test-azureappregistrationps1)
- [GitHub Copilot customization (.github/)](#github-copilot-customization-github)
- [Disclaimer](#disclaimer)
- [Assumptions & constraints](#assumptions--constraints)

---

## Repository layout

| Path | Purpose |
|---|---|
| `Update-UserFile.ps1` | Main script — resolves OneDrive items, uploads files, and grants permissions via Microsoft Graph |
| `Test-AzureAppRegistration.ps1` | Validates that the Azure AD app registration has the required Graph permissions (`Files.ReadWrite.All`, `User.Read.All`) |
| `UserInfo.csv` | Input CSV — Box collaboration data to process |
| `tests/Update-UserFile.Tests.ps1` | Pester 5 test suite for `Update-UserFile.ps1` (57 tests; all Graph/file-system calls are mocked) |
| `tests/Test-AzureAppRegistration.Tests.ps1` | Pester 5 test suite for `Test-AzureAppRegistration.ps1` |
| `tests/results/` | Pester test result output (NUnit XML), named by version |
| `LocalFiles/` | Per-user subfolders (named by UPN) containing files to upload with `-UploadFiles` |
| `logs/` | Timestamped log output (git-ignored) |
| `holding/` | Archived legacy files (`Box_Collaboration_Sample_Data.csv`, `setupcert.txt`) |
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
5. **Resolves each item** by path in OneDrive using the [`driveItem: get` API](https://learn.microsoft.com/graph/api/driveitem-get?view=graph-rest-1.0#access-a-driveitem-by-path).
6. **Validates** the collaborator's domain against `-AllowedDomains` to prevent unintended external sharing.
7. **Grants permissions** to each collaborator using the [`driveItem: invite` API](https://learn.microsoft.com/graph/api/driveitem-invite?view=graph-rest-1.0) with `sendInvitation = false` (no email notifications) and `requireSignIn = true`.
8. **Outputs** a structured `[pscustomobject]` result per CSV row to the pipeline for reporting.

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
  - [Register an application](https://learn.microsoft.com/entra/identity-platform/quickstart-register-app)
- A certificate uploaded to the app registration — [Certificate credentials](https://learn.microsoft.com/entra/identity-platform/certificate-credentials)
- Run `Test-AzureAppRegistration.ps1` to verify permissions before first use.

> **Important:** Do not load `PnP.PowerShell` in the same session. PnP.PowerShell bundles `Microsoft.Graph.Core` 1.x, which conflicts with the Graph SDK v2. Start a fresh `pwsh` session.

### Parameters

| Parameter | Type | Required | Default | Description |
|---|---|---|---|---|
| `-InputFile` | `FileInfo` | No | `UserInfo.csv` | Path to the input CSV file |
| `-UserToProcess` | `string` | No | — | Owner UPN to filter on. When omitted, all unique owners are processed. Accepts pipeline input. Aliases: `Owner Login`, `User`, `UPN`, `Account` |
| `-TenantId` | `string` | No | — | Microsoft 365 tenant ID (GUID). [Find your tenant ID](https://learn.microsoft.com/entra/fundamentals/how-to-find-tenant) |
| `-ClientId` | `string` | No | — | App registration client ID (GUID) |
| `-CertificateThumbprint` | `string` | **Yes** | — | SHA-1 thumbprint of a certificate in `Cert:\CurrentUser\My` |
| `-LogFolder` | `string` | No | `.\Logs` | Folder for timestamped log files |
| `-AllFilesDirectory` | `string` | No | `.\LocalFiles` | Root folder containing per-user subfolders of local files to upload |
| `-UploadFiles` | `switch` | No | — | When present, uploads local files to OneDrive before applying permissions |
| `-AllowedDomains` | `string[]` | No | — | List of allowed email domains for collaborators. Blocks external sharing to unlisted domains. |

Common parameters `-WhatIf`, `-Confirm`, and `-Verbose` are fully supported.

### Box role mapping

The script maps Box permissions to [Microsoft Graph sharing roles](https://learn.microsoft.com/graph/api/resources/permission?view=graph-rest-1.0#roles-property-values):

| Box Permission | Graph Role | Effect |
|---|---|---|
| `Co-owner` | `write` | View + edit access |
| `Editor` | `write` | View + edit access |
| `Viewer Uploader` | `write` | View + edit access |
| `Viewer` | `read` | View-only access |
| `Previewer Uploader` | *(skipped)* | No Graph equivalent |
| `Previewer` | *(skipped)* | No Graph equivalent |
| `Uploader` | *(skipped)* | No Graph equivalent |

Rows with unmapped roles are skipped with a log entry.

### Authentication

`Update-UserFile.ps1` uses certificate-based app-only authentication. The app registration
must have **Application** permissions `Files.ReadWrite.All` and `User.Read.All` granted with admin consent.

The script reuses an existing Graph session if the `TenantId` matches, avoiding redundant authentication.

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

- Folders are created first (parent before child, sorted by path depth) via `PATCH` with a folder body and `@microsoft.graph.conflictBehavior = replace`.
- Files are uploaded via `PUT /drives/{driveId}/root:/{path}:/content` ([simple upload API](https://learn.microsoft.com/graph/api/driveitem-put-content?view=graph-rest-1.0)).
- Files larger than 4 MB are rejected with a clear error message. A [resumable upload session](https://learn.microsoft.com/graph/api/driveitem-createuploadsession?view=graph-rest-1.0) would be required for larger files (not implemented).
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

- **Simple upload only:** Files larger than 4 MB require a resumable upload session, which is not implemented.
- **Path translation:** Box paths prefixed with `All Files/` are stripped automatically. If your migration tool renamed items or altered folder hierarchy, paths may not resolve correctly.
- **Box path encoding:** Special characters (spaces, parentheses, etc.) are URL-encoded per segment for the Graph API. Unusual characters may still cause issues.
- **Unmapped roles:** Permissions with no Graph equivalent (`Previewer`, `Uploader`, `Previewer Uploader`) are skipped.
- **No email notifications:** All permission grants use `sendInvitation = false`.
- **No idempotency check:** Re-running re-grants every permission without checking existing roles. Granting a different role to an existing collaborator may create a second permission entry.
- **No Graph batching:** Each permission grant is a separate HTTP request. Graph supports batching up to 20 requests (future enhancement).
- **PnP.PowerShell conflict:** Do not import `PnP.PowerShell` in the same session. It loads `Microsoft.Graph.Core` 1.x, which is incompatible with Graph SDK v2. Start a fresh `pwsh` session.
- Always run with `-WhatIf` first in a new tenant or against new datasets.

---

## Test-AzureAppRegistration.ps1

A validation script that connects to Microsoft Graph using the same certificate-based auth and verifies the app registration has the required permissions (`Files.ReadWrite.All`, `User.Read.All`) with admin consent. Run this before your first migration to catch permission issues early.

```powershell
.\Test-AzureAppRegistration.ps1 -TenantId '<tenant-guid>' -ClientId '<app-guid>' -CertificateThumbprint '<thumbprint>' -Verbose
```

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
| `prompts/run-tests.prompt.md` | **Reusable prompt** — runs the Pester 5 test suite and reports a pass/fail summary with failure details. |
| `prompts/version-bump.prompt.md` | **Reusable prompt** — bumps the script version (revision, minor, or major) and date in the `.SYNOPSIS` block. |

---

## Disclaimer

This repository includes [`DISCLAIMER.MD`](DISCLAIMER.MD). Please review it before using or modifying the scripts.

---

## Assumptions & constraints

See [`assumptions.md`](assumptions.md) for a comprehensive audit of security, performance, throttling, data handling, infrastructure, file upload, logging, and permission mapping assumptions.
