# NPSBox — Migrate Box collaboration permissions to OneDrive for Business

This repository contains PowerShell scripts that **translate Box collaboration/export data** into **OneDrive for Business item-level permissions**.

Two approaches are provided:

| Script | API surface | Best for |
|---|---|---|
| `Update-UserFile.ps1` | **Microsoft Graph** (PowerShell SDK) | New deployments, file uploads + permissions, unattended cert auth |
| `Set-BoxToOneDriveItemPermission.ps1` | **PnP PowerShell** (SharePoint CSOM) | Existing PnP-based workflows, legacy environments |

> **Note:** This is sample/illustration code. See `DISCLAIMER.MD`.

---

## Table of contents

- [Repository layout](#repository-layout)
- [Input CSV format](#input-csv-format)
- [Update-UserFile.ps1 (Microsoft Graph)](#update-userfileps1-microsoft-graph)
  - [What it does](#what-it-does)
  - [Prerequisites](#prerequisites)
  - [Parameters](#parameters)
  - [Box role mapping (Graph)](#box-role-mapping-graph)
  - [Authentication](#authentication)
  - [Usage examples](#usage-examples)
  - [File upload](#file-upload)
  - [Logging](#logging)
  - [Retry and throttle handling](#retry-and-throttle-handling)
  - [Known limitations](#known-limitations)
- [Set-BoxToOneDriveItemPermission.ps1 (PnP PowerShell)](#set-boxtoonedriveitemPermissionps1-pnp-powershell)
  - [What it does (PnP)](#what-it-does-pnp)
  - [Requirements (PnP)](#requirements-pnp)
  - [Box role mapping (PnP)](#box-role-mapping-pnp)
  - [Usage (PnP)](#usage-pnp)
  - [Authentication (PnP)](#authentication-pnp)
  - [Logging (PnP)](#logging-pnp)
  - [Known limitations (PnP)](#known-limitations-pnp)
- [Disclaimer](#disclaimer)

---

## Repository layout

| Path | Purpose |
|---|---|
| `Update-UserFile.ps1` | Main script — Microsoft Graph approach (resolve items, upload files, set permissions) |
| `Set-BoxToOneDriveItemPermission.ps1` | Main script — PnP PowerShell approach (resolve items + set permissions) |
| `Connect-SharePointOnlineAdmin.ps1` | Helper to connect to SPO admin center via PnP |
| `Add-AppCertificateCredential.ps1` | Helper to add a certificate credential to an app registration |
| `Add-SCA.ps1` | Helper to add a SharePoint site collection administrator |
| `Get-UserLocalFiles.ps1` | Utility to enumerate local per-user files |
| `Box_Collaboration_Sample_Data.csv` | Sample CSV structure |
| `DISCLAIMER.MD` | Disclaimer text |
| `tests/` | Pester tests |
| `tests/results/` | Pester test result output (NUnit XML) |
| `releases/`, `logs/`, `LocalFiles/` | Local/release assets |
| `alternatives/` | Alternative script implementations |
| `.github/` | GitHub Copilot customization files (instructions, agents, prompts) |

### `.github/` folder

The `.github/` directory contains [VS Code / GitHub Copilot customization files](https://code.visualstudio.com/docs/copilot/copilot-customization) that tailor AI-assisted development to this project's conventions.

| Path | Purpose |
|---|---|
| `.github/copilot-instructions.md` | **Global Copilot instructions** — loaded into every Copilot conversation for this repo. Defines the project role, key files, architecture notes, versioning convention, coding standards, API reference requirements, and validation workflow. |
| `.github/instructions/csv-data.instructions.md` | **CSV data schema instructions** — applied automatically when editing `*.csv` files. Documents the `UserInfo.csv` / `Box_Collaboration_Sample_Data.csv` column schemas, Box-role-to-Graph-permission mapping, and path normalization conventions. |
| `.github/instructions/pester-tests.instructions.md` | **Pester test instructions** — applied automatically when editing files under `tests/`. Covers Pester 5 test structure, the `New-CsvRow` helper, mocking conventions (all Graph/file-system calls are mocked), script invocation patterns, and assertion best practices. |
| `.github/instructions/powershell.instructions.md` | **PowerShell coding standards** — applied automatically when editing `*.ps1` files. Enforces Allman-style braces, `SupportsShouldProcess`, `Write-LogLine` logging, pipeline output conventions, Graph API URL encoding, and error-handling patterns. |
| `.github/agents/graph-api.agent.md` | **Graph API research agent** — a read-only specialist agent for researching Microsoft Graph REST endpoints, PowerShell SDK cmdlets, permissions, throttling, and error codes. Uses only official Microsoft documentation. |
| `.github/prompts/add-graph-operation.prompt.md` | **Reusable prompt** — scaffolds a new Graph API operation in `Update-UserFile.ps1` with retry logic, URL encoding, `ShouldProcess` guards, logging, and matching Pester tests. |
| `.github/prompts/run-tests.prompt.md` | **Reusable prompt** — runs the Pester 5 test suite and reports a pass/fail summary with failure details and fix suggestions. |
| `.github/prompts/version-bump.prompt.md` | **Reusable prompt** — bumps the script version (revision, minor, or major) and date in the `.SYNOPSIS` block following the project's versioning convention. |

---

## Input CSV format

Both scripts consume the same CSV structure. Required columns:

| Column | Description | Example |
|---|---|---|
| `Owner Login` | UPN of the file owner | `user@contoso.com` |
| `Path` | Box path to the item | `All Files/Documents` |
| `Item Name` | File or folder name | `Report.pdf` |
| `Collaborator Login` | UPN of the person to share with | `colleague@contoso.com` |
| `Collaborator Permission` | Box role | `Editor`, `Viewer`, `Co-owner` |

See `Box_Collaboration_Sample_Data.csv` for a working example.

---

## Update-UserFile.ps1 (Microsoft Graph)

### What it does

For a given Box owner (user), the script:

1. **Authenticates** to Microsoft Graph using certificate-based app-only auth.
2. **Reads the CSV** and identifies unique owners to process.
   If `-UserToProcess` is specified, only that user is processed;
   otherwise all unique owners in the CSV are processed.
3. **Looks up** the user's OneDrive drive via the [Graph API](https://learn.microsoft.com/graph/overview).
4. **Optionally uploads** local files and folders to the user's OneDrive (when `-UploadFiles` is specified).
5. **Resolves each item** by path in OneDrive using the [`driveItem: get` API](https://learn.microsoft.com/graph/api/driveitem-get?view=graph-rest-1.0#access-a-driveitem-by-path).
6. **Grants permissions** to each collaborator using the [`driveItem: invite` API](https://learn.microsoft.com/graph/api/driveitem-invite?view=graph-rest-1.0) with `sendInvitation = false` (no email notifications).
7. **Outputs** a structured result object per CSV row to the pipeline for reporting.

### Prerequisites

- **PowerShell 7.0 or later** — [Install PowerShell](https://learn.microsoft.com/powershell/scripting/install/installing-powershell)
- **Microsoft Graph PowerShell SDK** modules:

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module Microsoft.Graph.Users          -Scope CurrentUser
Install-Module Microsoft.Graph.Files          -Scope CurrentUser
```

- [Microsoft Graph PowerShell SDK installation guide](https://learn.microsoft.com/powershell/microsoftgraph/installation)
- **Azure AD App Registration** with the following **Application** permission
  granted with admin consent:
  - `Files.ReadWrite.All`
  - [Register an application](https://learn.microsoft.com/entra/identity-platform/quickstart-register-app)
- A certificate uploaded to the app registration — [Certificate credentials](https://learn.microsoft.com/entra/identity-platform/certificate-credentials)

> **Important:** Do not load `PnP.PowerShell` in the same session as `Update-UserFile.ps1`. PnP.PowerShell bundles `Microsoft.Graph.Core` 1.x, which conflicts with the Graph SDK v2. Start a fresh `pwsh` session.

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

Common parameters `-WhatIf`, `-Confirm`, and `-Verbose` are fully supported.

### Box role mapping (Graph)

The script maps Box permissions to [Microsoft Graph sharing roles](https://learn.microsoft.com/graph/api/resources/permission?view=graph-rest-1.0#roles-property-values):

| Box Permission | Graph Role | Effect |
|---|---|---|
| `Co-owner` | `write` | View + edit access |
| `Editor` | `write` | View + edit access |
| `Viewer Uploader` | `read` | View-only access |
| `Viewer` | `read` | View-only access |
| `Previewer Uploader` | *(skipped)* | No Graph equivalent |
| `Previewer` | *(skipped)* | No Graph equivalent |
| `Uploader` | *(skipped)* | No Graph equivalent |

Rows with unmapped roles are skipped with a log entry.

### Authentication

`Update-UserFile.ps1` uses certificate-based app-only authentication. The app registration
must have the **Application** permission `Files.ReadWrite.All` granted with admin consent.

The script reuses an existing Graph session if the `TenantId` matches, avoiding redundant authentication.

- [Connect-MgGraph reference](https://learn.microsoft.com/powershell/module/microsoft.graph.authentication/connect-mggraph)
- [Graph authentication overview](https://learn.microsoft.com/powershell/microsoftgraph/authentication-commands)

### Usage examples

#### 1) Dry-run (recommended first)

```powershell
$params = @{
    InputFile            = '.\Box_Collaboration_Sample_Data.csv'
    UserToProcess        = 'user@contoso.com'
    TenantId             = '<tenant-guid>'
    ClientId             = '<app-guid>'
    CertificateThumbprint = '<thumbprint>'
    Verbose              = $true
    WhatIf               = $true
}
.\Update-UserFile.ps1 @params
```

#### 2) Apply permissions for a specific user

```powershell
$params = @{
    InputFile            = '.\Box_Collaboration_Sample_Data.csv'
    UserToProcess        = 'user@contoso.com'
    TenantId             = '<tenant-guid>'
    ClientId             = '<app-guid>'
    CertificateThumbprint = '<thumbprint>'
    Verbose              = $true
}
.\Update-UserFile.ps1 @params
```

#### 3) Upload files and apply permissions

```powershell
$params = @{
    UserToProcess        = 'user@contoso.com'
    TenantId             = '<tenant-guid>'
    ClientId             = '<app-guid>'
    CertificateThumbprint = '<thumbprint>'
    UploadFiles          = $true
    Verbose              = $true
}
.\Update-UserFile.ps1 @params
```

#### 4) Export results to CSV

```powershell
$params = @{
    InputFile            = '.\Box_Collaboration_Sample_Data.csv'
    UserToProcess        = 'user@contoso.com'
    TenantId             = '<tenant-guid>'
    ClientId             = '<app-guid>'
    CertificateThumbprint = '<thumbprint>'
    Verbose              = $true
}
.\Update-UserFile.ps1 @params | Export-Csv -Path '.\permission_results.csv' -NoTypeInformation
```

#### 5) Process multiple users via pipeline

```powershell
$params = @{
    InputFile            = '.\Box_Collaboration_Sample_Data.csv'
    TenantId             = '<tenant-guid>'
    ClientId             = '<app-guid>'
    CertificateThumbprint = '<thumbprint>'
    Verbose              = $true
}
'user1@contoso.com', 'user2@contoso.com' | .\Update-UserFile.ps1 @params
```

#### 6) Process all users in the CSV

```powershell
$params = @{
    InputFile            = '.\Box_Collaboration_Sample_Data.csv'
    TenantId             = '<tenant-guid>'
    ClientId             = '<app-guid>'
    CertificateThumbprint = '<thumbprint>'
    Verbose              = $true
}
.\Update-UserFile.ps1 @params
```

### File upload

When `-UploadFiles` is specified, the script uploads local files from `<AllFilesDirectory>\<UserToProcess>\` to the user's OneDrive root before applying permissions.

- Folders are created first (parent before child) via `PATCH` with a folder body and `@microsoft.graph.conflictBehavior = replace`.
- Files are uploaded via `PUT /drives/{driveId}/root:/{path}:/content` ([simple upload API](https://learn.microsoft.com/graph/api/driveitem-put-content?view=graph-rest-1.0)).
- Files larger than 4 MB are **not supported** by the simple upload API; a [resumable upload session](https://learn.microsoft.com/graph/api/driveitem-createuploadsession?view=graph-rest-1.0) would be required (not implemented).
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
Update-UserFile_20260417_143022_123.log
```

Log entries include ISO 8601 timestamps, severity levels (`INFO`, `WARN`, `ERROR`), and contextual messages. Use `-Verbose` to also see log lines in the console.

### Retry and throttle handling

The script wraps all Graph API calls with automatic retry and exponential backoff:

- **Retried errors:** HTTP 429 (throttling), 500, 502, 503, 504, timeouts, and transient failures.
- **Not retried:** HTTP 401, 403, 404, and other permanent errors.
- **Defaults:** up to 4 attempts, starting at a 2-second delay, doubling each retry up to a 20-second maximum.

See [Microsoft Graph throttling guidance](https://learn.microsoft.com/graph/throttling) for details.

### Known limitations

- **PnP.PowerShell conflict:** Do not import `PnP.PowerShell` in the same session. It loads `Microsoft.Graph.Core` 1.x, which is incompatible with Graph SDK v2. Start a fresh `pwsh` session.
- **Simple upload only:** Files larger than 4 MB require a resumable upload session, which is not implemented.
- **Path translation:** Box paths prefixed with `All Files/` are stripped automatically. If your migration tool renamed items or altered folder hierarchy, paths may not resolve correctly.
- **Box path encoding:** Special characters (spaces, parentheses, etc.) are URL-encoded per segment for the Graph API. Unusual characters may still cause issues.
- **Unmapped roles:** Permissions with no Graph equivalent (`Previewer`, `Uploader`, `Previewer Uploader`) are skipped.
- **No email notifications:** All permission grants use `sendInvitation = false`.
- Always run with `-WhatIf` first in a new tenant or against new datasets.

---

## Set-BoxToOneDriveItemPermission.ps1 (PnP PowerShell)

### What it does (PnP)

Given a Box collaboration CSV export, for a specific Box owner (user):

1. Filters CSV rows to that owner (`Owner Login`).
2. Looks up the user's OneDrive personal site URL via SharePoint user profile properties.
3. Builds a OneDrive item URL and resolves it to the underlying **SharePoint list item ID**.
4. Maps Box roles to OneDrive/SharePoint roles.
5. Applies item-level permissions (supports `-WhatIf` and `-Confirm` via `SupportsShouldProcess`).
6. Emits a structured object per processed row (good for exporting reports).

### Requirements (PnP)

- **PowerShell 7+** recommended.
- **Modules:**

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
# Optional:
Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser
```

- The identity you authenticate with must be able to:
  - Read user profile properties (to find OneDrive personal site URL)
  - Access the target user's OneDrive site
  - Set item-level permissions in the target library

### Box role mapping (PnP)

| Box Permission | Normalized Role | SharePoint Role Definition |
|---|---|---|
| `Co-owner` | `Contributor` | `Contribute` |
| `Editor` | `Contributor` | `Contribute` |
| `Viewer Uploader` | `Viewer` | `Read` |
| `Viewer` | `Viewer` | `Read` |
| `Previewer Uploader` | `None` | *(skipped)* |
| `Previewer` | `None` | *(skipped)* |
| `Uploader` | `None` | *(skipped)* |

If you have unexpected Box roles, use `-AllowUnknownRole` to avoid failing fast.

### Usage (PnP)

#### 1) Dry-run (recommended first)
```powershell
.\Set-BoxToOneDriveItemPermission.ps1 -WhatIf -Verbose
```

#### 2) Process one user
```powershell
.\Set-BoxToOneDriveItemPermission.ps1 `
  -InputFile "C:\Path\To\Box_Collaboration_Export.csv" `
  -UserToProcess "user@contoso.com" `
  -MySiteHostUrl "https://contoso-my.sharepoint.com" `
  -SharePointOnlineAdminUrl "https://contoso-admin.sharepoint.com/" `
  -Verbose
```

#### 3) Export a report (without applying changes)
```powershell
.\Set-BoxToOneDriveItemPermission.ps1 `
  -InputFile "C:\Path\To\Box_Collaboration_Export.csv" `
  -UserToProcess "user@contoso.com" `
  -WhatIf -Verbose |
Export-Csv -Path "C:\Output\user_onedrive_permission_actions.csv" -NoTypeInformation
```

#### 4) Process multiple users via pipeline
```powershell
"user1@contoso.com","user2@contoso.com" |
  .\Set-BoxToOneDriveItemPermission.ps1 `
    -InputFile "C:\Path\To\Box_Collaboration_Export.csv" `
    -Verbose
```

Or derive users from the CSV:
```powershell
Import-Csv -Path "C:\Path\To\Box_Collaboration_Export.csv" |
  Select-Object -ExpandProperty "Owner Login" -Unique |
  Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
  .\Set-BoxToOneDriveItemPermission.ps1 -InputFile "C:\Path\To\Box_Collaboration_Export.csv" -Verbose
```

### Authentication (PnP)

`Set-BoxToOneDriveItemPermission.ps1` supports multiple authentication modes via `-AuthMode`:

- `Interactive`
- `ManagedIdentity`
- `WorkloadIdentity`
- `Certificate`
- `AccessToken`

You may also need to supply supporting parameters depending on mode:
- `-Tenant`
- `-ClientId`
- `-CertificateThumbprint` or `-CertificatePath` (+ optional `-CertificatePassword`)
- `-AccessToken`

> Tip: Start with `-AuthMode Interactive` for local runs, then move to `ManagedIdentity`/`WorkloadIdentity` for automation if your environment supports it.

### Logging (PnP)

Each run creates a file named like:
`Set-BoxToOneDriveItemPermission_<WindowsUser>_<yyyyMMdd_HHmmss_fff>.log`

Use `-Verbose` to also see log lines in the console.

### Known limitations (PnP)

- Item resolution is sensitive to **SharePoint URL rules**. If item names contain characters invalid for SharePoint URLs (e.g., `#`, `%`, `&`) or were renamed during migration, resolution may fail.
- Item resolution logic may not reflect full nested folder hierarchy if your migration preserves deep paths. Review how `ItemUrl` is constructed if needed.
- Roles mapped to `None` are intentionally skipped.
- Always run with `-WhatIf` first in a new tenant or against new datasets.

---

## Disclaimer

This repository includes `DISCLAIMER.MD`. Please review it before using or modifying the scripts.
# NPSBox — Synchronize Box collaboration permissions to OneDrive (SharePoint Online)

This repository contains PowerShell scripts that help **translate Box collaboration/export data** into **OneDrive for Business (SharePoint Online) item-level permissions**.

The primary script, `Set-BoxToOneDriveItemPermission.ps1`, reads a CSV export of Box collaboration data, resolves each Box item to a SharePoint/OneDrive list item, and (optionally) applies equivalent permissions using PnP PowerShell.

> Note: This is sample/illustration code. See `DISCLAIMER.MD`.

---

## What this does

Given a Box collaboration CSV export, for a specific Box owner (user):

1. Filters CSV rows to that owner (`Owner Login`).
2. Looks up the user’s OneDrive personal site URL via SharePoint user profile properties.
3. Builds a OneDrive item URL and resolves it to the underlying **SharePoint list item ID**.
4. Maps Box roles to OneDrive/SharePoint roles.
5. Applies item-level permissions (supports `-WhatIf` and `-Confirm` via `SupportsShouldProcess`).
6. Emits a structured object per processed row (good for exporting reports).

---

## Repository layout (high level)

- `Set-BoxToOneDriveItemPermission.ps1` — main script (resolve items + set permissions)
- `Connect-SharePointOnlineAdmin.ps1` — helper to connect to SPO admin center via PnP
- `Box_Collaboration_Sample_Data.csv` — sample CSV structure
- `DISCLAIMER.MD` — disclaimer text
- `tests/` — placeholder for Pester tests (recommended to add)
- `releases/`, `logs/`, `LocalFiles/` — local/release assets (use as needed)

---

## Requirements

### PowerShell
- PowerShell 7+ recommended.

### Modules
- `PnP.PowerShell`
- (Optional/legacy helper) `Microsoft.Online.SharePoint.PowerShell` (used by `Add-SCA.ps1`)

Install:
```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
# Optional:
Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser
```

### Permissions / access prerequisites
- The identity you authenticate with must be able to:
  - Read user profile properties (to find OneDrive personal site URL)
  - Access the target user’s OneDrive site
  - Set item-level permissions in the target library

---

## Input CSV format

Your CSV must include (at minimum) these columns:

- `Owner Login`
- `Item Name`
- `Item Type` (e.g., `File`, `Folder`)
- `Path`
- `Collaborator Login`
- `Collaborator Permission`

See: `Box_Collaboration_Sample_Data.csv`.

---

## Box role mapping

The script maps Box collaborator permissions to normalized OneDrive roles, then to SharePoint role definitions when applying permissions.

**Box → normalized OneDrive**
- `Co-owner` → `Contributor`
- `Editor` → `Contributor`
- `Viewer Uploader` → `Viewer`
- `Viewer` → `Viewer`
- `Previewer Uploader` → `None`
- `Previewer` → `None`
- `Uploader` → `None`

**Normalized OneDrive → SharePoint role definition**
- `Contributor` → `Contribute`
- `Viewer` → `Read`
- `None` → skipped (no assignment performed)

If you have unexpected Box roles, use `-AllowUnknownRole` to avoid failing fast.

---

## Usage

### 1) Dry-run (recommended first)
```powershell
.
Set-BoxToOneDriveItemPermission.ps1 -WhatIf -Verbose
```

### 2) Process one user
```powershell
.
Set-BoxToOneDriveItemPermission.ps1 `
  -InputFile "C:\Path\To\Box_Collaboration_Export.csv" `
  -UserToProcess "user@contoso.com" `
  -MySiteHostUrl "https://contoso-my.sharepoint.com" `
  -SharePointOnlineAdminUrl "https://contoso-admin.sharepoint.com/" `
  -Verbose
```

### 3) Export a report (without applying changes)
```powershell
.
Set-BoxToOneDriveItemPermission.ps1 `
  -InputFile "C:\Path\To\Box_Collaboration_Export.csv" `
  -UserToProcess "user@contoso.com" `
  -WhatIf -Verbose |
Export-Csv -Path "C:\Output\user_onedrive_permission_actions.csv" -NoTypeInformation
```

### 4) Process multiple users via pipeline
```powershell
"user1@contoso.com","user2@contoso.com" |
  .
Set-BoxToOneDriveItemPermission.ps1 `
    -InputFile "C:\Path\To\Box_Collaboration_Export.csv" `
    -Verbose
```

Or derive users from the CSV:
```powershell
Import-Csv -Path "C:\Path\To\Box_Collaboration_Export.csv" |
  Select-Object -ExpandProperty "Owner Login" -Unique |
  Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
  .
Set-BoxToOneDriveItemPermission.ps1 -InputFile "C:\Path\To\Box_Collaboration_Export.csv" -Verbose
```

---

## Authentication

`Set-BoxToOneDriveItemPermission.ps1` supports multiple authentication modes via `-AuthMode`:

- `Interactive`
- `ManagedIdentity`
- `WorkloadIdentity`
- `Certificate`
- `AccessToken`

You may also need to supply supporting parameters depending on mode:
- `-Tenant`
- `-ClientId`
- `-CertificateThumbprint` or `-CertificatePath` (+ optional `-CertificatePassword`)
- `-AccessToken`

> Tip: Start with `-AuthMode Interactive` for local runs, then move to `ManagedIdentity`/`WorkloadIdentity` for automation if your environment supports it.

---

## Logging

The script writes a log file per run to `-LogFolder` (default is currently a local path).

Each run creates a file named like:
`Set-BoxToOneDriveItemPermission_<WindowsUser>_<yyyyMMdd_HHmmss_fff>.log`

Use `-Verbose` to also see log lines in the console.

---

## Known limitations / important notes

- Item resolution is sensitive to **SharePoint URL rules**. If item names contain characters invalid for SharePoint URLs (e.g., `#`, `%`, `&`) or were renamed during migration, resolution may fail.
- Item resolution logic may not reflect full nested folder hierarchy if your migration preserves deep paths. Review how `ItemUrl` is constructed if needed.
- Roles mapped to `None` are intentionally skipped.
- Always run with `-WhatIf` first in a new tenant or against new datasets.

---

## Disclaimer

This repository includes `DISCLAIMER.MD`. Please review it before using or modifying the scripts.