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