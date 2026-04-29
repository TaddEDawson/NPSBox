# Assumptions & Constraints Audit

Reviewed: 2026-04-29
Script: `Update-UserFile.ps1` v1.2.0.15

---

## Security

| # | Finding | Risk | Status |
|---|---------|------|--------|
| S1 | **No collaborator domain validation** — CSV `Collaborator Login` goes directly into the invite API `recipients.email`. External addresses (e.g., `someone@gmail.com`) silently grant OneDrive access to external users. | **High** | Fixed in v1.2.0.2 — `Test-CollaboratorDomain` validates against `-AllowedDomains` parameter. |
| S2 | **`Files.ReadWrite.All` application permission** — Grants the app read/write to every user's OneDrive. A compromised certificate = full tenant file access. | Info | Documented (infrastructure control). `Assert-GraphPermissions` verifies the permission is granted. |
| S3 | **Hardcoded infrastructure identifiers** — `TenantId`, `ClientId`, `CertificateThumbprint` in param defaults reveal the app registration and tenant. | Low | Documented (convenience defaults) |
| S4 | **Silent permission grants** — `sendInvitation = $false`, `requireSignIn = $true` means collaborators receive no notification but must sign in to access items. | Info | By design |
| S5 | **Domain validation is case-insensitive** — Domain extracted after `@` and lowercased before comparison to `AllowedDomains`. `CONTOSO.COM` matches `contoso.com`. | Low | By design |
| S6 | **No RFC 5322 email validation** — Collaborator Login is checked for `@` to extract domain. Strings without `@` (e.g., `notanemail`) cause the row to be skipped. No full email format validation. | Low | Documented |
| S7 | **Certificate must be in `Cert:\CurrentUser\My`** — If the certificate with the specified thumbprint is not found in the user's personal cert store, `Connect-MgGraph` may silently fall back to delegated (interactive) auth, which lacks required application permissions. Validation catches this. | **High** | Documented (validated at startup) |
| S8 | **PnP.PowerShell loaded = terminating error** — Script checks for PnP.PowerShell in the same session and throws before any Graph calls. Known assembly conflict between Graph SDK v2 (3.x) and PnP's `Microsoft.Graph.Core` v1.x. Must run in a clean `pwsh` session. | **High** | By design — `Assert-GraphAssemblyCompatibility` enforces. |

## Performance

| # | Finding | Risk | Status |
|---|---------|------|--------|
| P1 | **No 4 MB file-size guard** — Comment documents the limit, but no pre-check on `$file.Length`. Files > 4 MB fail with a Graph error instead of a clear message. Resumable upload not implemented. | **Medium** | Fixed in v1.2.0.4 — explicit `$file.Length > 4MB` check throws with clear message. |
| P2 | **`ReadAllBytes` loads entire file into memory** — For many files near 4 MB, memory compounds. | Low | Mitigated by P1 guard |
| P3 | **CSV re-read on every pipeline input** — `Import-Csv` runs in `process` block. Piping multiple UPNs re-parses the entire CSV for each. | **Medium** | Fixed in v1.2.0.5 — CSV cached in `$script:CachedCsvRows` in `begin` block; reused in `process` block. |
| P4 | **O(n²) owner filtering** — `Where-Object` scans full array per owner. `Group-Object` would be O(n). | Low | Fixed in v1.2.0.8 — `Group-Object -AsHashTable -AsString` for O(1) owner lookup. |
| P5 | **No Graph JSON batching** — Each permission grant is a separate HTTP request. Graph supports batching up to 20 requests. | Low | Documented (future enhancement) |
| P6 | **No parallel processing** — Owners processed sequentially. `ForEach-Object -Parallel` could help. | Low | Documented (future enhancement) |
| P7 | **Entire CSV loaded into memory** — Large CSVs (100k+ rows) fully loaded into `$script:CachedCsvRows`. No streaming. Memory usage ≈ CSV size × 2–3×. | Medium | Documented |
| P8 | **No limit on output object count** — Every CSV row produces a `[pscustomobject]` result. Large CSVs (100k rows) = 100k objects in pipeline. Memory impact depends on downstream filtering. | Medium | Documented |

## Throttling

| # | Finding | Risk | Status |
|---|---------|------|--------|
| T1 | **Retry ignores `Retry-After` header** — Graph 429 responses include `Retry-After` specifying exactly how long to wait. Script uses its own exponential backoff instead. | **High** | Fixed in v1.2.0.3 — `Get-RetryAfterSeconds` parses the `Retry-After` header and honors it in the retry loop. |
| T2 | **Max retry window too short** — 4 attempts × exponential backoff ≈ 30 s. Graph can throttle for minutes under load. | **Medium** | Fixed in v1.2.0.7 — increased to 6 attempts, `MaxDelaySeconds = 60`. |
| T3 | **No proactive self-throttling** — Requests fire as fast as possible, virtually guaranteeing 429s for large CSVs. | Low | Documented (future enhancement) |
| T4 | **Max retry delay capped at 60 seconds** — Exponential backoff stops at 60 s even if `Retry-After` header says 120 s. Very long throttle windows may exhaust all 6 attempts. | Low | Documented |
| T5 | **Transient errors detected by regex pattern matching** — Retry logic uses regex on error messages (`timeout|throttl|too many requests|429|5\d{2}|…`). Changes in Graph SDK error message formatting could break detection. | Medium | Documented |

## Data & Idempotency

| # | Finding | Risk | Status |
|---|---------|------|--------|
| D1 | **No duplicate detection** — Duplicate CSV rows each make a separate API call. The invite API handles it gracefully but wastes calls. | Low | Fixed in v1.2.0.8 — `Sort-Object -Unique` deduplicates on `(Path, ItemName, CollaboratorLogin, CollaboratorPermission)`. |
| D2 | **No idempotency on re-runs** — Re-running re-grants every permission without checking existing roles. | Low | Documented (future enhancement) |
| D3 | **Folder `conflictBehavior = 'replace'`** — Overwrites folder metadata on every upload, even if folder exists. | Low | Documented (by design for migration) |
| D4 | **No permission role update** — If a permission already exists with a different role (e.g., `read`), re-granting with `write` may create a second permission instead of updating. The invite API is idempotent for exact matches, not for role changes. | Medium | Documented (future enhancement) |
| D5 | **CSV column names must match exactly** — Script expects: `'Owner Login'`, `'Path'`, `'Item Name'`, `'Collaborator Login'`, `'Collaborator Permission'`. Misspelled or renamed columns cause null reference errors. | **High** | Documented |
| D6 | **Empty/whitespace cells cause row skip or fail** — Empty `Owner Login` is filtered out; empty `Collaborator Login` throws; empty `Path` throws. No unified empty-cell handling. | Medium | Documented |
| D7 | **Deduplication matches on 4-tuple key** — Duplicates identified by `(Path, ItemName, CollaboratorLogin, CollaboratorPermission)`. Same collaborator granted a different role on the same item is kept (not merged). | Low | By design |

## Infrastructure

| # | Finding | Risk | Status |
|---|---------|------|--------|
| I1 | **Assumes OneDrive is provisioned** — `Get-MgUserDrive` fails with generic error if user hasn't been provisioned. No proactive check or useful message. | **Medium** | Fixed in v1.2.0.6 — detects provisioning errors and provides actionable message with `Request-SPOPersonalSite` and portal link. |
| I2 | **Path normalization assumes `All Files/` prefix** — If Box export format changes, normalization silently passes raw path through, leading to Graph 404s. | Low | Documented |
| I3 | **Single auth method (certificate-only)** — No client secret, managed identity, or interactive auth. Makes local testing harder. | Low | Documented (future enhancement) |
| I4 | **PowerShell 7+ required** — `#Requires -Version 7.0`. Script uses ternary operators and improved module handling. Users on PowerShell 5.1 will fail to load. | Medium | By design |
| I5 | **Graph API v1.0 endpoint hardcoded** — All calls use `https://graph.microsoft.com/v1.0/...`. No beta endpoint or version negotiation. | Low | By design |
| I6 | **Graph module imports are version-specific** — Script imports `Microsoft.Graph.Authentication`, `.Users`, `.Files`, `.Applications`. If any module is uninstalled or a major version mismatch occurs, `Assert-RequiredModules` throws at startup. | Medium | By design |
| I7 | **Assembly compatibility checked at startup** — If `Microsoft.Graph.Core` v1.x is already loaded in the AppDomain (e.g., from PnP or a prior script in ISE), Graph SDK v2 calls will fail. Script detects and advises starting a fresh `pwsh` session. | Medium | By design |

## File Upload (`-UploadFiles`)

| # | Finding | Risk | Status |
|---|---------|------|--------|
| U1 | **Local folder must be named by user's UPN** — e.g., `AllFilesDirectory\user@contoso.com\`. If the folder doesn't exist, upload skips silently (no error thrown). | **High** | Documented |
| U2 | **Only simple upload for files ≤ 4 MB** — Resumable upload (for files > 4 MB) is not implemented. Files exceeding the limit throw with a clear message. | **High** | Documented (P1 guard enforces) |
| U3 | **Folders created with PATCH (folder body)** — Folders uploaded via `PATCH` to `root:/{path}` with `{ folder: {}, conflictBehavior: 'replace' }`. File overwrite behavior is replace (not skip). | Low | By design |
| U4 | **Folder processing order: depth-first by path length** — Folders sorted by `FullName.Length` so parent folders are created before children. Prevents children-before-parent failures. | Low | By design |

## Logging & Output

| # | Finding | Risk | Status |
|---|---------|------|--------|
| L1 | **Log file timestamps use local time** — Log filenames use `Get-Date -Format` with default locale, not UTC. Timestamps are local to the machine running the script. | Low | By design |
| L2 | **`-WhatIf` still writes log files** — `$WhatIfPreference` is temporarily disabled for logging operations. Ensures audit trail exists even in preview mode. | Low | By design |
| L3 | **All output objects use `[pscustomobject]`** — Results for permissions, uploads, and errors include properties: `OwnerLogin`, `ItemName`, `CollaboratorLogin`, `Status`, `Action`, `GraphRole`, `Error`. Callers can filter and export to CSV. | Low | By design |

## Permission Mapping

| # | Finding | Risk | Status |
|---|---------|------|--------|
| M1 | **Box Previewer/Uploader roles silently skipped** — Roles that map to `$null` (no Graph equivalent) are skipped with `Status='Skipped'`. No API call is made. | Low | By design |
| M2 | **Invite API uses `requireSignIn = $true`** — Recipients must sign in with a Microsoft account to access shared items. Anonymous sharing is not supported. | Low | By design |

## Pipeline Behavior

| # | Finding | Risk | Status |
|---|---------|------|--------|
| B1 | **Three-block pipeline structure (`begin`/`process`/`end`)** — `begin` block runs once (caches CSV, authenticates, defines helpers). `process` block runs for each `-UserToProcess` pipeline input. `end` block disconnects Graph. | Low | By design |
| B2 | **`-WhatIf` and `-Confirm` fully supported** — `SupportsShouldProcess = $true`, `ConfirmImpact = 'Medium'`. Callers can use `-WhatIf` to preview all operations without making changes. | Low | By design |
| B3 | **Backslash paths normalized to forward slashes** — Windows-style `Path\SubPath` converted to `Path/SubPath` for Graph API compatibility. | Low | By design |
| B4 | **Path segments URL-encoded separately** — Each `/`-delimited segment is encoded independently (spaces → `%20`, parentheses → `%28%29`). The `/` separators remain unencoded. | Low | By design |
