# Assumptions & Constraints Audit

Reviewed: 2026-04-29
Script: `Update-UserFile.ps1` v1.2.0.1

---

## Security

| # | Finding | Risk | Status |
|---|---------|------|--------|
| S1 | **No collaborator domain validation** — CSV `Collaborator Login` goes directly into the invite API `recipients.email`. External addresses (e.g., `someone@gmail.com`) silently grant OneDrive access to external users. | **High** | Fixed in v1.2.0.2 |
| S2 | **`Files.ReadWrite.All` application permission** — Grants the app read/write to every user's OneDrive. A compromised certificate = full tenant file access. | Info | Documented (infrastructure control) |
| S3 | **Hardcoded infrastructure identifiers** — `TenantId`, `ClientId`, `CertificateThumbprint` in param defaults reveal the app registration and tenant. | Low | Documented (convenience defaults) |
| S4 | **Silent permission grants** — `sendInvitation = $false` means collaborators receive no notification. | Info | By design |

## Performance

| # | Finding | Risk | Status |
|---|---------|------|--------|
| P1 | **No 4 MB file-size guard** — Comment documents the limit, but no pre-check on `$file.Length`. Files > 4 MB fail with a Graph error instead of a clear message. Resumable upload not implemented. | **Medium** | Fixed in v1.2.0.4 |
| P2 | **`ReadAllBytes` loads entire file into memory** — For many files near 4 MB, memory compounds. | Low | Mitigated by P1 guard |
| P3 | **CSV re-read on every pipeline input** — `Import-Csv` runs in `process` block. Piping multiple UPNs re-parses the entire CSV for each. | **Medium** | Fixed in v1.2.0.5 |
| P4 | **O(n²) owner filtering** — `Where-Object` scans full array per owner. `Group-Object` would be O(n). | Low | Fixed in v1.2.0.8 |
| P5 | **No Graph JSON batching** — Each permission grant is a separate HTTP request. Graph supports batching up to 20 requests. | Low | Documented (future enhancement) |
| P6 | **No parallel processing** — Owners processed sequentially. `ForEach-Object -Parallel` could help. | Low | Documented (future enhancement) |

## Throttling

| # | Finding | Risk | Status |
|---|---------|------|--------|
| T1 | **Retry ignores `Retry-After` header** — Graph 429 responses include `Retry-After` specifying exactly how long to wait. Script uses its own exponential backoff instead. | **High** | Fixed in v1.2.0.3 |
| T2 | **Max retry window too short** — 4 attempts × exponential backoff ≈ 30 s. Graph can throttle for minutes under load. | **Medium** | Fixed in v1.2.0.7 |
| T3 | **No proactive self-throttling** — Requests fire as fast as possible, virtually guaranteeing 429s for large CSVs. | Low | Documented (future enhancement) |

## Data & Idempotency

| # | Finding | Risk | Status |
|---|---------|------|--------|
| D1 | **No duplicate detection** — Duplicate CSV rows each make a separate API call. The invite API handles it gracefully but wastes calls. | Low | Fixed in v1.2.0.8 |
| D2 | **No idempotency on re-runs** — Re-running re-grants every permission without checking existing roles. | Low | Documented (future enhancement) |
| D3 | **Folder `conflictBehavior = 'replace'`** — Overwrites folder metadata on every upload, even if folder exists. | Low | Documented (by design for migration) |

## Infrastructure

| # | Finding | Risk | Status |
|---|---------|------|--------|
| I1 | **Assumes OneDrive is provisioned** — `Get-MgUserDrive` fails with generic error if user hasn't been provisioned. No proactive check or useful message. | **Medium** | Fixed in v1.2.0.6 |
| I2 | **Path normalization assumes `All Files/` prefix** — If Box export format changes, normalization silently passes raw path through, leading to Graph 404s. | Low | Documented |
| I3 | **Single auth method (certificate-only)** — No client secret, managed identity, or interactive auth. Makes local testing harder. | Low | Documented (future enhancement) |
