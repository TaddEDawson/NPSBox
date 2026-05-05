# NPSBox — Top 10 Lists

## Top 10 Things the App Does

1. **Migrates Box permissions to OneDrive** — Reads a CSV of Box collaboration data and translates each row into an equivalent OneDrive sharing permission via Microsoft Graph.

2. **Authenticates with certificate-based app-only auth** — Uses `Connect-MgGraph` with `TenantId`, `ClientId`, and `CertificateThumbprint` for unattended, secure execution without user interaction.

3. **Uploads local files to OneDrive** — With `-UploadFiles`, uploads per-user local folders (named by UPN) to the owner's OneDrive root using the Graph simple upload API (≤ 4 MB).

4. **Maps Box roles to Graph sharing roles** — Converts Box permissions (Co-owner, Editor → `write`; Viewer, Viewer Uploader → `read`) and gracefully skips roles with no Graph equivalent (Previewer, Uploader).

5. **Validates collaborator domains** — The `-AllowedDomains` parameter blocks sharing to external domains, preventing accidental data exposure to unauthorized recipients.

6. **Retries transient Graph errors with exponential backoff** — `Invoke-WithGraphRetry` honors the `Retry-After` header on 429 responses and retries on 5xx/timeout errors (up to 6 attempts, max 60 s delay).

7. **Validates environment at startup** — Pre-flight checks confirm required Graph modules are installed (`Assert-RequiredModules`), no PnP assembly conflicts exist (`Assert-GraphAssemblyCompatibility`), and the app has the necessary API permissions (`Assert-GraphPermissions`).

8. **Supports `-WhatIf` and `-Confirm`** — Every destructive operation (permission grant, file upload) is gated by `ShouldProcess`, allowing full dry-run previews before making any changes.

9. **Deduplicates and groups CSV rows** — Removes duplicate rows by 4-tuple key and groups by owner using `Group-Object -AsHashTable` for O(1) lookup, reducing redundant API calls.

10. **Outputs structured pipeline objects** — Every processed row emits a `[pscustomobject]` with `OwnerLogin`, `ItemName`, `CollaboratorLogin`, `Status`, `Action`, `GraphRole`, and `Error` — ready for filtering, export to CSV, or downstream pipeline consumption.

---

## Top 10 Wish List

1. **Resumable upload for files > 4 MB** — Implement [upload sessions](https://learn.microsoft.com/graph/api/driveitem-createuploadsession?view=graph-rest-1.0) to handle large files up to 250 GB instead of failing at the 4 MB boundary.

2. **Documented folder creation method** — Replace the undocumented `PATCH` upsert (with `conflictBehavior: 'replace'` that risks deleting child files) with the documented `POST` to `/children` + 409 conflict handling.

3. **CSV column validation at startup** — Validate that all required column names exist in the CSV header before processing begins, with a clear error listing any missing columns.

4. **Graph JSON batching** — Combine up to 20 permission-grant requests per HTTP call using [Graph batching](https://learn.microsoft.com/graph/json-batching), dramatically reducing total requests and throttling for large CSVs.

5. **Parallel owner processing** — Use `ForEach-Object -Parallel` to process multiple owners concurrently, reducing total runtime for multi-owner migrations.

6. **Proactive self-throttling** — Add a configurable delay between requests (e.g., `-ThrottleDelayMs`) to stay under Graph rate limits instead of relying entirely on retry-after-429 recovery.

7. **Idempotent re-runs with permission diffing** — Before granting a permission, check existing permissions on the item and skip if the exact role is already in place. Update the role if it differs rather than creating duplicate permissions.

8. **Managed identity and client secret auth** — Support Azure Managed Identity for runbook/automation scenarios and client secret for environments where certificate deployment is impractical.

9. **Streaming CSV processing** — Replace the full in-memory `Import-Csv` load with streaming/chunked reads to support 100k+ row CSVs without proportional memory growth.

10. **Remove hardcoded infrastructure defaults** — Strip dev/test `TenantId` and `ClientId` defaults from the param block so production operators must explicitly supply their own values, preventing accidental runs against the wrong tenant.
