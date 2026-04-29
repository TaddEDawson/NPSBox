---
description: "Scaffold a new Microsoft Graph API operation with retry logic and error handling"
agent: "agent"
argument-hint: "Describe the Graph operation (e.g., 'list drive items in a folder')"
tools: [read, edit, search, web]
---

Add a new Graph API operation to `Update-UserFile.ps1`:

1. Read the existing script to understand the helper function patterns (`Invoke-WithGraphRetry`, `Write-LogLine`, URL encoding).
2. Create the new function in the `begin` block following existing conventions:
   - Allman-style braces with descriptive closing comments.
   - Wrap the Graph call in `Invoke-WithGraphRetry`.
   - URL-encode path segments with `[System.Uri]::EscapeDataString()`.
   - Use `Write-LogLine` for logging.
   - Guard state-changing calls with `$PSCmdlet.ShouldProcess()`.
   - Return a `[pscustomobject]` with relevant properties.
3. Cite the official Graph API documentation link for the endpoint.
4. Add corresponding Pester tests in `tests/Update-UserFile.Tests.ps1` with appropriate mocks.
5. Bump the script version (revision increment) and update the date.
