---
description: "Use when writing or editing PowerShell scripts (.ps1). Covers Allman-style braces, ShouldProcess, logging, and Graph SDK conventions for this project."
applyTo: "**/*.ps1"
---

# PowerShell Coding Standards

- PowerShell 7+ (`#Requires -Version 7.0`).
- Allman-style braces with descriptive closing comments: `} # if`, `} # foreach — row`.
- `[CmdletBinding(SupportsShouldProcess)]` on any function that modifies state.
- Guard destructive operations with `$PSCmdlet.ShouldProcess()` for `-WhatIf`/`-Confirm`.
- Temporarily disable `$WhatIfPreference` only for read-only/logging operations.

## Logging

- Log via `Write-LogLine` (writes to Verbose stream + log file). Never use `Write-Host`.
- Levels: `INFO`, `WARN`, `ERROR`.

## Pipeline Output

- Output `[pscustomobject]` results to the pipeline for each processed row.
- Include properties: `OwnerLogin`, `ItemName`, `CollaboratorLogin`, `Status`, `Action`, `GraphRole`, `Error`.

## Graph API

- URL-encode Graph API paths per-segment: `[System.Uri]::EscapeDataString()`.
- Keep module imports minimal: `Microsoft.Graph.Authentication`, `Microsoft.Graph.Users`, `Microsoft.Graph.Files`.
- Wrap Graph calls in `Invoke-WithGraphRetry` for exponential backoff on 429/5xx.
- PnP.PowerShell must NOT be loaded in the same session (assembly conflict with Graph SDK v2).

## Error Handling

- Use `try`/`catch`/`finally` with Allman braces.
- Close each block with a descriptive comment: `} # try`, `} # catch`.
