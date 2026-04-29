---
description: "Use when writing or editing Pester tests. Covers mocking conventions, test structure, and patterns used in this project's test suite."
applyTo: "tests/**"
---

# Pester 5 Test Conventions

## Structure

- `Describe` block per script under test.
- `Context` blocks group related scenarios (e.g., "Permission Mapping", "Path Handling", "Error Handling").
- `BeforeAll` at the top creates a `#Requires`-stripped copy of the script in `$TestDrive` so module-loading issues don't block unit tests.

## Helper Function

Use `New-CsvRow` to build test CSV data with sensible defaults:

```powershell
New-CsvRow -ItemName 'Doc1.txt' -CollaboratorPermission 'Editor'
```

## Mocking

Mock ALL external dependencies — no real Graph or file-system calls in tests:

- `Assert-RequiredModules`, `Assert-GraphAssemblyCompatibility`, `Connect-Graph` — mock as no-op.
- `Connect-MgGraph`, `Disconnect-MgGraph` — mock as no-op.
- `Get-MgUserDrive` — return `[PSCustomObject]@{ Id = $driveId; WebUrl = $webUrl }`.
- `Invoke-MgGraphRequest` — branch on `$Uri` and `$Method` to return appropriate mock data.

## Script Invocation Pattern

Run the script in a child scope and capture output + verbose stream:

```powershell
$results = & {
    . $script:ScriptUnderTest -InputFile $script:TestCsv -UserToProcess $script:DefaultOwner `
        -CertificateThumbprint $script:DefaultThumbprint -LogFolder $script:LogFolder -Verbose:$false
} 6>&1
```

## Assertions

- Filter results by `ItemName` to test specific rows: `$results | Where-Object { $_.ItemName -eq 'Doc1.txt' }`.
- Verify `Status`, `GraphRole`, `Action`, `Error`, `NormalizedPath` properties on result objects.
