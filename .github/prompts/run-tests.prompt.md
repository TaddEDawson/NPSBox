---
description: "Run the Pester test suite and report results"
agent: "agent"
tools: [execute, read]
---

Run the Pester 5 test suite at `tests/Update-UserFile.Tests.ps1`.

1. Execute the tests using the VS Code test runner or `Invoke-Pester`.
2. Report a summary: total tests, passed, failed, skipped.
3. For any failures, show the test name, expected vs. actual values, and the relevant line in the test file.
4. If failures are caused by recent script changes, suggest fixes.
