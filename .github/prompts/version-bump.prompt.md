---
description: "Bump the script version following the project's versioning convention"
agent: "agent"
argument-hint: "revision | minor | major"
tools: [read, edit]
---

Bump the version in `Update-UserFile.ps1` following the project convention:

- **revision** (default): Increment the fourth octet by 1. Example: `1.2.0.1` → `1.2.0.2`.
- **minor**: Bump the third octet and reset the fourth. Example: `1.2.0.5` → `1.2.1.0`.
- **major**: Bump the second octet and reset lower octets. Example: `1.2.1.0` → `1.3.0.0`.

Always update the `Date:` field to today's date in `yyyy-MM-dd` format.

The version and date are in the `.SYNOPSIS` block of `Update-UserFile.ps1`:

```
    Version: X.Y.Z.R
    Date:    yyyy-MM-dd
```

Apply the bump type specified by the user (default: revision), then confirm the old and new values.
