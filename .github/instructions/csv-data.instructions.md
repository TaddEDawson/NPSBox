---
description: "Use when working with CSV input files for Box-to-OneDrive migration. Describes column schemas and data conventions."
applyTo: "**/*.csv"
---

# CSV Data Schemas

## UserInfo.csv / Box_Collaboration_Sample_Data.csv

| Column | Type | Description |
|--------|------|-------------|
| `Owner Login` | UPN (string) | OneDrive file owner, e.g. `user@tenant.onmicrosoft.com` |
| `Path` | string | Box source path, always prefixed with `All Files/`. Normalized by stripping that prefix. |
| `Item Name` | string | File or folder name (may contain spaces, parentheses, special characters) |
| `Item Type` | string | `File` or `Folder` |
| `Collaborator Login` | UPN (string) | User to grant access to. Empty values cause row failure. |
| `Collaborator Permission` | string | Box role — see mapping below |

## Box Role → Graph Permission Mapping

| Box Role | Graph Role | Action |
|----------|------------|--------|
| Editor | `write` | Applied |
| Co-owner | `write` | Applied |
| Viewer | `read` | Applied |
| Viewer Uploader | `write` | Applied |
| Previewer | *(none)* | Skipped |
| Uploader | *(none)* | Skipped |

## Path Conventions

- Paths use forward slashes.
- `All Files/` prefix is stripped during normalization.
- Trailing slashes on folder paths are trimmed.
- Backslash paths (e.g., `Documents\Subfolder`) are normalized to forward slashes.
