# TODO

## Current Priority
- Resolve upload of all files from a given folder during migration.

## Action Items
- Reproduce the folder upload issue with `-UploadFiles` using a test user under `LocalFiles/`.
- Verify recursive file discovery logic in `Update-UserFile.ps1` (including nested folders).
- Confirm path normalization and Graph path encoding for folder/file names with spaces and special characters.
- Validate that each discovered file is uploaded to the correct OneDrive target path.
- Add or update Pester coverage in `tests/Update-UserFile.Tests.ps1` for multi-file and nested-folder uploads.
- Run full test suite and confirm no regressions.
- Document expected folder upload behavior and edge cases in `README.MD`.

## Definition of Done
- All files in a provided source folder (including subfolders) are uploaded successfully.
- Upload results are logged and emitted as pipeline output objects.
- Tests cover success and failure cases for folder-wide upload behavior.
