# Contributing

Issues and pull requests are welcome. Please keep changes consistent with how the kit is built:

- **Windows PowerShell 5.1 compatible.** No PowerShell 7-only syntax or modules.
- **ASCII-only** script source, saved as UTF-8 with BOM, CRLF line endings.
- **Safe by default:** every state-changing action supports `-WhatIf`, is idempotent or checkpointed, and never bypasses execution policy.
- **No secrets or tenant data** in code, examples, issues or logs. Use `contoso` placeholders and the `11111111-2222-3333-4444-555555555555` GUID.
- Keep `-Help` output accurate for any parameter you add.
- Run the checks locally before opening a PR (the same checks run in GitHub Actions):

  ```powershell
  Get-ChildItem *.ps1 | ForEach-Object { $e = $null; [void][System.Management.Automation.Language.Parser]::ParseFile($_.FullName, [ref]$null, [ref]$e); if ($e) { $_.Name; $e } }
  Invoke-ScriptAnalyzer -Path . -Severity Error
  ```

- Validate hosted changes in a test tenant and say in the PR what you tested.

Report security issues privately; see [SECURITY.md](SECURITY.md).
