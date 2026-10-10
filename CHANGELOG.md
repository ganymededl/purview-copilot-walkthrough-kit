# Changelog

## 2.2.1 – 2026-10-10

### Fixed
- Launcher startup error (`Join-Path … empty string`): scripts resolve their own folder when `$PSScriptRoot` is empty; launchers pass their folder in.
- HTTP 400 on the first video upload: the upload-session request no longer repeats the file name in the body; one fallback with an empty body for new files.
- Wizard accepted a bare answer (such as `y`) as the work folder; full local paths are now required and relative paths resolve against the PowerShell location.
- Wizard validated answers only at the end; slug (auto-lowercased), hostname, tenant GUID are now checked as they are entered.

### Added
- `Publish-WalkthroughHome.ps1`: hosts the interactive guide in the site library, builds per-card and all-cards PDFs with headless Edge (light print theme), turns video and export buttons into SharePoint links, rewrites and publishes `Home.aspx` (previous layout backed up). Reads the target site from the work folder.
- Deploy runs `Publish-WalkthroughHome.ps1` automatically at the end of a hosted install (`-SkipHomePage` to opt out).
- `Remove-VideoWalkthrough.ps1`: state-bound site deletion and recycle-bin purge with typed confirmation and `-WhatIf`.
- Graph errors now include the Graph error code and message (URLs and tokens removed).
- Clear sign-in requirement notice: one SharePoint/Global Administrator account for both sign-ins, which becomes the primary owner. The owner prompt was removed from the wizard.

## 2.2 – 2026-10-09
- Initial administrator package: offline walkthrough, optional SharePoint communication site, updater.
