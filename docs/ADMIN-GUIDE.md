# Video Walkthrough administrator package

Open **QUICKSTART.html** inside the extracted package for the yellow action steps.

## What changed in v2.2.1

- **Launcher startup error fixed.** v2.2 failed at start with `Join-Path : Cannot bind argument to parameter 'Path' because it is an empty string`. The scripts now find their own folder even when `$PSScriptRoot` is empty, and both `.cmd` launchers pass their folder in.
- **Video upload HTTP 400 fixed.** v2.2 created the site and asset folder, then failed on the first video upload. The upload request now sends only the conflict setting, and a new file that still gets a 400 is retried once with an empty body.
- **Graph errors are diagnosable.** Failures now include the Graph error code and message, with URLs and tokens removed.
- **Work-folder prompt hardened.** The wizard now requires a full path. A bare answer such as `y` used to create a folder named `y` under your user profile.
- Videos and `Content.zip` are unchanged from v2.2 (payload SHA-256 `80a4ae0b9e3d02f00f2ae6fe7b09be1d2125621fafeae9e252bec701e9934826`).

## Start in Windows PowerShell

1. Extract all files to a local folder outside OneDrive, such as **C:\VideoWalkthroughPackage**. Review the scripts and use your approved execution/signing policy.
2. Open **64-bit Windows PowerShell 5.1**. Run the help check with the content path explicitly supplied:

```powershell
Set-Location 'C:\VideoWalkthroughPackage'
.\Deploy-VideoWalkthrough.ps1 -ContentZipPath 'C:\VideoWalkthroughPackage\Content.zip' -Help
```

The first line should read `Deploy-VideoWalkthrough v2.2.1`. The v2.2.1 launcher fix has not yet been retested on Windows, so use these commands instead of double-clicking `Start-VideoWalkthrough.cmd`.

3. Run **one** command for your actual cloud:

```powershell
# GCC:
.\Deploy-VideoWalkthrough.ps1 -ContentZipPath 'C:\VideoWalkthroughPackage\Content.zip' -Cloud GCC

# Commercial:
.\Deploy-VideoWalkthrough.ps1 -ContentZipPath 'C:\VideoWalkthroughPackage\Content.zip' -Cloud Commercial
```

Use your actual extracted folder in both paths if different.

4. Choose **S**. Enter organization, site title, slug, primary SharePoint hostname and tenant GUID. At **Cloud**, press Enter to keep the command's choice. **Account rules:** both browser sign-ins (Graph, then SharePoint admin) use the same account, which must hold SharePoint Administrator or Global Administrator and becomes the primary owner. The wizard has no owner prompt; the signed-in admin becomes the owner automatically (`-OwnerUpn` remains for `-NonInteractive` runs and must equal the signed-in account). Business owners without an admin role are added after deployment. Add co-owners after deployment: `Set-SPOUser -Site https://<tenant>.sharepoint.com/sites/<slug> -LoginName <UPN> -IsSiteCollectionAdmin $true` (after `Connect-SPOService`). Leave the optional application ID blank unless your administrator supplied one. Choose **N** for embedded HTML unless you need a large offline copy.
5. At **Work folder**, press **Enter** for the full path shown, or type a full local path. **Do not enter Y.** Keep that folder and its deployment state for updates.
6. Complete Microsoft sign-in with the intended tenant administrator. Use **Alt+Tab** if the window is behind the terminal. Success ends with `SharePoint page: ... (state: published)` and `videoCount : 11`. Grant approved viewers Read access and test playback as a viewer. Do not restart a run just because a sign-in window is hidden.

7. **If a run fails after the site is created, resume it. Do not start a new site.** Rerun with the same answers as the failed run (including the Cloud value), the same work folder, and `-ReuseExistingSite`. Add `-EmbedVideos` if the first run created the embedded HTML.

```powershell
.\Deploy-VideoWalkthrough.ps1 -NonInteractive -OrgName '<org>' -SiteName '<site title>' -SiteSlug <slug> -SharePointHost <tenant>.sharepoint.com -TenantId <tenant GUID> -OwnerUpn <admin UPN> -Cloud <GCC|Commercial> -WorkFolder '<work folder>' -ReuseExistingSite
```

## Publish the interactive guide and Home page

Deploy runs `Publish-WalkthroughHome.ps1` automatically at the end of a hosted install (skip with `-SkipHomePage`). To run it again, for example after a content update:

```powershell
.\Publish-WalkthroughHome.ps1 -WorkFolder '<work folder>' -Force
```

It uploads the interactive walkthrough next to the hosted videos, builds the card PDFs locally with Microsoft Edge and uploads them to `VideoWalkthrough/pdf`, and rebuilds Home.aspx with a link to the walkthrough plus direct video links. In the SharePoint file viewer, inline video and browser printing are blocked, so video buttons open the SharePoint player and export buttons open the PDFs. The previous Home layout is saved as `home-backup-<UTC>.json` in the work folder. Use `-SkipPdf` if Edge is unavailable and `-Help` for all parameters. Rerun it after any content update.

## Remove a deployment and rebuild

`Remove-VideoWalkthrough.ps1` deletes only the site recorded in the work folder's `deployment-state.json` (communication site created by this package), purges it from the SharePoint recycle bin so the URL can be reused, and renames the work folder to `<work folder>-removed-<UTC>`. It asks you to type the site slug first. **The purge is permanent.** Preview with `-WhatIf`; use `-KeepInRecycleBin` to keep the site restorable for 93 days (the URL stays reserved).

```powershell
.\Remove-VideoWalkthrough.ps1 -WorkFolder '<work folder>' -WhatIf
.\Remove-VideoWalkthrough.ps1 -WorkFolder '<work folder>' -Force
```

Wait a few minutes, then rebuild with Deploy (section above, full work-folder path) followed by `Publish-WalkthroughHome.ps1`. Manual equivalent: `Remove-SPOSite` then `Remove-SPODeletedSite` for `https://<tenant>.sharepoint.com/sites/<slug>`, and rename the old work folder.

## Update after installation

Back up the work folder; replace same-named MP4s in **WorkFolder\videos**, not `content\videos`. Run `.\Update-VideoWalkthrough.ps1 -WorkFolder '<work folder>'` (preview first with `-WhatIf`) and choose **R**. Review changes before confirming Y. Revised source folders or Content.zip use **F/Z**; regenerate the hosted page only when explicitly intended. Keep pending journals/history if interrupted. The updater has not yet been tested on Windows.

## Prerequisites and release gate

Windows PowerShell 5.1 x64, .NET 4.7.2+, approved SPO/Graph modules, SharePoint Administrator access for installation and approved delegated `Sites.ReadWrite.All` consent. No policy bypass or password collection. Commercial/GCC hosting only; GCC High/DoD are blocked. **Verified 10 Oct 2026:** a full install on a GCC-Moderate demo tenant (site created, 11 videos uploaded, page published), resumed with `-ReuseExistingSite` after the v2.2 upload failure. **Still required before customer release:** a same-name video replacement with the updater (then rerun `Publish-WalkthroughHome.ps1`), and a retest of the double-click launchers. The interactive guide, video links and PDF exports were verified on the same tenant. Tenant-wide site lookups can be slow.
