# Content package (`Content.zip`) format

`Deploy-VideoWalkthrough.ps1` and `Update-VideoWalkthrough.ps1` accept only a package that passes these checks (implemented in `VideoWalkthrough.Common.ps1`: `Get-VWZipManifest`, `Get-VWData`, `Assert-VWPortableHtml`).

## Layout

```
Content.zip
├── manifest.json
├── source.html
└── videos/
    ├── 01-example-walkthrough.mp4
    └── ...
```

- 2–100 entries, each at most 100 MB, at most 250 MiB uncompressed in total, no suspicious compression ratios.
- Only `source.html` and `videos/<name>.mp4` (letters, digits, `_ . -`). No folders entries, symlinks, absolute paths, `..`, reserved Windows names or duplicate/case-colliding names.

## `manifest.json`

```json
{
  "schemaVersion": 1,
  "sourceHtml": "source.html",
  "generatedUtc": "2026-10-09T00:00:00Z",
  "files": [
    { "path": "source.html", "bytes": 2502941, "sha256": "<64 lowercase hex>" },
    { "path": "videos/01-example-walkthrough.mp4", "bytes": 5683792, "sha256": "<64 lowercase hex>" }
  ]
}
```

Every file except `manifest.json` must be listed with its exact size and SHA-256, and nothing else may be in the ZIP.

## `source.html`

A single self-contained page with:

- an `<h1>` heading and a `<p class="sub">` dated subtitle;
- a script block containing `const DATA = { ... };` (on its own line) with:
  - `cards`: array of objects with `id`, `num`, `title`, `q`, `verdict`, `status`, `licsum`, and optional `who`, `setup[]`, `check[]`, `gotchas[]`, `license[]`, `wait`, `examples[]`, `figs[]` (`{ src, cap }`), `learn[]` (`[title, https-url]`), `ref`;
  - `videos`: map of card `id` → `{ "src": "videos/<file>.mp4", "len": "2 min 45 s" }`; every listed video must be in the ZIP and every video in the ZIP must be listed;
  - `walk: true` to show the video panels.
- No inline event handlers (`onclick=` …), no `javascript:` URLs, no off-host scripts, stylesheets, fonts or images, no personal OneDrive links. Screenshots must be embedded `data:image/png|jpeg|webp;base64` images. Reference links must be credential-free `https://` URLs.

`Publish-WalkthroughHome.ps1` additionally expects the export buttons, print stylesheet and video template of the reference walkthrough; if they are missing it stops with a message rather than publishing a broken page (or use `-SkipPdf`).

## Building a package

From a folder that contains `source.html` and `videos\`:

```powershell
Set-Location 'C:\VideoWalkthroughPackage'
. .\VideoWalkthrough.Common.ps1
New-VWPayloadZip -SourceHtmlPath 'C:\WalkthroughRelease\source.html' -SourceFolder 'C:\WalkthroughRelease' -DestinationZip 'C:\WalkthroughRelease\Content.zip'
```

The function validates the HTML, hashes every file, writes `manifest.json`, and re-validates the finished ZIP. To publish a revised package to an existing site, use `Update-VideoWalkthrough.ps1 -SourceFolder` or `-ContentZipPath`, then rerun `Publish-WalkthroughHome.ps1`.
