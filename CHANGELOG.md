# Changelog

All notable changes to the delegated (per-user) edition of **M365AutoLink** (`M365AutoLink.ps1`) are documented here.

## [1.5.1]

### Shortcut names
- **No more doubled names.** OneDrive now puts the site title in front of the name a new shortcut is created with, so shortcuts came out as `Finance - Finance` or `Site - Site - Library`.
- **Faster reading of existing shortcuts.** Their targets now come with the folder listing in one call. The old bulk call read the OneDrive root instead of the `AutoLink` folder, so every shortcut needed its own lookup. 
- A shortcut that already exists elsewhere in the user's OneDrive is reported as "already exists" again instead of as an error on Windows PowerShell 5.1. A non-JSON error response (e.g. from a proxy) no longer hides the real error on PowerShell 7.

### Recovery without a restart
- **Failed runs retry automatically** after 2, 5, 15, 30 and then every 60 minutes while M365AutoLink is in the tray, so a logon run that fails because OneDrive, sign-in or the network isn't ready yet recovers by itself. Only the first failure in a row shows a notification; the tray tooltip shows when the next retry is.
- **Manage shortcuts is always available.** Before, it stayed greyed out until a run had succeeded. When nothing is loaded yet it now runs first and opens the window when that run is done. Clicking it, or the first-run notification, during a run opens it once the run finishes instead of being ignored.
- The "nothing to manage yet" message is a notification instead of a modal box, which could open behind other windows and block M365AutoLink until it was restarted.
- The Manage shortcuts window comes to the front when it opens, also after clicking a notification.
- The local OneDrive folder is looked up again on every run, so clicking the tray icon opens the `AutoLink` folder once OneDrive is signed in after M365AutoLink started.

### Settings safety
- If your settings (`Apps/M365AutoLink/config.json`) can't be read from OneDrive, the run stops and retries later. Before, it continued with default settings, re-linked excluded and held-back libraries, and saved the defaults over your settings. Manage shortcuts loads the settings again in that case.
- Creating the settings folders no longer replaces a folder that another device created at the same moment.
- Manage shortcuts logs what happened (closed, no changes, saved), so a click is never silent in the log.

## [1.5.0]

### Sync budget
- **OneDrive counts towards the total.** The item count of the user's own OneDrive is now added to the linked libraries everywhere a total is shown or checked. **Manage shortcuts** shows it as a greyed-out top row that can't be excluded.
- **Yellow and orange levels.** Besides red at 1,000,000 (fine for modern physical PCs), the total now turns yellow at 100,000 and orange at 250,000, since virtual desktops (VDI) struggle well before 1,000,000. Configurable via `$totalItemCountYellowThreshold` and `$totalItemCountOrangeThreshold`. They replace `$totalItemCountWarningRatio` and the amber "approaching" level. Yellow only colors the tray icon; orange and red also notify after each run.
- The capacity bar in **Manage shortcuts** has tick marks at the yellow and orange levels, and its total now covers every library, not only the rows the filter shows.

### First-run limit
- **The first run links at most 100,000 items.** When the `AutoLink` folder doesn't exist yet, libraries are linked smallest first (the most libraries that fit) until the total, including OneDrive, would pass `$totalItemCountYellowThreshold`. The rest is held back, so new users don't get a sudden flood of sync metadata.
- A notification says not all shortcuts were created; clicking it opens **Manage shortcuts**, where held-back libraries show as **Held back** with their item count. Unticking one includes it. Later runs keep them held back and only log them.
- The selection is saved to the user's OneDrive config before any shortcut is created, and an interrupted first run is still treated as the first run next time.
- Disable with `$LimitFirstRun = $false`.

## [1.4.0]

### Security & authentication
- **Silent Windows sign-in via WAM.** On Entra devices the script now tries the Windows Web Account Manager (WAM) before ever opening a browser, reusing the device Primary Refresh Token. The approach mirrors [dirkjanm/askWAM](https://github.com/dirkjanm/askWAM): a `WebAuthenticationCoreManager` scope-mode `WebTokenRequest` (`wam_compat=2.0`) resolved through `GetTokenSilentlyAsync`.
- **Safe, transparent fallback.** WAM is attempted on the global cloud. If WAM is unavailable, the device has no PRT, or the request cannot be brokered, the script silently falls back to the existing refresh-token / browser authorization-code flow.

## [1.3.1]
- **Logging** Better logging to help troubleshoot shortcut retrieval debugging

## [1.3.0]

### Security & authentication
- **PKCE + `state`** added to the browser authorization-code flow. Any callback whose `state` does not match is ignored, so another local process can no longer inject an authorization code.
- **v2 (v2.0) Entra endpoints** are now used for both the authorization and token requests (`scope=` instead of the legacy `resource=`), aligning with the OAuth 2.1 public-client baseline.
- **Loopback listener hardened**: switched from a raw `TcpListener` on a fixed port to `System.Net.HttpListener` on an **ephemeral port** (fixes collisions on multi-session RDS/AVD hosts and races between instances). All reflected error text is HTML-encoded, and a small branded landing page is shown after sign-in.
- **TLS 1.2/1.3** is enforced at startup on Windows PowerShell 5.1.
- **Uninstall now deletes the refresh-token cache and log** (`RefreshToken.xml` is a usable credential and must never survive an uninstall).
- Consent/sign-in failures no longer call `Exit` from deep in the request path (which silently killed the tray). They surface as an error balloon with the admin-consent URL and keep the tray alive for a retry.

### Reliability
- **Access-token caching fixed**: a valid cached access token is now reused instead of performing a full refresh-token exchange (+ DPAPI encrypt + disk write) on *every* API call. Token lifetime comes from `expires_in` with a 5-minute renewal margin; the refresh token is written to disk only when Entra rotates it.
- **Deletion circuit breaker**: the delete phase is skipped (with a warning + balloon) when SharePoint Search returns incomplete results or the desired set shrank by more than `$DeletionSafetyRatio`. A tombstone model requires a target to be absent on `$DeletionTombstoneRuns` consecutive runs before deletion. `$ForceReconcile` overrides both.
- **Single-instance guard**: a session mutex prevents two tray icons / token-rotation races / log contention. A second launch signals the running instance to run and exits.
- **Metadata lookup no longer aborts the whole run** on one transient error; failed items are skipped with a warning.
- **Unified retry pipeline**: `Invoke-GraphRaw` (config load/save) now retries transient 429/5xx/network errors, with correct `Retry-After` handling on both PowerShell 5.1 and 7.
- **Logging** is written directly to `lastRun.log` (no longer dependent on `Start-Transcript`) and rotated (`run-<timestamp>.log`, keeping `$LogHistoryCount`).
- **Pre-flight checks** warn (in one balloon) when OneDrive is not set up for a work account or the identity provider is unreachable.

### Performance
- Existing-shortcut metadata is fetched in a **single `RenderListDataAsStream` call** (with an automatic per-item fallback) instead of one call per shortcut.
- The fixed `Start-Sleep -Milliseconds 500` after every create/rename/delete was removed; throttling is handled by the shared retry/back-off.

### UX
- **High-DPI**: the process declares per-monitor-v2 DPI awareness and the floating progress bar scales with the display, so the UI is crisp at 125–200% scaling.
- **Manage shortcuts** dialog: text filter, click-to-sort columns, Exclude-all/Include-all, and the window is resizable and remembers its size.
- **Periodic auto-refresh** (`$AutoRefreshHours`, default off) re-runs on an interval while resident in the tray and shortly after the device resumes from sleep.
- **First-run onboarding** balloon explains where the shortcuts are and points at Manage shortcuts.
- Reduced the logon **console flash** on Windows 11 by launching through `conhost.exe --headless`.

### Compatibility & correctness
- PowerShell 7: large SharePoint JSON payloads use `ConvertFrom-Json -AsHashtable` instead of the .NET-Framework-only `JavaScriptSerializer`; `Add-Type -AssemblyName` replaces the deprecated `LoadWithPartialName`.
- Wildcard exclusion/inclusion now uses PowerShell `-like` (correct anchor semantics). **Note:** exclusions tighten slightly — a trailing `*` is required for prefix matches (e.g. `*/sites/HR*`).
- Window-drag handler resolves the form at event time (dragging borderless dialogs works reliably).
- Dead-code sweep (unused launch-mode set, per-page blocking GC, stale unique-name state).

### Documentation
- README: corrected delegated Graph permissions, documented previously-undocumented settings, added an Intune deployment guide and a troubleshooting matrix.
