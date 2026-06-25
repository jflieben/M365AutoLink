# M365AutoLink — Centralized (Admin) Version

`M365AutoLink_Centralized.ps1` is the **server-side, admin-driven** version of M365AutoLink. It creates OneDrive shortcuts to Microsoft Teams and SharePoint document libraries for **one or many users** without anyone signing in, using **application permissions** (Managed Identity or a certificate-based app registration — no user impersonation).

It's built for **scheduled tasks, Azure Automation runbooks, and Azure Functions**.

> **Just want to link sites for yourself?** Use the user version instead: it runs as the signed-in user with delegated auth and needs no admin setup. See **[README.md](README.md)**.

## User version vs. Centralized version

| | **[M365AutoLink.ps1](README.md)** (User) | **M365AutoLink_Centralized.ps1** (Admin) |
|---|---|---|
| **Max users**| Unlimited | ~500-1000 (also depends on site count) |
| **Runs as** | The signed-in user (delegated) | Managed Identity or app registration (application) |
| **Intended for** | Logon scripts / end-user self-service | Scheduled tasks, Azure Automation, Azure Functions |
| **Site discovery** | SharePoint Search (what *you* can see) | Enumerates **all** tenant sites + per-user permission check |
| **Auth** | Browser OAuth / cached refresh token | Managed Identity, or certificate-based client credentials |
| **End-user UI** | Tray icon, progress bar, Manage-shortcuts window | None (runs headless) |

## Features
- **No user interaction** — runs entirely with application permissions.
- **Targets groups, user lists, or the whole tenant.**
- **True permission-awareness** — checks each user's *effective* SharePoint permissions per library (`getUserEffectivePermissions`), so it also catches sites shared via direct permissions, not just group membership.
- **`View` vs `Edit` gating** — choose whether read access is enough or contribute/edit is required before a shortcut is created.
- **Automatic cleanup** — removes shortcuts when a user loses access.
- **High-scale optimizations** — adaptive throttling, SharePoint REST `$batch`-ing, parallel runspaces, and an optional recursive permission pre-check to skip expensive per-user API calls.
- **Library compatibility safety** — skips libraries that don't behave well as shortcuts (required check-out, custom columns) and logs exactly what it skipped.
- **Dry-run support** — preview every create/rename/delete without touching any OneDrive.

---

## Quick Start
1. Download [`M365AutoLink_Centralized.ps1`](https://github.com/jflieben/M365AutoLink/blob/main/M365AutoLink_Centralized.ps1).
2. Set up authentication — Managed Identity (recommended on Azure) or a certificate-based app registration (see [Authentication](#authentication)).
3. Configure target users and filters in the configuration block.
4. Run the script:
   ```powershell
   .\M365AutoLink_Centralized.ps1
   ```

> The script defaults to `$DryRun = $true`. Review `lastRun.log`, then set `$DryRun = $false` to make actual changes.

## How It Works
1. **Pre-fetches all tenant sites** via `sites/getAllSites` and enumerates their document libraries.
2. **Filters** sites/libraries using include/exclude wildcards, file-count limits, archived/locked status, and check-out requirements.
3. **For each target user**, checks effective SharePoint permissions on every pre-cached library, then creates/updates/deletes that user's OneDrive shortcuts accordingly.

Because permissions are evaluated per user and per library, shortcuts are created only for libraries each user genuinely has access to — including direct shares, not just group membership.

---

## Configuration
Edit the `##########START CONFIGURATION##########` block at the top of the script.

### Targeting & filtering
| Variable | Description | Default |
|---|---|---|
| `$FolderName` | OneDrive folder that houses shortcuts (auto-created per user). | `"AutoLink"` |
| `$CloudType` | Cloud environment: `global`, `usgov`, `usdod`, `china`. | `"global"` |
| `$TargetMode` | Which users to process: `"Group"`, `"UserList"`, or `"All"`. | `"Group"` |
| `$TargetGroupId` | M365 Group Object ID (when `$TargetMode = "Group"`). | *(sample GUID — change it)* |
| `$TargetUsers` | Array of UPNs (when `$TargetMode = "UserList"`). | `@()` |
| `$MinimumPermissionLevel` | `"View"` = read access is enough; `"Edit"` = require contribute/edit. | `"Edit"` |
| `$excludedSitesByWildcard` | URL patterns to exclude (`*` matches one or more characters). | *(curated default list)* |
| `$includedSitesByWildcard` | URL patterns to include (only matching sites are processed). | `"https://*.sharepoint.com/sites/*"` |
| `$onlyConnectedSites` | Only process M365 group / Team-connected sites. | `$false` |
| `$maxFileCount` | Skip libraries with more files than this. | `300000` |
| `$minFileCount` | Skip libraries with fewer files than this. | `0` |
| `$linkNameReplacements` | Find/replace patterns for shortcut names (see [Link Name Cleanup](#link-name-cleanup)). | *(see script)* |

### Authentication (certificate fallback)
Only needed when **not** using Managed Identity — see [Authentication](#authentication).

| Variable | Description |
|---|---|
| `$ClientId` | Application (client) ID of your app registration. |
| `$TenantId` | Tenant ID/domain (GUID or `contoso.onmicrosoft.com`). |
| `$CertificateThumbprint` | SHA1 thumbprint (cert in `CurrentUser\My` or `LocalMachine\My`). |
| `$CertificatePath` | Path to a `.pfx` file (alternative to thumbprint). |
| `$CertificatePassword` | PFX password (if any). |

### Performance & behavior
| Variable | Description | Default |
|---|---|---|
| `$DryRun` | Preview mode — no writes/deletes/renames. | `$true` |
| `$InitialParallelLimit` | Initial runspace concurrency for enumeration/permission phases. | `10` |
| `$MaxParallelLimit` | Adaptive concurrency ceiling. | `25` |
| `$BatchSize` | SharePoint REST `$batch` chunk size (max 100). | `25` |
| `$ShortcutActionParallelLimit` | Parallel create/delete shortcut operations. | `8` |
| `$GraphMutationMaxAttempts` | Retry attempts for Graph create/move/rename/delete. | `5` |
| `$EnableRecursivePermissionPreCheck` | Resolve site/list/group memberships recursively to skip the per-user permission matrix where possible. | `$true` |
| `$PreCheckIncludeEveryoneClaims` | Treat broad Everyone/All-users claims as allow-all during pre-check. | `$true` |
| `$PreCheckVerboseDiagnostics` | Verbose diagnostics for the recursive pre-check. | `$false` |
| `$PreCheckMaxRecursionDepth` | Max recursion depth for principal expansion. | `8` |
| `$PreCheckMaxExpandedPrincipals` | Safety cap on expanded principals during recursion. | `5000` |
| `$MaxSitesToProcessForTesting` | When > 0, only process the first N filtered sites (testing). | `0` |

### Library Compatibility Rules
The script intentionally skips libraries where OneDrive shortcuts are known to misbehave:
- Libraries with **required check-out** enabled (`ForceCheckout = true`).
- Libraries with **custom columns** (non-base, non-hidden, non-sealed, non-readonly fields).

Skipped libraries are recorded in `lastRun.log` with a warning so you can see exactly what was excluded and why.

---

## Authentication

Authentication methods are tried in this order (first success wins):
1. **Client certificate** — only if `$ClientId`, `$TenantId`, and a certificate are configured.
2. **Azure Functions / App Service** identity endpoint (`$env:IDENTITY_ENDPOINT`).
3. **Azure VM** Instance Metadata Service (IMDS).
4. **Az PowerShell** module (`Connect-AzAccount -Identity`).

### Option 1 — Managed Identity (recommended)
When running on Azure (VM, Automation Account, Function, App Service), the script acquires tokens via the Managed Identity automatically. No certificate config needed — just grant the [required permissions](#required-permissions-application) to the MI's service principal.

### Option 2 — App Registration with Certificate
For non-Azure environments (on-prem server, workstation), use a certificate-based app registration.

> **Client secrets are not supported — certificate auth is required.**

#### 1. Generate a self-signed certificate
Run PowerShell **as Administrator**:
```powershell
$cert = New-SelfSignedCertificate `
    -Subject "CN=M365AutoLink" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -KeyLength 2048 `
    -KeyAlgorithm RSA `
    -HashAlgorithm SHA256 `
    -NotAfter (Get-Date).AddYears(2)

Write-Output "Thumbprint: $($cert.Thumbprint)"
```

#### 2. Export the certificate
```powershell
# Public key (.cer) — uploaded to Entra ID
Export-Certificate -Cert $cert -FilePath "C:\certs\M365AutoLink.cer"

# Private key (.pfx) — stays on the machine running the script
$pfxPassword = ConvertTo-SecureString -String "YourPfxPassword" -Force -AsPlainText
Export-PfxCertificate -Cert $cert -FilePath "C:\certs\M365AutoLink.pfx" -Password $pfxPassword
```

#### 3. Create the App Registration
1. [Azure portal](https://portal.azure.com) → **Microsoft Entra ID** → **App registrations** → **New registration**.
2. Name it (e.g. `M365AutoLink`), supported account types: **this organizational directory only**, then **Register**.
3. Copy the **Application (client) ID** and **Directory (tenant) ID**.

#### 4. Upload the certificate
In the app registration → **Certificates & secrets** → **Certificates** → **Upload certificate** → select the `.cer` from step 2 → **Add**.

#### 5. Grant API permissions
Add the [permissions below](#required-permissions-application), then **Grant admin consent**.

#### 6a. Configure the script — using the thumbprint
Use this when the certificate is installed in the Windows certificate store on the machine that runs the script.
```powershell
$ClientId = "your-application-client-id"
$TenantId = "your-tenant-id-or-contoso.onmicrosoft.com"
$CertificateThumbprint = "A1B2C3D4E5F6..."  # from step 1
$CertificatePath = ""
$CertificatePassword = ""
```
> Find an installed cert's thumbprint:
> ```powershell
> Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Subject -like "*M365AutoLink*" } | Select-Object Thumbprint, Subject, NotAfter
> ```

#### 6b. Configure the script — using a PFX file
Use this when you can't (or don't want to) install the certificate in the store.
```powershell
$ClientId = "your-application-client-id"
$TenantId = "your-tenant-id-or-contoso.onmicrosoft.com"
$CertificateThumbprint = ""
$CertificatePath = "C:\certs\M365AutoLink.pfx"
$CertificatePassword = "YourPfxPassword"  # leave empty if the PFX has no password
```

> To use the thumbprint approach with a PFX you received, import it first:
> ```powershell
> $pfxPassword = ConvertTo-SecureString -String "YourPfxPassword" -Force -AsPlainText
> Import-PfxCertificate -FilePath "C:\certs\M365AutoLink.pfx" -CertStoreLocation "Cert:\CurrentUser\My" -Password $pfxPassword
> ```

### Required Permissions (Application)
Grant these **application** permissions to your Managed Identity or app registration:

**Microsoft Graph**
- `Sites.Read.All` — read SharePoint site information.
- `Files.ReadWrite.All` — create shortcuts in users' OneDrive.
- `User.Read.All` — read user profiles for target enumeration.
- `GroupMember.Read.All` — read group membership (only needed for `$TargetMode = "Group"`).

**SharePoint**
- `Sites.FullControl.All` — access SharePoint REST APIs for site metadata and permission checks.

You can use [SPNRoleMgr](https://lieben.nu/tools/SPNRoleMgr/) to configure these permissions quickly.

---

## Link Name Cleanup
`$linkNameReplacements` tidies up shortcut names. Each entry has a `Pattern` (text to find) and a `Replacement`. Patterns are applied in order and the final name is trimmed.

```powershell
$linkNameReplacements = @(
    @{ Pattern = " - Documents"; Replacement = "" }
    @{ Pattern = "- Documents"; Replacement = "" }
)
```

This turns `Marketing - Documents` into `Marketing`. Renaming also applies to existing shortcuts on each run, so changing these patterns updates shortcuts that already exist.

## Copyright / License
https://www.lieben.nu/liebensraum/commercial-use/
(Commercial (re)use is not allowed without prior written consent by the author; otherwise free to use/modify as long as headers are kept intact.)

## Support / Risk
Best-effort support, use at your own risk.
When reporting issues on GitHub, please include `lastRun.log` from `%APPDATA%\M365AutoLink\`.
Before reporting, make sure that:
- The user can actually open the site in a browser.
- The site is **not** excluded from Search (`/_layouts/15/srchvis.aspx`).
- Restricted Content (Copilot) is **not** enabled for the site.
- You're not already syncing the site some other way (e.g. direct OneDrive sync).

## Author
Jos Lieben — https://www.lieben.nu
