<#
.SYNOPSIS
    Automatically links all Microsoft Teams and Normal SharePoint sites to a user's OneDrive so they become client-side navigable

.DESCRIPTION
    This script authenticates to Microsoft Graph using cached tokens when possible,
    retrieves all Microsoft Teams and Sharepoint sites the user has access to, and creates shortcuts 
    to them in the user's OneDrive under an "AutoLink" (configurable) folder.

.REQUIREMENTS
    - PowerShell 5.x or 7.x
    - Microsoft 365 (licensed) account
    - Automatic or Manual app registration (see below)
    - Sites should be included in search (which is default but can be overridden at site level)

.APP REGISTRATION REQUIREMENTS
    AUTOMATIC: 
    
        go to https://login.microsoftonline.com/organizations/adminconsent?client_id=ae7727e4-0471-4690-b155-76cbf5fdcb30
        and sign in as an admin to provide consent for the Lieben Consultancy public client app registration.
        Lieben Consultancy will in no way be able to access your data, the app registration is only used
        for OAuth authentication purposes (delegated).

    MANUAL / PRIVATE:
     
        You can also create your own app registration in your Azure AD tenant:

        1. In Azure Portal > App Registrations > Your App:
        a. Go to "Authentication" blade
        b. Under "Platform configurations", click "Add a platform" > "Mobile and desktop applications"
        c. Check the box for: https://login.microsoftonline.com/common/oauth2/nativeclient
        d. Also add: http://localhost (for browser callback)
        e. Enable "Allow public client flows" (set to Yes)
        
        2. Grant Admin Consent (one-time, eliminates consent prompts for all users):
        a. Go to "API permissions" blade
        b. Click "Grant admin consent for <tenant>"

        3. Replace the $ClientID variable in this script with your App Registration's Application (client) ID

.PERMISSIONS REQUIRED (Delegated)
    Microsoft Graph:
    - Files.ReadWrite.All     - Create/rename/remove the OneDrive shortcuts and read/write the app's config.json
    - Sites.Read.All          - Read SharePoint site information
    SharePoint (Office 365 SharePoint Online):
    - AllSites.Read           - Run the SharePoint Search discovery query and read list metadata via the REST API
    NOTE: Teams permissions are NOT required (discovery moved to SharePoint Search); do not add Team.ReadBasic.All.

.AUTHENTICATION FLOW
    1. Cached Refresh Token - From previous successful authentication (completely silent)
    2. Silent Browser Auth - Opens browser in the background to get tokens silently (if SSO is properly configured)
    3. Interactive Browser Auth - Opens browser for user to sign in (first time only)
    
    After first authentication, the refresh token is cached and all subsequent runs are silent until the token expires

.NOTES
    Author: Jos Lieben
    Version: see $ScriptVersion in the configuration block below
    Updates/Git: https://github.com/jflieben/M365AutoLink
    Copyright/License: https://www.lieben.nu/liebensraum/commercial-use/ (Commercial (re)use not allowed without prior written consent by the author, otherwise free to use/modify as long as header are kept intact)
    Microsoft doc: https://support.microsoft.com/en-us/office/add-shortcuts-to-shared-folders-in-onedrive-d66b1347-99b7-4470-9360-ffc048d35a33
    Always test carefully, use at your own risk, author takes no responsibility for this script
    
.EXAMPLE
    .\M365AutoLink.ps1
#>

##########START CONFIGURATION#############################
$ScriptVersion = "1.3.0" #single source of truth for the version (shown in the log + tray tooltip)
$FolderName = "AutoLink" #this is the folder created in onedrive to house all links this tool will create. Feel free to change this to something localized, the tool will auto-create it if it does not exist
#WARNING: Any pre-existing folders in above folder will be deleted!
$CloudType = "global" #global, usgov, usdod, china
$ClientID = "ae7727e4-0471-4690-b155-76cbf5fdcb30" #Lieben Consultancy public client ID, you can also create your own (see APP REGISTRATION REQUIREMENTS above)
$WindowStyle = "Normal" #Normal, Hidden, Minimized, Maximized - this controls the browser window style during authentication, Hidden will not show the browser but the user then won't be able to sign in if SSO is not working

# Dry-run mode: when $true, no shortcuts are created, deleted, or renamed. The script only shows what it would do.
$DryRun = $false

# Auto Launch mode valid values: Desktop, Start Menu, AtLogon
# Do not configure this if you want to run 100% manual or e.g. use this as a logon script in Group Policy
# Not configured would look like this: 
# $LaunchModes = @()
$LaunchModes = @('AtLogon','Desktop')

# When the script is deployed through Intune it runs from a temporary location that is deleted again right after execution. 
# Any persistence (see $LaunchModes) would then create shortcuts pointing at a path that no longer exists.
# Set $deployToPath to a permanent location and on first run the script copies itself there
#
# - Use a full file path (ending in .ps1) or a folder (the script keeps its M365AutoLink.ps1 name).
# - Environment variables are supported in both $env:NAME (PowerShell) and %NAME% (Windows) form.
# - The target folder is created automatically if it does not exist.
#
# Examples:
#   $deployToPath = "$env:APPDATA\M365AutoLink\M365AutoLink.ps1"        # roaming AppData (per-user, roams with the profile)
#   $deployToPath = "$env:OneDrive\Apps\M365AutoLink\M365AutoLink.ps1"  # OneDrive (per-user, survives device reset/reinstall)
#
# Leave as $Null to never copy the script (it runs and persists from wherever it currently is).
$deployToPath = $Null

# Uninstall mode: when $true the script removes ALL persistence, to cleanly remove M365AutoLink from a device.
$Uninstall = $false

#excluded sites will not be added a link if below pattern occurs in the site's URL. Use a * to match 1 or more characters
#the default list is recommended
#e.g. https://contoso.sharepoint.com/sites/HR*" would exclude all sites where the name starts with HR"
$excludedSitesByWildcard = @(
    "*/groupforanswersinvivaengagedonotdelete*"
    "*/sites/Streamvideo*"
    "*/portals/personal/*"
    "*/sites/AllCompany*"
    "*/personal/*"
    "*/contentstorage/*"
    "*/sites/contentTypeHub*"
    "*/sites/pwa"
    "*/sites/AppCatalog*"
)
#if you define included site, only sites matching one of the patterns you enter will be linked. Use a * to match 1 or more characters
#e.g. https://contoso.sharepoint.com/sites/HR*" would include all sites where the name starts with HR"
$includedSitesByWildcard = @(
    "https://*.sharepoint.com/sites/*"
)

#link name cleanup patterns - applied to shortcut names after creation and to existing shortcuts on each run
#each entry has a Pattern (string to find) and Replacement (string to replace with)
#patterns are applied in order, final name is trimmed of leading/trailing whitespace
$linkNameReplacements = @(
    @{ Pattern = " - Documents"; Replacement = "" }
    @{ Pattern = "- Documents"; Replacement = "" }
    @{ Pattern = "- Documenten"; Replacement = "" }
)

#below variables can be used to filter based on the number of existing files in the target location before creating a link
$maxFileCount = 300000
$minFileCount = 0

# Combined item-count guidance. When the total number of items across ALL linked libraries crosses
# these thresholds, Windows Explorer may not reliably show all folders/links and (rarely) sync breaks.
# The tool only WARNS about this (icon color, tooltip, dialogs) - it never blocks going over the limit.
$totalItemCountWarningThreshold = 1000000   # red "over limit" once the combined total reaches this
$totalItemCountWarningRatio     = 0.9       # amber "approaching" once the total reaches this fraction of the threshold
# Knowledgebase article opened when the user clicks the over/approaching-limit tray balloon notification.
$ItemCountHelpLink = "https://support.microsoft.com/en-US/onedrive/restrictions-and-limitations-in-onedrive-and-sharepoint#numberitemscanbesynced"

# Basic floating progress bar (bottom-right)
$ShowProgressBar = $true
$ProgressBarColor = "#00A3FF"
$ProgressBarText = "M365AutoLink is updating your shortcuts..."

# System tray behavior
$EnableSystemTrayIcon = $true
$KeepRunningInTray = $true # keeps process alive so tray can trigger runs and manage excluded sites
$TrayHelpLink = "https://support.microsoft.com/en-us/office/add-shortcuts-to-shared-folders-in-onedrive-d66b1347-99b7-4470-9360-ffc048d35a33"
$TrayCopyrightText = "Copyright (c) Lieben Consultancy"
$TrayCopyrightLink = "https://www.lieben.nu/liebensraum/commercial-use/"

# Device name inclusion filter - only run on devices where the name contains a specific string. Leave as $Null to run on all devices the script is deployed to
$DeviceNameIncludeFilter = $Null

# Periodic auto-refresh. When > 0, a run is triggered automatically every N hours while the tray
# process is alive (in addition to logon and manual "Run now"), plus shortly after the device resumes
# from sleep. 0 = off (only logon + manual). Requires $KeepRunningInTray.
$AutoRefreshHours = 0

# Deletion circuit breaker. To protect against a partial SharePoint Search outage causing mass
# deletion of valid shortcuts, the delete phase is SKIPPED for this run when the desired set shrank by
# more than this fraction versus the last successful run, or when search paging ended abnormally.
# Set to 1 to effectively disable the ratio guard. 
$DeletionSafetyRatio = 0.40

# Logging (C5). How many timestamped previous run logs to keep alongside lastRun.log.
$LogHistoryCount = 5

#system libraries that should never become OneDrive shortcuts even if returned by search
$excludedLibrariesByWildcard = @(
    "*style library*"
    "*stijlbibliotheek*"
    "*site assets*"
    "*siteactiva*"
    "*site pages*"
    "*form templates*"
    "*preservation hold library*"
)

# Additional exact title exclusions (case-insensitive) and feature IDs.
$ExcludedListTitles = @(
    "Access Requests","App Packages","appdata","appfiles","Apps in Testing","Cache Profiles","Composed Looks","Content and Structure Reports","Content type publishing error log","Converted Forms",
    "Device Channels","Form Templates","fpdatasources","Get started with Apps for Office and SharePoint","List Template Gallery", "Long Running Operation Status","Maintenance Log Library", "Images", "site collection images",
    "Master Docs","Master Page Gallery","MicroFeed","NintexFormXml","Quick Deploy Items","Relationships List","Reusable Content","Reporting Metadata", "Reporting Templates", "Search Config List","Site Assets","Preservation Hold Library",
    "Site Pages", "Solution Gallery","Style Library","Suggested Content Browser Locations","Theme Gallery", "TaxonomyHiddenList","User Information List","Web Part Gallery","wfpub","wfsvc","Workflow History","Workflow Tasks", "Pages"
)

$ExcludedListFeatureIDs = @(
    "00000000-0000-0000-0000-000000000000",
    "a0e5a010-1329-49d4-9e09-f280cdbed37d",
    "d11bc7d4-96c6-40e3-837d-3eb861805bfa",
    "00bfea71-c796-4402-9f2f-0eb9a6e71b18",
    "de12eebe-9114-4a4a-b7da-7585dc36a907"
)

##########END CONFIGURATION#############################

#region External configuration (F1)
# Layered configuration so admins can manage settings WITHOUT editing the script (which every update would
# overwrite - especially painful with $deployToPath self-copy). Precedence, first match wins per setting:
#   1. HKLM\Software\Policies\Lieben\M365AutoLink   (Intune/GPO manageable + lockable, most authoritative)
#   2. HKCU\Software\Policies\Lieben\M365AutoLink
#   3. M365AutoLink.config.json next to the script  (admin-managed, survives script updates)
#   4. the in-script default assigned above         (zero-config still works unchanged)
$script:externalConfigJson = $null
try {
    $selfConfigPath = if(-not [string]::IsNullOrWhiteSpace($PSCommandPath)) { $PSCommandPath }
                      elseif($MyInvocation.MyCommand -and -not [string]::IsNullOrWhiteSpace($MyInvocation.MyCommand.Path)) { $MyInvocation.MyCommand.Path }
                      else { $null }
    if($selfConfigPath) {
        $externalConfigFile = Join-Path -Path ([System.IO.Path]::GetDirectoryName($selfConfigPath)) -ChildPath 'M365AutoLink.config.json'
        if(Test-Path -LiteralPath $externalConfigFile) {
            $script:externalConfigJson = (Get-Content -LiteralPath $externalConfigFile -Raw | ConvertFrom-Json)
            Write-Host "Loaded external configuration from $externalConfigFile"
        }
    }
} catch {
    Write-Host "Failed to read M365AutoLink.config.json, using in-script defaults: $($_.Exception.Message)"
}

function Convert-SettingValue {
    # Coerce an external (string/registry/json) value to the type of the in-script default template.
    param($Value, $Template)
    try {
        if($Template -is [bool]) {
            if($Value -is [bool]) { return $Value }
            $text = ([string]$Value).Trim().ToLowerInvariant()
            return ($text -eq 'true' -or $text -eq '1' -or $text -eq 'yes' -or $text -eq 'on')
        }
        if($Template -is [int] -or $Template -is [long]) { return [int]$Value }
        if($Template -is [double]) { return [double]$Value }
        if($Template -is [array]) { return @($Value) }
        return [string]$Value
    } catch { return $Template }
}

function Resolve-Setting {
    param([Parameter(Mandatory = $true)][string]$Name, $Default)

    foreach($policyRoot in @('HKLM:\Software\Policies\Lieben\M365AutoLink', 'HKCU:\Software\Policies\Lieben\M365AutoLink')) {
        try {
            if(Test-Path -LiteralPath $policyRoot) {
                $property = Get-ItemProperty -Path $policyRoot -Name $Name -ErrorAction SilentlyContinue
                if($property -and $null -ne $property.$Name) {
                    return (Convert-SettingValue -Value $property.$Name -Template $Default)
                }
            }
        } catch {}
    }

    if($script:externalConfigJson) {
        try {
            $jsonProperty = $script:externalConfigJson.PSObject.Properties[$Name]
            if($jsonProperty -and $null -ne $jsonProperty.Value) {
                return (Convert-SettingValue -Value $jsonProperty.Value -Template $Default)
            }
        } catch {}
    }

    return $Default
}

# Apply external overrides to the configurable settings (no-op when nothing external is present).
$FolderName                     = Resolve-Setting -Name 'FolderName' -Default $FolderName
$CloudType                      = Resolve-Setting -Name 'CloudType' -Default $CloudType
$ClientID                       = Resolve-Setting -Name 'ClientID' -Default $ClientID
$WindowStyle                    = Resolve-Setting -Name 'WindowStyle' -Default $WindowStyle
$DryRun                         = Resolve-Setting -Name 'DryRun' -Default $DryRun
$LaunchModes                    = Resolve-Setting -Name 'LaunchModes' -Default $LaunchModes
$deployToPath                   = Resolve-Setting -Name 'deployToPath' -Default $deployToPath
$Uninstall                      = Resolve-Setting -Name 'Uninstall' -Default $Uninstall
$excludedSitesByWildcard        = Resolve-Setting -Name 'excludedSitesByWildcard' -Default $excludedSitesByWildcard
$includedSitesByWildcard        = Resolve-Setting -Name 'includedSitesByWildcard' -Default $includedSitesByWildcard
$maxFileCount                   = Resolve-Setting -Name 'maxFileCount' -Default $maxFileCount
$minFileCount                   = Resolve-Setting -Name 'minFileCount' -Default $minFileCount
$totalItemCountWarningThreshold = Resolve-Setting -Name 'totalItemCountWarningThreshold' -Default $totalItemCountWarningThreshold
$ShowProgressBar                = Resolve-Setting -Name 'ShowProgressBar' -Default $ShowProgressBar
$EnableSystemTrayIcon           = Resolve-Setting -Name 'EnableSystemTrayIcon' -Default $EnableSystemTrayIcon
$KeepRunningInTray              = Resolve-Setting -Name 'KeepRunningInTray' -Default $KeepRunningInTray
$DeviceNameIncludeFilter        = Resolve-Setting -Name 'DeviceNameIncludeFilter' -Default $DeviceNameIncludeFilter
$AutoRefreshHours               = Resolve-Setting -Name 'AutoRefreshHours' -Default $AutoRefreshHours
$DeletionSafetyRatio            = Resolve-Setting -Name 'DeletionSafetyRatio' -Default $DeletionSafetyRatio
$LogHistoryCount                = Resolve-Setting -Name 'LogHistoryCount' -Default $LogHistoryCount

# Normalize the wildcard pattern lists: trim each entry and drop empties. Leading/trailing whitespace in a
# pattern (common when it comes from a hand-edited config) is otherwise matched LITERALLY by -like, so e.g.
# "*kennisportaal* " would never match a URL that doesn't actually end in a space.
$excludedSitesByWildcard     = @($excludedSitesByWildcard     | ForEach-Object { ([string]$_).Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
$includedSitesByWildcard     = @($includedSitesByWildcard     | ForEach-Object { ([string]$_).Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
$excludedLibrariesByWildcard = @($excludedLibrariesByWildcard | ForEach-Object { ([string]$_).Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
#endregion

if($DeviceNameIncludeFilter -and -not $env:COMPUTERNAME.ToLowerInvariant().Contains($DeviceNameIncludeFilter.ToLowerInvariant())) {
    Write-Host "Device name '$($env:COMPUTERNAME)' does not match filter '$DeviceNameIncludeFilter'; exiting."
    return
}

#base vars
$global:octo = @{}
$global:octo.LCRefreshToken = $Null
$global:octo.LCCachedTokens = @{}
$global:octo.LCClientId = $ClientID
$global:octo.TokenCachePath = "$env:APPDATA\M365AutoLink\RefreshToken.xml"
$global:octo.LogPath = "$env:APPDATA\M365AutoLink\lastRun.log"

$script:traySync = $null
$script:trayRunspace = $null
$script:trayPS = $null
$script:userConfig = $null
$script:lastMappedLibraryOptions = @()
$script:lastAlreadyExistingShortcuts = @()
# Set for one run when the user deliberately changes exclusions in Manage shortcuts, so the deletion
# ratio guard (meant to catch a partial Search outage) doesn't block their intentional shrink.
$script:bypassDeletionRatioOnce = $false
$script:localOneDriveRootPath = $null
$script:localShortcutFolderPath = $null
$script:effectiveScriptPath = $null

#determine URLs based on where the tenant resides
switch($CloudType){
    'global' {
        $global:octo.idpUrl = "https://login.microsoftonline.com"
        $global:octo.graphUrl = "https://graph.microsoft.com"
        $global:octo.sharepointUrl = "https://www.sharepoint.com"
    }
    'usgov' {
        $global:octo.idpUrl = "https://login.microsoftonline.us"
        $global:octo.graphUrl = "https://graph.microsoft.us"
        $global:octo.sharepointUrl = "https://www.sharepoint.us"
    }
    'usdod' {
        $global:octo.idpUrl = "https://login.microsoftonline.us"
        $global:octo.graphUrl = "https://dod-graph.microsoft.us"
        $global:octo.sharepointUrl = "https://www.sharepoint-mil.us"
    }
    'china' {
        $global:octo.idpUrl = "https://login.chinacloudapi.cn"
        $global:octo.graphUrl = "https://microsoftgraph.chinacloudapi.cn"
        $global:octo.sharepointUrl = "https://www.sharepoint.cn"
    }
}

# OAuth 2.0 v2 endpoints (B2). The v2 authorize/token endpoints take scopes instead of the legacy
# v1 resource= parameter and are the current public-client baseline (PKCE, OAuth 2.1).
$global:octo.authorizeUrl = "$($global:octo.idpUrl)/common/oauth2/v2.0/authorize"
$global:octo.tokenUrl = "$($global:octo.idpUrl)/common/oauth2/v2.0/token"

# Windows PowerShell 5.1 on older/unmanaged machines can still negotiate TLS 1.0. Force modern TLS
# once at startup (guarded so we never fail on an older enum that lacks Tls13).
try {
    $desiredProtocols = [Net.SecurityProtocolType]::Tls12
    try { $desiredProtocols = $desiredProtocols -bor [Net.SecurityProtocolType]::Tls13 } catch {}
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor $desiredProtocols
} catch {}

function Get-RetryAfterSeconds {
    # read the Retry-After header across both the PS 5.1 (HttpWebResponse) and PS 7
    # (HttpResponseMessage) exception shapes. Returns 0 when no usable value is present.
    param($ErrorRecord)

    if($null -eq $ErrorRecord) { return 0 }
    $response = $null
    try { $response = $ErrorRecord.Exception.Response } catch {}
    if($null -eq $response) { return 0 }

    # PS 7 / HttpResponseMessage: Headers.RetryAfter.Delta (a TimeSpan) or .Date.
    try {
        $retryAfter = $response.Headers.RetryAfter
        if($retryAfter) {
            if($retryAfter.Delta -and $retryAfter.Delta.TotalSeconds -gt 0) {
                return [int][math]::Ceiling($retryAfter.Delta.TotalSeconds)
            }
            if($retryAfter.Date) {
                $seconds = ($retryAfter.Date.UtcDateTime - [DateTime]::UtcNow).TotalSeconds
                if($seconds -gt 0) { return [int][math]::Ceiling($seconds) }
            }
        }
    } catch {}

    # PS 5.1 / HttpWebResponse: Headers.GetValues("Retry-After") returns a numeric string.
    try {
        $values = $response.Headers.GetValues("Retry-After")
        if($values -and $values.Count -gt 0 -and $values[0] -match '^\d+$') {
            return [int]$values[0]
        }
    } catch {}

    return 0
}

function Get-HttpStatusCode {
    param($ErrorRecord)
    if($null -eq $ErrorRecord) { return $null }
    try { return [int]$ErrorRecord.Exception.Response.StatusCode } catch {}
    return $null
}

function Test-IsTransientHttpError {
    # Shared classifier used by every retry loop: throttling (429) or transport-level blips.
    param($ErrorRecord)

    $statusCode = Get-HttpStatusCode -ErrorRecord $ErrorRecord
    $message = ""
    try { $message = [string]$ErrorRecord.Exception.Message } catch {}

    $is429 = ($statusCode -eq 429) -or ($message -like "*429*")
    if($is429) { return $true }
    # Retry 5xx server errors too - they are frequently transient on the SharePoint search endpoint.
    if($null -ne $statusCode -and $statusCode -ge 500 -and $statusCode -lt 600) { return $true }

    $isTransientNetwork = $message -like "*No such host is known*" -or $message -like "*name or service not known*" -or $message -like "*network is unreachable*" -or $message -like "*connection was forcibly closed*" -or $message -like "*An existing connection was forcibly closed*" -or $message -like "*The operation has timed out*" -or $message -like "*Unable to connect to the remote server*"
    return [bool]$isTransientNetwork
}

function Invoke-RestWithRetry {
    # single retry/throttle core shared by Invoke-GraphRaw (and available to any caller). Honors
    # Retry-After (A10), retries only 429/5xx/transport errors, and fails fast on other HTTP errors.
    param(
        [Parameter(Mandatory = $true)][string]$Method,
        [Parameter(Mandatory = $true)][string]$Uri,
        [hashtable]$Headers,
        $Body,
        [string]$ContentType = 'application/json; charset=utf-8',
        [int]$MaxAttempts = 5,
        [int]$TimeoutSec = 120
    )

    $attempt = 0
    while($true) {
        $attempt++
        try {
            $params = @{
                Method      = $Method
                Uri         = $Uri
                ErrorAction = 'Stop'
                TimeoutSec  = $TimeoutSec
                UserAgent   = "ISV|LiebenConsultancy|M365AutoLink|$ScriptVersion"
                Verbose     = $false
            }
            if($Headers) { $params.Headers = $Headers }
            if($PSBoundParameters.ContainsKey('Body') -and $null -ne $Body) {
                $params.Body = $Body
                $params.ContentType = $ContentType
            }
            return Invoke-RestMethod @params
        } catch {
            if($attempt -ge $MaxAttempts -or -not (Test-IsTransientHttpError -ErrorRecord $_)) {
                throw
            }
            $delay = Get-RetryAfterSeconds -ErrorRecord $_
            if($delay -le 0) { $delay = [math]::Min(15, 2 * $attempt) }
            Write-Log "Transient error on attempt $attempt/$MaxAttempts, retrying in $($delay)s: $($_.Exception.Message)" -Level "WARN"
            Start-Sleep -Seconds (1 + $delay)
        }
    }
}

function Get-ScopeForResource {
    # Faithful v2 translation of the v1 "resource=<url>" parameter: request every statically consented
    # permission for that resource (.default) plus a refresh token (offline_access).
    param([Parameter(Mandatory = $true)][string]$Resource)
    return ("{0}/.default offline_access" -f $Resource.TrimEnd('/'))
}

function Save-RefreshToken {
    # DPAPI-protected (current user) persistence of the refresh token. Only called when the token
    # actually changed (A1) so we are not encrypting + writing to disk on every single API call.
    param([Parameter(Mandatory = $true)][string]$RefreshToken)
    try {
        $tokenDir = [System.IO.Path]::GetDirectoryName($global:octo.TokenCachePath)
        if(!(Test-Path $tokenDir)){ New-Item -ItemType Directory -Path $tokenDir -Force | Out-Null }
        $secureToken = ConvertTo-SecureString -String $RefreshToken -AsPlainText -Force
        $credential = New-Object System.Management.Automation.PSCredential("RefreshToken", $secureToken)
        $credential | Export-Clixml -Path $global:octo.TokenCachePath -Force
    } catch {
        Write-Warning "Could not cache refresh token: $($_.Exception.Message)"
    }
}

function New-PkceCodeVerifier {
    # RFC 7636 code_verifier: 32 random bytes, base64url-encoded (43 chars, no padding).
    $bytes = New-Object byte[] 32
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try { $rng.GetBytes($bytes) } finally { $rng.Dispose() }
    return (([Convert]::ToBase64String($bytes)) -replace '\+','-' -replace '/','_' -replace '=','')
}

function New-PkceCodeChallenge {
    # S256 challenge = base64url(SHA256(code_verifier)).
    param([Parameter(Mandatory = $true)][string]$Verifier)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $hash = $sha.ComputeHash([System.Text.Encoding]::ASCII.GetBytes($Verifier))
    } finally { $sha.Dispose() }
    return (([Convert]::ToBase64String($hash)) -replace '\+','-' -replace '/','_' -replace '=','')
}

#region Helper Functions
function Set-M365ProcessDpiAwareness {
    # declare the process per-monitor-v2 DPI aware BEFORE any window is created, so WinForms renders
    # crisp (not bitmap-stretched/blurry) at 125-200% display scaling. Falls back through the older APIs
    # for down-level Windows, and is a no-op if already set.
    try {
        if(-not ("M365AutoLink.DpiNative" -as [type])) {
            Add-Type -Namespace "M365AutoLink" -Name "DpiNative" -MemberDefinition @"
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool SetProcessDpiAwarenessContext(System.IntPtr value);
[System.Runtime.InteropServices.DllImport("shcore.dll")]
public static extern int SetProcessDpiAwareness(int value);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool SetProcessDPIAware();
"@ -ErrorAction Stop
        }
    } catch { return }

    # DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = -4
    try { if([M365AutoLink.DpiNative]::SetProcessDpiAwarenessContext([System.IntPtr](-4))) { return } } catch {}
    # PROCESS_PER_MONITOR_DPI_AWARE = 2 (Windows 8.1+)
    try { if([M365AutoLink.DpiNative]::SetProcessDpiAwareness(2) -eq 0) { return } } catch {}
    # System-DPI aware (Vista+)
    try { [void][M365AutoLink.DpiNative]::SetProcessDPIAware() } catch {}
}

function Get-TotalItemCountStatus {
    param(
        [long]$TotalItemCount,
        [long]$Threshold = $totalItemCountWarningThreshold,
        [double]$WarningRatio = $totalItemCountWarningRatio
    )

    if($Threshold -le 0) { return "ok" }
    if($TotalItemCount -ge $Threshold) { return "over" }
    if($TotalItemCount -ge [long]($Threshold * $WarningRatio)) { return "approaching" }
    return "ok"
}

function Get-ItemCountSummaryText {
    param(
        [long]$TotalItemCount,
        [string]$Status,
        [long]$Threshold = $totalItemCountWarningThreshold
    )

    $formattedTotal = '{0:N0}' -f $TotalItemCount
    switch($Status) {
        "over" {
            return ("{0} / {1} items - over limit" -f $formattedTotal, ('{0:N0}' -f $Threshold))
        }
        "approaching" {
            return ("{0} / {1} items - approaching limit" -f $formattedTotal, ('{0:N0}' -f $Threshold))
        }
        default {
            return ("{0} items linked" -f $formattedTotal)
        }
    }
}

function Get-CleanedShortcutName {
    param([string]$Name)
    $cleanedName = $Name
    foreach($replacement in $linkNameReplacements) {
        $cleanedName = $cleanedName.Replace($replacement.Pattern, $replacement.Replacement)
    }
    $cleanedName = $cleanedName.Trim()
    if([string]::IsNullOrWhiteSpace($cleanedName)){
        $cleanedName = $Name.Trim()
    }
    return $cleanedName
}

function Get-SafeDriveItemName {
    param([string]$Name)

    $safeName = $Name
    # OneDrive/SharePoint invalid filename characters.
    $safeName = $safeName -replace '[\\/:*?"<>|]', '-'
    $safeName = $safeName.Trim()
    $safeName = $safeName.TrimEnd('.')

    if([string]::IsNullOrWhiteSpace($safeName)) {
        $safeName = "Unnamed Shortcut"
    }

    return $safeName
}

function Test-IsExcludedLibraryName {
    param([string]$ListName)

    if([string]::IsNullOrWhiteSpace($ListName)) { return $false }
    foreach($pattern in $excludedLibrariesByWildcard) {
        # PowerShell -like is case-insensitive and anchors both ends correctly, unlike the old
        # start-anchored regex which over-matched (e.g. "*/sites/pwa" also hit "/sites/pwa-archive").
        if($ListName -like $pattern) {
            return $true
        }
    }

    return $false
}

function Normalize-GuidString {
    param([string]$Value)

    if([string]::IsNullOrWhiteSpace($Value)) { return $null }
    try {
        return ([guid]$Value).ToString().ToLowerInvariant()
    } catch {
        return $Value.Trim('{}').ToLowerInvariant()
    }
}

function Get-DefaultUserConfig {
    return @{
        version = 1
        preferences = @{
            excludedSiteUrls = @()
            excludedLibraryKeys = @()
        }
        diagnostics = @{
            lastAlreadyExisting = @()
            totalItemCount = 0
            lastDesiredCount = 0
        }
        cache = @{
            staticExcludedLibraries = @()
        }
    }
}

function Get-NormalizedSiteUrl {
    param([string]$SiteUrl)

    if([string]::IsNullOrWhiteSpace($SiteUrl)) { return $null }
    return $SiteUrl.Trim().TrimEnd('/').ToLowerInvariant()
}

function Get-LocalOneDriveRootPath {
    $candidates = [System.Collections.Generic.List[string]]::new()

    foreach($envValue in @($env:OneDriveCommercial, $env:OneDriveConsumer, $env:OneDrive)) {
        if(-not [string]::IsNullOrWhiteSpace($envValue) -and -not $candidates.Contains([string]$envValue)) {
            $candidates.Add([string]$envValue)
        }
    }

    if($candidates.Count -eq 0) {
        try {
            $profilePath = [Environment]::GetFolderPath("UserProfile")
            foreach($dir in Get-ChildItem -Path $profilePath -Directory -ErrorAction SilentlyContinue) {
                if($dir.Name -like "OneDrive*") {
                    $fullPath = $dir.FullName
                    if(-not $candidates.Contains($fullPath)) {
                        $candidates.Add($fullPath)
                    }
                }
            }
        } catch {}
    }

    foreach($candidate in $candidates) {
        if(Test-Path $candidate) {
            return $candidate
        }
    }

    return $null
}

function Get-LocalShortcutFolderPath {
    param([string]$FolderName)

    if([string]::IsNullOrWhiteSpace($FolderName)) { return $null }
    $oneDriveRoot = Get-LocalOneDriveRootPath
    if([string]::IsNullOrWhiteSpace($oneDriveRoot)) { return $null }

    return [System.IO.Path]::Combine($oneDriveRoot, $FolderName)
}

function Get-WebUrlFromListPath {
    param([string]$ListPath)

    if([string]::IsNullOrWhiteSpace($ListPath)) { return $null }

    try {
        $listPathUri = [System.Uri]::new($ListPath)
        $rawPath = [System.Uri]::UnescapeDataString($listPathUri.AbsolutePath)
        $webPath = $rawPath -replace '/Forms/[^/]+\.aspx$', ''
        $webSegments = @($webPath.Trim('/').Split('/'))
        if($webSegments.Count -ge 2) {
            return ("{0}://{1}/{2}" -f $listPathUri.Scheme, $listPathUri.Host, ($webSegments[0..($webSegments.Count - 2)] -join '/'))
        }

        return ("{0}://{1}" -f $listPathUri.Scheme, $listPathUri.Host)
    } catch {
        return $null
    }
}

function Get-NormalizedLaunchModes {
    param([object]$LaunchModes)

    $normalizedLaunchModes = [System.Collections.Generic.List[string]]::new()

    foreach($launchMode in @($LaunchModes)) {
        $launchModeText = [string]$launchMode
        if([string]::IsNullOrWhiteSpace($launchModeText)) { continue }

        $canonicalLaunchMode = switch -Regex ($launchModeText.Trim()) {
            '^desktop$' { 'Desktop' }
            '^start\s*menu$' { 'Start Menu' }
            '^startmenu$' { 'Start Menu' }
            '^atlogon$' { 'AtLogon' }
            default { $null }
        }

        if([string]::IsNullOrWhiteSpace($canonicalLaunchMode)) { continue }
        if(-not $normalizedLaunchModes.Contains($canonicalLaunchMode)) {
            $normalizedLaunchModes.Add($canonicalLaunchMode)
        }
    }

    return @($normalizedLaunchModes)
}

function Get-RawScriptPath {
    # The physical path the script is currently executing from (a temporary location when deployed via Intune).
    if(-not [string]::IsNullOrWhiteSpace($PSCommandPath)) {
        return $PSCommandPath
    }

    try {
        if($MyInvocation.MyCommand -and -not [string]::IsNullOrWhiteSpace($MyInvocation.MyCommand.Path)) {
            return $MyInvocation.MyCommand.Path
        }
    } catch {}

    return $null
}

function Get-M365AutoLinkScriptPath {
    # Once self-deployment has resolved a permanent location, persistence must target that copy rather
    # than the (possibly temporary) location the current process is running from.
    if(-not [string]::IsNullOrWhiteSpace($script:effectiveScriptPath)) {
        return $script:effectiveScriptPath
    }

    return Get-RawScriptPath
}

function Get-DeployTargetPath {
    param([string]$DeployToPath)

    if([string]::IsNullOrWhiteSpace($DeployToPath)) { return $null }

    # Expand %NAME% style environment variables. $env:NAME style is already expanded at assignment time.
    $expandedPath = [System.Environment]::ExpandEnvironmentVariables($DeployToPath.Trim())
    if([string]::IsNullOrWhiteSpace($expandedPath)) { return $null }

    # A path ending in .ps1 is treated as the full target file; anything else is treated as a folder
    # into which the script is copied under its standard M365AutoLink.ps1 name.
    if($expandedPath.TrimEnd('\', '/').ToLowerInvariant().EndsWith('.ps1')) {
        $targetFile = $expandedPath
    } else {
        $targetFile = [System.IO.Path]::Combine($expandedPath, 'M365AutoLink.ps1')
    }

    try {
        return [System.IO.Path]::GetFullPath($targetFile)
    } catch {
        return $targetFile
    }
}

function Invoke-SelfDeployment {
    param([string]$DeployToPath)

    # Resolves the path persistence should target. With no deployment configured (or if it cannot be
    # performed) this is simply the current script path.
    $currentScriptPath = Get-RawScriptPath

    $targetScriptPath = Get-DeployTargetPath -DeployToPath $DeployToPath
    if([string]::IsNullOrWhiteSpace($targetScriptPath)) {
        return $currentScriptPath
    }

    if([string]::IsNullOrWhiteSpace($currentScriptPath)) {
        Write-Log "deployToPath is set but the current script path could not be resolved; skipping self-deployment." "WARN"
        return $currentScriptPath
    }

    $resolvedCurrentPath = $currentScriptPath
    try { $resolvedCurrentPath = [System.IO.Path]::GetFullPath($currentScriptPath) } catch {}

    # Already running from the deploy location: nothing to copy, and we must not rewrite the file each run.
    if([string]::Equals($resolvedCurrentPath, $targetScriptPath, [System.StringComparison]::OrdinalIgnoreCase)) {
        Write-Log "Running from the configured deploy location ($targetScriptPath); no copy needed." "INFO"
        return $targetScriptPath
    }

    try {
        $targetDirectory = [System.IO.Path]::GetDirectoryName($targetScriptPath)
        if(-not [string]::IsNullOrWhiteSpace($targetDirectory) -and -not (Test-Path -LiteralPath $targetDirectory)) {
            New-Item -ItemType Directory -Path $targetDirectory -Force | Out-Null
        }

        Copy-Item -LiteralPath $currentScriptPath -Destination $targetScriptPath -Force -ErrorAction Stop
        Write-Log "Deployed script copy to permanent location: $targetScriptPath" "SUCCESS"
        return $targetScriptPath
    } catch {
        Write-Log "Failed to deploy script to '$targetScriptPath', persistence will target the current location instead: $($_.Exception.Message)" "WARN"
        return $currentScriptPath
    }
}

function Test-PathUnderIntune {
    param([string]$Path)

    if([string]::IsNullOrWhiteSpace($Path)) { return $false }

    # Well-known locations the Intune Management Extension uses to stage and run PowerShell scripts and
    # Win32 app payloads. Anything running from here is temporary and gets cleaned up after execution.
    $intunePathFragments = @(
        'Microsoft Intune Management Extension\Policies\Scripts',
        'Microsoft Intune Management Extension\Content',
        '\IMECache\'
    )

    foreach($fragment in $intunePathFragments) {
        if($Path.IndexOf($fragment, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
            return $true
        }
    }

    return $false
}

function Test-RunningUnderIntune {
    # Intune enforces a timeout on PowerShell scripts and considers the deployment failed if the process
    # does not exit in time (which the tray "keep running" loop would trigger). Detect Intune either by
    # the temporary location the script is launched from, or by walking the parent process chain for the
    # Intune agent processes that spawn it.
    try {
        $rawScriptPath = Get-RawScriptPath
        if(Test-PathUnderIntune -Path $rawScriptPath) {
            return $true
        }
    } catch {}

    $intuneProcessNames = @(
        'intunemanagementextension',
        'agentexecutor',
        'microsoft.management.services.intunewindowsagent'
    )

    try {
        $currentProcessId = $PID
        $depth = 0
        while($currentProcessId -and $depth -lt 8) {
            $processInfo = Get-CimInstance -ClassName Win32_Process -Filter "ProcessId = $currentProcessId" -ErrorAction Stop
            if(-not $processInfo) { break }

            $parentProcessId = [int]$processInfo.ParentProcessId
            if($parentProcessId -le 0) { break }

            $parentInfo = Get-CimInstance -ClassName Win32_Process -Filter "ProcessId = $parentProcessId" -ErrorAction SilentlyContinue
            if(-not $parentInfo) { break }

            $parentName = [System.IO.Path]::GetFileNameWithoutExtension([string]$parentInfo.Name).ToLowerInvariant()
            if($intuneProcessNames -contains $parentName) {
                return $true
            }

            $currentProcessId = $parentProcessId
            $depth++
        }
    } catch {}

    return $false
}

function Start-DetachedM365AutoLinkRun {
    param([Parameter(Mandatory = $true)][string]$ScriptPath)

    if([string]::IsNullOrWhiteSpace($ScriptPath) -or -not (Test-Path -LiteralPath $ScriptPath)) {
        Write-Log "Cannot start a detached run, script path '$ScriptPath' is invalid or missing." "WARN"
        return $false
    }

    $powerShellExe = Get-PowerShellExecutablePath
    $launchCommand = Get-PowerShellLaunchCommand -ScriptPath $ScriptPath -PowerShellExe $powerShellExe -HiddenWindow
    $commandLine = '"{0}" {1}' -f $launchCommand.TargetPath, $launchCommand.Arguments

    # Preferred: spawn via WMI so the new process is owned by the WMI provider host and is NOT part of the
    # Intune agent's process/job tree. That way it keeps running (tray icon + mapping) after this
    # Intune-launched process exits immediately.
    try {
        $result = Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments @{ CommandLine = $commandLine } -ErrorAction Stop
        if($result -and $result.ReturnValue -eq 0 -and $result.ProcessId) {
            Write-Log "Started detached M365AutoLink run (PID $($result.ProcessId)) via WMI." "SUCCESS"
            return $true
        }
        Write-Log "WMI process creation returned code $($result.ReturnValue); falling back to Start-Process." "WARN"
    } catch {
        Write-Log "WMI process creation failed, falling back to Start-Process: $($_.Exception.Message)" "WARN"
    }

    try {
        Start-Process -FilePath $launchCommand.TargetPath -ArgumentList $launchCommand.Arguments -WindowStyle Hidden -ErrorAction Stop | Out-Null
        Write-Log "Started detached M365AutoLink run via Start-Process." "SUCCESS"
        return $true
    } catch {
        Write-Log "Failed to start a detached M365AutoLink run: $($_.Exception.Message)" "WARN"
        return $false
    }
}

function Get-PowerShellExecutablePath {
    $systemPowerShell = Join-Path -Path $env:SystemRoot -ChildPath 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if(Test-Path $systemPowerShell) {
        return $systemPowerShell
    }

    try {
        $powershellCommand = Get-Command powershell.exe -ErrorAction SilentlyContinue
        if($powershellCommand -and -not [string]::IsNullOrWhiteSpace($powershellCommand.Source)) {
            return $powershellCommand.Source
        }
    } catch {}

    return 'powershell.exe'
}

function Test-IsWindows11OrLater {
    # Windows 11 is build 22000+; used to gate the conhost --headless launch (D5).
    try { return ([Environment]::OSVersion.Version.Build -ge 22000) } catch { return $false }
}

function Get-PowerShellLaunchCommand {
    param(
        [Parameter(Mandatory = $true)][string]$ScriptPath,
        [Parameter(Mandatory = $true)][string]$PowerShellExe,
        [switch]$HiddenWindow
    )

    $arguments = '-NoLogo -NoProfile -ExecutionPolicy Bypass -Sta -File "{0}"' -f $ScriptPath
    if($HiddenWindow) {
        $arguments = '-NoLogo -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Sta -File "{0}"' -f $ScriptPath

        # on Windows 11, launching through "conhost.exe --headless" avoids the brief console window
        # flash that "-WindowStyle Hidden" alone still shows at logon.
        if(Test-IsWindows11OrLater) {
            $conhostPath = Join-Path -Path $env:SystemRoot -ChildPath 'System32\conhost.exe'
            if(Test-Path -LiteralPath $conhostPath) {
                return @{
                    TargetPath = $conhostPath
                    Arguments = ('--headless "{0}" {1}' -f $PowerShellExe, $arguments)
                }
            }
        }
    }

    return @{
        TargetPath = $PowerShellExe
        Arguments = $arguments
    }
}

function Get-ShortcutPathSuffix {
    param([string]$SiteUrl)

    if([string]::IsNullOrWhiteSpace($SiteUrl)) { return $null }

    try {
        $siteUri = [System.Uri]::new($SiteUrl)
        $segments = @($siteUri.AbsolutePath.Trim('/').Split('/') | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if($segments.Count -le 2) { return $null }

        $suffixSegments = @($segments | Select-Object -Skip 2)
        if($suffixSegments.Count -eq 0) { return $null }

        return ($suffixSegments -join ' - ')
    } catch {
        return $null
    }
}

function Get-UniqueShortcutName {
    param(
        [Parameter(Mandatory = $true)][string]$BaseName,
        [string]$SiteUrl,
        [System.Collections.Generic.HashSet[string]]$ExistingNames,
        [System.Collections.Generic.HashSet[string]]$ReservedNames
    )

    $baseShortcutName = Get-SafeDriveItemName -Name (Get-CleanedShortcutName -Name $BaseName)
    if([string]::IsNullOrWhiteSpace($baseShortcutName)) {
        $baseShortcutName = 'Shortcut'
    }

    $isNameInUse = $false
    if($ExistingNames -and $ExistingNames.Contains($baseShortcutName)) { $isNameInUse = $true }
    if($ReservedNames -and $ReservedNames.Contains($baseShortcutName)) { $isNameInUse = $true }

    if(-not $isNameInUse) {
        return $baseShortcutName
    }

    $suffix = Get-ShortcutPathSuffix -SiteUrl $SiteUrl
    if(-not [string]::IsNullOrWhiteSpace($suffix)) {
        $suffixedName = Get-SafeDriveItemName -Name ("{0} - {1}" -f $baseShortcutName, $suffix)
        $isSuffixNameInUse = $false
        if($ExistingNames -and $ExistingNames.Contains($suffixedName)) { $isSuffixNameInUse = $true }
        if($ReservedNames -and $ReservedNames.Contains($suffixedName)) { $isSuffixNameInUse = $true }

        if(-not $isSuffixNameInUse) {
            return $suffixedName
        }
    }

    $counter = 2
    while($counter -lt 1000) {
        $numberedName = Get-SafeDriveItemName -Name ("{0} ({1})" -f $baseShortcutName, $counter)
        $isNumberedNameInUse = $false
        if($ExistingNames -and $ExistingNames.Contains($numberedName)) { $isNumberedNameInUse = $true }
        if($ReservedNames -and $ReservedNames.Contains($numberedName)) { $isNumberedNameInUse = $true }

        if(-not $isNumberedNameInUse) {
            return $numberedName
        }

        $counter++
    }

    return ("{0} - {1}" -f $baseShortcutName, [Guid]::NewGuid().ToString('N').Substring(0, 6))
}

function Set-ShortcutLinkFile {
    param(
        [Parameter(Mandatory = $true)][string]$ShortcutPath,
        [Parameter(Mandatory = $true)][string]$Description,
        [switch]$Remove
    )

    if($Remove) {
        if(Test-Path $ShortcutPath) {
            Remove-Item -Path $ShortcutPath -Force -ErrorAction Stop
        }
        return
    }

    $scriptPath = Get-M365AutoLinkScriptPath
    if([string]::IsNullOrWhiteSpace($scriptPath)) {
        throw 'Could not resolve the current script path for shortcut creation.'
    }

    $launchCommand = Get-PowerShellLaunchCommand -ScriptPath $scriptPath -PowerShellExe (Get-PowerShellExecutablePath) -HiddenWindow

    $shortcutDirectory = Split-Path -Path $ShortcutPath -Parent
    if(-not (Test-Path $shortcutDirectory)) {
        New-Item -ItemType Directory -Path $shortcutDirectory -Force | Out-Null
    }

    $shell = New-Object -ComObject WScript.Shell
    try {
        $shortcut = $shell.CreateShortcut($ShortcutPath)
        $shortcut.TargetPath = $launchCommand.TargetPath
        $shortcut.Arguments = $launchCommand.Arguments
        $shortcut.WorkingDirectory = Split-Path -Path $scriptPath -Parent
        $shortcut.Description = $Description
        try {
            $shortcut.IconLocation = "$($launchCommand.TargetPath),0"
        } catch {}
        $shortcut.Save()
    } finally {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) } catch {}
    }
}

function Set-AtLogonPersistence {
    param(
        [Parameter(Mandatory = $true)][string]$Mode,
        [switch]$Remove
    )

    $taskName = 'M365AutoLink'
    $runKeyPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
    $startupShortcutPath = Join-Path -Path ([Environment]::GetFolderPath('Startup')) -ChildPath 'M365AutoLink.lnk'
    $scriptPath = Get-M365AutoLinkScriptPath
    $launchCommand = Get-PowerShellLaunchCommand -ScriptPath $scriptPath -PowerShellExe (Get-PowerShellExecutablePath) -HiddenWindow
    $commandLine = '"{0}" {1}' -f $launchCommand.TargetPath, $launchCommand.Arguments

    if($Remove) {
        $removedAny = $false
        try {
            if(Get-Command Unregister-ScheduledTask -ErrorAction SilentlyContinue) {
                Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction Stop | Out-Null
                $removedAny = $true
            } elseif(Get-Command schtasks.exe -ErrorAction SilentlyContinue) {
                $deleteArgs = @('/Delete', '/TN', $taskName, '/F')
                $deleteProcess = Start-Process -FilePath 'schtasks.exe' -ArgumentList $deleteArgs -Wait -PassThru -NoNewWindow -ErrorAction Stop
                if($deleteProcess.ExitCode -eq 0) { $removedAny = $true }
            }
        } catch {
            Write-Log "At-logon persistence task removal failed: $($_.Exception.Message)" "WARN"
        }

        try {
            if(Test-Path $runKeyPath) {
                Remove-ItemProperty -Path $runKeyPath -Name $taskName -ErrorAction Stop
                $removedAny = $true
            }
        } catch {
            Write-Log "At-logon Run-key removal failed: $($_.Exception.Message)" "WARN"
        }

        try {
            if(Test-Path $startupShortcutPath) {
                Remove-Item -Path $startupShortcutPath -Force -ErrorAction Stop
                $removedAny = $true
            }
        } catch {
            Write-Log "At-logon startup shortcut removal failed: $($_.Exception.Message)" "WARN"
        }

        return $removedAny
    }

    if(Get-Command Register-ScheduledTask -ErrorAction SilentlyContinue) {
        try {
            $principalUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
            $action = New-ScheduledTaskAction -Execute $launchCommand.TargetPath -Argument $launchCommand.Arguments
            $trigger = New-ScheduledTaskTrigger -AtLogOn -User $principalUser
            $principal = New-ScheduledTaskPrincipal -UserId $principalUser -LogonType Interactive -RunLevel Limited
            # the default 72h ExecutionTimeLimit would kill the long-lived tray process; disable it.
            # MultipleInstances IgnoreNew avoids a second logon-triggered instance racing the running one.
            $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -StartWhenAvailable -ExecutionTimeLimit ([TimeSpan]::Zero) -MultipleInstances IgnoreNew
            $task = New-ScheduledTask -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description 'Launch M365AutoLink at logon'
            Register-ScheduledTask -TaskName $taskName -InputObject $task -Force -ErrorAction Stop | Out-Null
            return 'ScheduledTask'
        } catch {
            Write-Log "Scheduled task persistence is unavailable, falling back: $($_.Exception.Message)" "WARN"
        }
    }

    if(Test-Path $runKeyPath) {
        try {
            New-ItemProperty -Path $runKeyPath -Name $taskName -Value $commandLine -PropertyType String -Force -ErrorAction Stop | Out-Null
            return 'RunKey'
        } catch {
            Write-Log "Run-key persistence is unavailable, falling back: $($_.Exception.Message)" "WARN"
        }
    }

    try {
        Set-ShortcutLinkFile -ShortcutPath $startupShortcutPath -Description 'Launch M365AutoLink at logon'
        return 'StartupShortcut'
    } catch {
        Write-Log "Startup-folder persistence is unavailable: $($_.Exception.Message)" "WARN"
    }

    return $null
}

function Sync-LaunchPersistence {
    $configuredLaunchModes = @($LaunchModes)

    $normalizedLaunchModes = @(Get-NormalizedLaunchModes -LaunchModes $configuredLaunchModes)

    $launchModeSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach($launchMode in $normalizedLaunchModes) {
        [void]$launchModeSet.Add([string]$launchMode)
    }

    $desktopShortcutPath = Join-Path -Path ([Environment]::GetFolderPath('DesktopDirectory')) -ChildPath 'M365AutoLink.lnk'
    $startMenuShortcutPath = Join-Path -Path ([Environment]::GetFolderPath('Programs')) -ChildPath 'M365AutoLink.lnk'

    try {
        if($launchModeSet.Contains('Desktop')) {
            Set-ShortcutLinkFile -ShortcutPath $desktopShortcutPath -Description 'Launch M365AutoLink from the desktop'
            Write-Log "Desktop launch shortcut is configured" "INFO"
        } else {
            Set-ShortcutLinkFile -ShortcutPath $desktopShortcutPath -Description 'Launch M365AutoLink from the desktop' -Remove
        }
    } catch {
        Write-Log "Desktop launch shortcut sync failed: $($_.Exception.Message)" "WARN"
    }

    try {
        if($launchModeSet.Contains('Start Menu')) {
            Set-ShortcutLinkFile -ShortcutPath $startMenuShortcutPath -Description 'Launch M365AutoLink from the Start Menu'
            Write-Log "Start Menu launch shortcut is configured" "INFO"
        } else {
            Set-ShortcutLinkFile -ShortcutPath $startMenuShortcutPath -Description 'Launch M365AutoLink from the Start Menu' -Remove
        }
    } catch {
        Write-Log "Start Menu launch shortcut sync failed: $($_.Exception.Message)" "WARN"
    }

    try {
        if($launchModeSet.Contains('AtLogon')) {
            $configuredMethod = Set-AtLogonPersistence -Mode 'AtLogon'
            if([string]::IsNullOrWhiteSpace($configuredMethod)) {
                Write-Log 'At-logon persistence could not be configured on this device.' "WARN"
            } else {
                Write-Log "At-logon persistence configured via $configuredMethod" "INFO"
            }
        } else {
            $null = Set-AtLogonPersistence -Mode 'AtLogon' -Remove
        }
    } catch {
        Write-Log "At-logon persistence sync failed: $($_.Exception.Message)" "WARN"
    }
}

function Invoke-Uninstall {
    param([string]$DeployToPath)

    Write-Log "=== M365AutoLink Uninstall requested ===" "INFO"

    $desktopShortcutPath = Join-Path -Path ([Environment]::GetFolderPath('DesktopDirectory')) -ChildPath 'M365AutoLink.lnk'
    $startMenuShortcutPath = Join-Path -Path ([Environment]::GetFolderPath('Programs')) -ChildPath 'M365AutoLink.lnk'

    # Remove the Desktop + Start Menu launch shortcuts regardless of the configured $LaunchModes.
    foreach($shortcut in @(
        @{ Path = $desktopShortcutPath; Label = 'Desktop launch shortcut' },
        @{ Path = $startMenuShortcutPath; Label = 'Start Menu launch shortcut' }
    )) {
        try {
            if(Test-Path -LiteralPath $shortcut.Path) {
                Remove-Item -LiteralPath $shortcut.Path -Force -ErrorAction Stop
                Write-Log "Removed $($shortcut.Label)" "INFO"
            }
        } catch {
            Write-Log "Failed to remove $($shortcut.Label): $($_.Exception.Message)" "WARN"
        }
    }

    # Remove at-logon persistence (scheduled task + Run key + startup-folder shortcut).
    try {
        $removedAtLogon = Set-AtLogonPersistence -Mode 'AtLogon' -Remove
        if($removedAtLogon) {
            Write-Log "Removed at-logon persistence" "INFO"
        } else {
            Write-Log "No at-logon persistence found to remove" "INFO"
        }
    } catch {
        Write-Log "Failed to remove at-logon persistence: $($_.Exception.Message)" "WARN"
    }

    # always remove M365AutoLink's own credential + log files. RefreshToken.xml is a usable,
    # long-lived credential and must never survive an uninstall. Deleting the cache is also the
    # practical way to end the session (there is no refresh-token logout endpoint).
    foreach($ownFile in @(
        @{ Path = $global:octo.TokenCachePath; Label = 'refresh token cache' },
        @{ Path = $global:octo.LogPath; Label = 'run log' }
    )) {
        try {
            if($ownFile.Path -and (Test-Path -LiteralPath $ownFile.Path)) {
                Remove-Item -LiteralPath $ownFile.Path -Force -ErrorAction Stop
                Write-Log "Removed $($ownFile.Label): $($ownFile.Path)" "INFO"
            }
        } catch {
            Write-Log "Failed to remove $($ownFile.Label) '$($ownFile.Path)': $($_.Exception.Message)" "WARN"
        }
    }
    # Also drop the in-memory token so nothing re-writes the cache during this process.
    $global:octo.LCRefreshToken = $Null
    $global:octo.LCCachedTokens = @{}

    # Remove the deployed copy of the script from disk - only when a deploy location is configured.
    # A running .ps1 is not locked on Windows, so this also works when the current process IS that copy.
    $targetScriptPath = Get-DeployTargetPath -DeployToPath $DeployToPath
    if([string]::IsNullOrWhiteSpace($targetScriptPath)) {
        Write-Log "deployToPath is not configured; leaving the script file in place." "INFO"
    } else {
        try {
            if(Test-Path -LiteralPath $targetScriptPath) {
                Remove-Item -LiteralPath $targetScriptPath -Force -ErrorAction Stop
                Write-Log "Removed deployed script: $targetScriptPath" "SUCCESS"
            } else {
                Write-Log "Deployed script not found at $targetScriptPath; nothing to remove." "INFO"
            }

            # Clean up the deploy folder too, but only when it is now empty so we never delete
            # unrelated data (e.g. the token cache / log when they share the folder).
            $targetDirectory = [System.IO.Path]::GetDirectoryName($targetScriptPath)
            if(-not [string]::IsNullOrWhiteSpace($targetDirectory) -and (Test-Path -LiteralPath $targetDirectory)) {
                if(-not (Get-ChildItem -LiteralPath $targetDirectory -Force -ErrorAction SilentlyContinue)) {
                    Remove-Item -LiteralPath $targetDirectory -Force -ErrorAction Stop
                    Write-Log "Removed empty deploy folder: $targetDirectory" "INFO"
                }
            }
        } catch {
            Write-Log "Failed to remove deployed script '$targetScriptPath': $($_.Exception.Message)" "WARN"
        }
    }

    Write-Log "=== M365AutoLink Uninstall complete ===" "SUCCESS"
}

function Set-RoundedFormRegion {
    param(
        [Parameter(Mandatory = $true)]$Form,
        [int]$Radius = 10
    )

    if($Radius -lt 2) { $Radius = 2 }

    $applyRegion = {
        param($targetForm, $cornerRadius)

        if(-not $targetForm -or $targetForm.IsDisposed) { return }
        if($targetForm.ClientSize.Width -lt 4 -or $targetForm.ClientSize.Height -lt 4) { return }

        $path = New-Object Drawing.Drawing2D.GraphicsPath
        $diameter = $cornerRadius * 2
        $width = $targetForm.ClientSize.Width
        $height = $targetForm.ClientSize.Height

        $path.AddArc(0, 0, $diameter, $diameter, 180, 90)
        $path.AddArc($width - $diameter, 0, $diameter, $diameter, 270, 90)
        $path.AddArc($width - $diameter, $height - $diameter, $diameter, $diameter, 0, 90)
        $path.AddArc(0, $height - $diameter, $diameter, $diameter, 90, 90)
        $path.CloseFigure()

        if($targetForm.Region) {
            try { $targetForm.Region.Dispose() } catch {}
        }
        $targetForm.Region = New-Object Drawing.Region($path)
        $path.Dispose()
    }

    & $applyRegion $Form $Radius
}

function Enable-FormDrag {
    param(
        [Parameter(Mandatory = $true)]$Form,
        [Parameter(Mandatory = $true)][array]$DragControls
    )

    try {
        if(-not ("Win32.NativeMethods" -as [type])) {
            Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

namespace Win32 {
    public static class NativeMethods {
        [DllImport("user32.dll")]
        public static extern bool ReleaseCapture();

        [DllImport("user32.dll")]
        public static extern IntPtr SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
    }
}
"@ -Language CSharp -ErrorAction Stop
        }
    } catch {
        Write-Log "Enable-FormDrag initialization failed: $($_.Exception.Message)" "WARN"
        return
    }

    foreach($control in $DragControls) {
        if($null -eq $control) { continue }
        $control.Add_MouseDown({
            param($sender, $e)
            if($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
                # resolve the form from the sender at event time; the captured $Form is out of scope
                # once Enable-FormDrag has returned, so dragging would otherwise fail silently.
                $topLevelForm = $null
                try { $topLevelForm = $sender.FindForm() } catch {}
                if($null -eq $topLevelForm){
                    try { $topLevelForm = ($sender -as [System.Windows.Forms.Control]).TopLevelControl } catch {}
                }
                if($null -ne $topLevelForm){
                    [void][Win32.NativeMethods]::ReleaseCapture()
                    [void][Win32.NativeMethods]::SendMessage($topLevelForm.Handle, 0xA1, 0x2, 0)
                }
            }
        })
    }
}

function Invoke-GraphRaw {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('GET','POST','PATCH','DELETE','PUT')][string]$Method,
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $false)]$Body,
        [string]$ContentType = 'application/json; charset=utf-8'
    )

    $token = Get-AccessToken -resource $global:octo.graphUrl
    $headers = @{ Authorization = "Bearer $token" }

    # config load/save (used at the start and end of every run) now survives a transient 429/5xx
    # instead of dying, because it goes through the shared retry core rather than a bare Invoke-RestMethod.
    if($PSBoundParameters.ContainsKey('Body')) {
        return Invoke-RestWithRetry -Method $Method -Uri $Uri -Headers $headers -Body $Body -ContentType $ContentType
    }

    return Invoke-RestWithRetry -Method $Method -Uri $Uri -Headers $headers
}

function Get-OneDriveFolder {
    param(
        [Parameter(Mandatory = $true)][string]$FolderPath,
        [Parameter(Mandatory = $true)][string]$FolderName,
        [Parameter(Mandatory = $true)][string]$ParentChildrenUri
    )

    $folderUri = "$($global:octo.graphUrl)/v1.0/me/drive/root:/$FolderPath"
    try {
        return Invoke-GraphRaw -Method GET -Uri $folderUri
    } catch {
        $statusCode = $null
        try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {}
        if($statusCode -ne 404) { throw }
    }

    $folderBody = @{
        name = $FolderName
        folder = @{}
        "@microsoft.graph.conflictBehavior" = "replace"
    } | ConvertTo-Json -Depth 3

    [void](Invoke-GraphRaw -Method POST -Uri $ParentChildrenUri -Body $folderBody)
    return Invoke-GraphRaw -Method GET -Uri $folderUri
}

function ConvertTo-UserConfig {
    param($ConfigObject)

    $defaultConfig = Get-DefaultUserConfig
    if($null -eq $ConfigObject) { return $defaultConfig }

    $config = @{
        version = 1
        preferences = @{
            excludedSiteUrls = @()
            excludedLibraryKeys = @()
        }
        diagnostics = @{
            lastAlreadyExisting = @()
            totalItemCount = 0
            lastDesiredCount = 0
        }
        cache = @{
            staticExcludedLibraries = @()
        }
    }

    try { if($ConfigObject.version) { $config.version = [int]$ConfigObject.version } } catch {}

    $excluded = @()
    try {
        if($ConfigObject.preferences -and $ConfigObject.preferences.excludedSiteUrls) {
            $excluded = @($ConfigObject.preferences.excludedSiteUrls)
        }
    } catch {}

    $normalizedExcluded = [System.Collections.Generic.List[string]]::new()
    foreach($siteUrl in $excluded) {
        $normalizedSiteUrl = Get-NormalizedSiteUrl -SiteUrl ([string]$siteUrl)
        if(-not [string]::IsNullOrWhiteSpace($normalizedSiteUrl) -and -not $normalizedExcluded.Contains($normalizedSiteUrl)) {
            $normalizedExcluded.Add($normalizedSiteUrl)
        }
    }
    $config.preferences.excludedSiteUrls = @($normalizedExcluded)

    $excludedLibraries = @()
    try {
        if($ConfigObject.preferences -and $ConfigObject.preferences.excludedLibraryKeys) {
            $excludedLibraries = @($ConfigObject.preferences.excludedLibraryKeys)
        }
    } catch {}

    $normalizedExcludedLibraries = [System.Collections.Generic.List[string]]::new()
    foreach($libraryKey in $excludedLibraries) {
        $libraryKeyText = ([string]$libraryKey).Trim().ToLowerInvariant()
        if(-not [string]::IsNullOrWhiteSpace($libraryKeyText) -and -not $normalizedExcludedLibraries.Contains($libraryKeyText)) {
            $normalizedExcludedLibraries.Add($libraryKeyText)
        }
    }
    $config.preferences.excludedLibraryKeys = @($normalizedExcludedLibraries)

    $alreadyExisting = @()
    try {
        if($ConfigObject.diagnostics -and $ConfigObject.diagnostics.lastAlreadyExisting) {
            $alreadyExisting = @($ConfigObject.diagnostics.lastAlreadyExisting)
        }
    } catch {}

    $normalizedExisting = [System.Collections.Generic.List[hashtable]]::new()
    foreach($entry in $alreadyExisting) {
        $entryItemCount = 0
        try { $entryItemCount = [long]$entry.itemCount } catch {}
        $normalizedExisting.Add(@{
            siteUrl = [string]$entry.siteUrl
            listName = [string]$entry.listName
            reason = [string]$entry.reason
            timestamp = [string]$entry.timestamp
            itemCount = $entryItemCount
        })
    }
    $config.diagnostics.lastAlreadyExisting = @($normalizedExisting)

    try { if($ConfigObject.diagnostics -and $null -ne $ConfigObject.diagnostics.totalItemCount) { $config.diagnostics.totalItemCount = [long]$ConfigObject.diagnostics.totalItemCount } } catch {}
    try { if($ConfigObject.diagnostics -and $null -ne $ConfigObject.diagnostics.lastDesiredCount) { $config.diagnostics.lastDesiredCount = [int]$ConfigObject.diagnostics.lastDesiredCount } } catch {}

    $staticExcludedLibraries = @()
    try {
        if($ConfigObject.cache -and $ConfigObject.cache.staticExcludedLibraries) {
            $staticExcludedLibraries = @($ConfigObject.cache.staticExcludedLibraries)
        }
    } catch {}

    $normalizedStaticExcludedLibraries = [System.Collections.Generic.List[hashtable]]::new()
    $seenStaticLibraryKeys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach($entry in $staticExcludedLibraries) {
        $entryKey = [string]$entry.key
        if([string]::IsNullOrWhiteSpace($entryKey)) {
            $entryKey = Get-ShortcutTargetKey -SiteId ([string]$entry.siteId) -WebId ([string]$entry.webId) -ListId ([string]$entry.listId)
        }

        if([string]::IsNullOrWhiteSpace($entryKey)) { continue }
        if(-not $seenStaticLibraryKeys.Add($entryKey)) { continue }

        $normalizedStaticExcludedLibraries.Add(@{
            key = $entryKey
            siteId = [string]$entry.siteId
            webId = [string]$entry.webId
            listId = [string]$entry.listId
            listName = [string]$entry.listName
            reason = [string]$entry.reason
            lastSeenUtc = [string]$entry.lastSeenUtc
        })
    }
    $config.cache.staticExcludedLibraries = @($normalizedStaticExcludedLibraries)

    return $config
}

function Save-OneDriveUserConfig {
    param([Parameter(Mandatory = $true)]$Config)

    $configJson = $Config | ConvertTo-Json -Depth 8
    $contentUri = "$($global:octo.graphUrl)/v1.0/me/drive/root:/Apps/M365AutoLink/config.json:/content"
    [void](Invoke-GraphRaw -Method PUT -Uri $contentUri -Body $configJson -ContentType 'application/json; charset=utf-8')
}

function Get-OneDriveUserConfig {
    $defaultConfig = Get-DefaultUserConfig

    [void](Get-OneDriveFolder -FolderPath "Apps" -FolderName "Apps" -ParentChildrenUri "$($global:octo.graphUrl)/v1.0/me/drive/root/children")
    [void](Get-OneDriveFolder -FolderPath "Apps/M365AutoLink" -FolderName "M365AutoLink" -ParentChildrenUri "$($global:octo.graphUrl)/v1.0/me/drive/root:/Apps:/children")

    $contentUri = "$($global:octo.graphUrl)/v1.0/me/drive/root:/Apps/M365AutoLink/config.json:/content"
    $rawConfig = $null
    try {
        $rawConfig = Invoke-GraphRaw -Method GET -Uri $contentUri
    } catch {
        $statusCode = $null
        try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {}
        if($statusCode -ne 404) { throw }

        Save-OneDriveUserConfig -Config $defaultConfig
        return $defaultConfig
    }

    $configObject = $null
    if($rawConfig -is [string]) {
        try {
            $configObject = $rawConfig | ConvertFrom-Json -ErrorAction Stop
        } catch {
            $configObject = $null
        }
    } else {
        $configObject = $rawConfig
    }

    $normalizedConfig = ConvertTo-UserConfig -ConfigObject $configObject
    return $normalizedConfig
}

function Get-DisplaySiteUrl {
    # Strips the "https://host/" prefix purely for on-screen display so the path is easier to scan.
    param([string]$SiteUrl)

    if([string]::IsNullOrWhiteSpace($SiteUrl)) { return $SiteUrl }
    $trimmed = $SiteUrl -replace '^https?://[^/]+', ''
    $trimmed = $trimmed.TrimStart('/')
    if([string]::IsNullOrWhiteSpace($trimmed)) { return $SiteUrl }
    return $trimmed
}

function Get-ManageDialogSizePath {
    # remember the Manage-shortcuts window size in a small local file (avoids OneDrive round-trips).
    $dir = [System.IO.Path]::GetDirectoryName($global:octo.LogPath)
    if([string]::IsNullOrWhiteSpace($dir)) { return $null }
    return (Join-Path -Path $dir -ChildPath "manage-ui.json")
}

function Get-SavedManageDialogSize {
    try {
        $path = Get-ManageDialogSizePath
        if($path -and (Test-Path -LiteralPath $path)) {
            $saved = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
            $width = [int]$saved.width
            $height = [int]$saved.height
            if($width -ge 700 -and $height -ge 400) {
                return @{ Width = $width; Height = $height }
            }
        }
    } catch {}
    return $null
}

function Save-ManageDialogSize {
    param([int]$Width, [int]$Height)
    try {
        $path = Get-ManageDialogSizePath
        if([string]::IsNullOrWhiteSpace($path)) { return }
        $dir = [System.IO.Path]::GetDirectoryName($path)
        if(-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        (@{ width = $Width; height = $Height } | ConvertTo-Json) | Set-Content -LiteralPath $path -Encoding UTF8
    } catch {}
}

function Show-InfoDialog {
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][string]$Message
    )

    try {
        Add-Type -AssemblyName System.Windows.Forms
        [void][System.Windows.Forms.MessageBox]::Show($Message, $Title, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } catch {
        Update-TrayState -ShowBalloon -BalloonTitle $Title -BalloonMessage $Message -BalloonIcon "Info"
    }
}

function Invoke-ManageShortcuts {
    if(-not $script:lastMappedLibraryOptions -or @($script:lastMappedLibraryOptions).Count -eq 0) {
        Show-InfoDialog -Title "M365AutoLink" -Message "No shortcuts to manage yet.`r`n`r`nRun a mapping first, then open Manage shortcuts again."
        return
    }

    try {
        Update-TrayState -Text "M365AutoLink - Loading config" -ProgressText "Opening shortcut manager"
        if(-not $script:userConfig) {
            $script:userConfig = Get-OneDriveUserConfig
        }

        # Originally-excluded libraries = those the run marked isExcluded (covers both the per-library
        # list and any legacy per-site exclusions that were applied during the run).
        $originalSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach($option in @($script:lastMappedLibraryOptions)) {
            if($option.isExcluded) {
                $optionKey = ([string]$option.key).Trim().ToLowerInvariant()
                if(-not [string]::IsNullOrWhiteSpace($optionKey)) { [void]$originalSet.Add($optionKey) }
            }
        }

        $selectionResult = Show-ManageShortcutsDialog -LibraryOptions @($script:lastMappedLibraryOptions)
        if($selectionResult.isCanceled) {
            Update-TrayState -Text "M365AutoLink - Idle" -ProgressText "No changes"
            return
        }

        $chosenSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach($libraryKey in @($selectionResult.excludedLibraryKeys)) {
            $keyText = ([string]$libraryKey).Trim().ToLowerInvariant()
            if(-not [string]::IsNullOrWhiteSpace($keyText)) { [void]$chosenSet.Add($keyText) }
        }

        $exclusionsChanged = -not $originalSet.SetEquals($chosenSet)
        # Even with no exclusion change, migrate any legacy per-site exclusions into the per-library
        # model so they stop being a separate, invisible mechanism.
        $hasLegacySiteExclusions = (@($script:userConfig.preferences.excludedSiteUrls).Count -gt 0)

        if(-not $exclusionsChanged -and -not $hasLegacySiteExclusions) {
            Update-TrayState -Text "M365AutoLink - Idle" -ProgressText "No changes" -ShowBalloon -BalloonMessage "No changes to apply." -BalloonIcon "Info"
            return
        }

        $script:userConfig.preferences.excludedLibraryKeys = @($chosenSet)
        $script:userConfig.preferences.excludedSiteUrls = @()
        Save-OneDriveUserConfig -Config $script:userConfig

        if($exclusionsChanged) {
            # Exclusions changed: re-run automatically instead of asking the user to click Run now.
            # This shrink is intentional, so let the upcoming run bypass the deletion ratio safety guard.
            $script:bypassDeletionRatioOnce = $true
            if($script:traySync) { $script:traySync.RequestRerun = $true }
            Update-TrayState -Text "M365AutoLink - Applying changes" -ProgressText "Re-running" -ShowBalloon -BalloonMessage "Saved $($chosenSet.Count) excluded librar$(if($chosenSet.Count -eq 1){'y'}else{'ies'}). Re-running now to apply..." -BalloonIcon "Info"
        } else {
            Update-TrayState -Text "M365AutoLink - Idle" -ProgressText "Saved" -ShowBalloon -BalloonMessage "Exclusions saved." -BalloonIcon "Info"
        }
    } catch {
        Write-Log "Failed to open/save shortcut manager: $($_.Exception.Message)" "ERROR"
        Update-TrayState -Text "M365AutoLink - Error" -ProgressText "Failed to update shortcuts" -ShowBalloon -BalloonMessage $_.Exception.Message -BalloonIcon "Error"
    }
}

function Show-ManageShortcutsDialog {
    param(
        [Parameter(Mandatory = $true)][array]$LibraryOptions
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $excludedForeColor = [Drawing.Color]::FromArgb(150, 158, 168)

    # Build the master item list once; the visible ListView is (re)populated from this by filter + sort.
    $allItems = [System.Collections.Generic.List[object]]::new()

    foreach($option in $LibraryOptions) {
        $libraryValue = [string]$option.listName
        if([string]::IsNullOrWhiteSpace($libraryValue)) { $libraryValue = "-" }
        $optionItemCount = [long]0
        try { $optionItemCount = [long]$option.itemCount } catch {}
        $isExcluded = [bool]$option.isExcluded

        $itemsValue = if($isExcluded) { "-" } elseif($optionItemCount -gt 0) { '{0:N0}' -f $optionItemCount } else { "0" }
        $statusValue = if($isExcluded) { "Excluded" } else { "Linked" }
        $reasonValue = if($isExcluded) { "Excluded by you" } else { "" }

        $item = New-Object Windows.Forms.ListViewItem("")
        [void]$item.SubItems.Add($libraryValue)
        [void]$item.SubItems.Add((Get-DisplaySiteUrl -SiteUrl ([string]$option.siteUrl)))
        [void]$item.SubItems.Add($itemsValue)
        [void]$item.SubItems.Add($statusValue)
        [void]$item.SubItems.Add($reasonValue)
        $item.ToolTipText = [string]$option.siteUrl
        $item.Tag = @{ key = [string]$option.key; itemCount = $optionItemCount; autoSkipped = $false }
        $item.Checked = $isExcluded
        if($isExcluded) { $item.ForeColor = $excludedForeColor }
        $allItems.Add($item)
    }

    $form = New-Object Windows.Forms.Form
    $form.Text = "M365AutoLink - Manage shortcuts"
    $form.StartPosition = "CenterScreen"
    $form.AutoScaleMode = [Windows.Forms.AutoScaleMode]::None
    # a normal sizable window (resizable + remembers its size), rather than the old borderless one.
    $form.FormBorderStyle = [Windows.Forms.FormBorderStyle]::Sizable
    $form.MaximizeBox = $true
    $form.MinimumSize = New-Object Drawing.Size(760, 460)
    $savedSize = Get-SavedManageDialogSize
    if($savedSize) {
        $form.ClientSize = New-Object Drawing.Size([int]$savedSize.Width, [int]$savedSize.Height)
    } else {
        $form.ClientSize = New-Object Drawing.Size(1040, 600)
    }
    $form.BackColor = [Drawing.Color]::FromArgb(246, 248, 252)

    $pad = 12
    $footerH = 48
    $filterH = 30
    $clientW = $form.ClientSize.Width
    $clientH = $form.ClientSize.Height
    $contentWidth = $clientW - ($pad * 2)
    $rightAnchor = [Windows.Forms.AnchorStyles]::Top -bor [Windows.Forms.AnchorStyles]::Left -bor [Windows.Forms.AnchorStyles]::Right

    $headerPanel = New-Object Windows.Forms.Panel
    $headerPanel.Location = New-Object Drawing.Point(0, 0)
    $headerPanel.Size = New-Object Drawing.Size($clientW, 66)
    $headerPanel.BackColor = [Drawing.Color]::FromArgb(33, 37, 43)
    $headerPanel.Anchor = $rightAnchor

    $titleLabel = New-Object Windows.Forms.Label
    $titleLabel.Location = New-Object Drawing.Point($pad, 9)
    $titleLabel.Size = New-Object Drawing.Size(($clientW - ($pad * 2)), 22)
    $titleLabel.Font = New-Object Drawing.Font("Segoe UI", 11, [Drawing.FontStyle]::Bold)
    $titleLabel.ForeColor = [Drawing.Color]::FromArgb(237, 244, 252)
    $titleLabel.Text = "Manage shortcuts"
    $titleLabel.Anchor = $rightAnchor

    $subLabel = New-Object Windows.Forms.Label
    $subLabel.Location = New-Object Drawing.Point($pad, 34)
    $subLabel.Size = New-Object Drawing.Size(($clientW - ($pad * 2)), 26)
    $subLabel.Font = New-Object Drawing.Font("Segoe UI", 9)
    $subLabel.ForeColor = [Drawing.Color]::FromArgb(191, 205, 223)
    $subLabel.Text = "Tick Exclude to stop syncing a library. Type to filter, click a column to sort. Saving re-runs automatically."
    $subLabel.Anchor = $rightAnchor

    $headerPanel.Controls.Add($titleLabel)
    $headerPanel.Controls.Add($subLabel)

    # Filter row: a search box plus quick Exclude-all / Include-all buttons.
    $filterLabel = New-Object Windows.Forms.Label
    $filterLabel.Location = New-Object Drawing.Point($pad, 74)
    $filterLabel.Size = New-Object Drawing.Size(40, $filterH)
    $filterLabel.Text = "Filter"
    $filterLabel.TextAlign = [Drawing.ContentAlignment]::MiddleLeft
    $filterLabel.Font = New-Object Drawing.Font("Segoe UI", 9)

    $filterBox = New-Object Windows.Forms.TextBox
    $filterBox.Location = New-Object Drawing.Point(($pad + 44), 76)
    $filterBox.Size = New-Object Drawing.Size(($contentWidth - 44 - 220), 24)
    $filterBox.Font = New-Object Drawing.Font("Segoe UI", 9)
    $filterBox.Anchor = $rightAnchor

    $excludeAllButton = New-Object Windows.Forms.Button
    $excludeAllButton.Text = "Exclude all"
    $excludeAllButton.Size = New-Object Drawing.Size(100, 26)
    $excludeAllButton.Location = New-Object Drawing.Point(($clientW - $pad - 208), 75)
    $excludeAllButton.FlatStyle = [Windows.Forms.FlatStyle]::Flat
    $excludeAllButton.BackColor = [Drawing.Color]::FromArgb(231, 236, 244)
    $excludeAllButton.Anchor = [Windows.Forms.AnchorStyles]::Top -bor [Windows.Forms.AnchorStyles]::Right

    $includeAllButton = New-Object Windows.Forms.Button
    $includeAllButton.Text = "Include all"
    $includeAllButton.Size = New-Object Drawing.Size(100, 26)
    $includeAllButton.Location = New-Object Drawing.Point(($clientW - $pad - 100), 75)
    $includeAllButton.FlatStyle = [Windows.Forms.FlatStyle]::Flat
    $includeAllButton.BackColor = [Drawing.Color]::FromArgb(231, 236, 244)
    $includeAllButton.Anchor = [Windows.Forms.AnchorStyles]::Top -bor [Windows.Forms.AnchorStyles]::Right

    # Capacity bar: shows how much of the sync "budget" the currently INCLUDED libraries consume.
    $capPanel = New-Object Windows.Forms.Panel
    $capPanel.Location = New-Object Drawing.Point($pad, 112)
    $capPanel.Size = New-Object Drawing.Size($contentWidth, 44)
    $capPanel.BackColor = [Drawing.Color]::FromArgb(246, 248, 252)
    $capPanel.Anchor = $rightAnchor

    $capLabel = New-Object Windows.Forms.Label
    $capLabel.Location = New-Object Drawing.Point(0, 0)
    $capLabel.Size = New-Object Drawing.Size($contentWidth, 18)
    $capLabel.Font = New-Object Drawing.Font("Segoe UI", 9, [Drawing.FontStyle]::Bold)
    $capLabel.TextAlign = [Drawing.ContentAlignment]::MiddleLeft
    $capLabel.Anchor = $rightAnchor

    $capTrack = New-Object Windows.Forms.Panel
    $capTrack.Location = New-Object Drawing.Point(0, 22)
    $capTrack.Size = New-Object Drawing.Size($contentWidth, 14)
    $capTrack.BackColor = [Drawing.Color]::FromArgb(225, 230, 238)
    $capTrack.Anchor = $rightAnchor

    $capFill = New-Object Windows.Forms.Panel
    $capFill.Location = New-Object Drawing.Point(0, 0)
    $capFill.Size = New-Object Drawing.Size(0, 14)
    $capFill.BackColor = [Drawing.Color]::FromArgb(31, 122, 49)
    $capTrack.Controls.Add($capFill)

    $capPanel.Controls.Add($capLabel)
    $capPanel.Controls.Add($capTrack)

    $listTop = 164
    $listView = New-Object Windows.Forms.ListView
    $listView.Location = New-Object Drawing.Point($pad, $listTop)
    $listView.Size = New-Object Drawing.Size($contentWidth, ($clientH - $listTop - $footerH - 8))
    $listView.View = [Windows.Forms.View]::Details
    $listView.FullRowSelect = $true
    $listView.GridLines = $true
    $listView.MultiSelect = $false
    $listView.CheckBoxes = $true
    $listView.ShowItemToolTips = $true
    $listView.Font = New-Object Drawing.Font("Segoe UI", 9)
    $listView.Anchor = [Windows.Forms.AnchorStyles]::Top -bor [Windows.Forms.AnchorStyles]::Bottom -bor [Windows.Forms.AnchorStyles]::Left -bor [Windows.Forms.AnchorStyles]::Right
    $siteColumnWidth = [Math]::Max(200, $contentWidth - 584 - 22)
    [void]$listView.Columns.Add("Exclude", 64)
    [void]$listView.Columns.Add("Library", 180)
    [void]$listView.Columns.Add("Site", $siteColumnWidth)
    [void]$listView.Columns.Add("Items", 100, [Windows.Forms.HorizontalAlignment]::Right)
    [void]$listView.Columns.Add("Status", 90)
    [void]$listView.Columns.Add("Reason", 150)

    # Suppress capacity recompute while we bulk-repopulate the list (filter/sort), then recompute once.
    $script:mgSuspend = $false

    $refreshCapacity = {
        if($script:mgSuspend) { return }
      try {
        $includedTotal = [long]0
        foreach($row in $listView.Items) {
            # During the ListView handle-creation ItemChecked storm the enumeration can briefly yield a
            # null/partial row; indexing $null.SubItems is what threw "Cannot index into a null array".
            if($null -eq $row -or $null -eq $row.SubItems -or $row.SubItems.Count -lt 6) { continue }
            if($row.Tag.autoSkipped) { continue }
            $rowCount = [long]0
            try { $rowCount = [long]$row.Tag.itemCount } catch {}
            if($row.Checked) {
                if([string]$row.SubItems[4].Text -ne "Excluded") { $row.SubItems[4].Text = "Excluded" }
                if([string]$row.SubItems[5].Text -ne "Excluded by you") { $row.SubItems[5].Text = "Excluded by you" }
                $row.ForeColor = $excludedForeColor
            } else {
                if([string]$row.SubItems[4].Text -ne "Linked") { $row.SubItems[4].Text = "Linked" }
                if([string]$row.SubItems[5].Text -ne "") { $row.SubItems[5].Text = "" }
                $row.ForeColor = $listView.ForeColor
                $includedTotal += $rowCount
            }
        }

        $status = Get-TotalItemCountStatus -TotalItemCount $includedTotal
        $accent = switch($status) {
            "over"        { [Drawing.Color]::FromArgb(196, 43, 28) }
            "approaching" { [Drawing.Color]::FromArgb(176, 110, 0) }
            default       { [Drawing.Color]::FromArgb(31, 122, 49) }
        }
        $capLabel.ForeColor = $accent
        $capFill.BackColor = $accent

        if($totalItemCountWarningThreshold -gt 0) {
            $ratio = [double]$includedTotal / [double]$totalItemCountWarningThreshold
            if($ratio -gt 1) { $ratio = 1 }
            if($ratio -lt 0) { $ratio = 0 }
            $capFill.Width = [int]($capTrack.Width * $ratio)
            $remaining = $totalItemCountWarningThreshold - $includedTotal
            if($remaining -lt 0) {
                $capLabel.Text = "{0:N0} of {1:N0} items synced  -  OVER by {2:N0}" -f $includedTotal, $totalItemCountWarningThreshold, [math]::Abs($remaining)
            } else {
                $capLabel.Text = "{0:N0} of {1:N0} items synced  -  {2:N0} remaining" -f $includedTotal, $totalItemCountWarningThreshold, $remaining
            }
        } else {
            $capFill.Width = 0
            $capLabel.Text = "{0:N0} items synced  (limit warning disabled)" -f $includedTotal
        }
      } catch {
        # Never let a stray handler error surface as a WinForms Continue/Quit ThreadException popup.
        Write-Log "Manage-shortcuts capacity refresh failed: $($_.Exception.Message)" "WARN"
      }
    }

    # (Re)build the visible rows from $allItems using the current filter text + sort column/direction.
    $applyView = {
      try {
        $filterText = ([string]$filterBox.Text).Trim().ToLowerInvariant()
        $rows = @($allItems)
        if(-not [string]::IsNullOrWhiteSpace($filterText)) {
            $rows = @($rows | Where-Object {
                (([string]$_.SubItems[1].Text).ToLowerInvariant().Contains($filterText)) -or
                (([string]$_.SubItems[2].Text).ToLowerInvariant().Contains($filterText)) -or
                (([string]$_.ToolTipText).ToLowerInvariant().Contains($filterText)) -or
                (([string]$_.SubItems[5].Text).ToLowerInvariant().Contains($filterText))
            })
        }

        $col = $script:mgSortColumn
        if($null -ne $col) {
            if($col -eq 0) {
                $rows = @($rows | Sort-Object @{ Expression = { [bool]$_.Checked } })
            } elseif($col -eq 3) {
                $rows = @($rows | Sort-Object @{ Expression = { [long]$_.Tag.itemCount } })
            } else {
                $rows = @($rows | Sort-Object @{ Expression = { [string]$_.SubItems[$col].Text } })
            }
            if(-not $script:mgSortAsc) { [array]::Reverse($rows) }
        }

        $script:mgSuspend = $true
        $listView.BeginUpdate()
        $listView.Items.Clear()
        foreach($r in $rows) { [void]$listView.Items.Add($r) }
        $listView.EndUpdate()
        $script:mgSuspend = $false
        & $refreshCapacity
      } catch {
        $script:mgSuspend = $false
        Write-Log "Manage-shortcuts view refresh failed: $($_.Exception.Message)" "WARN"
      }
    }

    $listView.Add_ItemChecked({ & $refreshCapacity })
    $listView.Add_ColumnClick({
        param($s, $e)
        if($script:mgSortColumn -eq $e.Column) { $script:mgSortAsc = -not $script:mgSortAsc }
        else { $script:mgSortColumn = $e.Column; $script:mgSortAsc = $true }
        & $applyView
    })
    $filterBox.Add_TextChanged({ & $applyView })

    $excludeAllButton.Add_Click({
        try {
            $script:mgSuspend = $true
            foreach($row in $listView.Items) { if(-not $row.Tag.autoSkipped -and -not $row.Checked) { $row.Checked = $true } }
        } catch {} finally { $script:mgSuspend = $false }
        & $refreshCapacity
    })
    $includeAllButton.Add_Click({
        try {
            $script:mgSuspend = $true
            foreach($row in $listView.Items) { if(-not $row.Tag.autoSkipped -and $row.Checked) { $row.Checked = $false } }
        } catch {} finally { $script:mgSuspend = $false }
        & $refreshCapacity
    })

    # Default sort: site/library ascending.
    $script:mgSortColumn = 2
    $script:mgSortAsc = $true
    & $applyView

    $footerPanel = New-Object Windows.Forms.Panel
    $footerPanel.Location = New-Object Drawing.Point(0, ($clientH - $footerH))
    $footerPanel.Size = New-Object Drawing.Size($clientW, $footerH)
    $footerPanel.BackColor = [Drawing.Color]::FromArgb(241, 245, 251)
    $footerPanel.Anchor = [Windows.Forms.AnchorStyles]::Bottom -bor [Windows.Forms.AnchorStyles]::Left -bor [Windows.Forms.AnchorStyles]::Right

    $buttonTop = [int](($footerH - 30) / 2)
    $saveButton = New-Object Windows.Forms.Button
    $saveButton.Text = "Save"
    $saveButton.Location = New-Object Drawing.Point(($clientW - $pad - 176), $buttonTop)
    $saveButton.Size = New-Object Drawing.Size(80, 30)
    $saveButton.FlatStyle = [Windows.Forms.FlatStyle]::Flat
    $saveButton.BackColor = [Drawing.Color]::FromArgb(0, 163, 255)
    $saveButton.ForeColor = [Drawing.Color]::White
    $saveButton.FlatAppearance.BorderSize = 0
    $saveButton.DialogResult = [Windows.Forms.DialogResult]::OK
    $saveButton.Anchor = [Windows.Forms.AnchorStyles]::Top -bor [Windows.Forms.AnchorStyles]::Right

    $cancelButton = New-Object Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object Drawing.Point(($clientW - $pad - 88), $buttonTop)
    $cancelButton.Size = New-Object Drawing.Size(80, 30)
    $cancelButton.FlatStyle = [Windows.Forms.FlatStyle]::Flat
    $cancelButton.BackColor = [Drawing.Color]::FromArgb(231, 236, 244)
    $cancelButton.ForeColor = [Drawing.Color]::FromArgb(33, 37, 43)
    $cancelButton.FlatAppearance.BorderColor = [Drawing.Color]::FromArgb(210, 218, 230)
    $cancelButton.DialogResult = [Windows.Forms.DialogResult]::Cancel
    $cancelButton.Anchor = [Windows.Forms.AnchorStyles]::Top -bor [Windows.Forms.AnchorStyles]::Right

    $form.AcceptButton = $saveButton
    $form.CancelButton = $cancelButton
    $footerPanel.Controls.Add($saveButton)
    $footerPanel.Controls.Add($cancelButton)

    $form.Controls.Add($headerPanel)
    $form.Controls.Add($filterLabel)
    $form.Controls.Add($filterBox)
    $form.Controls.Add($excludeAllButton)
    $form.Controls.Add($includeAllButton)
    $form.Controls.Add($capPanel)
    $form.Controls.Add($listView)
    $form.Controls.Add($footerPanel)
    $footerPanel.BringToFront()

    try {
        $dialogResult = $form.ShowDialog()
    } catch {
        Write-Log "Failed to open manage-shortcuts dialog: $($_.Exception.Message)" "ERROR"
        try { $form.Dispose() } catch {}
        return @{ isCanceled = $true; excludedLibraryKeys = @() }
    }

    # remember the window size for next time.
    try { Save-ManageDialogSize -Width $form.ClientSize.Width -Height $form.ClientSize.Height } catch {}

    if($dialogResult -ne [Windows.Forms.DialogResult]::OK) {
        $form.Dispose()
        return @{ isCanceled = $true; excludedLibraryKeys = @() }
    }

    $selected = [System.Collections.Generic.List[string]]::new()
    foreach($row in $allItems) {
        if($row.Tag.autoSkipped) { continue }
        if(-not $row.Checked) { continue }
        $rowKey = [string]$row.Tag.key
        if(-not [string]::IsNullOrWhiteSpace($rowKey) -and -not $selected.Contains($rowKey)) {
            $selected.Add($rowKey)
        }
    }

    $form.Dispose()
    return @{ isCanceled = $false; excludedLibraryKeys = @($selected) }
}

function Get-ShortcutTargetKey {
    param(
        [string]$SiteId,
        [string]$WebId,
        [string]$ListId
    )

    $normalizedSiteId = Normalize-GuidString -Value $SiteId
    if([string]::IsNullOrWhiteSpace($normalizedSiteId)) { $normalizedSiteId = [string]$SiteId }

    $normalizedWebId = Normalize-GuidString -Value $WebId
    if([string]::IsNullOrWhiteSpace($normalizedWebId)) { $normalizedWebId = [string]$WebId }

    $normalizedListId = Normalize-GuidString -Value $ListId
    if([string]::IsNullOrWhiteSpace($normalizedListId)) { $normalizedListId = [string]$ListId }

    if([string]::IsNullOrWhiteSpace($normalizedSiteId) -or [string]::IsNullOrWhiteSpace($normalizedWebId) -or [string]::IsNullOrWhiteSpace($normalizedListId)) {
        return $null
    }

    return "{0}|{1}|{2}" -f $normalizedSiteId.Trim('{}').ToLowerInvariant(), $normalizedWebId.Trim('{}').ToLowerInvariant(), $normalizedListId.Trim('{}').ToLowerInvariant()
}

function Get-ListFeatureId {
    param($ListMetadata)

    $featureIdCandidates = @(
        $ListMetadata.FeatureId,
        $ListMetadata.featureid,
        $ListMetadata.TemplateFeatureId,
        $ListMetadata.templatefeatureid
    )

    foreach($candidate in $featureIdCandidates) {
        $normalized = Normalize-GuidString -Value ([string]$candidate)
        if(-not [string]::IsNullOrWhiteSpace($normalized)) {
            return $normalized
        }
    }

    return $null
}

$script:ExcludedListTitleSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach($title in $ExcludedListTitles) {
    if(-not [string]::IsNullOrWhiteSpace($title)) {
        [void]$script:ExcludedListTitleSet.Add($title)
    }
}

$script:ExcludedFeatureIdSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach($featureId in $ExcludedListFeatureIDs) {
    $normalizedId = Normalize-GuidString -Value $featureId
    if(-not [string]::IsNullOrWhiteSpace($normalizedId)) {
        [void]$script:ExcludedFeatureIdSet.Add($normalizedId)
    }
}

function Invoke-RefreshTokenExchange {
    # Redeems the cached refresh token for an access token for a specific resource (v2 endpoint, B2).
    # Persists the refresh token only when Entra rotated it (A1), so the hot path never writes to disk.
    param([Parameter(Mandatory = $true)][string]$Resource)

    $body = @{
        client_id     = $global:octo.LCClientId
        grant_type    = "refresh_token"
        refresh_token = $global:octo.LCRefreshToken
        scope         = (Get-ScopeForResource -Resource $Resource)
    }

    $response = Invoke-RestMethod -Uri $global:octo.tokenUrl -Method POST -Body $body -ErrorAction Stop -Verbose:$false

    if($response.refresh_token -and $response.refresh_token -ne $global:octo.LCRefreshToken){
        $global:octo.LCRefreshToken = $response.refresh_token
        Save-RefreshToken -RefreshToken $response.refresh_token
    }

    return $response
}

function get-AccessToken{
    Param(
        [Parameter(Mandatory=$true)]$resource,
        [Switch]$returnHeader
    )

    # Try to load refresh token from disk (once per process)
    if(!$global:octo.LCRefreshToken -and (Test-Path $global:octo.TokenCachePath)){
        try {
            $global:octo.LCRefreshToken = (Import-Clixml $global:octo.TokenCachePath).GetNetworkCredential().Password
            Write-Verbose "Loaded refresh token from local storage"
        } catch {
            Write-Warning "Failed to load cached token, proceeding to authentication..."
            Remove-Item $global:octo.TokenCachePath -ErrorAction SilentlyContinue
        }
    }

    # hot path: serve a still-valid cached access token WITHOUT any network/DPAPI/disk work.
    # Renew 5 minutes ahead of the real expiry reported by Entra.
    $cached = $global:octo.LCCachedTokens[$resource]
    if($cached -and $cached.accessToken -and $cached.expiresOn -gt (Get-Date).AddMinutes(5)){
        if($returnHeader){ return @{ "Authorization" = "Bearer $($cached.accessToken)" } }
        return $cached.accessToken
    }

    # No usable cached access token: make sure we have a refresh token, then exchange it.
    if(!$global:octo.LCRefreshToken){
        $global:octo.LCRefreshToken = Get-BrowserAuthorizationCode
    }

    $response = $null
    try {
        $response = Invoke-RefreshTokenExchange -Resource $resource
    } catch {
        # Refresh token invalid/expired/revoked -> drop it and fall back to an interactive sign-in once.
        Write-Warning "Cached refresh token invalid or expired, will re-authenticate..."
        $global:octo.LCRefreshToken = $Null
        Remove-Item $global:octo.TokenCachePath -ErrorAction SilentlyContinue
        $global:octo.LCRefreshToken = Get-BrowserAuthorizationCode
        $response = Invoke-RefreshTokenExchange -Resource $resource
    }

    if(!$response -or !$response.access_token){
        throw "Failed to retrieve access token!"
    }

    # Cache the access token against its real lifetime (expires_in seconds), defaulting to ~55 minutes.
    $expiresOn = (Get-Date).AddSeconds(3300)
    if($response.expires_in){
        try { $expiresOn = (Get-Date).AddSeconds([double]$response.expires_in) } catch {}
    }
    $global:octo.LCCachedTokens[$resource] = @{ accessToken = $response.access_token; expiresOn = $expiresOn }

    if($returnHeader){ return @{ "Authorization" = "Bearer $($response.access_token)" } }
    return $response.access_token
}

function Send-AuthListenerResponse {
    # Writes an HTML response to an HttpListener callback and closes it. Everything reflected from the
    # query string is HTML-encoded by the caller before it reaches here (B3).
    param(
        [Parameter(Mandatory = $true)]$Context,
        [Parameter(Mandatory = $true)][string]$Html,
        [int]$StatusCode = 200
    )
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Html)
        $Context.Response.StatusCode = $StatusCode
        $Context.Response.ContentType = "text/html; charset=utf-8"
        $Context.Response.ContentLength64 = $bytes.Length
        $Context.Response.OutputStream.Write($bytes, 0, $bytes.Length)
    } catch {
    } finally {
        try { $Context.Response.OutputStream.Close() } catch {}
        try { $Context.Response.Close() } catch {}
    }
}

function Get-AuthLandingPage {
    # Branded, self-closing landing page shown in the browser after the callback (D6).
    param(
        [Parameter(Mandatory = $true)][ValidateSet('success','failure')][string]$Kind,
        [string]$Detail = ""
    )
    $accent = if($Kind -eq 'success'){ "#107c10" } else { "#a80000" }
    $heading = if($Kind -eq 'success'){ "&#10004; You're signed in" } else { "&#10006; Sign-in failed" }
    $message = if($Kind -eq 'success'){ "M365AutoLink has what it needs. You can safely close this tab." } else { "M365AutoLink could not complete sign-in." }
    $detailHtml = if([string]::IsNullOrWhiteSpace($Detail)){ "" } else { "<p class='detail'>$Detail</p>" }
    return @"
<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>M365AutoLink</title>
<style>
 body{margin:0;font-family:Segoe UI,Roboto,Helvetica,Arial,sans-serif;background:#f3f2f1;color:#201f1e;display:flex;min-height:100vh;align-items:center;justify-content:center}
 .card{background:#fff;border-radius:10px;box-shadow:0 6px 24px rgba(0,0,0,.12);padding:36px 40px;max-width:420px;text-align:center}
 .accent{color:$accent;font-size:22px;font-weight:600;margin:0 0 10px}
 p{margin:6px 0;line-height:1.5}
 .brand{margin-top:18px;font-size:12px;color:#605e5c}
 .detail{font-size:13px;color:#605e5c;word-break:break-word}
</style></head><body>
<div class="card">
 <p class="accent">$heading</p>
 <p>$message</p>
 $detailHtml
 <p class="brand">M365AutoLink &middot; Lieben Consultancy</p>
</div>
<script>setTimeout(function(){window.close();},1500);</script>
</body></html>
"@
}

function Get-BrowserAuthorizationCode {
    # pick a free ephemeral loopback port (fixed ports collide on multi-session hosts and races).
    $portFinder = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, 0)
    $portFinder.Start()
    $port = ([System.Net.IPEndPoint]$portFinder.LocalEndpoint).Port
    $portFinder.Stop()

    $redirectUri = "http://localhost:$port/"

    # HttpListener handles HTTP correctly (vs. the raw TcpListener that only parsed the first line).
    # Binding to localhost does not require admin rights.
    $listener = [System.Net.HttpListener]::new()
    $listener.Prefixes.Add($redirectUri)
    try {
        $listener.Start()
    } catch {
        throw "Could not start the local sign-in listener on port $port : $($_.Exception.Message)"
    }

    # PKCE (S256) + anti-forgery state so any other local process cannot inject a code.
    $codeVerifier = New-PkceCodeVerifier
    $codeChallenge = New-PkceCodeChallenge -Verifier $codeVerifier
    $state = [Guid]::NewGuid().ToString("N")
    $scope = "$($global:octo.graphUrl)/.default offline_access"

    # v2 authorize endpoint (scope=, not resource=).
    $authUrl = "$($global:octo.authorizeUrl)?" +
        "client_id=$([System.Uri]::EscapeDataString($global:octo.LCClientId))" +
        "&response_type=code" +
        "&redirect_uri=$([System.Uri]::EscapeDataString($redirectUri))" +
        "&response_mode=query" +
        "&scope=$([System.Uri]::EscapeDataString($scope))" +
        "&state=$([System.Uri]::EscapeDataString($state))" +
        "&code_challenge=$([System.Uri]::EscapeDataString($codeChallenge))" +
        "&code_challenge_method=S256"

    Write-Host ""
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host " First-time authentication required" -ForegroundColor Cyan
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Opening browser for sign-in..." -ForegroundColor Yellow
    Write-Host "(After signing in once, future runs will be silent)" -ForegroundColor DarkGray
    Write-Host ""

    try {
        Start-Process $authUrl -WindowStyle $WindowStyle | Out-Null
    } catch {
        Write-Host "Could not open browser automatically." -ForegroundColor Yellow
        Write-Host "Please open this URL manually:" -ForegroundColor Yellow
        Write-Host $authUrl -ForegroundColor White
    }

    $code = $null
    $authError = $null
    $deadline = (Get-Date).AddMinutes(5)

    try {
        while((Get-Date) -lt $deadline){
            $remainingMs = [int][Math]::Max(1000, ($deadline - (Get-Date)).TotalMilliseconds)
            $contextTask = $listener.GetContextAsync()
            if(-not $contextTask.Wait($remainingMs)){ break }

            $context = $contextTask.Result
            $query = $context.Request.QueryString
            $returnedState = [string]$query["state"]

            # ignore anything whose state doesn't match ours (junk, races, injection attempts)
            # instead of terminating the wait.
            if($returnedState -ne $state){
                Send-AuthListenerResponse -Context $context -StatusCode 400 -Html (Get-AuthLandingPage -Kind 'failure' -Detail 'Unexpected request ignored.')
                continue
            }

            if($query["error"]){
                $errorCode = [string]$query["error"]
                $errorDesc = [string]$query["error_description"]
                $authError = "$errorCode - $errorDesc"
                $safeDetail = [System.Net.WebUtility]::HtmlEncode("$($errorCode): $errorDesc")
                Send-AuthListenerResponse -Context $context -Html (Get-AuthLandingPage -Kind 'failure' -Detail $safeDetail)
                break
            }

            if($query["code"]){
                $code = [string]$query["code"]
                Send-AuthListenerResponse -Context $context -Html (Get-AuthLandingPage -Kind 'success')
                break
            }

            # State matched but neither code nor error present: ignore and keep waiting.
            Send-AuthListenerResponse -Context $context -StatusCode 400 -Html (Get-AuthLandingPage -Kind 'failure' -Detail 'Incomplete request ignored.')
        }
    } finally {
        try { $listener.Stop() } catch {}
        try { $listener.Close() } catch {}
    }

    if($authError){
        throw "Authentication error: $authError"
    }
    if([string]::IsNullOrWhiteSpace($code)){
        throw "Authentication timed out - no valid response received within 5 minutes"
    }

    Write-Host "Authorization code received, exchanging for tokens..." -ForegroundColor Cyan

    # v2 token endpoint. include the PKCE code_verifier in the exchange.
    $tokenBody = @{
        grant_type    = "authorization_code"
        client_id     = $global:octo.LCClientId
        code          = $code
        redirect_uri  = $redirectUri
        scope         = $scope
        code_verifier = $codeVerifier
    }

    $response = Invoke-RestMethod -Uri $global:octo.tokenUrl -Method POST -Body $tokenBody -ErrorAction Stop

    if ($response.refresh_token) {
        Save-RefreshToken -RefreshToken $response.refresh_token
        Write-Host ""
        Write-Host "Authentication successful! Token cached for future use." -ForegroundColor Green
        Write-Host ""
        return $response.refresh_token
    }

    throw "No refresh token received from Entra ID"
}

function New-GraphQuery {
  
    Param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Uri,

        [Parameter(Mandatory = $true)]
        [ValidateSet('GET', 'POST', 'PATCH', 'DELETE', 'PUT')]
        [string]$Method,    
        
        [Parameter(Mandatory = $false)]
        [string]$Body,        

        [Parameter(Mandatory = $false)]
        [int]$MaxAttempts = 5, 

        [Parameter(Mandatory = $false)]
        [String]$resource = "https://graph.microsoft.com",

        [Parameter(Mandatory = $false)]
        [String]$ContentType = 'application/json; charset=utf-8'
    )

    function get-resourceHeaders{
        Param(
            [string]$resource
        )

        try{
            $token = Get-AccessToken -resource $resource
        }catch{
            # do NOT Exit here - that would silently kill the tray process mid-run with no balloon.
            # Throw a recognizable error instead so Invoke-M365AutoLinkRun's catch can surface it as an
            # Error balloon (with the consent URL) and keep the tray alive for a retry after consent.
            $consentUrl = "$($global:octo.idpUrl)/organizations/adminconsent?client_id=$($global:octo.LCClientId)"
            Write-Log "Token acquisition failed for '$resource': $($_.Exception.Message)" -Level "ERROR"
            Write-Log "Possible fix: an admin still needs to approve this application at $consentUrl" -Level "ERROR"
            throw "M365AutoLink needs admin consent (or sign-in failed). Ask IT to approve: $consentUrl"
        }
        $headers = @{
            "Authorization" = "Bearer $token"
        }
        $headers['Accept-Language'] = "en-US"

        if($resource -like "*sharepoint.com*"){
            $headers['Accept'] = "application/json;odata=nometadata"
        }    

        if($resource -like "*outlook.office365.com*"){
            $headers['Accept'] = "application/json;odata.metadata=minimal"
        }

        return $headers
    }

    $nextURL = $uri

    if($Method -in ('POST', 'PATCH', 'PUT')){
        try {
            $attempts = 0
            while ($attempts -lt $MaxAttempts) {
                $attempts++
                try {
                    $headers = get-resourceHeaders -resource $resource
                    $Data = $Null; $Data = (Invoke-RestMethod -Uri $nextURL -Method $Method -Headers $headers -Body $Body -ContentType $ContentType -Verbose:$False -ErrorAction Stop -UserAgent "ISV|LiebenConsultancy|M365AutoLink|1.0")
                    $attempts = $MaxAttempts
                }catch {
                    $statusCode = $null
                    try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {}
                    $is429 = $statusCode -eq 429 -or $_.Exception.Message -like "*429*"
                    $isTransientNetwork = $_.Exception.Message -like "*No such host is known*" -or $_.Exception.Message -like "*name or service not known*" -or $_.Exception.Message -like "*network is unreachable*" -or $_.Exception.Message -like "*connection was forcibly closed*" -or $_.Exception.Message -like "*An existing connection was forcibly closed*"

                    # Fail fast on all HTTP errors except 429 (including 500/403/404).
                    if($null -ne $statusCode -and -not $is429){
                        $nextUrl = $Null
                        throw $_
                    }

                    # Retry only throttling or transport-level transient failures.
                    if(-not $is429 -and -not $isTransientNetwork){
                        $nextUrl = $Null
                        throw $_
                    }

                    if ($attempts -ge $MaxAttempts) { 
                        Throw $_
                    }

                    $delay = 0
                    if ($is429){
                        try {
                            $retryAfter = $_.Exception.Response.Headers.GetValues("Retry-After")
                            if ($retryAfter -and $retryAfter.Count -gt 0) {
                                $retryAfterValue = $retryAfter[0]
                                if ($retryAfterValue -match '^\d+$') {
                                    $delay = [int]$retryAfterValue
                                }
                            }
                        }catch {}
                        if($delay -le 0){
                            $delay = [math]::Min(15, 2 * [math]::Max(1, $attempts))
                        }
                    }
                    if($delay -le 0 -and $isTransientNetwork){
                        $delay = [math]::Min(5, $attempts)
                    }
                    Write-Log "Transient error on attempt $attempts/$MaxAttempts, retrying in $($delay)s: $($_.Exception.Message)" -Level "WARN"
                    Start-Sleep -Seconds (1 + $delay)
                }     
            }
        }catch {
            $Message = ($_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue).error.message
            if ($null -eq $Message) { $Message = $($_.Exception.Message) }
            throw $Message
        }                               
        return $Data
    }else{
        $ReturnedData = @()
        $totalResults = 0     
           
        while($Null -ne $nextUrl -and $nextUrl.indexOf("http") -eq 0){
            try {
                $attempts = 0
                while ($attempts -lt $MaxAttempts) {
                    $attempts ++
                    try {
                        $headers = get-resourceHeaders -resource $resource
                        $Data = $Null; $Data = (Invoke-RestMethod -Uri $nextURL -Method $Method -Headers $headers -ContentType $ContentType -Verbose:$False -ErrorAction Stop -UserAgent "ISV|LiebenConsultancy|M365AutoLink|1.0")
                        $attempts = $MaxAttempts
                    }catch {                 
                        $statusCode = $null
                        try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {}
                        $is429 = $statusCode -eq 429 -or $_.Exception.Message -like "*429*"
                        $isTransientNetwork = $_.Exception.Message -like "*No such host is known*" -or $_.Exception.Message -like "*name or service not known*" -or $_.Exception.Message -like "*network is unreachable*" -or $_.Exception.Message -like "*connection was forcibly closed*" -or $_.Exception.Message -like "*An existing connection was forcibly closed*"

                        # Fail fast on all HTTP errors except 429 (including 500/403/404).
                        if($null -ne $statusCode -and -not $is429){
                            $nextUrl = $Null
                            throw $_
                        }

                        # Retry only throttling or transport-level transient failures.
                        if(-not $is429 -and -not $isTransientNetwork){
                            $nextUrl = $Null
                            throw $_
                        }
              
                        if ($attempts -ge $MaxAttempts) { 
                            $nextURL = $null
                            Throw $_
                        }
                       
                        $delay = 0
                        if ($is429){
                            try {
                                $retryAfter = $_.Exception.Response.Headers.GetValues("Retry-After")
                                if ($retryAfter -and $retryAfter.Count -gt 0) {
                                    $retryAfterValue = $retryAfter[0]
                                    if ($retryAfterValue -match '^\d+$') {
                                        $delay = [int]$retryAfterValue
                                    }
                                }
                            }catch {}
                            if($delay -le 0){
                                $delay = [math]::Min(15, 2 * [math]::Max(1, $attempts))
                            }
                        }
                        if($delay -le 0 -and $isTransientNetwork){
                            $delay = [math]::Min(5, $attempts)
                        }
                        Write-Log "Transient error on attempt $attempts/$MaxAttempts, retrying in $($delay)s: $($_.Exception.Message)" -Level "WARN"
                        Start-Sleep -Seconds (1 + $delay)
                    }
                }

                if($resource -like "*sharepoint.com*"){
                    if($Data -and $Data.PSObject.TypeNames -notcontains "System.Management.Automation.PSCustomObject"){
                        # on PowerShell 7 System.Web.Extensions (JavaScriptSerializer) does not exist,
                        # so use ConvertFrom-Json -AsHashtable (handles large payloads and returns a
                        # hashtable natively). Fall back to the serializer only on Windows PowerShell 5.1.
                        if($PSVersionTable.PSVersion.Major -ge 6){
                            $Data = ($Data | Out-String | ConvertFrom-Json -AsHashtable)
                        } else {
                            $Null = Add-Type -AssemblyName System.Web.Extensions
                            $serializer = New-Object System.Web.Script.Serialization.JavaScriptSerializer
                            $serializer.MaxJsonLength = 2147483647
                            $jsonContent = $serializer.DeserializeObject($Data)
                            if ($jsonContent -is [System.Collections.IDictionary]) {
                                $Data = New-Object Hashtable $jsonContent
                            } else {
                                $Data = $jsonContent
                            }
                        }
                    }
                }

                $pageItems = $null
                if($Data.psobject.properties.name -icontains 'value' -or ($Data.PSObject.BaseObject -is [hashtable] -and $Data.Keys -icontains 'value')){ # Added check for hashtable
                    $pageItems = $Data.value
                }else{
                    # This case handles responses where the data is the root object (e.g., an array of items directly)
                    $pageItems = $Data
                }

                if ($null -ne $pageItems) {
                    # Ensure $pageItems is treated as a collection for .Count, even if it's a single object
                    $pageItemCount = @($pageItems).Count
                    $totalResults += $pageItemCount

                    if ($pageItemCount -eq 1 -and -not ($pageItems -is [array])) {
                        $ReturnedData += @($pageItems) # Add single item as an array element
                    } elseif ($pageItemCount -gt 0) {
                            $ReturnedData += $pageItems # Add array of items
                    }
                }     
                
                if($Data.'@odata.nextLink'){
                    $nextURL = $Data.'@odata.nextLink'  
                }elseif($Data.'odata.nextLink'){
                    $nextURL = $Data.'odata.nextLink'                      
                }elseif($Data.nextLink){
                    $nextURL = $Data.nextLink
                }else{
                    $nextURL = $null
                }            
            }catch {
                throw $_
            }
        }

        if ($ReturnedData -and !$ReturnedData.value -and $ReturnedData.PSObject.Properties["value"]) { return $null }
        return $ReturnedData
    }
}

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$timestamp] [$Level] $Message"
    $color = switch($Level) {
        "ERROR" { "Red" }
        "WARN"  { "Yellow" }
        "SUCCESS" { "Green" }
        default { "White" }
    }
    Write-Host $line -ForegroundColor $color

    # write to the log file directly so we never depend on Start-Transcript (which fights between
    # two instances and disappears if it fails to start). Best-effort; never let logging break a run.
    try {
        $logPath = $global:octo.LogPath
        if(-not [string]::IsNullOrWhiteSpace($logPath)) {
            $logDir = [System.IO.Path]::GetDirectoryName($logPath)
            if(-not [string]::IsNullOrWhiteSpace($logDir) -and -not (Test-Path -LiteralPath $logDir)) {
                New-Item -ItemType Directory -Path $logDir -Force | Out-Null
            }
            Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8 -ErrorAction Stop
        }
    } catch {}
}

function Invoke-LogRotation {
    # roll the previous lastRun.log to run-<timestamp>.log and prune to $LogHistoryCount files.
    try {
        $logPath = $global:octo.LogPath
        if([string]::IsNullOrWhiteSpace($logPath)) { return }
        $logDir = [System.IO.Path]::GetDirectoryName($logPath)
        if([string]::IsNullOrWhiteSpace($logDir)) { return }
        if(-not (Test-Path -LiteralPath $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        }

        if(Test-Path -LiteralPath $logPath) {
            $stamp = (Get-Item -LiteralPath $logPath).LastWriteTime.ToString("yyyyMMdd-HHmmss")
            $archivePath = Join-Path -Path $logDir -ChildPath "run-$stamp.log"
            try { Move-Item -LiteralPath $logPath -Destination $archivePath -Force -ErrorAction Stop } catch {}
        }

        $keep = [int]$LogHistoryCount
        if($keep -lt 0) { $keep = 0 }
        Get-ChildItem -Path $logDir -Filter "run-*.log" -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending |
            Select-Object -Skip $keep |
            ForEach-Object { try { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop } catch {} }
    } catch {}
}

function Initialize-ProgressBar {
    if(-not $ShowProgressBar -or -not $script:traySync) { return }
    $script:traySync.ProgressVisible = $true
}

function Stop-ProgressBar {
    if(-not $script:traySync) { return }
    $script:traySync.ProgressVisible = $false
}

function Update-TrayState {
    param(
        [string]$Text,
        [int]$Percent,
        [string]$ProgressText,
        [switch]$ShowBalloon,
        [string]$BalloonTitle = "M365AutoLink",
        [string]$BalloonMessage = "",
        [ValidateSet("Info", "Warning", "Error")]
        [string]$BalloonIcon = "Info",
        [string]$BalloonClickUrl = "",
        [long]$TotalItemCount,
        [ValidateSet("ok", "approaching", "over")]
        [string]$ItemCountStatus,
        [switch]$IsRunning
    )

    if(-not $script:traySync) { return }

    if($PSBoundParameters.ContainsKey("TotalItemCount")) {
        $script:traySync.TotalItemCount = $TotalItemCount
        $resolvedStatus = if($PSBoundParameters.ContainsKey("ItemCountStatus")) { $ItemCountStatus } else { Get-TotalItemCountStatus -TotalItemCount $TotalItemCount }
        $script:traySync.ItemCountStatus = $resolvedStatus
        $script:traySync.ItemCountText = "M365AutoLink - " + (Get-ItemCountSummaryText -TotalItemCount $TotalItemCount -Status $resolvedStatus)
    } elseif($PSBoundParameters.ContainsKey("ItemCountStatus")) {
        $script:traySync.ItemCountStatus = $ItemCountStatus
    }

    if($PSBoundParameters.ContainsKey("Text") -and $Text) {
        $trimmedText = $Text
        if($trimmedText.Length -gt 63) {
            $trimmedText = $trimmedText.Substring(0, 63)
        }
        $script:traySync.Text = $trimmedText
    }

    if($PSBoundParameters.ContainsKey("Percent")) {
        $script:traySync.ProgressPercent = [Math]::Max(0, [Math]::Min(100, $Percent))
    }

    if($PSBoundParameters.ContainsKey("ProgressText")) {
        $script:traySync.ProgressText = $ProgressText
    }

    if($PSBoundParameters.ContainsKey("IsRunning")) {
        $script:traySync.IsRunning = [bool]$IsRunning
    }

    if($ShowBalloon) {
        $script:traySync.BalloonTitle = $BalloonTitle
        $script:traySync.BalloonMsg = $BalloonMessage
        $script:traySync.BalloonIcon = $BalloonIcon
        # Set the click target for this balloon (or clear it, so a previous link never carries over).
        $script:traySync.BalloonClickUrl = $BalloonClickUrl
        $script:traySync.ShowBalloon = $true
    }
}

function Initialize-TrayIcon {
    # This runspace owns ALL long-lived WinForms UI (tray icon AND progress form) and pumps
    # messages continuously via Application.Run().
    if($script:traySync) { return }
    if(-not $EnableSystemTrayIcon -and -not $ShowProgressBar) { return }

    try {
        $script:traySync = [hashtable]::Synchronized(@{
            Text            = "M365AutoLink - Starting"
            ProgressPercent = 0
            ProgressText    = "Idle"
            ProgressVisible = $false
            ProgressBarText = $ProgressBarText
            ProgressBarColor = $ProgressBarColor
            EnableTrayIcon  = [bool]$EnableSystemTrayIcon
            TrayReady       = $false
            IsRunning       = $false
            HasCompletedRun = $false
            HasMappedSites  = $false
            HasExistingConflicts = $false
            RequestRerun    = $false
            RequestManageShortcuts = $false
            ExitRequested   = $false
            RefreshIconRequested = $false
            ResumeDetected  = $false
            BalloonTitle    = ""
            BalloonMsg      = ""
            BalloonIcon     = "Info"
            BalloonClickUrl = ""
            ShowBalloon     = $false
            TotalItemCount  = 0
            ItemCountThreshold = $totalItemCountWarningThreshold
            ItemCountStatus = "ok"
            ItemCountText   = ""
            ItemCountHelpLink = $ItemCountHelpLink
            LogFile         = $global:octo.LogPath
            OneDriveRootPath = $script:localOneDriveRootPath
            LocalShortcutFolderPath = $script:localShortcutFolderPath
            ShortcutFolderName = $FolderName
            HelpLink        = $TrayHelpLink
            CopyrightText   = $TrayCopyrightText
            CopyrightLink   = $TrayCopyrightLink
            Version         = [string]$ScriptVersion
        })

        $script:trayRunspace = [runspacefactory]::CreateRunspace()
        $script:trayRunspace.ApartmentState = "STA"
        $script:trayRunspace.ThreadOptions = "ReuseThread"
        $script:trayRunspace.Open()
        $script:trayRunspace.SessionStateProxy.SetVariable("sync", $script:traySync)

        $script:trayPS = [powershell]::Create().AddScript({
            Add-Type -AssemblyName System.Drawing
            Add-Type -AssemblyName System.Windows.Forms

            try {
                $powerModeHandler = {
                    param($powerSender, $powerArgs)
                    if($powerArgs.Mode -eq [Microsoft.Win32.PowerModes]::Resume) {
                        try { $sync.RefreshIconRequested = $true } catch {}
                        # let the main loop trigger a refresh shortly after resume (if auto-refresh is on).
                        try { $sync.ResumeDetected = $true } catch {}
                    }
                }
                [Microsoft.Win32.SystemEvents]::add_PowerModeChanged($powerModeHandler)
            } catch {}

            $icon = New-Object Windows.Forms.NotifyIcon

            # Draws the cloud+arrow glyph tinted by item-count status so the icon itself signals
            # whether the combined linked-library item count is ok (blue), approaching (amber) or
            # over (red) the limit. Returns a fresh Icon handle the caller is responsible for.
            function New-TrayIconHandle {
                param([string]$Status)

                $accent = switch($Status) {
                    "over"        { [Drawing.Color]::FromArgb(220, 53, 69) }   # red
                    "approaching" { [Drawing.Color]::FromArgb(245, 158, 11) }  # amber
                    default       { [Drawing.Color]::FromArgb(0, 120, 215) }   # blue
                }

                $bmp = New-Object Drawing.Bitmap(16, 16)
                $g = [Drawing.Graphics]::FromImage($bmp)
                $g.SmoothingMode = "AntiAlias"
                $g.Clear([Drawing.Color]::Transparent)
                $fill = New-Object Drawing.SolidBrush($accent)
                $g.FillEllipse($fill, 2, 5, 12, 9)
                $g.FillEllipse($fill, 4, 2, 8, 8)
                $g.FillEllipse($fill, 1, 6, 6, 7)
                $g.FillEllipse($fill, 9, 6, 6, 7)
                $pen = New-Object Drawing.Pen([Drawing.Color]::White, 1.6)
                $pen.StartCap = $pen.EndCap = [Drawing.Drawing2D.LineCap]::Round
                $g.DrawLine($pen, 8, 12, 8, 7)
                $g.DrawLine($pen, 5.5, 9.5, 8, 7)
                $g.DrawLine($pen, 10.5, 9.5, 8, 7)
                $pen.Dispose(); $fill.Dispose(); $g.Dispose()
                $iconHandle = [Drawing.Icon]::FromHandle($bmp.GetHicon())
                $bmp.Dispose()
                return $iconHandle
            }

            $script:lastIconStatus = [string]$sync.ItemCountStatus
            if([string]::IsNullOrWhiteSpace($script:lastIconStatus)) { $script:lastIconStatus = "ok" }
            $icon.Icon = New-TrayIconHandle -Status $script:lastIconStatus

            # When only the progress bar is enabled, this runspace still runs but the icon stays hidden.
            $icon.Visible = [bool]$sync.EnableTrayIcon
            $icon.Text = if([string]::IsNullOrWhiteSpace([string]$sync.Version)) { "M365AutoLink" } else { "M365AutoLink v$($sync.Version)" }

            # Clicking the over/approaching-limit balloon opens the configured knowledgebase article.
            $icon.Add_BalloonTipClicked({
                try {
                    $balloonUrl = [string]$sync.BalloonClickUrl
                    if(-not [string]::IsNullOrWhiteSpace($balloonUrl)) {
                        Start-Process $balloonUrl
                    }
                } catch {}
            })

            $icon.Add_MouseClick({
                param($sender, $e)

                if($e.Button -ne [Windows.Forms.MouseButtons]::Left) { return }

                $targetPath = [string]$sync.LocalShortcutFolderPath
                if(-not [string]::IsNullOrWhiteSpace($targetPath) -and (Test-Path $targetPath)) {
                    Start-Process explorer.exe $targetPath
                    return
                }

                $oneDriveRoot = [string]$sync.OneDriveRootPath
                if(-not [string]::IsNullOrWhiteSpace($oneDriveRoot) -and (Test-Path $oneDriveRoot)) {
                    Start-Process explorer.exe $oneDriveRoot
                    $folderName = [string]$sync.ShortcutFolderName
                    if([string]::IsNullOrWhiteSpace($folderName)) { $folderName = "AutoLink" }
                    $icon.ShowBalloonTip(2000, "M365AutoLink", "Local folder '$folderName' is not available yet. Opened your OneDrive root instead.", [Windows.Forms.ToolTipIcon]::Info)
                    return
                }

                $icon.ShowBalloonTip(2500, "M365AutoLink", "Could not locate a local OneDrive folder on this device.", [Windows.Forms.ToolTipIcon]::Warning)
            })

            $menu = New-Object Windows.Forms.ContextMenuStrip
            $menu.AutoClose = $true

            $menu.Add_Opening({
                try {
                    if($sync.IsRunning) {
                        $remapItem.Enabled = $false
                        $manageShortcutsItem.Enabled = $false
                    } else {
                        $remapItem.Enabled = $true
                        $manageShortcutsItem.Enabled = [bool]$sync.HasMappedSites
                    }

                    # Refresh the read-only combined item-count line (hidden until a run produced data).
                    $itemCountText = [string]$sync.ItemCountText
                    $hasItemCount = -not [string]::IsNullOrWhiteSpace($itemCountText)
                    $itemCountInfoItem.Visible = $hasItemCount
                    $itemCountInfoSeparator.Visible = $hasItemCount
                    if($hasItemCount) {
                        $itemCountInfoItem.Text = $itemCountText -replace '^M365AutoLink - ', ''
                        $itemCountInfoItem.ForeColor = switch([string]$sync.ItemCountStatus) {
                            "over"        { [Drawing.Color]::FromArgb(196, 43, 28) }
                            "approaching" { [Drawing.Color]::FromArgb(176, 110, 0) }
                            default       { [Drawing.Color]::FromArgb(72, 82, 94) }
                        }
                    }
                } catch {}
            })

            $itemCountInfoItem = New-Object Windows.Forms.ToolStripMenuItem("")
            $itemCountInfoItem.Enabled = $false
            $itemCountInfoItem.Visible = $false
            $itemCountInfoSeparator = New-Object Windows.Forms.ToolStripSeparator
            $itemCountInfoSeparator.Visible = $false

            $remapItem = New-Object Windows.Forms.ToolStripMenuItem("Run now")
            $remapItem.Add_Click({
                try {
                    if(-not $sync.IsRunning) {
                        $sync.RequestRerun = $true
                    }
                } catch {}
            })

            $manageShortcutsItem = New-Object Windows.Forms.ToolStripMenuItem("Manage shortcuts")
            $manageShortcutsItem.Add_Click({
                try {
                    if(-not $sync.IsRunning) {
                        $sync.RequestManageShortcuts = $true
                        $icon.ShowBalloonTip(1500, "M365AutoLink", "Opening Manage shortcuts...", [Windows.Forms.ToolTipIcon]::Info)
                    }
                } catch {
                    try { $icon.ShowBalloonTip(2000, "M365AutoLink", "Tray action failed. Please try again.", [Windows.Forms.ToolTipIcon]::Warning) } catch {}
                }
            })
            $manageShortcutsItem.Enabled = $false

            $helpItem = New-Object Windows.Forms.ToolStripMenuItem("Open help")
            $helpItem.Add_Click({
                if($sync.HelpLink) { Start-Process $sync.HelpLink }
            })

            $copyrightItem = New-Object Windows.Forms.ToolStripMenuItem($sync.CopyrightText)
            $copyrightItem.Add_Click({
                if($sync.CopyrightLink) { Start-Process $sync.CopyrightLink }
            })

            $openLogItem = New-Object Windows.Forms.ToolStripMenuItem("Open log file")
            $openLogItem.Add_Click({
                $lf = $sync.LogFile
                if ($lf -and (Test-Path $lf)) { Start-Process notepad.exe $lf }
            })

            $exitItem = New-Object Windows.Forms.ToolStripMenuItem("Exit M365AutoLink")
            $exitItem.Add_Click({
                try { $sync.ExitRequested = $true } catch {}
            })

            # show the running version (single source of truth: $ScriptVersion).
            $versionItem = New-Object Windows.Forms.ToolStripMenuItem("M365AutoLink v$([string]$sync.Version)")
            $versionItem.Enabled = $false

            [void]$menu.Items.Add($versionItem)
            [void]$menu.Items.Add((New-Object Windows.Forms.ToolStripSeparator))
            [void]$menu.Items.Add($itemCountInfoItem)
            [void]$menu.Items.Add($itemCountInfoSeparator)
            [void]$menu.Items.Add($remapItem)
            [void]$menu.Items.Add($manageShortcutsItem)
            [void]$menu.Items.Add((New-Object Windows.Forms.ToolStripSeparator))
            [void]$menu.Items.Add($helpItem)
            [void]$menu.Items.Add($copyrightItem)
            [void]$menu.Items.Add($openLogItem)
            [void]$menu.Items.Add((New-Object Windows.Forms.ToolStripSeparator))
            [void]$menu.Items.Add($exitItem)
            $icon.ContextMenuStrip = $menu

            $script:progressForm = $null
            $script:progressFill = $null
            $script:progressLabel = $null
            $script:progressTrackWidth = 0
            $script:progressFormBroken = $false
            $script:lastProgressPercent = -1
            $script:lastProgressText = $null
            $script:lastIconText = ""

            function New-M365ProgressForm {
                # scale the hand-laid-out floating bar by the primary monitor's DPI so it stays the
                # right physical size and crisp (the process is per-monitor DPI aware).
                $scale = 1.0
                try {
                    $screenGraphics = [Drawing.Graphics]::FromHwnd([IntPtr]::Zero)
                    if($screenGraphics.DpiX -gt 0) { $scale = $screenGraphics.DpiX / 96.0 }
                    $screenGraphics.Dispose()
                } catch {}
                function ds { param($v) return [int][math]::Round($v * $scale) }

                $w = ds 360
                $h = ds 50
                $pad = ds 8
                $iconBox = ds 24
                $trackH = [Math]::Max(2, (ds 5))
                $gap = ds 8
                $script:progressTrackWidth = $w - ($pad * 2) - $iconBox - $gap

                $form = New-Object Windows.Forms.Form
                $form.Text = "M365AutoLink"
                $form.AutoScaleMode = "None"
                $form.Size = New-Object Drawing.Size($w, $h)
                $form.MaximumSize = $form.Size
                $form.MinimumSize = $form.Size
                $form.BackColor = [Drawing.Color]::FromArgb(33, 37, 43)
                $form.ControlBox = $false
                $form.FormBorderStyle = "None"
                $form.ShowInTaskbar = $false
                $form.StartPosition = "Manual"
                $form.TopMost = $true
                $form.Opacity = 0.90

                $radius = ds 8
                $gp = New-Object Drawing.Drawing2D.GraphicsPath
                $gp.AddArc(0, 0, $radius * 2, $radius * 2, 180, 90)
                $gp.AddArc($w - $radius * 2 - 1, 0, $radius * 2, $radius * 2, 270, 90)
                $gp.AddArc($w - $radius * 2 - 1, $h - $radius * 2 - 1, $radius * 2, $radius * 2, 0, 90)
                $gp.AddArc(0, $h - $radius * 2 - 1, $radius * 2, $radius * 2, 90, 90)
                $gp.CloseFigure()
                $form.Region = New-Object Drawing.Region($gp)
                $gp.Dispose()

                $iconPanel = New-Object Windows.Forms.Panel
                $iconPanel.Location = New-Object Drawing.Point($pad, (ds 7))
                $iconPanel.Size = New-Object Drawing.Size($iconBox, (ds 24))
                $iconPanel.BackColor = [Drawing.Color]::Transparent

                $iconBitmap = New-Object Drawing.Bitmap($iconBox, (ds 24))
                $ig = [Drawing.Graphics]::FromImage($iconBitmap)
                $ig.SmoothingMode = "AntiAlias"
                $ig.Clear([Drawing.Color]::Transparent)
                $iconFill = New-Object Drawing.SolidBrush([Drawing.Color]::FromArgb(0, 163, 255))
                $ig.FillEllipse($iconFill, (ds 3), (ds 9), (ds 19), (ds 11))
                $ig.FillEllipse($iconFill, (ds 7), (ds 4), (ds 13), (ds 10))
                $ig.FillEllipse($iconFill, (ds 1), (ds 10), (ds 10), (ds 9))
                $ig.FillEllipse($iconFill, (ds 14), (ds 10), (ds 11), (ds 9))
                $iconPen = New-Object Drawing.Pen([Drawing.Color]::White, [float](1.7 * $scale))
                $iconPen.StartCap = $iconPen.EndCap = [Drawing.Drawing2D.LineCap]::Round
                $ig.DrawLine($iconPen, (ds 13), (ds 17), (ds 13), (ds 11))
                $ig.DrawLine($iconPen, (ds 10), (ds 13), (ds 13), (ds 11))
                $ig.DrawLine($iconPen, (ds 16), (ds 13), (ds 13), (ds 11))
                $iconPen.Dispose(); $iconFill.Dispose(); $ig.Dispose()

                $iconPicture = New-Object Windows.Forms.PictureBox
                $iconPicture.Location = New-Object Drawing.Point(0, 0)
                $iconPicture.Size = New-Object Drawing.Size($iconBox, (ds 24))
                $iconPicture.BackColor = [Drawing.Color]::Transparent
                $iconPicture.Image = $iconBitmap
                $iconPicture.SizeMode = "CenterImage"
                $iconPanel.Controls.Add($iconPicture)

                $label = New-Object Windows.Forms.Label
                $label.Text = [string]$sync.ProgressBarText
                $label.Location = New-Object Drawing.Point(($pad + $iconBox + $gap), (ds 7))
                $label.Size = New-Object Drawing.Size(($w - ($pad * 2) - $iconBox - $gap), (ds 15))
                $label.Font = New-Object Drawing.Font("Segoe UI", 9)
                $label.ForeColor = [Drawing.Color]::FromArgb(237, 244, 252)
                $label.BackColor = [Drawing.Color]::Transparent
                $label.AutoEllipsis = $true

                $track = New-Object Windows.Forms.Panel
                $track.Location = New-Object Drawing.Point(($pad + $iconBox + $gap), (ds 32))
                $track.Size = New-Object Drawing.Size($script:progressTrackWidth, $trackH)
                $track.BackColor = [Drawing.Color]::FromArgb(72, 82, 94)

                $fill = New-Object Windows.Forms.Panel
                $fill.Location = New-Object Drawing.Point(0, 0)
                $fill.Size = New-Object Drawing.Size(0, $trackH)
                try {
                    $fill.BackColor = [Drawing.ColorTranslator]::FromHtml([string]$sync.ProgressBarColor)
                } catch {
                    $fill.BackColor = [Drawing.Color]::FromArgb(0, 163, 255)
                }

                $track.Controls.Add($fill)
                $form.Controls.AddRange(@($iconPanel, $label, $track))

                [void]$form.Show()
                $screen = [Windows.Forms.Screen]::PrimaryScreen.WorkingArea
                $form.SetDesktopLocation(($screen.Right - $w - (ds 12)), ($screen.Bottom - $h - (ds 12)))

                $script:progressForm = $form
                $script:progressFill = $fill
                $script:progressLabel = $label
                $script:lastProgressPercent = -1
                $script:lastProgressText = $null
            }

            $timer = New-Object Windows.Forms.Timer
            $timer.Interval = 200
            $timer.Add_Tick({
                try {
                    if($sync.ExitRequested) {
                        $timer.Stop()
                        if($script:progressForm) {
                            try { $script:progressForm.Close(); $script:progressForm.Dispose() } catch {}
                            $script:progressForm = $null
                        }
                        try {
                            $icon.Visible = $false
                            $icon.Dispose()
                        } catch {}
                        [Windows.Forms.Application]::ExitThread()
                        return
                    }

                    if($sync.ProgressVisible) {
                        if((-not $script:progressForm -or $script:progressForm.IsDisposed) -and -not $script:progressFormBroken) {
                            try {
                                New-M365ProgressForm
                            } catch {
                                # Creation failed; don't retry on every tick.
                                $script:progressFormBroken = $true
                                $script:progressForm = $null
                            }
                        }
                        if($script:progressForm -and -not $script:progressForm.IsDisposed) {
                            $percent = [Math]::Max(0, [Math]::Min(100, [int]$sync.ProgressPercent))
                            if($percent -ne $script:lastProgressPercent) {
                                $script:lastProgressPercent = $percent
                                $targetWidth = [int]($script:progressTrackWidth * $percent / 100)
                                if($percent -gt 0 -and $targetWidth -lt 2) { $targetWidth = 2 }
                                $script:progressFill.Width = $targetWidth
                            }
                            $progressText = [string]$sync.ProgressText
                            if($progressText -ne $script:lastProgressText) {
                                $script:lastProgressText = $progressText
                                if([string]::IsNullOrWhiteSpace($progressText)) {
                                    $script:progressLabel.Text = [string]$sync.ProgressBarText
                                } else {
                                    $script:progressLabel.Text = "M365AutoLink: $progressText"
                                }
                            }
                        }
                    } elseif($script:progressForm) {
                        try { $script:progressForm.Close(); $script:progressForm.Dispose() } catch {}
                        $script:progressForm = $null
                        $script:progressFill = $null
                        $script:progressLabel = $null
                    }

                    # Recolor the tray icon when the combined item-count status changes.
                    $currentItemStatus = [string]$sync.ItemCountStatus
                    if([string]::IsNullOrWhiteSpace($currentItemStatus)) { $currentItemStatus = "ok" }
                    if($currentItemStatus -ne $script:lastIconStatus) {
                        $script:lastIconStatus = $currentItemStatus
                        try {
                            $oldIconHandle = $icon.Icon
                            $icon.Icon = New-TrayIconHandle -Status $currentItemStatus
                            if($oldIconHandle) { try { $oldIconHandle.Dispose() } catch {} }
                        } catch {}
                    }

                    if(-not $menu.Visible) {
                        # When near/over the limit, surface the item-count warning in the tooltip itself.
                        $iconText = [string]$sync.Text
                        if(($currentItemStatus -eq "over" -or $currentItemStatus -eq "approaching") -and -not [string]::IsNullOrWhiteSpace([string]$sync.ItemCountText)) {
                            $iconText = [string]$sync.ItemCountText
                        }
                        if($iconText.Length -gt 63) { $iconText = $iconText.Substring(0, 63) }
                        if($iconText -and $iconText -ne $script:lastIconText) {
                            $script:lastIconText = $iconText
                            try { $icon.Text = $iconText } catch {}
                        }

                        if($sync.RefreshIconRequested) {
                            $sync.RefreshIconRequested = $false
                            if($sync.EnableTrayIcon) {
                                try {
                                    $icon.Visible = $false
                                    $icon.Visible = $true
                                } catch {}
                            }
                        }

                        if($sync.ShowBalloon) {
                            $sync.ShowBalloon = $false
                            if($sync.EnableTrayIcon -and -not [string]::IsNullOrWhiteSpace([string]$sync.BalloonMsg)) {
                                $tipIcon = switch ([string]$sync.BalloonIcon) {
                                    "Warning" { [Windows.Forms.ToolTipIcon]::Warning }
                                    "Error"   { [Windows.Forms.ToolTipIcon]::Error }
                                    default    { [Windows.Forms.ToolTipIcon]::Info }
                                }
                                $icon.ShowBalloonTip(3000, [string]$sync.BalloonTitle, [string]$sync.BalloonMsg, $tipIcon)
                            }
                        }
                    }
                } catch {
                    # Never let a timer tick exception kill tray responsiveness.
                }
            })
            $timer.Start()
            $sync.TrayReady = $true

            [Windows.Forms.Application]::Run()
        })

        $script:trayPS.Runspace = $script:trayRunspace
        $null = $script:trayPS.BeginInvoke()
        $readyDeadline = [DateTime]::UtcNow.AddSeconds(5)
        while(-not $script:traySync.TrayReady -and [DateTime]::UtcNow -lt $readyDeadline) {
            Start-Sleep -Milliseconds 25
        }
    } catch {
        $script:traySync = $null
    }
}

function Stop-TrayIcon {
    if($script:traySync) {
        $script:traySync.ExitRequested = $true
    }

    if($script:trayPS) {
        $stopDeadline = [DateTime]::UtcNow.AddSeconds(2)
        while($script:trayPS.InvocationStateInfo.State -eq 'Running' -and [DateTime]::UtcNow -lt $stopDeadline) {
            Start-Sleep -Milliseconds 50
        }
        if($script:trayPS.InvocationStateInfo.State -ne 'Running') {
            try { $script:trayPS.Dispose() } catch {}
            if($script:trayRunspace) {
                try { $script:trayRunspace.Close() } catch {}
                try { $script:trayRunspace.Dispose() } catch {}
            }
        }
        $script:trayPS = $null
        $script:trayRunspace = $null
    }
}

function Convert-SearchRowToMap {
    param([Parameter(Mandatory = $true)]$Row)

    $rowMap = @{}
    $cells = @()

    if($Row.PSObject.Properties['Cells']) {
        $cells = @($Row.Cells)
    } elseif($Row.PSObject.Properties['cells']) {
        $cells = @($Row.cells)
    }

    foreach($cell in $cells) {
        $key = $cell.Key
        if([string]::IsNullOrWhiteSpace($key)) { $key = $cell.key }
        if([string]::IsNullOrWhiteSpace($key)) { continue }

        $value = $cell.Value
        if($null -eq $value) { $value = $cell.value }

        $rowMap[$key] = $value
    }

    return $rowMap
}

function Get-SearchResultRows {
    param([Parameter(Mandatory = $true)]$SearchResponse)

    $rows = @()

    if($SearchResponse.PSObject.Properties['PrimaryQueryResult']) {
        $rows = @($SearchResponse.PrimaryQueryResult.RelevantResults.Table.Rows)
    } elseif($SearchResponse.PSObject.Properties['d']) {
        $rows = @($SearchResponse.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results)
    }

    return @($rows)
}

function Get-SharePointDocumentLibrariesFromSearch {
    param([Parameter(Mandatory = $true)][string]$SearchRootUrl)

    $rowLimit = 500
    $startRow = 0
    $foundLibraries = [System.Collections.Generic.List[hashtable]]::new()
    # track whether discovery completed normally. If a page keeps failing after retries we treat the
    # whole result set as INCOMPLETE and let the caller skip the destructive delete phase this run rather
    # than deleting valid shortcuts based on a partial view.
    $script:searchIncomplete = $false

    while($true) {
        $queryText = [System.Web.HttpUtility]::UrlEncode("contentclass:STS_List_DocumentLibrary")
        $selectProperties = [System.Web.HttpUtility]::UrlEncode("Title,Path,ListId,SiteId,WebId,SPWebUrl,SPSiteUrl,SiteName")
        $queryUri = "$SearchRootUrl/_api/search/query?querytext='$queryText'&trimduplicates=false&rowlimit=$rowLimit&startrow=$startRow&selectproperties='$selectProperties'"

        try {
            $searchResponse = New-GraphQuery -resource $global:octo.sharepointUrl -Uri $queryUri -Method GET -MaxAttempts 3
        } catch {
            Write-Log "SharePoint Search paging failed at startRow=$startRow after retries; treating results as INCOMPLETE so the delete phase is skipped this run: $($_.Exception.Message)" "WARN"
            $script:searchIncomplete = $true
            break
        }
        $rows = @(Get-SearchResultRows -SearchResponse $searchResponse)

        if(@($rows).Count -eq 0) {
            break
        }

        foreach($row in $rows) {
            $map = Convert-SearchRowToMap -Row $row

            $listId = [string]$map.ListId
            $siteId = [string]$map.SiteId
            $webId = [string]$map.WebId
            $listPath = [string]$map.Path
            $siteWebUrl = [string]$map.SPWebUrl
            $derivedWebUrl = Get-WebUrlFromListPath -ListPath $listPath
            if(-not [string]::IsNullOrWhiteSpace($derivedWebUrl)) {
                $siteWebUrl = $derivedWebUrl
            }
            if([string]::IsNullOrWhiteSpace($siteWebUrl)) {
                $siteWebUrl = [string]$map.SPSiteUrl
            }
            if([string]::IsNullOrWhiteSpace($siteWebUrl)) {
                $siteWebUrl = $listPath
            }

            if([string]::IsNullOrWhiteSpace($listId) -or [string]::IsNullOrWhiteSpace($siteId) -or [string]::IsNullOrWhiteSpace($webId) -or [string]::IsNullOrWhiteSpace($siteWebUrl)) {
                continue
            }

            $foundLibraries.Add(@{
                listId = $listId.Trim("{}")
                siteId = $siteId.Trim("{}")
                webId = $webId.Trim("{}")
                siteWebUrl = $siteWebUrl.TrimEnd('/')
                siteCollectionUrl = [string]$map.SPSiteUrl
                siteName = [string]$map.SiteName
                listName = [string]$map.Title
                listPath = $listPath
            })
        }

        if(@($rows).Count -lt $rowLimit) {
            break
        }

        $startRow += $rowLimit
    }

    return @($foundLibraries)
}

function Get-ListMetadataWithFallback {
    param(
        [Parameter(Mandatory = $true)][hashtable]$Library,
        [Parameter(Mandatory = $true)][string]$PrimarySiteUrl
    )

    $listId = $Library.listId
    $listPath = [string]$Library.listPath
    $siteCollectionUrl = [string]$Library.siteCollectionUrl
    $derivedWebUrl = Get-WebUrlFromListPath -ListPath $listPath

    $candidateBaseUrls = [System.Collections.Generic.List[string]]::new()
    if(-not [string]::IsNullOrWhiteSpace($derivedWebUrl)) {
        $candidateBaseUrls.Add($derivedWebUrl.TrimEnd('/'))
    }
    if(-not [string]::IsNullOrWhiteSpace($siteCollectionUrl)) {
        $candidateBaseUrls.Add($siteCollectionUrl.TrimEnd('/'))
    }
    if(-not [string]::IsNullOrWhiteSpace($PrimarySiteUrl)) {
        $candidateBaseUrls.Add($PrimarySiteUrl.TrimEnd('/'))
    }

    foreach($baseUrl in $candidateBaseUrls | Select-Object -Unique) {
        try {
            return New-GraphQuery -resource $global:octo.sharepointUrl -Uri "$baseUrl/_api/lists/GetById('$listId')" -Method GET
        } catch {
            $isNotFound = $_.Exception.Message -like "*404*" -or $_.Exception.Message -like "*Not Found*"
            if(-not $isNotFound) {
                throw
            }
        }
    }

    throw "List metadata lookup failed for list '$listId' using all GetById fallbacks"
}

function Get-ShortcutMetadataMap {
    # fetch the target metadata (the hidden A2OD* remote-item fields) for EVERY existing shortcut in
    # the AutoLink folder in a SINGLE RenderListDataAsStream call, instead of one GetItemByUniqueId per
    # shortcut (the old N+1 loop). Returns a map keyed by normalized item UniqueId. The caller falls back
    # to a per-item lookup for any shortcut this bulk call did not fully resolve, so correctness is
    # preserved even if the tenant does not surface the hidden fields via RenderListDataAsStream.
    param(
        [Parameter(Mandatory = $true)][string]$WebUrl,
        [Parameter(Mandatory = $true)][string]$ListId
    )

    $map = @{}
    $viewXml = "<View><ViewFields><FieldRef Name='UniqueId'/><FieldRef Name='FileLeafRef'/><FieldRef Name='A2ODRemoteItemSiteId'/><FieldRef Name='A2ODRemoteItemWebId'/><FieldRef Name='A2ODRemoteItemListId'/><FieldRef Name='A2ODRemoteItemUniqueId'/></ViewFields><RowLimit>5000</RowLimit></View>"
    $body = @{ parameters = @{ RenderOptions = 2; ViewXml = $viewXml } } | ConvertTo-Json -Depth 5
    $uri = "$WebUrl/_api/web/lists('$ListId')/RenderListDataAsStream"

    $response = New-GraphQuery -resource $global:octo.sharepointUrl -Uri $uri -Method POST -Body $body
    foreach($row in @($response.Row)) {
        $uniqueId = ([string]$row.UniqueId).Trim('{}').ToLowerInvariant()
        if([string]::IsNullOrWhiteSpace($uniqueId)) { continue }
        $map[$uniqueId] = @{
            Name = [string]$row.FileLeafRef
            targetSiteId = [string]$row.A2ODRemoteItemSiteId
            targetWebId = [string]$row.A2ODRemoteItemWebId
            targetListId = [string]$row.A2ODRemoteItemListId
            targetItemUniqueId = [string]$row.A2ODRemoteItemUniqueId
        }
    }
    return $map
}


function Invoke-PreflightChecks {
    # cheap, actionable checks before touching Graph. Warnings are logged and surfaced in one balloon;
    # they never block the run (the consent probe below is the only thing that can, via a friendly error).
    $warnings = [System.Collections.Generic.List[string]]::new()

    # 1) Is the OneDrive client configured for a work/school account on this device?
    $oneDriveRoot = Get-LocalOneDriveRootPath
    if([string]::IsNullOrWhiteSpace($oneDriveRoot)) {
        $msg = "OneDrive does not appear to be set up for a work/school account here. Shortcuts will be created in your cloud OneDrive but may not appear in File Explorer until OneDrive is signed in and syncing."
        Write-Log "Pre-flight: $msg" "WARN"
        $warnings.Add($msg)
    } else {
        Write-Log "Pre-flight: local OneDrive folder found at $oneDriveRoot" "INFO"
    }

    # 2) Is the identity provider reachable? (ICMP is often blocked, so fall back to a TCP 443 probe.)
    try {
        $idpHost = ([System.Uri]$global:octo.idpUrl).Host
        $reachable = $false
        try { $reachable = [bool](Test-Connection -ComputerName $idpHost -Count 1 -Quiet -ErrorAction SilentlyContinue) } catch {}
        if(-not $reachable) {
            try {
                $tcpClient = New-Object System.Net.Sockets.TcpClient
                $asyncResult = $tcpClient.BeginConnect($idpHost, 443, $null, $null)
                if($asyncResult.AsyncWaitHandle.WaitOne(3000)) { $tcpClient.EndConnect($asyncResult); $reachable = $true }
                $tcpClient.Close()
            } catch {}
        }
        if(-not $reachable) {
            $msg = "The sign-in service ($idpHost) is not reachable. Check your network or proxy connection."
            Write-Log "Pre-flight: $msg" "WARN"
            $warnings.Add($msg)
        } else {
            Write-Log "Pre-flight: identity provider $idpHost is reachable" "INFO"
        }
    } catch {}

    if($warnings.Count -gt 0) {
        Update-TrayState -ShowBalloon -BalloonTitle "M365AutoLink" -BalloonMessage ($warnings -join " ") -BalloonIcon "Warning"
    }
    return @($warnings)
}

#endregion

#region Main Script

function Invoke-M365AutoLinkRun {
    try {
        Initialize-ProgressBar

        # rotate logs and start a fresh lastRun.log. Write-Log appends to the file directly, so we
        # no longer depend on Start-Transcript.
        Invoke-LogRotation

        Update-TrayState -Text "M365AutoLink - Starting mapping" -Percent 1 -ProgressText "Starting" -IsRunning

        Write-Log "=== M365AutoLink v$ScriptVersion Started ===" "INFO"
        if($DryRun) { Write-Log "*** DRY RUN MODE - no changes will be made ***" "WARN" }

        Add-Type -AssemblyName System.Web

        # pre-flight checks (OneDrive present, IdP reachable) before we touch Graph.
        Update-TrayState -Text "M365AutoLink - Checking prerequisites" -Percent 3 -ProgressText "Checking prerequisites" -IsRunning
        [void](Invoke-PreflightChecks)

        # Pre populate the token cache (this doubles as the admin-consent probe).
        Update-TrayState -Text "M365AutoLink - Authenticating" -Percent 5 -ProgressText "Authenticating" -IsRunning
        try {
            [void](Get-AccessToken -resource $global:octo.graphUrl)
        } catch {
            $consentUrl = "$($global:octo.idpUrl)/organizations/adminconsent?client_id=$($global:octo.LCClientId)"
            Write-Log "Sign-in/consent probe failed: $($_.Exception.Message)" "ERROR"
            Write-Log "If this persists, an admin may still need to approve M365AutoLink: $consentUrl" "ERROR"
            throw "M365AutoLink could not sign in. If this keeps happening, ask IT to approve access: $consentUrl"
        }

        $script:userConfig = $null
        $configuredExcludedSiteSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $userExcludedLibraryKeySet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $cachedStaticExcludedLibraryKeySet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $cachedStaticExcludedLibraryMap = @{}
        try {
            $script:userConfig = Get-OneDriveUserConfig
            foreach($siteUrl in @($script:userConfig.preferences.excludedSiteUrls)) {
                $normalizedSiteUrl = Get-NormalizedSiteUrl -SiteUrl ([string]$siteUrl)
                if(-not [string]::IsNullOrWhiteSpace($normalizedSiteUrl)) {
                    [void]$configuredExcludedSiteSet.Add($normalizedSiteUrl)
                }
            }

            foreach($libraryKey in @($script:userConfig.preferences.excludedLibraryKeys)) {
                $libraryKeyText = ([string]$libraryKey).Trim().ToLowerInvariant()
                if(-not [string]::IsNullOrWhiteSpace($libraryKeyText)) {
                    [void]$userExcludedLibraryKeySet.Add($libraryKeyText)
                }
            }

            foreach($cachedEntry in @($script:userConfig.cache.staticExcludedLibraries)) {
                $cacheKey = [string]$cachedEntry.key
                if([string]::IsNullOrWhiteSpace($cacheKey)) {
                    $cacheKey = Get-ShortcutTargetKey -SiteId ([string]$cachedEntry.siteId) -WebId ([string]$cachedEntry.webId) -ListId ([string]$cachedEntry.listId)
                }
                if([string]::IsNullOrWhiteSpace($cacheKey)) { continue }

                if($cachedStaticExcludedLibraryKeySet.Add($cacheKey)) {
                    $cachedStaticExcludedLibraryMap[$cacheKey] = @{
                        key = $cacheKey
                        siteId = [string]$cachedEntry.siteId
                        webId = [string]$cachedEntry.webId
                        listId = [string]$cachedEntry.listId
                        listName = [string]$cachedEntry.listName
                        reason = [string]$cachedEntry.reason
                        lastSeenUtc = [string]$cachedEntry.lastSeenUtc
                    }
                }
            }

            if($configuredExcludedSiteSet.Count -gt 0) {
                Write-Log "Loaded $($configuredExcludedSiteSet.Count) legacy user-configured excluded site(s) from OneDrive config" "INFO"
            }
            if($userExcludedLibraryKeySet.Count -gt 0) {
                Write-Log "Loaded $($userExcludedLibraryKeySet.Count) user-excluded librar$(if($userExcludedLibraryKeySet.Count -eq 1){'y'}else{'ies'}) from OneDrive config" "INFO"
            }
            if($cachedStaticExcludedLibraryKeySet.Count -gt 0) {
                Write-Log "Loaded $($cachedStaticExcludedLibraryKeySet.Count) cached static excluded librar$(if($cachedStaticExcludedLibraryKeySet.Count -eq 1){'y'}else{'ies'})" "INFO"
            }
        } catch {
            Write-Log "Failed to load OneDrive config, continuing with defaults: $($_.Exception.Message)" "WARN"
            $script:userConfig = Get-DefaultUserConfig
        }

        try {
            Sync-LaunchPersistence
            Save-OneDriveUserConfig -Config $script:userConfig
        } catch {
            Write-Log "Failed to synchronize launch persistence preferences: $($_.Exception.Message)" "WARN"
        }

        $registerStaticExcludedLibrary = {
            param(
                [hashtable]$Library,
                [string]$Reason,
                [string]$ListDisplayName
            )

            $cacheKey = Get-ShortcutTargetKey -SiteId ([string]$Library.siteId) -WebId ([string]$Library.webId) -ListId ([string]$Library.listId)
            if([string]::IsNullOrWhiteSpace($cacheKey)) { return }

            $resolvedListName = $ListDisplayName
            if([string]::IsNullOrWhiteSpace($resolvedListName)) { $resolvedListName = [string]$Library.listName }

            $entry = @{
                key = $cacheKey
                siteId = [string]$Library.siteId
                webId = [string]$Library.webId
                listId = [string]$Library.listId
                listName = [string]$resolvedListName
                reason = [string]$Reason
                lastSeenUtc = [DateTime]::UtcNow.ToString("o")
            }

            if($cachedStaticExcludedLibraryKeySet.Add($cacheKey) -or -not $cachedStaticExcludedLibraryMap.ContainsKey($cacheKey)) {
                $cachedStaticExcludedLibraryMap[$cacheKey] = $entry
            } else {
                $cachedStaticExcludedLibraryMap[$cacheKey] = $entry
            }
        }

        # Check if target folder exists, create if not
        Write-Log "Checking for '$FolderName' folder in OneDrive..." "INFO"
        Update-TrayState -Text "M365AutoLink - Preparing folder" -Percent 10 -ProgressText "Preparing OneDrive folder" -IsRunning
        $targetFolder = $null

        try {
            $targetFolder = New-GraphQuery -Uri "$($global:octo.graphUrl)/v1.0/me/drive/root:/$($FolderName)?`$expand=listItem" -Method "GET"
            Write-Log "Folder '$FolderName' already exists" "INFO"
        } catch {
            if ($_.Exception.Response.StatusCode -eq 404) {
                Write-Log "Creating folder '$FolderName'..." "INFO"

                $folderBody = @{
                    name = $FolderName
                    folder = @{}
                    "@microsoft.graph.conflictBehavior" = "rename"
                } | ConvertTo-Json -Depth 3

                $targetFolder = New-GraphQuery -Uri "$($global:octo.graphUrl)/v1.0/me/drive/root/children?`$expand=listItem" -Method POST -Body $folderBody
                Write-Log "Folder created successfully" "SUCCESS"
                $targetFolder = New-GraphQuery -Uri "$($global:octo.graphUrl)/v1.0/me/drive/root:/$($FolderName)?`$expand=listItem" -Method "GET"
            } else {
                throw $_
            }
        }

        #determine root onedrive url by list item:
        $urlParts = $targetFolder.listItem.webUrl -split "/personal/"
        $rootUrl = $urlParts[0]
        $userComponent = $urlParts[1].Split('/')[0]
        $libraryName = $urlParts[1].Split('/')[1]

        $rootUri = [System.Uri]::new($rootUrl)
        $tenantHost = $rootUri.Host -replace '(^[^.]+)-my(\.)', '$1$2'
        $searchRootUrl = "$($rootUri.Scheme)://$tenantHost"

        Write-Log "Discovering document libraries with SharePoint Search..." "INFO"
        Write-Log "Search root: $searchRootUrl" "INFO"
        Update-TrayState -Text "M365AutoLink - Discovering libraries" -Percent 18 -ProgressText "Discovering libraries" -IsRunning
        $discoveredLibraries = @(Get-SharePointDocumentLibrariesFromSearch -SearchRootUrl $searchRootUrl)

        if(!$discoveredLibraries -or $discoveredLibraries.Count -eq 0) {
            Write-Log "No searchable document libraries found for this user" "WARN"
            Update-TrayState -Text "M365AutoLink - No libraries found" -Percent 100 -ProgressText "Nothing to map" -IsRunning:$false -ShowBalloon -BalloonMessage "No searchable document libraries found" -BalloonIcon "Warning"
            return @{
                successCount = 0
                renameCount = 0
                skipCount = 0
                deletedCount = 0
                errorCount = 0
                existingConflictCount = 0
            }
        }

        Write-Log "Discovered $($discoveredLibraries.Count) document library search hits" "SUCCESS"

        Write-Log "Will apply the following exclusion patterns:" "INFO"
        foreach($pattern in $excludedSitesByWildcard){
            Write-Log "  - $pattern" "INFO"
        }

        Write-Log "Will apply the following inclusion patterns later (if defined):" "INFO"
        foreach($pattern in $includedSitesByWildcard){
            Write-Log "  - $pattern" "INFO"
        }

        $docLibrary = (New-GraphQuery -Uri "$($global:octo.graphUrl)/v1.0/sites/$($targetFolder.parentReference.siteId)/lists" -Method "GET") | Where-Object { $_.list.template -eq "mySiteDocumentLibrary" -and !$_.list.hidden}

        $currentShortCuts = @()

        #retrieve current shortcuts
        Write-Log "Getting target info for all current shortcuts...." "INFO"
        Update-TrayState -Text "M365AutoLink - Reading current shortcuts" -Percent 25 -ProgressText "Reading current shortcuts" -IsRunning
        $folderContents = New-GraphQuery -resource $global:octo.sharepointUrl -Uri "$rootUrl/personal/$userComponent/_api/web/GetFolderByServerRelativeUrl('/personal/$userComponent/$libraryName/$FolderName')/Files?`$top=5000&`$format=json&`$expand=listItem" -Method GET

        #sometimes, e.g. when a library is changed to sync-blocked, onedrive changes it to a folder. These should be wiped as they would only confuse the user
        New-GraphQuery -resource $global:octo.sharepointUrl -Uri "$rootUrl/personal/$userComponent/_api/web/GetFolderByServerRelativeUrl('/personal/$userComponent/$libraryName/$FolderName')/Folders?`$top=5000&`$format=json&`$expand=listItem" -Method GET | ForEach-Object {
            if($_.UniqueId){
                New-GraphQuery -resource $global:octo.sharepointUrl -Uri "$rootUrl/personal/$userComponent/_api/web/GetFolderById('$($_.UniqueId)')/DeleteObject()" -Method POST
                Write-Log "Found and deleted an unexpected folder where only links should exist. Name: $($_.Name)" "ERROR"
            }
        }

        # try to fetch all shortcut metadata in one RenderListDataAsStream call. Falls back to per-item.
        $shortcutMetadataMap = @{}
        try {
            $shortcutMetadataMap = Get-ShortcutMetadataMap -WebUrl "$rootUrl/personal/$userComponent" -ListId $docLibrary.id
            Write-Log "Fetched metadata for $($shortcutMetadataMap.Count) shortcut(s) in a single call (RenderListDataAsStream)" "INFO"
        } catch {
            Write-Log "Bulk shortcut metadata call failed, falling back to per-item lookups: $($_.Exception.Message)" "WARN"
        }

        $shortcutMetadataErrorCount = 0
        foreach($shortCut in $folderContents){
            $normalizedUniqueId = ([string]$shortCut.UniqueId).Trim('{}').ToLowerInvariant()
            $meta = $null
            if($shortcutMetadataMap.ContainsKey($normalizedUniqueId)) {
                $meta = $shortcutMetadataMap[$normalizedUniqueId]
            }

            # Fall back to a per-item lookup when the bulk call did not return this row or it lacks the
            # target fields. one transient error must not abort the whole run - skip+warn instead.
            if(-not $meta -or [string]::IsNullOrWhiteSpace([string]$meta.targetSiteId) -or [string]::IsNullOrWhiteSpace([string]$meta.targetListId)) {
                try {
                    $shortCutMetaData = (New-GraphQuery -resource $global:octo.sharepointUrl -Uri "$rootUrl/personal/$userComponent/_api/web/lists('$($docLibrary.id)')/GetItemByUniqueId('$($shortCut.UniqueId)')?`$expand=FieldValuesAsText" -Method GET)
                    $meta = @{
                        Name = $shortCutMetaData.FieldValuesAsText.FileLeafRef
                        targetSiteId = $shortCutMetaData.FieldValuesAsText.A2ODRemoteItemSiteId
                        targetWebId = $shortCutMetaData.FieldValuesAsText.A2ODRemoteItemWebId
                        targetListId = $shortCutMetaData.FieldValuesAsText.A2ODRemoteItemListId
                        targetItemUniqueId = $shortCutMetaData.FieldValuesAsText.A2ODRemoteItemUniqueId
                    }
                } catch {
                    $shortcutMetadataErrorCount++
                    Write-Log "  Could not read metadata for existing shortcut '$($shortCut.UniqueId)', skipping it this run: $($_.Exception.Message)" "WARN"
                    continue
                }
            }

            $currentShortCuts += @{
                "ID" = $shortCut.uniqueId
                "Name" = $meta.Name
                "targetSiteId" = $meta.targetSiteId
                "targetWebId" = $meta.targetWebId
                "targetListId" = $meta.targetListId
                "targetItemUniqueId" = $meta.targetItemUniqueId
            }
        }
        if($shortcutMetadataErrorCount -gt 0){
            Write-Log "Skipped $shortcutMetadataErrorCount existing shortcut(s) whose metadata could not be read this run." "WARN"
        }

        Write-Log "You currently have $($currentShortCuts.count) shortcuts" "INFO"

        $desiredShortcuts = @()
        $alreadyExistingShortcuts = [System.Collections.Generic.List[hashtable]]::new()

        # Process each site
        $successCount = 0
        $skipCount = 0
        $errorCount = 0

        Write-Log "Evaluating discovered libraries against site and library rules..." "INFO"
        $siteEvaluationCache = @{}
        $manageableLibraryTable = [ordered]@{}
        $seenLibraryKeys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $libraryTotal = [Math]::Max(1, $discoveredLibraries.Count)
        $libraryIndex = 0

        foreach($library in $discoveredLibraries) {
            $libraryIndex++
            $evalPercent = [int](25 + (30 * ($libraryIndex / $libraryTotal)))
            Update-TrayState -Text "M365AutoLink - Evaluating..." -Percent $evalPercent -ProgressText ("Evaluating {0}/{1}" -f $libraryIndex, $libraryTotal) -IsRunning

            $siteUrl = $library.siteWebUrl
            if([string]::IsNullOrWhiteSpace($siteUrl)) { continue }

            if(-not $siteEvaluationCache.ContainsKey($siteUrl)) {
                Write-Log "Checking site: $siteUrl..." "INFO"

                $siteState = @{
                    Include = $false
                    SkipReason = $null
                    SiteName = $library.siteName
                }

                $isExcluded = $false
                foreach($pattern in $excludedSitesByWildcard){
                    # -like has correct anchor semantics (a trailing * is required for prefix matches).
                    if($siteUrl -like $pattern){
                        Write-Log "  Site URL '$siteUrl' matches exclusion pattern '$pattern', skipping..." "WARN"
                        $isExcluded = $true
                        break
                    }
                }
                if($isExcluded){
                    $siteState.SkipReason = "Excluded by wildcard pattern"
                    $siteEvaluationCache[$siteUrl] = $siteState
                    continue
                }

                $isIncluded = $false
                foreach($pattern in $includedSitesByWildcard){
                    # -like has correct anchor semantics (a trailing * is required for prefix matches).
                    if($siteUrl -like $pattern){
                        $isIncluded = $true
                        break
                    }
                }
                if(-not $isIncluded){
                    Write-Log "  Site URL '$siteUrl' does not match any inclusion pattern, skipping..." "WARN"
                    $siteState.SkipReason = "Not included by wildcard pattern"
                    $siteEvaluationCache[$siteUrl] = $siteState
                    continue
                }

                try {
                # Get more site info to determine if the site is archived or read-only or other blocking properties using the sharepoint API
                    $siteDetails = New-GraphQuery -resource $global:octo.sharepointUrl -Uri "$siteUrl/_api/site" -Method "GET" -MaxAttempts 1
                    if($siteDetails.WriteLocked -or $siteDetails.ReadOnly){
                    Write-Log "  Site is locked or read only, skipping..." "WARN"
                        $siteState.SkipReason = "Site locked/read-only"
                        $siteEvaluationCache[$siteUrl] = $siteState
                        continue
                    }

                    if([string]::IsNullOrWhiteSpace($siteState.SiteName)) {
                        $siteState.SiteName = [string]$siteDetails.Title
                    }

                    $siteState.Include = $true
                    $siteEvaluationCache[$siteUrl] = $siteState
                } catch {
                    Write-Log "  Failed to evaluate site '$siteUrl': $($_.Exception.Message)" "ERROR"
                    $siteState.SkipReason = "Site evaluation error"
                    $siteEvaluationCache[$siteUrl] = $siteState
                    $errorCount++
                    continue
                }
            }

            $cachedSiteState = $siteEvaluationCache[$siteUrl]
            if(-not $cachedSiteState.Include) {
                continue
            }

            $libraryKey = "$($library.siteId)|$($library.webId)|$($library.listId)"
            if($seenLibraryKeys.Contains($libraryKey)) {
                continue
            }

            $cachedLibraryKey = Get-ShortcutTargetKey -SiteId ([string]$library.siteId) -WebId ([string]$library.webId) -ListId ([string]$library.listId)
            if(-not [string]::IsNullOrWhiteSpace($cachedLibraryKey) -and $cachedStaticExcludedLibraryKeySet.Contains($cachedLibraryKey)) {
                Write-Log "  $($library.listName) is statically excluded (cached), skipping metadata lookup..." "INFO"
                continue
            }

            $normalizedSiteUrl = Get-NormalizedSiteUrl -SiteUrl $siteUrl

            # Per-library user exclusion. The legacy per-site exclusion list is honoured here too so
            # existing site exclusions keep working until the user next saves (which migrates them).
            $isUserExcludedLibrary = ($cachedLibraryKey -and $userExcludedLibraryKeySet.Contains($cachedLibraryKey)) -or (-not [string]::IsNullOrWhiteSpace($normalizedSiteUrl) -and $configuredExcludedSiteSet.Contains($normalizedSiteUrl))
            if($isUserExcludedLibrary) {
                # Record it so it can be re-included from Manage shortcuts, but skip metadata and linking.
                Write-Log "  $($library.listName) is excluded by the user, skipping..." "INFO"
                [void]$seenLibraryKeys.Add($libraryKey)
                if(-not [string]::IsNullOrWhiteSpace($cachedLibraryKey) -and -not $manageableLibraryTable.Contains($cachedLibraryKey)) {
                    $excludedSiteName = if([string]::IsNullOrWhiteSpace([string]$cachedSiteState.SiteName)) { $siteUrl } else { [string]$cachedSiteState.SiteName }
                    $manageableLibraryTable[$cachedLibraryKey] = @{
                        key = $cachedLibraryKey
                        siteUrl = $normalizedSiteUrl
                        siteName = $excludedSiteName
                        listName = [string]$library.listName
                        itemCount = [long]0
                        isExcluded = $true
                    }
                }
                continue
            }

            try {
                $listMetaData = Get-ListMetadataWithFallback -Library $library -PrimarySiteUrl $siteUrl

                $listDisplayName = [string]$listMetaData.Title
                if([string]::IsNullOrWhiteSpace($listDisplayName)) {
                    $listDisplayName = [string]$library.listName
                }

                if($listMetaData.Hidden){
                    Write-Log "  $listDisplayName is hidden, skipping..." "WARN"
                    & $registerStaticExcludedLibrary -Library $library -Reason "Hidden library" -ListDisplayName $listDisplayName
                    continue
                }

                # Only include classic document libraries (BaseTemplate 101).
                $baseTemplate = $null
                try { $baseTemplate = [int]$listMetaData.BaseTemplate } catch {}
                if($null -ne $baseTemplate -and $baseTemplate -ne 101) {
                    Write-Log "  $listDisplayName is not a standard document library (BaseTemplate=$baseTemplate), skipping..." "WARN"
                    continue
                }

                # Exclude catalog/system libraries and known non-user libraries.
                if($listMetaData.IsCatalog -eq $true -or $listMetaData.IsSystemList -eq $true -or (Test-IsExcludedLibraryName -ListName $listDisplayName)) {
                    Write-Log "  $listDisplayName is a system/catalog library, skipping..." "WARN"
                    & $registerStaticExcludedLibrary -Library $library -Reason "System/catalog library" -ListDisplayName $listDisplayName
                    continue
                }

                if($script:ExcludedListTitleSet.Contains($listDisplayName)) {
                    Write-Log "  $listDisplayName is an excluded system title, skipping..." "WARN"
                    & $registerStaticExcludedLibrary -Library $library -Reason "Excluded list title" -ListDisplayName $listDisplayName
                    continue
                }

                $listFeatureId = Get-ListFeatureId -ListMetadata $listMetaData
                if($listFeatureId -and $script:ExcludedFeatureIdSet.Contains($listFeatureId)) {
                    Write-Log "  $listDisplayName uses excluded FeatureId '$listFeatureId', skipping..." "WARN"
                    & $registerStaticExcludedLibrary -Library $library -Reason ("Excluded FeatureId {0}" -f $listFeatureId) -ListDisplayName $listDisplayName
                    continue
                }

                # Skip libraries that are not suitable for OneDrive shortcuts.
                if($listMetaData.ForceCheckout -eq $true -or $listMetaData.ExcludeFromOfflineClient -eq $true){
                    Write-Log "  $listDisplayName requires lockout/check-out settings, skipping..." "WARN"
                    continue
                }

                if($listMetaData.ItemCount -gt $maxFileCount){
                    Write-Log "  $listDisplayName has more than $($maxFileCount) files, skipping..." "WARN"
                    continue
                }

                if($listMetaData.ItemCount -lt $minFileCount){
                    Write-Log "  $listDisplayName has less than $($minFileCount) files, skipping..." "WARN"
                    continue
                }

                [void]$seenLibraryKeys.Add($libraryKey)

                $resolvedListName = $library.listName
                if([string]::IsNullOrWhiteSpace($resolvedListName)) {
                    $resolvedListName = $listDisplayName
                }
                $resolvedListName = Get-SafeDriveItemName -Name $resolvedListName

                $resolvedSiteName = $cachedSiteState.SiteName
                if([string]::IsNullOrWhiteSpace($resolvedSiteName)) {
                    $resolvedSiteName = $siteUrl
                }

                $resolvedItemCount = 0
                try { $resolvedItemCount = [long]$listMetaData.ItemCount } catch {}

                # Record this library as a linkable candidate for the Manage shortcuts dialog.
                if(-not [string]::IsNullOrWhiteSpace($cachedLibraryKey) -and -not $manageableLibraryTable.Contains($cachedLibraryKey)) {
                    $manageableLibraryTable[$cachedLibraryKey] = @{
                        key = $cachedLibraryKey
                        siteUrl = $normalizedSiteUrl
                        siteName = $resolvedSiteName
                        listName = $resolvedListName
                        itemCount = $resolvedItemCount
                        isExcluded = $false
                    }
                }

                # Extract SharePoint IDs from search results and parent site context.
                $desiredShortcuts += @{
                    shortCut = @{
                        siteId = $library.siteId
                        siteUrl = $siteUrl
                        webId = $library.webId
                        listId = $library.listId
                        listItemUniqueId = "root"
                    }
                    siteName = $resolvedSiteName
                    listName = $resolvedListName
                    itemCount = $resolvedItemCount
                }
            }catch{
                Write-Log "  Failed to evaluate library '$($library.listName)' on '$siteUrl': $($_.Exception.Message)" "ERROR"
                $errorCount++
                continue
            }
        }

        $existingShortcutTargetKeys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach($existing in $currentShortCuts) {
            $existingTargetKey = Get-ShortcutTargetKey -SiteId ([string]$existing.targetSiteId) -WebId ([string]$existing.targetWebId) -ListId ([string]$existing.targetListId)
            if(-not [string]::IsNullOrWhiteSpace($existingTargetKey)) {
                [void]$existingShortcutTargetKeys.Add($existingTargetKey)
            }
        }

        $existingShortcutNameSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach($existing in $currentShortCuts) {
            if(-not [string]::IsNullOrWhiteSpace([string]$existing.Name)) {
                [void]$existingShortcutNameSet.Add([string]$existing.Name)
            }
        }

        $reservedShortcutNameSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach($existingName in $existingShortcutNameSet) {
            [void]$reservedShortcutNameSet.Add([string]$existingName)
        }

        $dedupedDesiredShortcuts = [System.Collections.Generic.List[hashtable]]::new()
        $desiredShortcutTargetKeys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach($desiredShortcut in $desiredShortcuts) {
            $desiredTargetKey = Get-ShortcutTargetKey -SiteId ([string]$desiredShortcut.shortCut.siteId) -WebId ([string]$desiredShortcut.shortCut.webId) -ListId ([string]$desiredShortcut.shortCut.listId)
            if([string]::IsNullOrWhiteSpace($desiredTargetKey)) {
                $dedupedDesiredShortcuts.Add($desiredShortcut)
                continue
            }
            if($desiredShortcutTargetKeys.Add($desiredTargetKey)) {
                $dedupedDesiredShortcuts.Add($desiredShortcut)
            } else {
                Write-Log "  Duplicate desired target for '$($desiredShortcut.shortcut.siteUrl)' detected in this run, skipping duplicate target..." "WARN"
                $skipCount++
            }
        }

        $desiredShortcuts = @($dedupedDesiredShortcuts)

        # Grand total of items across all libraries that will actually be linked (for the tray + bar).
        $totalLinkedItemCount = [long]0
        foreach($desiredShortcut in $desiredShortcuts) {
            try { $totalLinkedItemCount += [long]$desiredShortcut.itemCount } catch {}
        }
        Write-Log "Combined item count across $($desiredShortcuts.Count) linked librar$(if($desiredShortcuts.Count -eq 1){'y'}else{'ies'}): $('{0:N0}' -f $totalLinkedItemCount)" "INFO"

        # Per-library options feed the Manage shortcuts dialog (one row per candidate library).
        $script:lastMappedLibraryOptions = @($manageableLibraryTable.Values)
        $excludedLibraryCount = @($script:lastMappedLibraryOptions | Where-Object { $_.isExcluded }).Count
        Write-Log "Manageable libraries: $(@($script:lastMappedLibraryOptions).Count) ($excludedLibraryCount currently excluded by the user)" "INFO"
        if($script:traySync) {
            $script:traySync.HasMappedSites = (@($script:lastMappedLibraryOptions).Count -gt 0)
        }

        $createTotal = [Math]::Max(1, $desiredShortcuts.Count)
        $createIndex = 0
        foreach($desiredShortcut in $desiredShortcuts) {
            $createIndex++
            # reset per iteration so the shortcutAlreadyExists catch path can never register a stale
            # name from a previous iteration into $reservedShortcutNameSet.
            $safeShortcutName = $null
            $createPercent = [int](55 + (25 * ($createIndex / $createTotal)))
            Update-TrayState -Text "M365AutoLink - Creating shortcuts" -Percent $createPercent -ProgressText ("Creating shortcuts {0}/{1}" -f $createIndex, $createTotal) -IsRunning

            # Check if shortcut already exists
            $desiredTargetKey = Get-ShortcutTargetKey -SiteId ([string]$desiredShortcut.shortcut.siteId) -WebId ([string]$desiredShortcut.shortcut.webId) -ListId ([string]$desiredShortcut.shortcut.listId)
            $exists = $false
            if(-not [string]::IsNullOrWhiteSpace($desiredTargetKey)) {
                $exists = $existingShortcutTargetKeys.Contains($desiredTargetKey)
            }

            if ($exists) {
                Write-Log "  Shortcut already exists for '$($desiredShortcut.shortcut.siteUrl)', skipping..." "SUCCESS"
                $alreadyExistingShortcuts.Add(@{
                    siteUrl = [string]$desiredShortcut.shortcut.siteUrl
                    listName = [string]$desiredShortcut.listName
                    reason = "Already exists in current mapped shortcuts"
                    timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                    itemCount = [long]$desiredShortcut.itemCount
                })
                $skipCount++
                continue
            }

            if($DryRun) {
                Write-Log "  [DRY RUN] Would create shortcut for '$($desiredShortcut.shortcut.siteUrl)'" "INFO"
                $successCount++
                continue
            }

            try {
                # Create the shortcut
                $safeShortcutName = Get-UniqueShortcutName -BaseName $desiredShortcut.listName -SiteUrl ([string]$desiredShortcut.shortcut.siteUrl) -ExistingNames $existingShortcutNameSet -ReservedNames $reservedShortcutNameSet
                $baseShortcutName = Get-SafeDriveItemName -Name (Get-CleanedShortcutName -Name $desiredShortcut.listName)
                if($safeShortcutName -ne $baseShortcutName) {
                    Write-Log "  Adjusted shortcut name '$baseShortcutName' to '$safeShortcutName' to keep it unique" "WARN"
                }

                [void]$reservedShortcutNameSet.Add($safeShortcutName)

                $shortcutBody = @{
                    name = $safeShortcutName
                    remoteItem = @{
                        sharepointIds = $desiredShortcut.shortcut
                    }
                    "@microsoft.graph.conflictBehavior" = "rename"
                } | ConvertTo-Json -Depth 3

                Write-Log "  Creating shortcut $($desiredShortcut.shortcut.siteUrl)..." "INFO"
                $newShortCut = $Null; $newShortCut = New-GraphQuery -Uri "$($global:octo.graphUrl)/v1.0/me/drive/root/children" -Method POST -Body $shortcutBody

                # Graph only allows reliable shortcut creation in root; move it into the target folder afterward.
                if($newShortCut.id -and $targetFolder.id){
                    $moveBody = @{
                        parentReference = @{
                            id = $targetFolder.id
                        }
                    } | ConvertTo-Json -Depth 3
                    $newShortCut = New-GraphQuery -Uri "$($global:octo.graphUrl)/v1.0/me/drive/items/$($newShortCut.id)" -Method PATCH -Body $moveBody
                    Write-Log "  Moved shortcut into '$FolderName' folder" "INFO"
                }

                # Rename the shortcut if the created name differs from our desired name (Graph may append suffix)
                $cleanName = Get-SafeDriveItemName -Name (Get-CleanedShortcutName -Name $newShortCut.name)
                if($newShortCut.id -and $cleanName -ne $newShortCut.name){
                    try {
                        $renameBody = @{ name = $cleanName } | ConvertTo-Json
                        $Null = New-GraphQuery -Uri "$($global:octo.graphUrl)/v1.0/me/drive/items/$($newShortCut.id)" -Method PATCH -Body $renameBody
                        Write-Log "  Renamed shortcut from '$($newShortCut.name)' to '$cleanName'" "INFO"
                    } catch {
                        Write-Log "  Failed to rename shortcut '$($newShortCut.name)': $($_.Exception.Message)" "WARN"
                    }
                }
                # no fixed sleep - the shared retry core already backs off on 429, so steady-state runs
                # are not artificially slowed by a per-item delay.
                Write-Log "  Successfully created shortcut for '$($desiredShortcut.shortcut.siteUrl)'" "SUCCESS"
                [void]$existingShortcutNameSet.Add($safeShortcutName)
                [void]$reservedShortcutNameSet.Add($safeShortcutName)
                if(-not [string]::IsNullOrWhiteSpace($desiredTargetKey)) {
                    [void]$existingShortcutTargetKeys.Add($desiredTargetKey)
                }
                $successCount++
            }catch{
                $errorMessage = [string]$_.Exception.Message
                $isShortcutExists = $errorMessage -like "*shortcut already exists*" -or $errorMessage -like "*That shortcut already exists*"

                if($isShortcutExists) {
                    Write-Log "  Shortcut already exists for '$($desiredShortcut.shortcut.siteUrl)' (Graph duplicate detection), skipping..." "WARN"
                    $alreadyExistingShortcuts.Add(@{
                        siteUrl = [string]$desiredShortcut.shortcut.siteUrl
                        listName = [string]$desiredShortcut.listName
                        reason = "Graph reported shortcutAlreadyExists during create"
                        timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                        itemCount = [long]$desiredShortcut.itemCount
                    })
                    if(-not [string]::IsNullOrWhiteSpace($desiredTargetKey)) {
                        [void]$existingShortcutTargetKeys.Add($desiredTargetKey)
                    }
                    if(-not [string]::IsNullOrWhiteSpace($safeShortcutName)) {
                        [void]$reservedShortcutNameSet.Add($safeShortcutName)
                    }
                    $skipCount++
                    continue
                }

                Write-Log "  Failed to create shortcut for '$($desiredShortcut.shortcut.siteUrl)': $errorMessage" "ERROR"
                $errorCount++
            }
        }

        # Rename existing shortcuts if link name cleanup patterns apply
        $renameCount = 0
        if($linkNameReplacements.Count -gt 0) {
            Write-Log "Checking existing shortcuts for name cleanup..." "INFO"
            Update-TrayState -Text "M365AutoLink - Renaming shortcuts" -Percent 85 -ProgressText "Renaming shortcuts" -IsRunning
            foreach($existing in $currentShortCuts) {
                if(-not $existing.Name) { continue }
                $cleanedName = Get-SafeDriveItemName -Name (Get-CleanedShortcutName -Name $existing.Name)
                if($cleanedName -ne $existing.Name -and $currentShortCuts.Name -notcontains $cleanedName) {
                    if($DryRun) {
                        Write-Log "  [DRY RUN] Would rename '$($existing.Name)' to '$cleanedName'" "INFO"
                        $renameCount++
                        continue
                    }
                    try {
                        $renameBody = @{ name = $cleanedName } | ConvertTo-Json
                        $Null = New-GraphQuery -Uri "$($global:octo.graphUrl)/v1.0/me/drive/items/$($existing.ID)" -Method PATCH -Body $renameBody
                        Write-Log "  Renamed '$($existing.Name)' to '$cleanedName'" "SUCCESS"
                        $renameCount++
                    } catch {
                        Write-Log "  Failed to rename '$($existing.Name)': $($_.Exception.Message)" "WARN"
                    }
                }
            }
        }

        #delete shortcuts user should no longer have access to
        $deletedCount = 0
        $deleteTotal = [Math]::Max(1, $currentShortCuts.Count)
        $deleteIndex = 0
        $desiredTargetKeySet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach($desired in $desiredShortcuts) {
            $desiredTargetKey = Get-ShortcutTargetKey -SiteId ([string]$desired.shortcut.siteId) -WebId ([string]$desired.shortcut.webId) -ListId ([string]$desired.shortcut.listId)
            if(-not [string]::IsNullOrWhiteSpace($desiredTargetKey)) {
                [void]$desiredTargetKeySet.Add($desiredTargetKey)
            }
        }

        $lastDesiredCount = 0
        try { $lastDesiredCount = [int]$script:userConfig.diagnostics.lastDesiredCount } catch {}
        $desiredCount = $desiredTargetKeySet.Count

        # Consume the one-shot bypass: a shrink the user just made on purpose (via Manage shortcuts) is not
        # the partial-Search-outage scenario the ratio guard protects against, so skip that check this run.
        $userInitiatedShrink = $script:bypassDeletionRatioOnce
        $script:bypassDeletionRatioOnce = $false

        $skipDeletion = $false
        $skipReason = $null
        if($script:searchIncomplete) {
            $skipDeletion = $true
            $skipReason = "SharePoint Search returned incomplete results"
        } elseif(-not $userInitiatedShrink -and $lastDesiredCount -gt 0 -and $desiredCount -lt [int][math]::Floor($lastDesiredCount * (1 - $DeletionSafetyRatio))) {
            $skipDeletion = $true
            $skipReason = ("the desired set shrank from {0} to {1} (more than {2}%)" -f $lastDesiredCount, $desiredCount, [int]($DeletionSafetyRatio * 100))
        }

        if($skipDeletion) {
            Write-Log "Deletion phase SKIPPED as a safety measure: $skipReason. No shortcuts were removed this run." "WARN"
            Update-TrayState -ShowBalloon -BalloonTitle "M365AutoLink" -BalloonMessage "Skipped removing shortcuts this run as a safety measure ($skipReason). Nothing was deleted." -BalloonIcon "Warning"
        } else {
            foreach($existing in $currentShortCuts) {
                $deleteIndex++
                $deletePercent = [int](90 + (8 * ($deleteIndex / $deleteTotal)))
                Update-TrayState -Text "M365AutoLink - Removing obsolete shortcuts" -Percent $deletePercent -ProgressText ("Removing obsolete shortcuts {0}/{1}" -f $deleteIndex, $deleteTotal) -IsRunning

                $shouldExist = $false
                $existingTargetKey = Get-ShortcutTargetKey -SiteId ([string]$existing.targetSiteId) -WebId ([string]$existing.targetWebId) -ListId ([string]$existing.targetListId)
                if(-not [string]::IsNullOrWhiteSpace($existingTargetKey)) {
                    $shouldExist = $desiredTargetKeySet.Contains($existingTargetKey)
                } else {
                    foreach($desired in $desiredShortcuts) {
                        if ($existing.targetSiteId -eq $desired.shortcut.siteId -and $existing.targetWebId -eq $desired.shortcut.webId -and $existing.targetListId -eq $desired.shortcut.listId) {
                            $shouldExist = $true
                            break
                        }
                    }
                }

                if($shouldExist) {
                    # Target is back / still valid
                    continue
                }

                if($DryRun) {
                    Write-Log "  [DRY RUN] Would delete obsolete shortcut '$($existing.Name)'" "INFO"
                    $deletedCount++
                    continue
                }
                try {
                    Write-Log "  Deleting obsolete shortcut '$($existing.Name)' (ID: $($existing.ID))..." "INFO"
                    New-GraphQuery -Uri "$($global:octo.graphUrl)/v1.0/me/drive/items/$($existing.ID)" -Method DELETE
                    $deletedCount++
                    Write-Log "  Successfully deleted obsolete shortcut" "SUCCESS"
                } catch {
                    Write-Log "  Failed to delete obsolete shortcut '$($existing.Name)': $($_.Exception.Message)" "ERROR"
                    $errorCount++
                }
            }
        }

        # Summary
        $modeLabel = if($DryRun) { " (DRY RUN)" } else { "" }
        Write-Log "=== Summary$modeLabel ===" "INFO"
        Write-Log "Shortcuts Created: $successCount" "SUCCESS"
        Write-Log "Shortcuts Renamed: $renameCount" "SUCCESS"
        Write-Log "Shortcuts Skipped: $skipCount" "INFO"
        Write-Log "Shortcuts Deleted: $deletedCount" "SUCCESS"
        if($errorCount -gt 0){
            Write-Log "Errors: $errorCount" "ERROR"
        }else{
            Write-Log "Errors: $errorCount" "SUCCESS"
        }

        if($alreadyExistingShortcuts.Count -gt 0) {
            Write-Log "Already existing: $($alreadyExistingShortcuts.Count)" "SUCCESS"
        }

        $script:lastAlreadyExistingShortcuts = @($alreadyExistingShortcuts)
        if($script:traySync) {
            $script:traySync.HasExistingConflicts = (@($script:lastAlreadyExistingShortcuts).Count -gt 0)
        }
        if(-not $script:userConfig) {
            $script:userConfig = Get-DefaultUserConfig
        }
        if(-not $script:userConfig.cache) {
            $script:userConfig.cache = @{ staticExcludedLibraries = @() }
        }
        $script:userConfig.cache.staticExcludedLibraries = @($cachedStaticExcludedLibraryMap.Values | Sort-Object key)
        $script:userConfig.diagnostics.lastAlreadyExisting = @($script:lastAlreadyExistingShortcuts)
        $script:userConfig.diagnostics.totalItemCount = $totalLinkedItemCount

        if(-not $script:searchIncomplete) {
            $script:userConfig.diagnostics.lastDesiredCount = [int]$desiredCount
        }
        try {
            Save-OneDriveUserConfig -Config $script:userConfig
        } catch {
            Write-Log "Failed to save diagnostics to OneDrive config: $($_.Exception.Message)" "WARN"
        }

        Update-TrayState -Text "M365AutoLink - Mapping complete" -Percent 100 -ProgressText "Completed" -IsRunning:$false

        # Reflect the combined item count on the tray (icon color + tooltip) and warn once if we are
        # near/over the limit where Explorer may misbehave. The balloon links to the KB article.
        $totalItemCountStatus = Get-TotalItemCountStatus -TotalItemCount $totalLinkedItemCount
        Update-TrayState -TotalItemCount $totalLinkedItemCount -ItemCountStatus $totalItemCountStatus
        if($totalItemCountStatus -ne "ok") {
            $itemCountBalloon = if($totalItemCountStatus -eq "over") {
                "Your linked libraries now hold $('{0:N0}' -f $totalLinkedItemCount) items, over the $('{0:N0}' -f $totalItemCountWarningThreshold) limit. Explorer may not show all folders. Click to learn how to reduce this."
            } else {
                "Your linked libraries hold $('{0:N0}' -f $totalLinkedItemCount) items, approaching the $('{0:N0}' -f $totalItemCountWarningThreshold) limit. Click to learn more."
            }
            $itemCountBalloonIcon = if($totalItemCountStatus -eq "over") { "Warning" } else { "Info" }
            Update-TrayState -ShowBalloon -BalloonTitle "M365AutoLink" -BalloonMessage $itemCountBalloon -BalloonIcon $itemCountBalloonIcon -BalloonClickUrl $ItemCountHelpLink
            Write-Log "Combined linked item count is $totalItemCountStatus the limit ($('{0:N0}' -f $totalLinkedItemCount)/$('{0:N0}' -f $totalItemCountWarningThreshold))" "WARN"
        }

        Write-Log "=== Script Completed ===" "SUCCESS"

        return @{
            successCount = $successCount
            renameCount = $renameCount
            skipCount = $skipCount
            deletedCount = $deletedCount
            errorCount = $errorCount
            existingConflictCount = @($alreadyExistingShortcuts).Count
        }
    } catch {
        Update-TrayState -Text "M365AutoLink - Error" -ProgressText "Failed" -IsRunning:$false
        Write-Log "Fatal error: $($_.Exception.Message)" "ERROR"
        Write-Log $_.ScriptStackTrace "ERROR"
        throw
    } finally {
        Stop-ProgressBar
    }
}

$runInTrayMode = $EnableSystemTrayIcon -and $KeepRunningInTray

if($Uninstall) {
    # Uninstall short-circuits everything else: no deployment, no tray, no mapping run.
    Invoke-Uninstall -DeployToPath $deployToPath
    return
}

if(Test-RunningUnderIntune) {
    # Running under Intune: Intune times out PowerShell scripts and would flag M365AutoLink as failed
    # because the tray "keep running" loop never returns. Instead we only copy the script to its
    # permanent home, apply the configured persistence, and kick off a fully detached run - then exit
    # right away. Intune records a quick, successful execution while the user still gets the tray icon,
    # shortcuts and mapping from the detached process.
    Write-Log "=== M365AutoLink started under Intune - bootstrapping persistence + detached run ===" "INFO"

    try {
        $script:effectiveScriptPath = Invoke-SelfDeployment -DeployToPath $deployToPath
    } catch {
        Write-Log "Self-deployment under Intune failed: $($_.Exception.Message)" "WARN"
    }

    try {
        Sync-LaunchPersistence
    } catch {
        Write-Log "Persistence sync under Intune failed: $($_.Exception.Message)" "WARN"
    }

    $detachedScriptPath = Get-M365AutoLinkScriptPath
    if(Test-PathUnderIntune -Path $detachedScriptPath) {
        # No permanent copy was made (deployToPath not configured), so the only path we have is the
        # temporary Intune one which gets deleted right after we exit. Launching from there would be
        # unreliable and could re-trigger this same Intune branch, so we skip it and tell the admin how
        # to fix it.
        Write-Log "The script is still running from a temporary Intune location because `$deployToPath is not configured. Set `$deployToPath to a permanent path so M365AutoLink can copy itself there; skipping the detached run to avoid launching from a location Intune will delete." "WARN"
    } else {
        [void](Start-DetachedM365AutoLinkRun -ScriptPath $detachedScriptPath)
    }

    Write-Log "=== Intune bootstrap complete - exiting so Intune does not time out ===" "SUCCESS"
    return
}

$script:instanceMutex = $null
$script:mutexAcquired = $false
$script:runNowEvent = $null
try {
    # single-instance guard for this user session. If another instance already owns the mutex, signal
    # it to run (via a named event its tray loop polls) and exit cleanly - avoids two tray icons, duplicate
    # refresh-token rotation, log contention and auth-port collisions. "Local\" scopes it per session so
    # multi-session hosts (RDS/AVD) still get one instance per user.
    try {
        $createdNew = $false
        $script:instanceMutex = New-Object System.Threading.Mutex($false, "Local\M365AutoLink", [ref]$createdNew)
        try {
            $script:mutexAcquired = $script:instanceMutex.WaitOne(0)
        } catch [System.Threading.AbandonedMutexException] {
            # Previous owner crashed without releasing; we now own it.
            $script:mutexAcquired = $true
        }
    } catch {
        # If the mutex machinery is unavailable for any reason, don't block the run.
        $script:mutexAcquired = $true
    }

    try {
        $script:runNowEvent = New-Object System.Threading.EventWaitHandle($false, [System.Threading.EventResetMode]::AutoReset, "Local\M365AutoLink_RunNow")
    } catch {}

    if(-not $script:mutexAcquired) {
        Write-Log "Another M365AutoLink instance is already running in this session; signaling it to run and exiting." "INFO"
        try { if($script:runNowEvent) { [void]$script:runNowEvent.Set() } } catch {}
        return
    }

    # must run before the tray runspace (same process) creates its first window.
    Set-M365ProcessDpiAwareness

    $script:localOneDriveRootPath = Get-LocalOneDriveRootPath
    $script:localShortcutFolderPath = Get-LocalShortcutFolderPath -FolderName $FolderName
    # Copy the script to its permanent home (if configured) before any persistence is created, so that
    # shortcuts/scheduled tasks/run keys point at the permanent copy rather than a temporary deploy path.
    $script:effectiveScriptPath = Invoke-SelfDeployment -DeployToPath $deployToPath
    Initialize-TrayIcon
    Update-TrayState -Text "M365AutoLink - Ready" -Percent 0 -ProgressText "Waiting to start"

    # track the last run time so we can auto-refresh every $AutoRefreshHours while resident in the tray.
    $script:lastRunCompletedAt = $null

    $runRequested = $true
    while($true) {
        if($script:traySync -and $script:traySync.ExitRequested) {
            break
        }

        if(-not $runRequested) {
            if(-not $runInTrayMode) { break }
            if($script:traySync -and $script:traySync.RequestRerun) {
                $script:traySync.RequestRerun = $false
                $runRequested = $true
            } elseif($script:runNowEvent -and $script:runNowEvent.WaitOne(0)) {
                # another launch signaled us to run instead of starting a second instance.
                Write-Log "Run requested by another launch (single-instance signal)" "INFO"
                $runRequested = $true
            } elseif($script:traySync -and $script:traySync.RequestManageShortcuts) {
                $script:traySync.RequestManageShortcuts = $false
                Write-Log "Tray action received: Manage shortcuts" "INFO"
                Update-TrayState -Text "M365AutoLink - Opening shortcuts" -ProgressText "Opening shortcut manager"
                Invoke-ManageShortcuts
                Start-Sleep -Milliseconds 100
                continue
            } elseif($AutoRefreshHours -gt 0 -and $script:lastRunCompletedAt -and ((Get-Date) - $script:lastRunCompletedAt).TotalHours -ge $AutoRefreshHours) {
                # periodic auto-refresh interval elapsed.
                Write-Log "Auto-refresh interval ($AutoRefreshHours h) elapsed; starting a run." "INFO"
                $runRequested = $true
            } elseif($AutoRefreshHours -gt 0 -and $script:traySync -and $script:traySync.ResumeDetected) {
                # refresh shortly after the device resumes from sleep (only when auto-refresh is on).
                $script:traySync.ResumeDetected = $false
                Write-Log "Device resumed from sleep; starting a refresh run." "INFO"
                $runRequested = $true
            } else {
                if($script:traySync) { $script:traySync.ResumeDetected = $false }
                Start-Sleep -Milliseconds 250
                continue
            }
        }

        $runRequested = $false
        if($script:traySync) {
            $script:traySync.IsRunning = $true
        }

        try {
            $summary = Invoke-M365AutoLinkRun
            if($script:traySync) {
                $script:traySync.HasCompletedRun = $true
            }

            # one-time onboarding after the first successful run that actually resulted in shortcuts -
            # explain where they are, that OneDrive needs a moment to sync them, and where Manage lives.
            # If a run produced no shortcuts we skip (and do NOT drop the marker) so onboarding can still
            # fire on a later run that does create some.
            $shortcutsPresent = (([int]$summary.successCount) + ([int]$summary.existingConflictCount)) -gt 0
            try {
                $onboardMarker = Join-Path -Path ([System.IO.Path]::GetDirectoryName($global:octo.LogPath)) -ChildPath ".onboarded"
                if($shortcutsPresent -and -not (Test-Path -LiteralPath $onboardMarker)) {
                    $onboardMessage = "M365AutoLink created your shortcuts in the '$FolderName' folder in OneDrive. Allow a few minutes for OneDrive to sync them into File Explorer. Tip: right-click this tray icon > Manage shortcuts to include or exclude libraries."
                    $onboardClickTarget = if(-not [string]::IsNullOrWhiteSpace($script:localShortcutFolderPath)) { $script:localShortcutFolderPath } else { [string]$script:localOneDriveRootPath }
                    Update-TrayState -ShowBalloon -BalloonTitle "M365AutoLink is set up" -BalloonMessage $onboardMessage -BalloonIcon "Info" -BalloonClickUrl $onboardClickTarget
                    New-Item -ItemType File -Path $onboardMarker -Force | Out-Null
                }
            } catch {}

            # Individual per-library/per-item errors during mapping only increment $summary.errorCount and are
            # recorded in the log - we deliberately do NOT toast the user about them. Only critical failures
            # (e.g. no Graph access, nothing works) throw and surface the Error balloon in the catch block below.
            $hasIssues = ($summary.errorCount -gt 0)
            $idleProgressText = if($hasIssues) { "Idle - completed with errors (see log)" } else { "Idle - click Run now" }
            Update-TrayState -Text "M365AutoLink - Idle" -Percent 100 -ProgressText $idleProgressText -IsRunning:$false
        } catch {
            if($script:traySync) {
                $script:traySync.HasCompletedRun = $true
            }
            Update-TrayState -Text "M365AutoLink - Error" -Percent 0 -ProgressText "Run failed - check log" -IsRunning:$false -ShowBalloon -BalloonMessage $_.Exception.Message -BalloonIcon "Error"
            if(-not $runInTrayMode) {
                throw
            }
        }

        $script:lastRunCompletedAt = Get-Date

        if(-not $runInTrayMode) {
            break
        }
    }
} finally {
    Stop-TrayIcon
    # release + dispose the single-instance mutex so the next launch can become primary.
    if($script:mutexAcquired -and $script:instanceMutex) {
        try { $script:instanceMutex.ReleaseMutex() } catch {}
    }
    if($script:instanceMutex) { try { $script:instanceMutex.Dispose() } catch {} }
    if($script:runNowEvent) { try { $script:runNowEvent.Dispose() } catch {} }
}

#endregion
