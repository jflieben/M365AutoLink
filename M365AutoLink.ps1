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

.GRAPH PERMISSIONS REQUIRED (Delegated)
    - Files.ReadWrite.All     - Create shortcuts in OneDrive
    - Team.ReadBasic.All      - Read Teams membership
    - Sites.Read.All          - Read SharePoint site information

.AUTHENTICATION FLOW
    1. Cached Refresh Token - From previous successful authentication (completely silent)
    2. Silent Browser Auth - Opens browser in the background to get tokens silently (if SSO is properly configured)
    3. Interactive Browser Auth - Opens browser for user to sign in (first time only)
    
    After first authentication, the refresh token is cached and all subsequent runs are silent until the token expires

.NOTES
    Author: Jos Lieben
    Date: 2026-03-22
    Copyright/License: https://www.lieben.nu/liebensraum/commercial-use/ (Commercial (re)use not allowed without prior written consent by the author, otherwise free to use/modify as long as header are kept intact)
    Microsoft doc: https://support.microsoft.com/en-us/office/add-shortcuts-to-shared-folders-in-onedrive-d66b1347-99b7-4470-9360-ffc048d35a33
    Always test carefully, use at your own risk, author takes no responsibility for this script
    
.EXAMPLE
    .\M365AutoLink.ps1
#>

##########START CONFIGURATION#############################
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
$LaunchModes = @('AtLogon')

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
$script:localOneDriveRootPath = $null
$script:localShortcutFolderPath = $null
$script:effectiveScriptPath = $null
$script:allowedLaunchModes = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach($launchMode in @('Desktop', 'Start Menu', 'AtLogon')) {
    [void]$script:allowedLaunchModes.Add($launchMode)
}

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

#region Helper Functions
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
        $wildcardPattern = "^" + [regex]::Escape($pattern) -replace "\\\*", ".*"
        if($ListName.ToLowerInvariant() -match $wildcardPattern) {
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

function Get-PowerShellLaunchCommand {
    param(
        [Parameter(Mandatory = $true)][string]$ScriptPath,
        [Parameter(Mandatory = $true)][string]$PowerShellExe,
        [switch]$HiddenWindow
    )

    $arguments = '-NoLogo -NoProfile -ExecutionPolicy Bypass -Sta -File "{0}"' -f $ScriptPath
    if($HiddenWindow) {
        $arguments = '-NoLogo -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Sta -File "{0}"' -f $ScriptPath
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
            $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -StartWhenAvailable
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
                [void][Win32.NativeMethods]::ReleaseCapture()
                [void][Win32.NativeMethods]::SendMessage($Form.Handle, 0xA1, 0x2, 0)
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

    if($PSBoundParameters.ContainsKey('Body')) {
        return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -Body $Body -ContentType $ContentType -ErrorAction Stop -TimeoutSec 120 -UserAgent "ISV|LiebenConsultancy|M365AutoLink|1.0"
    }

    return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -ErrorAction Stop -TimeoutSec 120 -UserAgent "ISV|LiebenConsultancy|M365AutoLink|1.0"
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

function Show-InfoDialog {
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][string]$Message
    )

    try {
        [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
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

    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

    $form = New-Object Windows.Forms.Form
    $form.Text = "M365AutoLink - Manage shortcuts"
    $form.StartPosition = "CenterScreen"
    $form.MaximizeBox = $false
    $form.AutoScaleMode = [Windows.Forms.AutoScaleMode]::None
    # Set the borderless style BEFORE the size so ClientSize reflects the borderless area. If the size
    # is applied while the default (sizable) border is still active, the client area ends up narrower
    # and shorter than expected, which pushes the table/buttons off the right and bottom edges.
    $form.FormBorderStyle = [Windows.Forms.FormBorderStyle]::None
    $form.ClientSize = New-Object Drawing.Size(1040, 600)
    $form.MinimumSize = $form.Size
    $form.TopMost = $true
    $form.BackColor = [Drawing.Color]::FromArgb(246, 248, 252)
    $form.Padding = New-Object Windows.Forms.Padding(1)

    Set-RoundedFormRegion -Form $form -Radius 10

    # Drive all geometry from the measured client area so nothing overflows the window edges.
    $pad = 12
    $footerH = 48
    $clientW = $form.ClientSize.Width
    $clientH = $form.ClientSize.Height
    $contentWidth = $clientW - ($pad * 2)

    $headerPanel = New-Object Windows.Forms.Panel
    $headerPanel.Location = New-Object Drawing.Point(0, 0)
    $headerPanel.Size = New-Object Drawing.Size($clientW, 74)
    $headerPanel.BackColor = [Drawing.Color]::FromArgb(33, 37, 43)

    $titleLabel = New-Object Windows.Forms.Label
    $titleLabel.Location = New-Object Drawing.Point($pad, 10)
    $titleLabel.Size = New-Object Drawing.Size(($clientW - 56), 24)
    $titleLabel.Font = New-Object Drawing.Font("Segoe UI", 11, [Drawing.FontStyle]::Bold)
    $titleLabel.ForeColor = [Drawing.Color]::FromArgb(237, 244, 252)
    $titleLabel.Text = "Manage shortcuts"

    $subLabel = New-Object Windows.Forms.Label
    $subLabel.Location = New-Object Drawing.Point($pad, 35)
    $subLabel.Size = New-Object Drawing.Size(($clientW - 56), 28)
    $subLabel.Font = New-Object Drawing.Font("Segoe UI", 9)
    $subLabel.ForeColor = [Drawing.Color]::FromArgb(191, 205, 223)
    $subLabel.Text = "Tick the Exclude box to stop syncing a library. Saving applies your changes and re-runs automatically."

    $headerCloseButton = New-Object Windows.Forms.Button
    $headerCloseButton.Text = "[X]"
    $headerCloseButton.Location = New-Object Drawing.Point(($clientW - 54), 9)
    $headerCloseButton.Size = New-Object Drawing.Size(34, 24)
    $headerCloseButton.FlatStyle = [Windows.Forms.FlatStyle]::Flat
    $headerCloseButton.FlatAppearance.BorderSize = 1
    $headerCloseButton.FlatAppearance.BorderColor = [Drawing.Color]::FromArgb(95, 108, 124)
    $headerCloseButton.BackColor = [Drawing.Color]::FromArgb(58, 65, 75)
    $headerCloseButton.ForeColor = [Drawing.Color]::FromArgb(237, 244, 252)
    $headerCloseButton.Font = New-Object Drawing.Font("Segoe UI", 8.5, [Drawing.FontStyle]::Bold)
    $headerCloseButton.UseVisualStyleBackColor = $false
    $headerCloseButton.Cursor = [Windows.Forms.Cursors]::Hand
    $headerCloseButton.Anchor = [Windows.Forms.AnchorStyles]::Top -bor [Windows.Forms.AnchorStyles]::Right
    $headerCloseButton.Add_Click({ $form.DialogResult = [Windows.Forms.DialogResult]::Cancel; $form.Close() })

    $headerPanel.Controls.Add($titleLabel)
    $headerPanel.Controls.Add($subLabel)
    $headerPanel.Controls.Add($headerCloseButton)
    $headerCloseButton.BringToFront()

    # Capacity bar: shows how much of the 1M sync "budget" the currently INCLUDED sites consume.
    $capPanel = New-Object Windows.Forms.Panel
    $capPanel.Location = New-Object Drawing.Point($pad, 82)
    $capPanel.Size = New-Object Drawing.Size($contentWidth, 48)
    $capPanel.BackColor = [Drawing.Color]::FromArgb(246, 248, 252)

    $capLabel = New-Object Windows.Forms.Label
    $capLabel.Location = New-Object Drawing.Point(0, 0)
    $capLabel.Size = New-Object Drawing.Size($contentWidth, 18)
    $capLabel.Font = New-Object Drawing.Font("Segoe UI", 9, [Drawing.FontStyle]::Bold)
    $capLabel.TextAlign = [Drawing.ContentAlignment]::MiddleLeft

    $capTrackWidth = $contentWidth
    $capTrack = New-Object Windows.Forms.Panel
    $capTrack.Location = New-Object Drawing.Point(0, 24)
    $capTrack.Size = New-Object Drawing.Size($capTrackWidth, 16)
    $capTrack.BackColor = [Drawing.Color]::FromArgb(225, 230, 238)

    $capFill = New-Object Windows.Forms.Panel
    $capFill.Location = New-Object Drawing.Point(0, 0)
    $capFill.Size = New-Object Drawing.Size(0, 16)
    $capFill.BackColor = [Drawing.Color]::FromArgb(31, 122, 49)
    $capTrack.Controls.Add($capFill)

    $capPanel.Controls.Add($capLabel)
    $capPanel.Controls.Add($capTrack)

    $listView = New-Object Windows.Forms.ListView
    $listView.Location = New-Object Drawing.Point($pad, 138)
    $listView.Size = New-Object Drawing.Size($contentWidth, ($clientH - 138 - $footerH - 8))
    $listView.View = [Windows.Forms.View]::Details
    $listView.FullRowSelect = $true
    $listView.GridLines = $true
    $listView.MultiSelect = $false
    $listView.CheckBoxes = $true
    $listView.ShowItemToolTips = $true
    $listView.Font = New-Object Drawing.Font("Segoe UI", 9)
    $listView.Anchor = [Windows.Forms.AnchorStyles]::Top -bor [Windows.Forms.AnchorStyles]::Left
    # Fixed columns total 584px; the Site column flexes to fill the rest so the columns span the full
    # table width (less ~22px for the vertical scrollbar) - no dead space and nothing pushed off-screen.
    $siteColumnWidth = [Math]::Max(200, $contentWidth - 584 - 22)
    [void]$listView.Columns.Add("Exclude", 64)
    [void]$listView.Columns.Add("Library", 180)
    [void]$listView.Columns.Add("Site", $siteColumnWidth)
    [void]$listView.Columns.Add("Items", 100, [Windows.Forms.HorizontalAlignment]::Right)
    [void]$listView.Columns.Add("Status", 90)
    [void]$listView.Columns.Add("Reason", 150)

    $excludedForeColor = [Drawing.Color]::FromArgb(150, 158, 168)
    foreach($option in ($LibraryOptions | Sort-Object @{ Expression = { [string]$_.siteName } }, @{ Expression = { [string]$_.listName } })) {
        $libraryKeyValue = [string]$option.key
        $siteUrlValue = [string]$option.siteUrl
        $libraryValue = [string]$option.listName
        if([string]::IsNullOrWhiteSpace($libraryValue)) { $libraryValue = "-" }
        $optionItemCount = [long]0
        try { $optionItemCount = [long]$option.itemCount } catch {}
        $isExcluded = [bool]$option.isExcluded

        # Item counts are unknown for excluded libraries (we skip their metadata to save API calls).
        $itemsValue = if($isExcluded) { "-" } elseif($optionItemCount -gt 0) { '{0:N0}' -f $optionItemCount } else { "0" }
        $statusValue = if($isExcluded) { "Excluded" } else { "Linked" }
        $reasonValue = if($isExcluded) { "Excluded by you" } else { "" }

        $item = New-Object Windows.Forms.ListViewItem("")
        [void]$item.SubItems.Add($libraryValue)
        [void]$item.SubItems.Add((Get-DisplaySiteUrl -SiteUrl $siteUrlValue))
        [void]$item.SubItems.Add($itemsValue)
        [void]$item.SubItems.Add($statusValue)
        [void]$item.SubItems.Add($reasonValue)
        $item.ToolTipText = $siteUrlValue
        $item.Tag = @{ key = $libraryKeyValue; itemCount = $optionItemCount }
        $item.Checked = $isExcluded
        if($isExcluded) { $item.ForeColor = $excludedForeColor }
        [void]$listView.Items.Add($item)
    }

    # Recompute the capacity bar + per-row status from the current checkbox states.
    $refreshCapacity = {
        $includedTotal = [long]0
        foreach($row in $listView.Items) {
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
            $capFill.Width = [int]($capTrackWidth * $ratio)
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
    }

    $listView.Add_ItemChecked({ & $refreshCapacity })
    & $refreshCapacity

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

    Enable-FormDrag -Form $form -DragControls @($form, $headerPanel, $titleLabel, $subLabel)

    $form.Controls.Add($headerPanel)
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

    if($dialogResult -ne [Windows.Forms.DialogResult]::OK) {
        $form.Dispose()
        return @{ isCanceled = $true; excludedLibraryKeys = @() }
    }

    $selected = [System.Collections.Generic.List[string]]::new()
    foreach($row in $listView.Items) {
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

function get-AccessToken{    
    Param(
        [Parameter(Mandatory=$true)]$resource,
        [Switch]$returnHeader
    )   

    # Try to load refresh token from disk
    if(!$global:octo.LCRefreshToken -and (Test-Path $global:octo.TokenCachePath)){
        try {
            $global:octo.LCRefreshToken = (Import-Clixml $global:octo.TokenCachePath).GetNetworkCredential().Password
            Write-Verbose "Loaded refresh token from local storage"
        } catch {
            Write-Warning "Failed to load cached token, proceeding to authentication..."
            Remove-Item $global:octo.TokenCachePath -ErrorAction SilentlyContinue
        }     
    }

    # Use cached refresh token
    if($global:octo.LCRefreshToken){
        try {
            $response = Invoke-RestMethod "$($global:octo.idpUrl)/common/oauth2/token" -Method POST -Body "resource=$([System.Web.HttpUtility]::UrlEncode($resource))&grant_type=refresh_token&refresh_token=$($global:octo.LCRefreshToken)&client_id=$($global:octo.LCClientId)" -ErrorAction Stop -Verbose:$false
            
            if($response.access_token){
                if($response.refresh_token){ 
                    $global:octo.LCRefreshToken = $response.refresh_token
                    try {
                        $tokenDir = [System.IO.Path]::GetDirectoryName($global:octo.TokenCachePath)
                        if(!(Test-Path $tokenDir)){ New-Item -ItemType Directory -Path $tokenDir -Force | Out-Null }
                        $secureToken = ConvertTo-SecureString -String $response.refresh_token -AsPlainText -Force
                        $credential = New-Object System.Management.Automation.PSCredential("RefreshToken", $secureToken)
                        $credential | Export-Clixml -Path $global:octo.TokenCachePath -Force
                    } catch {}
                }
                
                if(!$global:octo.LCCachedTokens.$resource){
                    $global:octo.LCCachedTokens.$resource = @{ "validFrom" = Get-Date; "accessToken" = $Null }
                }
                $global:octo.LCCachedTokens.$($resource).accessToken = $response.access_token
                $global:octo.LCCachedTokens.$($resource).validFrom = Get-Date
                return $response.access_token
            }
        } catch {
            Write-Warning "Cached refresh token invalid or expired, will re-authenticate..."
            $global:octo.LCRefreshToken = $Null
            Remove-Item $global:octo.TokenCachePath -ErrorAction SilentlyContinue
        }
    }

    # Browser-based authentication
    if(!$global:octo.LCRefreshToken){
        $global:octo.LCRefreshToken = Get-BrowserAuthorizationCode
    }

    if(!$global:octo.LCCachedTokens.$resource){
        $global:octo.LCCachedTokens.$resource = @{ "validFrom" = Get-Date; "accessToken" = $Null }
    }

    if(!$global:octo.LCCachedTokens.$($resource).accessToken -or $global:octo.LCCachedTokens.$($resource).validFrom -lt (Get-Date).AddMinutes(-15)){
        $response = Invoke-RestMethod "$($global:octo.idpUrl)/common/oauth2/token" -Method POST -Body "resource=$([System.Web.HttpUtility]::UrlEncode($resource))&grant_type=refresh_token&refresh_token=$($global:octo.LCRefreshToken)&client_id=$($global:octo.LCClientId)" -ErrorAction Stop -Verbose:$false
        
        if($response.access_token){
            if($response.refresh_token){ 
                $global:octo.LCRefreshToken = $response.refresh_token
                try {
                    $tokenDir = [System.IO.Path]::GetDirectoryName($global:octo.TokenCachePath)
                    if(!(Test-Path $tokenDir)){ New-Item -ItemType Directory -Path $tokenDir -Force | Out-Null }
                    $secureToken = ConvertTo-SecureString -String $response.refresh_token -AsPlainText -Force
                    $credential = New-Object System.Management.Automation.PSCredential("RefreshToken", $secureToken)
                    $credential | Export-Clixml -Path $global:octo.TokenCachePath -Force
                } catch {}
            }
            $global:octo.LCCachedTokens.$($resource).accessToken = $response.access_token
            $global:octo.LCCachedTokens.$($resource).validFrom = Get-Date
        }else{
            throw "Failed to retrieve access token!"
        }
    }

    return $global:octo.LCCachedTokens.$($resource).accessToken
}

function Get-BrowserAuthorizationCode {
    $tcpListener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, 1985)
    $tcpListener.Start()

    $redirectUri = "http://localhost:1985"
    $authUrl = "$($global:octo.idpUrl)/common/oauth2/authorize?" +
        "client_id=$($global:octo.LCClientId)" +
        "&response_type=code" +
        "&redirect_uri=$([System.Web.HttpUtility]::UrlEncode($redirectUri))" +
        "&response_mode=query" +
        "&resource=$([System.Web.HttpUtility]::UrlEncode($global:octo.graphUrl))"

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

    $tcpListener.Server.ReceiveTimeout = 300000 # 5 minutes
    
    try {
        $client = $tcpListener.AcceptTcpClient()
    } catch {
        $tcpListener.Stop()
        throw "Authentication timed out - no response received within 5 minutes"
    }
    
    Start-Sleep -Milliseconds 500
    $stream = $client.GetStream()
    $reader = New-Object System.IO.StreamReader($stream)
    $writer = New-Object System.IO.StreamWriter($stream)
    $requestLine = $reader.ReadLine()
    
    # Check for errors
    if($requestLine -match "error=([^&\s]+)"){
        $errorCode = $matches[1]
        $errorDesc = ""
        if($requestLine -match "error_description=([^&\s]+)"){
            $errorDesc = [System.Web.HttpUtility]::UrlDecode($matches[1])
        }
        $writer.Write("HTTP/1.1 200 OK`r`nContent-Type: text/html`r`n`r`n<html><body><h2>Authentication for M365AutoLink failed</h2><p>$($errorCode): $errorDesc</p></body></html>")
        $writer.Flush()
        $writer.Close();$reader.Close();$client.Close();$tcpListener.Stop()
        throw "Authentication error: $errorCode - $errorDesc"
    }
    
    # Extract authorization code
    if($requestLine -match "code=([^&\s]+)"){
        $code = $matches[1]
    }else{
        $writer.Close();$reader.Close();$client.Close();$tcpListener.Stop()
        throw "Failed to receive authorization code"
    }
    
    # Send success response
    $writer.Write("HTTP/1.1 200 OK`r`nContent-Type: text/html`r`n`r`n<html><head><script>window.close();</script></head><body><h2 style='color:green'>&#10004; M365AutoLink authentication successful!</h2><p>You can close this window.</p></body></html>")
    $writer.Flush()
    Start-Sleep -Milliseconds 500
    $writer.Close();$reader.Close();$client.Close();$tcpListener.Stop()

    Write-Host "Authorization code received, exchanging for tokens..." -ForegroundColor Cyan

    # Exchange code for tokens
    $tokenBody = @{
        grant_type    = "authorization_code"
        client_id     = $global:octo.LCClientId
        code          = $code
        redirect_uri  = $redirectUri
        resource      = $global:octo.graphUrl
    }
    
    $response = Invoke-RestMethod -Uri "$($global:octo.idpUrl)/common/oauth2/token" -Method POST -Body $tokenBody -ErrorAction Stop
    
    if ($response.refresh_token) {
        try {
            $tokenDir = [System.IO.Path]::GetDirectoryName($global:octo.TokenCachePath)
            if(!(Test-Path $tokenDir)){ New-Item -ItemType Directory -Path $tokenDir -Force | Out-Null }
            $secureToken = ConvertTo-SecureString -String $response.refresh_token -AsPlainText -Force
            $credential = New-Object System.Management.Automation.PSCredential("RefreshToken", $secureToken)
            $credential | Export-Clixml -Path $global:octo.TokenCachePath -Force
            Write-Host ""
            Write-Host "Authentication successful! Token cached for future use." -ForegroundColor Green
            Write-Host ""
        } catch { Write-Warning "Could not cache token: $_" }
        
        return $response.refresh_token
    }
    
    throw "No refresh token received from Azure AD"
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
            Write-Log "Possible fix: an admin still needs to approve this application at https://login.microsoftonline.com/organizations/adminconsent?client_id=$($ClientID)" -Level "ERROR"
            Exit 1
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
                    Write-Log "[WARNING] Transient error on attempt $attempts/$MaxAttempts, retrying in $($delay)s: $($_.Exception.Message)" -ForegroundColor Yellow
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
                $Null = [System.GC]::GetTotalMemory($true)
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
                        Write-Log "[WARNING] Transient error on attempt $attempts/$MaxAttempts, retrying in $($delay)s: $($_.Exception.Message)" -ForegroundColor Yellow
                        Start-Sleep -Seconds (1 + $delay)
                    }
                }

                if($resource -like "*sharepoint.com*"){
                    if($Data -and $Data.PSObject.TypeNames -notcontains "System.Management.Automation.PSCustomObject"){
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
    $color = switch($Level) {
        "ERROR" { "Red" }
        "WARN"  { "Yellow" }
        "SUCCESS" { "Green" }
        default { "White" }
    }
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $color
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
        })

        $script:trayRunspace = [runspacefactory]::CreateRunspace()
        $script:trayRunspace.ApartmentState = "STA"
        $script:trayRunspace.ThreadOptions = "ReuseThread"
        $script:trayRunspace.Open()
        $script:trayRunspace.SessionStateProxy.SetVariable("sync", $script:traySync)

        $script:trayPS = [powershell]::Create().AddScript({
            [void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
            [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

            try {
                $powerModeHandler = {
                    param($powerSender, $powerArgs)
                    if($powerArgs.Mode -eq [Microsoft.Win32.PowerModes]::Resume) {
                        try { $sync.RefreshIconRequested = $true } catch {}
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
            $icon.Text = "M365AutoLink"

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
                $w = 360
                $h = 50
                $pad = 8
                $iconBox = 24
                $trackH = 5
                $script:progressTrackWidth = $w - ($pad * 2) - $iconBox - 8

                $form = New-Object Windows.Forms.Form
                $form.Text = "M365AutoLink"
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

                $radius = 8
                $gp = New-Object Drawing.Drawing2D.GraphicsPath
                $gp.AddArc(0, 0, $radius * 2, $radius * 2, 180, 90)
                $gp.AddArc($w - $radius * 2 - 1, 0, $radius * 2, $radius * 2, 270, 90)
                $gp.AddArc($w - $radius * 2 - 1, $h - $radius * 2 - 1, $radius * 2, $radius * 2, 0, 90)
                $gp.AddArc(0, $h - $radius * 2 - 1, $radius * 2, $radius * 2, 90, 90)
                $gp.CloseFigure()
                $form.Region = New-Object Drawing.Region($gp)
                $gp.Dispose()

                $iconPanel = New-Object Windows.Forms.Panel
                $iconPanel.Location = New-Object Drawing.Point($pad, 7)
                $iconPanel.Size = New-Object Drawing.Size($iconBox, 24)
                $iconPanel.BackColor = [Drawing.Color]::Transparent

                $iconBitmap = New-Object Drawing.Bitmap($iconBox, 24)
                $ig = [Drawing.Graphics]::FromImage($iconBitmap)
                $ig.SmoothingMode = "AntiAlias"
                $ig.Clear([Drawing.Color]::Transparent)
                $iconFill = New-Object Drawing.SolidBrush([Drawing.Color]::FromArgb(0, 163, 255))
                $ig.FillEllipse($iconFill, 3, 9, 19, 11)
                $ig.FillEllipse($iconFill, 7, 4, 13, 10)
                $ig.FillEllipse($iconFill, 1, 10, 10, 9)
                $ig.FillEllipse($iconFill, 14, 10, 11, 9)
                $iconPen = New-Object Drawing.Pen([Drawing.Color]::White, 1.7)
                $iconPen.StartCap = $iconPen.EndCap = [Drawing.Drawing2D.LineCap]::Round
                $ig.DrawLine($iconPen, 13, 17, 13, 11)
                $ig.DrawLine($iconPen, 10, 13, 13, 11)
                $ig.DrawLine($iconPen, 16, 13, 13, 11)
                $iconPen.Dispose(); $iconFill.Dispose(); $ig.Dispose()

                $iconPicture = New-Object Windows.Forms.PictureBox
                $iconPicture.Location = New-Object Drawing.Point(0, 0)
                $iconPicture.Size = New-Object Drawing.Size($iconBox, 24)
                $iconPicture.BackColor = [Drawing.Color]::Transparent
                $iconPicture.Image = $iconBitmap
                $iconPicture.SizeMode = "CenterImage"
                $iconPanel.Controls.Add($iconPicture)

                $label = New-Object Windows.Forms.Label
                $label.Text = [string]$sync.ProgressBarText
                $label.Location = New-Object Drawing.Point(($pad + $iconBox + 8), 7)
                $label.Size = New-Object Drawing.Size(($w - ($pad * 2) - $iconBox - 8), 15)
                $label.Font = New-Object Drawing.Font("Segoe UI", 9)
                $label.ForeColor = [Drawing.Color]::FromArgb(237, 244, 252)
                $label.BackColor = [Drawing.Color]::Transparent
                $label.AutoEllipsis = $true

                $track = New-Object Windows.Forms.Panel
                $track.Location = New-Object Drawing.Point(($pad + $iconBox + 8), 32)
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
                $form.SetDesktopLocation(($screen.Right - $w - 12), ($screen.Bottom - $h - 12))

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

    while($true) {
        $queryText = [System.Web.HttpUtility]::UrlEncode("contentclass:STS_List_DocumentLibrary")
        $selectProperties = [System.Web.HttpUtility]::UrlEncode("Title,Path,ListId,SiteId,WebId,SPWebUrl,SPSiteUrl,SiteName")
        $queryUri = "$SearchRootUrl/_api/search/query?querytext='$queryText'&trimduplicates=false&rowlimit=$rowLimit&startrow=$startRow&selectproperties='$selectProperties'"

        $searchResponse = New-GraphQuery -resource $global:octo.sharepointUrl -Uri $queryUri -Method GET -MaxAttempts 3
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


#endregion

#region Main Script

function Invoke-M365AutoLinkRun {
    $transcriptStarted = $false
    try {
        Initialize-ProgressBar

        $logDir = [System.IO.Path]::GetDirectoryName($global:octo.LogPath)
        if(!(Test-Path $logDir)){ New-Item -ItemType Directory -Path $logDir -Force | Out-Null }

        try {
            Start-Transcript -Path $global:octo.LogPath -Force | Out-Null
            $transcriptStarted = $true
        } catch {}

        Update-TrayState -Text "M365AutoLink - Starting mapping" -Percent 1 -ProgressText "Starting" -IsRunning

        Write-Log "=== M365AutoLink v1.2 Started ===" "INFO"
        if($DryRun) { Write-Log "*** DRY RUN MODE - no changes will be made ***" "WARN" }

        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Web")

        # Pre populate the token cache
        Update-TrayState -Text "M365AutoLink - Authenticating" -Percent 5 -ProgressText "Authenticating" -IsRunning
        [void](Get-AccessToken -resource $global:octo.graphUrl)

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

        foreach($shortCut in $folderContents){
            $shortCutMetaData = (New-GraphQuery -resource $global:octo.sharepointUrl -Uri "$rootUrl/personal/$userComponent/_api/web/lists('$($docLibrary.id)')/GetItemByUniqueId('$($shortCut.UniqueId)')?`$expand=FieldValuesAsText" -Method GET -MaxAttempts 1)
            $currentShortCuts += @{
                "ID" = $shortCut.uniqueId
                "Name" = $shortCutMetaData.FieldValuesAsText.FileLeafRef
                "targetSiteId" = $shortCutMetaData.FieldValuesAsText.A2ODRemoteItemSiteId
                "targetWebId" = $shortCutMetaData.FieldValuesAsText.A2ODRemoteItemWebId
                "targetListId" = $shortCutMetaData.FieldValuesAsText.A2ODRemoteItemListId
                "targetItemUniqueId" = $shortCutMetaData.FieldValuesAsText.A2ODRemoteItemUniqueId
            }
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
                    $wildcardPattern = "^" + [regex]::Escape($pattern) -replace "\\\*",".*"
                    if($siteUrl -match $wildcardPattern){
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
                    $wildcardPattern = "^" + [regex]::Escape($pattern) -replace "\\\*",".*"
                    if($siteUrl -match $wildcardPattern){
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
                Start-Sleep -Milliseconds 500
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
                    [void]$reservedShortcutNameSet.Add($safeShortcutName)
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
                        Start-Sleep -Milliseconds 500
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

            if (-not $shouldExist) {
                if($DryRun) {
                    Write-Log "  [DRY RUN] Would delete obsolete shortcut '$($existing.Name)'" "INFO"
                    $deletedCount++
                    continue
                }
                try {
                    Write-Log "  Deleting obsolete shortcut '$($existing.Name)' (ID: $($existing.ID))..." "INFO"
                    New-GraphQuery -Uri "$($global:octo.graphUrl)/v1.0/me/drive/items/$($existing.ID)" -Method DELETE
                    Start-Sleep -Milliseconds 500
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
        if($transcriptStarted) {
            try { Stop-Transcript | Out-Null } catch {}
        }
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

try {
    $script:localOneDriveRootPath = Get-LocalOneDriveRootPath
    $script:localShortcutFolderPath = Get-LocalShortcutFolderPath -FolderName $FolderName
    # Copy the script to its permanent home (if configured) before any persistence is created, so that
    # shortcuts/scheduled tasks/run keys point at the permanent copy rather than a temporary deploy path.
    $script:effectiveScriptPath = Invoke-SelfDeployment -DeployToPath $deployToPath
    Initialize-TrayIcon
    Update-TrayState -Text "M365AutoLink - Ready" -Percent 0 -ProgressText "Waiting to start"

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
            } elseif($script:traySync -and $script:traySync.RequestManageShortcuts) {
                $script:traySync.RequestManageShortcuts = $false
                Write-Log "Tray action received: Manage shortcuts" "INFO"
                Update-TrayState -Text "M365AutoLink - Opening shortcuts" -ProgressText "Opening shortcut manager"
                Invoke-ManageShortcuts
                Start-Sleep -Milliseconds 100
                continue
            } else {
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
            $hasIssues = ($summary.errorCount -gt 0)
            if($hasIssues) {
                $msg = "Issues detected: errors $($summary.errorCount). Open the log for details."
                Update-TrayState -Text "M365AutoLink - Idle" -Percent 100 -ProgressText "Idle - click Run now" -IsRunning:$false -ShowBalloon -BalloonMessage $msg -BalloonIcon "Warning"
            } else {
                Update-TrayState -Text "M365AutoLink - Idle" -Percent 100 -ProgressText "Idle - click Run now" -IsRunning:$false
            }
        } catch {
            if($script:traySync) {
                $script:traySync.HasCompletedRun = $true
            }
            Update-TrayState -Text "M365AutoLink - Error" -Percent 0 -ProgressText "Run failed - check log" -IsRunning:$false -ShowBalloon -BalloonMessage $_.Exception.Message -BalloonIcon "Error"
            if(-not $runInTrayMode) {
                throw
            }
        }

        if(-not $runInTrayMode) {
            break
        }
    }
} finally {
    Stop-TrayIcon
}

#endregion
