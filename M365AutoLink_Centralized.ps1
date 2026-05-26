<#
.SYNOPSIS
    Automatically links all Microsoft Teams and SharePoint sites to targeted users' OneDrive using Managed Identity or cert based auth (no user impersonation)

.DESCRIPTION
    This script authenticates to Microsoft Graph using a Managed Identity (Azure VM, Automation Account,
    Azure Function App, etc.) and creates OneDrive shortcuts for each target user's Microsoft Teams and 
    SharePoint sites, under an "AutoLink" (configurable) folder.

    The script enumerates ALL sites in the tenant and their document libraries, applying
    include/exclude wildcard filters, file count limits, and archived/locked checks.
    For each target user, it checks the user's effective permissions on each document library
    using the SharePoint getUserEffectivePermissions REST API, and creates shortcuts only for
    libraries the user actually has access to.

    The $MinimumPermissionLevel setting controls whether "View" (read-only) access is sufficient,
    or whether "Edit" (contribute) access is required before a shortcut is created.

.REQUIREMENTS
    - PowerShell 5.x or 7.x
    - Azure Managed Identity OR Entra ID App Registration with certificate
    - Microsoft Graph API application permissions (see MANAGED IDENTITY SETUP)

.MANAGED IDENTITY SETUP
    Grant the following APPLICATION permissions to your Managed Identity's service principal:

    Microsoft Graph (AppId: 00000003-0000-0000-c000-000000000000):
    - Sites.Read.All          - Read SharePoint site information
    - Files.ReadWrite.All     - Create shortcuts in users' OneDrive
    - User.Read.All           - Read user profiles for target user enumeration
    - GroupMember.Read.All    - Read group memberships if using Group-based targeting

    SharePoint (AppId: 00000003-0000-0ff1-ce00-000000000000):
    - Sites.FullControl.All   - Access SharePoint REST APIs for permission checks

    Example PowerShell to grant Graph permissions to a Managed Identity:

        $MIObjectId = "<Your-Managed-Identity-Object-Id>"
        $GraphAppId = "00000003-0000-0000-c000-000000000000"

        Connect-MgGraph -Scopes "AppRoleAssignment.ReadWrite.All"
        $graphSp = Get-MgServicePrincipal -Filter "appId eq '$GraphAppId'"
        $permissions = @("Sites.Read.All","Files.ReadWrite.All","User.Read.All")
        foreach($perm in $permissions){
            $role = $graphSp.AppRoles | Where-Object { $_.Value -eq $perm }
            New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $MIObjectId `
                -PrincipalId $MIObjectId -ResourceId $graphSp.Id -AppRoleId $role.Id
        }

        $SPAppId = "00000003-0000-0ff1-ce00-000000000000"
        $spSp = Get-MgServicePrincipal -Filter "appId eq '$SPAppId'"
        $spRole = $spSp.AppRoles | Where-Object { $_.Value -eq "Sites.FullControl.All" }
        New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $MIObjectId `
            -PrincipalId $MIObjectId -ResourceId $spSp.Id -AppRoleId $spRole.Id

.AUTHENTICATION FLOW
    Authentication methods are tried in order (first success wins):
    1. Client Certificate (client_id + certificate) — if $ClientId, $TenantId, and certificate are configured
    2. Azure Functions / App Service identity endpoint ($env:IDENTITY_ENDPOINT)
    3. Azure VM Instance Metadata Service (IMDS) at 169.254.169.254
    4. Az PowerShell module (Connect-AzAccount -Identity)

.NOTES
    Author: Jos Lieben
    Version: 3.0
    Date: 2026-03-25
    Copyright/License: https://www.lieben.nu/liebensraum/commercial-use/ (Commercial (re)use not allowed without prior written consent by the author, otherwise free to use/modify as long as header are kept intact)
    Microsoft doc: https://support.microsoft.com/en-us/office/add-shortcuts-to-shared-folders-in-onedrive-d66b1347-99b7-4470-9360-ffc048d35a33
    Always test carefully, use at your own risk, author takes no responsibility for this script

.EXAMPLE
    .\M365AutoLink_Centralized.ps1
#>

##########START CONFIGURATION#############################
$FolderName = "AutoLink" #this is the folder created in onedrive to house all links this tool will create. Feel free to change this to something localized, the tool will auto-create it if it does not exist
#WARNING: Any pre-existing folders in above folder will be deleted!
$CloudType = "global" #global, usgov, usdod, china

# Optional: if not using Managed Identity, fill out ClientId, TenantId and either CertificateThumbprint or CertificatePath to authenticate using an Entra ID app registration with a certificate
$ClientId = ""      # Application (client) ID of your Entra ID app registration
$TenantId = ""       # Tenant ID (e.g. "contoso.onmicrosoft.com" or a GUID)
$CertificateThumbprint = ""   # e.g. "A1B2C3D4E5F6..." - SHA1 thumbprint of the certificate
$CertificatePath = ""         # e.g. "C:\certs\app.pfx" - path to a .pfx file (alternative to thumbprint)
$CertificatePassword = ""     # Password for the PFX file (leave empty if PFX has no password)

# Target user selection mode: "Group", "UserList", or "All"
# - Group: Process all members of the specified M365 group
# - UserList: Process only the specified users (by UPN/email)
# - All: Process all enabled member users in the tenant (can be slow for large tenants)
$TargetMode = "Group"

# When TargetMode = "Group", specify the group Object ID
$TargetGroupId = "f0cd4b92-4776-4e3d-8428-2077a3c59fa2"

# When TargetMode = "UserList", specify an array of UPNs
$TargetUsers = @(
    # "user1@contoso.com"
    # "user2@contoso.com"
)

#excluded sites will not be added a link if below pattern occurs in the site's URL. Use a * to match 1 or more characters
$excludedSitesByWildcard = @(
    "*/groupforanswersinvivaengagedonotdelete*"
    "*/sites/Streamvideo*"
    "*/portals/personal/*"
    "*/sites/AllCompany*"
    "*/personal/*"
    "*/contentstorage/*"
    "*/sites/contentTypeHub*"
    "*/sites/pwa"
)
#if you define included sites, only sites matching one of the patterns you enter will be linked
$includedSitesByWildcard = @(
    "https://*.sharepoint.com/sites/*"
)

#Optionally, only target sites with a connected group (Office 365 Group OR Team)
$onlyConnectedSites = $False

#link name cleanup patterns - applied to shortcut names after creation and to existing shortcuts on each run
$linkNameReplacements = @(
    @{ Pattern = " - Documents"; Replacement = "" }
    @{ Pattern = "- Documents"; Replacement = "" }
)

#below variables can be used to filter based on the number of existing files in the target location before creating a link
$maxFileCount = 300000
$minFileCount = 0

# Permission level required before creating a shortcut
# "View" = create shortcut when user has View (read) or higher permissions
# "Edit" = create shortcut only when user has Edit (contribute) or higher permissions
$MinimumPermissionLevel = "Edit"

# --- Performance tuning ---

# Initial concurrency for site enumeration AND permission check runspaces
# The adaptive throttle will scale this up/down based on 429 rate
$InitialParallelLimit = 10

# Absolute maximum concurrent threads (adaptive throttle ceiling)
$MaxParallelLimit = 25

# SharePoint REST $batch size (max sub-requests per batch call, Microsoft limit is 100)
$BatchSize = 25

# Concurrency for create/delete shortcut mutations in Phase 3
$ShortcutActionParallelLimit = 8

# Retry attempts for Graph mutation calls (create/move/rename/delete)
$GraphMutationMaxAttempts = 5

# Dry-run mode: when $true, no shortcuts are created, deleted, or renamed
$DryRun = $true

# Optional optimization: recursively pre-resolve site/list owners and members to skip expensive
# getUserEffectivePermissions calls for users that are already known to have access.
# Disabled by default to preserve existing behavior unless explicitly enabled.
$EnableRecursivePermissionPreCheck = $true

# When recursive pre-check is enabled, treat broad SharePoint "Everyone/All users" claims as allow-all.
$PreCheckIncludeEveryoneClaims = $true

# Optional diagnostics for recursive pre-check (site/list pre-allow and fallback counts).
$PreCheckVerboseDiagnostics = $false

# Safety guards for recursive principal expansion.
$PreCheckMaxRecursionDepth = 8
$PreCheckMaxExpandedPrincipals = 5000

# Optional test limiter: when > 0, only process first N filtered sites.
$MaxSitesToProcessForTesting = 0

##########END CONFIGURATION#############################


#base vars
$global:octo = @{}
$global:octo.LCCachedTokens = @{}
$global:octo.EntraGroupTargetMemberCache = @{}
$logDir = if($env:APPDATA) { "$env:APPDATA\M365AutoLink" } elseif($env:TEMP) { "$env:TEMP\M365AutoLink" } else { ".\M365AutoLink" }
$global:octo.LogPath = Join-Path $logDir "lastRun.log"

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

function New-CaseInsensitiveStringSet {
    return ,([System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase))
}

function Ensure-StringSet {
    param([object]$Set)
    if($null -eq $Set -or $Set -isnot [System.Collections.Generic.HashSet[string]]) {
        return ,(New-CaseInsensitiveStringSet)
    }
    return ,$Set
}

function Get-PrincipalIdentityKey {
    param([object]$Principal)
    if($null -eq $Principal) { return "null-principal" }
    $loginName = $null
    $principalId = $null
    try { $loginName = [string]$Principal.LoginName } catch {}
    try { $principalId = [string]$Principal.Id } catch {}

    if(-not [string]::IsNullOrWhiteSpace($loginName)) { return "login:$($loginName.ToLowerInvariant())" }
    if(-not [string]::IsNullOrWhiteSpace($principalId)) { return "id:$principalId" }
    return "unknown:$([guid]::NewGuid().ToString())"
}

function Get-UpnFromLoginName {
    param([string]$LoginName)
    if([string]::IsNullOrWhiteSpace($LoginName)) { return $null }

    $lower = $LoginName.ToLowerInvariant()
    if($lower -like "*|membership|*") {
        return ($lower -split "\|")[-1]
    }
    if($lower -match '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
        return $lower
    }
    return $null
}

function Get-EntraGroupIdFromLoginName {
    param([string]$LoginName)
    if([string]::IsNullOrWhiteSpace($LoginName)) { return $null }

    $guidRegex = '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}'
    $m = [regex]::Match($LoginName, $guidRegex)
    if($m.Success) { return $m.Value.ToLowerInvariant() }
    return $null
}

function Test-IsSpecialAllUsersPrincipal {
    param([string]$LoginName)
    if([string]::IsNullOrWhiteSpace($LoginName)) { return $false }

    $normalized = $LoginName.ToLowerInvariant()
    if($normalized -like "*spo-grid-all-users*") { return $true }
    if($normalized -like "*rolemanager|spo-grid-all-users*") { return $true }
    if($normalized -like "*everyone*") { return $true }
    if($normalized -like "*all users*") { return $true }
    if($normalized -like "c:0*.s|true") { return $true }
    if($normalized -like "c:0(.s|true") { return $true }
    return $false
}

function Test-RoleBindingAllowsAccess {
    param(
        [object]$RoleBindings,
        [string]$MinimumPermissionLevel
    )

    if(-not $RoleBindings) { return $false }
    $bindings = @($RoleBindings)
    if($bindings.Count -eq 0) { return $false }

    foreach($binding in $bindings) {
        $allowView = $false
        $allowEdit = $false

        try {
            if($binding.BasePermissions -and $null -ne $binding.BasePermissions.Low) {
                $lowVal = [long]$binding.BasePermissions.Low
                $allowView = ($lowVal -band 0x1) -ne 0
                $allowEdit = ($lowVal -band 0x4) -ne 0
            }
        } catch {}

        if(-not $allowView -and -not $allowEdit) {
            $roleName = ""
            try { $roleName = [string]$binding.Name } catch {}
            $roleName = $roleName.ToLowerInvariant()

            if($roleName -like "*full control*" -or $roleName -like "*owner*" -or $roleName -like "*design*" -or $roleName -like "*edit*" -or $roleName -like "*contribute*" -or $roleName -like "*approve*") {
                $allowEdit = $true
                $allowView = $true
            } elseif($roleName -like "*read*" -or $roleName -like "*view*") {
                $allowView = $true
            }
        }

        if($MinimumPermissionLevel -eq "Edit" -and $allowEdit) { return $true }
        if($MinimumPermissionLevel -ne "Edit" -and ($allowView -or $allowEdit)) { return $true }
    }

    return $false
}

function Add-ResolvedTargetUserToSet {
    param(
        [string]$CandidateUpn,
        [string]$CandidateId,
        [hashtable]$TargetUserByUpn,
        [hashtable]$TargetUserById,
        [System.Collections.Generic.HashSet[string]]$OutputSet
    )

    if(-not [string]::IsNullOrWhiteSpace($CandidateUpn)) {
        $k = $CandidateUpn.ToLowerInvariant()
        if($TargetUserByUpn.ContainsKey($k)) {
            [void]$OutputSet.Add($TargetUserByUpn[$k])
            return $true
        }
    }
    if(-not [string]::IsNullOrWhiteSpace($CandidateId)) {
        $idKey = $CandidateId.ToLowerInvariant()
        if($TargetUserById.ContainsKey($idKey)) {
            [void]$OutputSet.Add($TargetUserById[$idKey])
            return $true
        }
    }
    return $false
}

function Resolve-EntraGroupTargetUsers {
    param(
        [string]$GroupId,
        [hashtable]$TargetUserByUpn,
        [hashtable]$TargetUserById,
        [hashtable]$EntraGroupCache,
        [hashtable]$PreCheckStats
    )

    $resultSet = New-CaseInsensitiveStringSet
    if([string]::IsNullOrWhiteSpace($GroupId)) { return $resultSet }

    $groupKey = $GroupId.ToLowerInvariant()
    if($EntraGroupCache.ContainsKey($groupKey)) {
        $PreCheckStats.CacheHits++
        $cachedUpns = @($EntraGroupCache[$groupKey])
        foreach($u in $cachedUpns) { [void]$resultSet.Add($u) }
        return $resultSet
    }

    $PreCheckStats.CacheMisses++
    try {
        $members = @(New-GraphQuery -Uri "$($global:octo.graphUrl)/v1.0/groups/$groupKey/transitiveMembers?`$select=id,userPrincipalName" -Method GET -MaxAttempts 3)
        foreach($member in $members) {
            if($member.'@odata.type' -ne '#microsoft.graph.user') { continue }
            $memberUpn = $null
            try { $memberUpn = [string]$member.userPrincipalName } catch {}
            $memberId = $null
            try { $memberId = [string]$member.id } catch {}
            [void](Add-ResolvedTargetUserToSet -CandidateUpn $memberUpn -CandidateId $memberId -TargetUserByUpn $TargetUserByUpn -TargetUserById $TargetUserById -OutputSet $resultSet)
        }
    } catch {
        Write-Log "Recursive pre-check: failed to resolve Entra group '$groupKey': $($_.Exception.Message)" "WARN"
    }

    # Cache only compact string list for this run to keep memory footprint low.
    $EntraGroupCache[$groupKey] = @($resultSet)
    return @($resultSet)
}

function Resolve-PrincipalToTargetUsersRecursive {
    param(
        [string]$SiteWebUrl,
        [object]$Principal,
        [hashtable]$TargetUserByUpn,
        [hashtable]$TargetUserById,
        [System.Collections.Generic.HashSet[string]]$OutputSet,
        [System.Collections.Generic.HashSet[string]]$VisitedPrincipalKeys,
        [System.Collections.Generic.HashSet[string]]$VisitedSpGroupIds,
        [hashtable]$State,
        [hashtable]$EntraGroupCache,
        [hashtable]$PreCheckStats,
        [bool]$IncludeEveryoneClaims,
        [int]$Depth,
        [int]$MaxDepth,
        [int]$MaxExpandedPrincipals
    )

    if($Depth -gt $MaxDepth) {
        $PreCheckStats.RecursionCutoffs++
        return
    }

    if($null -eq $Principal) {
        return
    }

    if($State.ExpandedPrincipals -ge $MaxExpandedPrincipals) {
        $PreCheckStats.RecursionCutoffs++
        return
    }
    $State.ExpandedPrincipals++

    $principalKey = Get-PrincipalIdentityKey -Principal $Principal
    if($VisitedPrincipalKeys.Contains($principalKey)) { return }
    [void]$VisitedPrincipalKeys.Add($principalKey)

    $loginName = ""
    $principalType = 0
    $principalId = $null
    try { $loginName = [string]$Principal.LoginName } catch {}
    try { $principalType = [int]$Principal.PrincipalType } catch {}
    try { $principalId = [string]$Principal.Id } catch {}

    if($IncludeEveryoneClaims -and (Test-IsSpecialAllUsersPrincipal -LoginName $loginName)) {
        $State.AllowAllUsers = $true
        return
    }

    $principalUpn = $null
    try { $principalUpn = [string]$Principal.UserPrincipalName } catch {}
    if([string]::IsNullOrWhiteSpace($principalUpn)) {
        try { $principalUpn = [string]$Principal.Email } catch {}
    }
    if([string]::IsNullOrWhiteSpace($principalUpn)) {
        $principalUpn = Get-UpnFromLoginName -LoginName $loginName
    }
    [void](Add-ResolvedTargetUserToSet -CandidateUpn $principalUpn -CandidateId $principalId -TargetUserByUpn $TargetUserByUpn -TargetUserById $TargetUserById -OutputSet $OutputSet)

    $entraGroupId = Get-EntraGroupIdFromLoginName -LoginName $loginName
    if(-not [string]::IsNullOrWhiteSpace($entraGroupId)) {
        $groupUsers = Resolve-EntraGroupTargetUsers -GroupId $entraGroupId -TargetUserByUpn $TargetUserByUpn -TargetUserById $TargetUserById -EntraGroupCache $EntraGroupCache -PreCheckStats $PreCheckStats
        foreach($upn in $groupUsers) { [void]$OutputSet.Add($upn) }
    }

    # SharePoint group bit flag = 8
    $isSpGroup = ($principalType -band 8) -ne 0
    if($isSpGroup -and -not [string]::IsNullOrWhiteSpace($principalId)) {
        $spGroupKey = $principalId.ToLowerInvariant()
        if($VisitedSpGroupIds.Contains($spGroupKey)) { return }
        [void]$VisitedSpGroupIds.Add($spGroupKey)

        try {
            $groupMembers = @(New-GraphQuery -Uri "$SiteWebUrl/_api/web/sitegroups/GetById($principalId)/Users?`$select=Id,LoginName,PrincipalType,UserPrincipalName,Email" -Method GET -MaxAttempts 3)
            foreach($member in $groupMembers) {
                Resolve-PrincipalToTargetUsersRecursive -SiteWebUrl $SiteWebUrl -Principal $member -TargetUserByUpn $TargetUserByUpn -TargetUserById $TargetUserById -OutputSet $OutputSet -VisitedPrincipalKeys $VisitedPrincipalKeys -VisitedSpGroupIds $VisitedSpGroupIds -State $State -EntraGroupCache $EntraGroupCache -PreCheckStats $PreCheckStats -IncludeEveryoneClaims $IncludeEveryoneClaims -Depth ($Depth + 1) -MaxDepth $MaxDepth -MaxExpandedPrincipals $MaxExpandedPrincipals
            }
        } catch {
            Write-Log "Recursive pre-check: failed to enumerate SharePoint group '$principalId' on '$SiteWebUrl': $($_.Exception.Message)" "WARN"
        }
    }
}

function Get-ScopeAllowedTargetUsers {
    param(
        [string]$SiteWebUrl,
        [ValidateSet('Web','List')]
        [string]$ScopeType,
        [string]$ListId,
        [string]$MinimumPermissionLevel,
        [hashtable]$TargetUserByUpn,
        [hashtable]$TargetUserById,
        [hashtable]$EntraGroupCache,
        [hashtable]$PreCheckStats,
        [bool]$IncludeEveryoneClaims,
        [int]$MaxDepth,
        [int]$MaxExpandedPrincipals
    )

    $allowedUsers = New-CaseInsensitiveStringSet
    $state = @{ AllowAllUsers = $false; ExpandedPrincipals = 0 }
    $visitedPrincipalKeys = New-CaseInsensitiveStringSet
    $visitedSpGroupIds = New-CaseInsensitiveStringSet

    $roleUri = if($ScopeType -eq 'Web') {
        "$SiteWebUrl/_api/web/RoleAssignments?`$expand=Member,RoleDefinitionBindings"
    } else {
        "$SiteWebUrl/_api/web/lists/GetById('$ListId')/RoleAssignments?`$expand=Member,RoleDefinitionBindings"
    }

    try {
        $roleAssignments = @(New-GraphQuery -Uri $roleUri -Method GET -MaxAttempts 3)
        foreach($ra in $roleAssignments) {
            if($null -eq $ra) { continue }

            $roleBindings = $null
            $member = $null
            try { $roleBindings = $ra.RoleDefinitionBindings } catch {}
            try { $member = $ra.Member } catch {}

            if(-not (Test-RoleBindingAllowsAccess -RoleBindings $roleBindings -MinimumPermissionLevel $MinimumPermissionLevel)) { continue }
            if($null -eq $member) { continue }

            Resolve-PrincipalToTargetUsersRecursive -SiteWebUrl $SiteWebUrl -Principal $member -TargetUserByUpn $TargetUserByUpn -TargetUserById $TargetUserById -OutputSet $allowedUsers -VisitedPrincipalKeys $visitedPrincipalKeys -VisitedSpGroupIds $visitedSpGroupIds -State $state -EntraGroupCache $EntraGroupCache -PreCheckStats $PreCheckStats -IncludeEveryoneClaims $IncludeEveryoneClaims -Depth 0 -MaxDepth $MaxDepth -MaxExpandedPrincipals $MaxExpandedPrincipals
        }
    } catch {
        Write-Log "Recursive pre-check: failed to read $ScopeType role assignments on '$SiteWebUrl': $($_.Exception.Message)" "WARN"
    }

    return [PSCustomObject]@{
        AllowedUsers  = $allowedUsers
        AllowAllUsers = [bool]$state.AllowAllUsers
    }
}

function Get-AccessToken {
    Param(
        [Parameter(Mandatory=$true)]
        [string]$resource
    )

    if($global:octo.LCCachedTokens.$resource -and
       $global:octo.LCCachedTokens.$resource.accessToken -and
       $global:octo.LCCachedTokens.$resource.validFrom -gt (Get-Date).AddMinutes(-25)){
        return $global:octo.LCCachedTokens.$resource.accessToken
    }

    $token = $null
    $encodedResource = [System.Web.HttpUtility]::UrlEncode($resource)

    # Method 1: Client Certificate
    if(-not $token -and $ClientId -and $TenantId -and ($CertificateThumbprint -or $CertificatePath)){
        try {
            $cert = $null
            if($CertificateThumbprint){
                $cert = Get-ChildItem -Path "Cert:\CurrentUser\My\$CertificateThumbprint" -ErrorAction SilentlyContinue
                if(-not $cert){ $cert = Get-ChildItem -Path "Cert:\LocalMachine\My\$CertificateThumbprint" -ErrorAction SilentlyContinue }
                if(-not $cert){ throw "Certificate with thumbprint '$CertificateThumbprint' not found in CurrentUser\My or LocalMachine\My" }
            } elseif($CertificatePath) {
                if(-not (Test-Path $CertificatePath)){ throw "PFX file not found at '$CertificatePath'" }
                if($CertificatePassword){
                    $secPwd = ConvertTo-SecureString -String $CertificatePassword -AsPlainText -Force
                    $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertificatePath, $secPwd, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::MachineKeySet)
                } else {
                    $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertificatePath, $null, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::MachineKeySet)
                }
            }

            $certHash = [System.Convert]::ToBase64String($cert.GetCertHash()) -replace '\+','-' -replace '/','_' -replace '=',''
            $jwtHeader = @{ alg = "RS256"; typ = "JWT"; x5t = $certHash } | ConvertTo-Json -Compress
            $jwtHeaderB64 = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($jwtHeader)) -replace '\+','-' -replace '/','_' -replace '=',''

            $now = [System.DateTimeOffset]::UtcNow
            $jwtPayload = @{
                aud = "$($global:octo.idpUrl)/$TenantId/oauth2/v2.0/token"
                iss = $ClientId
                sub = $ClientId
                jti = [guid]::NewGuid().ToString()
                nbf = [int]$now.ToUnixTimeSeconds()
                exp = [int]$now.AddMinutes(10).ToUnixTimeSeconds()
            } | ConvertTo-Json -Compress
            $jwtPayloadB64 = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($jwtPayload)) -replace '\+','-' -replace '/','_' -replace '=',''

            $dataToSign = [System.Text.Encoding]::UTF8.GetBytes("$jwtHeaderB64.$jwtPayloadB64")
            $rsaKey = $null
            try {
                $rsaKey = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
            } catch {}
            if (-not $rsaKey) { $rsaKey = $cert.PrivateKey }
            if (-not $rsaKey) { throw "Certificate does not contain a usable RSA private key" }
            $signature = $rsaKey.SignData($dataToSign, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
            $signatureB64 = [System.Convert]::ToBase64String($signature) -replace '\+','-' -replace '/','_' -replace '=',''

            $clientAssertion = "$jwtHeaderB64.$jwtPayloadB64.$signatureB64"

            $body = @{
                grant_type                = "client_credentials"
                client_id                 = $ClientId
                client_assertion_type     = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
                client_assertion          = $clientAssertion
                scope                     = "$resource/.default"
            }
            $response = Invoke-RestMethod -Uri "$($global:octo.idpUrl)/$TenantId/oauth2/v2.0/token" `
                -Method POST -Body $body -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop -Verbose:$false
            $token = $response.access_token
        } catch {
            Write-Warning "Client certificate authentication failed: $($_.Exception.Message)"
        }
    }

    # Method 2: Azure Functions / App Service managed identity
    if(-not $token -and $env:IDENTITY_ENDPOINT -and $env:IDENTITY_HEADER){
        try {
            $response = Invoke-RestMethod -Uri "$($env:IDENTITY_ENDPOINT)?resource=$encodedResource&api-version=2019-08-01" `
                -Headers @{"X-IDENTITY-HEADER"=$env:IDENTITY_HEADER} -Method GET -ErrorAction Stop -Verbose:$false
            $token = $response.access_token
        } catch {
            Write-Warning "App Service MI endpoint failed: $($_.Exception.Message)"
        }
    }

    # Method 3: Azure VM IMDS
    if(-not $token){
        try {
            $response = Invoke-RestMethod -Uri "http://169.254.169.254/metadata/identity/oauth2/token?api-version=2018-02-01&resource=$encodedResource" `
                -Headers @{Metadata="true"} -Method GET -ErrorAction Stop -Verbose:$false
            $token = $response.access_token
        } catch {
            Write-Verbose "IMDS endpoint not available: $($_.Exception.Message)"
        }
    }

    # Method 4: Az PowerShell module fallback
    if(-not $token){
        try {
            $null = Connect-AzAccount -Identity -ErrorAction Stop
            $tokenResponse = Get-AzAccessToken -ResourceUrl $resource -ErrorAction Stop
            if($tokenResponse.Token -is [string]){
                $token = $tokenResponse.Token
            }else{
                $token = $tokenResponse.Token | ConvertFrom-SecureString -AsPlainText
            }
        } catch {
            Write-Verbose "Az PowerShell MI not available: $($_.Exception.Message)"
        }
    }

    if(-not $token){
        $methods = "IMDS, App Service MI, Az PowerShell MI"
        if($ClientId -and $TenantId -and ($CertificateThumbprint -or $CertificatePath)){ $methods += ", Client Certificate" }
        throw "Failed to acquire token for resource '$resource'. Tried: $methods."
    }

    $global:octo.LCCachedTokens.$resource = @{
        accessToken = $token
        validFrom = Get-Date
    }

    return $token
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
        [String]$ContentType = 'application/json; charset=utf-8'
    )

    function get-resourceHeaders{
        Param([string]$apiUrl)
        $tokenResource = $global:octo.graphUrl
        $isSharePoint = $false
        try {
            $parsedUri = [System.Uri]$apiUrl
            if($parsedUri.Host -match "sharepoint"){
                $tokenResource = $global:octo.sharepointUrl
                $isSharePoint = $true
            }
        } catch {}

        try{
            $token = Get-AccessToken -resource $tokenResource
        }catch{
            $null = Write-Log "Failed to acquire token for '$tokenResource': $_" -Level "ERROR"
            throw
        }
        $headers = @{
            "Authorization" = "Bearer $token"
            "Accept-Language" = "en-US"
        }
        if($isSharePoint){
            $headers['Accept'] = "application/json;odata=nometadata"
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
                    $headers = get-resourceHeaders -apiUrl $nextURL
                    $Data = $Null; $Data = (Invoke-RestMethod -Uri $nextURL -Method $Method -Headers $headers -Body $Body -ContentType $ContentType -Verbose:$False -ErrorAction Stop -UserAgent "ISV|LiebenConsultancy|M365AutoLink|3.0")
                    $attempts = $MaxAttempts
                }catch {
                    if($_.Exception.Message -like "*404*" -or $_.Exception.Message -like "*Request_ResourceNotFound*" -or $_.Exception.Message -like "*Resource*does not exist*" -or $_.Exception.Message -like "*403*" -or $_.Exception.Message -like "*409*" -or $_.Exception.StatusCode -in (401,403,409,"Unauthorized",404,"NotFound","Conflict")){
                        $nextUrl = $Null
                        throw $_
                    }
                    $statusCode = $null
                    try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {}
                    $is429 = $statusCode -eq 429 -or $_.Exception.Message -like "*429*"
                    $is5xx = $statusCode -in 500,502,503,504 -or $_.Exception.Message -like "*500*" -or $_.Exception.Message -like "*502*" -or $_.Exception.Message -like "*503*" -or $_.Exception.Message -like "*504*"
                    # 429s always retry indefinitely — do not count against MaxAttempts
                    if ($is429) { $attempts-- }
                    if ($is5xx -and $attempts -lt $MaxAttempts) { $attempts-- }
                    if ($attempts -ge $MaxAttempts) {
                        Throw $_
                    }

                    $delay = 0
                    $isTransientNetwork = $_.Exception.Message -like "*No such host is known*" -or $_.Exception.Message -like "*name or service not known*" -or $_.Exception.Message -like "*network is unreachable*" -or $_.Exception.Message -like "*connection was forcibly closed*" -or $_.Exception.Message -like "*An existing connection was forcibly closed*"
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
                        if($delay -le 0) { $delay = [math]::Min(120, [math]::Pow(5, [math]::Max(1, $attempts))) }
                    }
                    if($delay -le 0 -and $is5xx){
                        $delay = [math]::Min(120, [math]::Pow(2, [math]::Max(1, $attempts)) + (Get-Random -Minimum 1 -Maximum 6))
                    }
                    if($delay -le 0 -and $isTransientNetwork){
                        $delay = [math]::Min(10, 2 * $attempts)
                    }
                    if($delay -le 0){
                        $delay = [math]::Pow(5, $attempts)
                    }
                    $apiFamily = if($isSharePoint) { 'SharePoint' } else { 'Graph' }
                    $null = Write-Log "[$apiFamily] [WARNING] Transient error on attempt $attempts/$MaxAttempts, retrying in $($delay)s: $($_.Exception.Message)" -ForegroundColor Yellow
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
                    $attempts++
                    try {
                        $headers = get-resourceHeaders -apiUrl $nextURL
                        $Data = $Null; $Data = (Invoke-RestMethod -Uri $nextURL -Method $Method -Headers $headers -ContentType $ContentType -Verbose:$False -ErrorAction Stop -UserAgent "ISV|LiebenConsultancy|M365AutoLink|3.0")
                        $attempts = $MaxAttempts
                    }catch {
                        if(($_.Exception -and $_.Exception.StatusCode -and $_.Exception.StatusCode -in (401,403,409,"Unauthorized",404,"NotFound","Conflict")) -or ($_.Exception.Message -like "*404*" -or $_.Exception.Message -like "*403*" -or $_.Exception.Message -like "*409*" -or $_.Exception.Message -like "*Request_ResourceNotFound*" -or $_.Exception.Message -like "*Resource*does not exist*")){
                            $nextUrl = $Null
                            throw $_
                        }

                        $statusCode = $null
                        try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {}
                        $is429 = $statusCode -eq 429 -or $_.Exception.Message -like "*429*"
                        $is5xx = $statusCode -in 500,502,503,504 -or $_.Exception.Message -like "*500*" -or $_.Exception.Message -like "*502*" -or $_.Exception.Message -like "*503*" -or $_.Exception.Message -like "*504*"
                        # 429s always retry indefinitely — do not count against MaxAttempts
                        if ($is429) { $attempts-- }
                        if ($is5xx -and $attempts -lt $MaxAttempts) { $attempts-- }
                        if ($attempts -ge $MaxAttempts) {
                            $nextURL = $null
                            Throw $_
                        }

                        $delay = 0
                        $isTransientNetwork = $_.Exception.Message -like "*No such host is known*" -or $_.Exception.Message -like "*name or service not known*" -or $_.Exception.Message -like "*network is unreachable*" -or $_.Exception.Message -like "*connection was forcibly closed*" -or $_.Exception.Message -like "*An existing connection was forcibly closed*"
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
                            if($delay -le 0) { $delay = [math]::Min(120, [math]::Pow(5, [math]::Max(1, $attempts))) }
                        }
                        if($delay -le 0 -and $is5xx){
                            $delay = [math]::Min(120, [math]::Pow(2, [math]::Max(1, $attempts)) + (Get-Random -Minimum 1 -Maximum 6))
                        }
                        if($delay -le 0 -and $isTransientNetwork){
                            $delay = [math]::Min(10, 2 * $attempts)
                        }
                        if($delay -le 0){
                            $delay = [math]::Pow(5, $attempts)
                        }
                        $apiFamily = if($nextURL -match "sharepoint") { 'SharePoint' } else { 'Graph' }
                        $null = Write-Log "[$apiFamily] [WARNING] Transient error on attempt $attempts/$MaxAttempts, retrying in $($delay)s: $($_.Exception.Message)" -ForegroundColor Yellow
                        Start-Sleep -Seconds (1 + $delay)
                    }
                }

                if($nextURL -match "sharepoint"){
                    if($Data -and $Data.PSObject.TypeNames -notcontains "System.Management.Automation.PSCustomObject"){
                        try {
                            $Data = $Data | ConvertFrom-Json -AsHashtable
                        } catch {
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

                if($null -eq $Data) {
                    $nextURL = $null
                    continue
                }

                $pageItems = $null
                if($Data.psobject.properties.name -icontains 'value' -or ($Data.PSObject.BaseObject -is [hashtable] -and $Data.Keys -icontains 'value')){
                    $pageItems = $Data.value
                }else{
                    $pageItems = $Data
                }

                if ($null -ne $pageItems) {
                    $pageItemCount = @($pageItems).Count
                    $totalResults += $pageItemCount

                    if ($pageItemCount -eq 1 -and -not ($pageItems -is [array])) {
                        $ReturnedData += @($pageItems)
                    } elseif ($pageItemCount -gt 0) {
                        $ReturnedData += $pageItems
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

# Detect Azure Automation Account environment (Write-Host is not supported there)
$global:IsAzureAutomation = ($env:AUTOMATION_ASSET_ACCOUNTID -or $env:AZUREPS_HOST_ENVIRONMENT -like 'AzureAutomation*')

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$timestamp] [$Level] $Message"
    if($global:IsAzureAutomation -and $global:octo -and $global:octo.RunspaceLogBuffer) {
        $global:octo.RunspaceLogBuffer.Add(@{
            Timestamp = $timestamp
            Level     = $Level
            Message   = $Message
            Line      = $line
        })
        return
    }
    if($global:IsAzureAutomation) {
        # Write-Host is not supported in Azure Automation.
        Write-Output $line
    } else {
        $color = switch($Level) {
            "ERROR" { "Red" }
            "WARN"  { "Yellow" }
            "SUCCESS" { "Green" }
            default { "White" }
        }
        Write-Host $line -ForegroundColor $color
    }
}


#endregion

#region SharePoint Batch Functions

function Invoke-SharePointBatch {
    <#
    .SYNOPSIS
        Sends a SharePoint REST $batch request containing multiple sub-requests and returns parsed results.
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$SiteUrl,

        [Parameter(Mandatory=$true)]
        [array]$SubRequests,

        [Parameter(Mandatory=$false)]
        [int]$MaxAttempts = 3
    )

    $batchGuid = [guid]::NewGuid().ToString()
    $boundary = "batch_$batchGuid"
    $changesetGuid = [guid]::NewGuid().ToString()

    # Build multipart batch body — all are GET so no changeset needed
    # SP $batch requires absolute URLs in each sub-request
    $siteUri = [System.Uri]::new($SiteUrl)
    $hostHeader = $siteUri.Host
    $bodyLines = [System.Collections.Generic.List[string]]::new()
    $requestIndex = 0
    foreach($sub in $SubRequests) {
        $bodyLines.Add("--$boundary")
        $bodyLines.Add("Content-Type: application/http")
        $bodyLines.Add("Content-Transfer-Encoding: binary")
        $bodyLines.Add("")
        $bodyLines.Add("GET $SiteUrl/$($sub.RelativeUrl) HTTP/1.1")
        $bodyLines.Add("Host: $hostHeader")
        $bodyLines.Add("Accept: application/json;odata=nometadata")
        $bodyLines.Add("")
        $requestIndex++
    }
    $bodyLines.Add("--$boundary--")
    $batchBody = $bodyLines -join "`r`n"

    $batchUrl = "$SiteUrl/_api/`$batch"
    $batchContentType = "multipart/mixed; boundary=$boundary"

    $attempts = 0
    while($attempts -lt $MaxAttempts) {
        $attempts++
        try {
            $tokenResource = $global:octo.sharepointUrl
            $token = Get-AccessToken -resource $tokenResource
            $headers = @{
                "Authorization" = "Bearer $token"
                "Accept" = "application/json"
            }

            $response = Invoke-WebRequest -Uri $batchUrl -Method POST -Headers $headers -Body $batchBody `
                -ContentType $batchContentType -UseBasicParsing -ErrorAction Stop -UserAgent "ISV|LiebenConsultancy|M365AutoLink|3.0"

            # Parse multipart response
            # In PS7, Invoke-WebRequest returns Content as byte[] for multipart/mixed — convert to string
            $responseContent = $response.Content
            if($responseContent -is [byte[]]) {
                $responseContent = [System.Text.Encoding]::UTF8.GetString($responseContent)
            }
            # In PS7, Headers values are string arrays — coerce to scalar
            $responseContentType = [string]$response.Headers['Content-Type']

            # Extract boundary from response Content-Type
            $responseBoundary = $null
            if($responseContentType -match 'boundary=(.+)'){
                $responseBoundary = $Matches[1].Trim()
            }

            if(-not $responseBoundary) {
                throw "Could not parse batch response boundary"
            }

            # Split response by boundary
            $parts = $responseContent -split "--$([regex]::Escape($responseBoundary))"
            $results = [System.Collections.Generic.List[object]]::new()

            foreach($part in $parts) {
                if([string]::IsNullOrWhiteSpace($part) -or $part.Trim() -eq '--') { continue }

                # Find the JSON body in each part (after the empty line separating HTTP headers from body)
                # Use \r?\n to handle both \r\n and \n line endings (PS7 may normalize)
                $segments = $part -split '\r?\n\r?\n'
                # The structure is: MIME headers \r\n\r\n HTTP status line + headers \r\n\r\n JSON body
                $jsonBody = $null
                $statusCode = 200
                for($si = 0; $si -lt $segments.Count; $si++) {
                    if($segments[$si] -match 'HTTP/1\.\d\s+(\d+)') {
                        $statusCode = [int]$Matches[1]
                    }
                    # Try to find JSON
                    $trimmed = $segments[$si].Trim()
                    if($trimmed.StartsWith('{') -or $trimmed.StartsWith('[')) {
                        $jsonBody = $trimmed
                    }
                }

                # Fallback: if segment splitting didn't isolate JSON, extract payload from first JSON token onward
                if($null -eq $jsonBody) {
                    $firstBrace = $part.IndexOf('{')
                    $firstBracket = $part.IndexOf('[')
                    $startIdx = -1
                    if($firstBrace -ge 0 -and $firstBracket -ge 0) {
                        $startIdx = [math]::Min($firstBrace, $firstBracket)
                    } elseif($firstBrace -ge 0) {
                        $startIdx = $firstBrace
                    } elseif($firstBracket -ge 0) {
                        $startIdx = $firstBracket
                    }
                    if($startIdx -ge 0) {
                        $jsonBody = $part.Substring($startIdx).Trim()
                    }
                }

                if($null -ne $jsonBody) {
                    try {
                        $parsed = $jsonBody | ConvertFrom-Json -ErrorAction Stop
                        $results.Add([PSCustomObject]@{
                            StatusCode = $statusCode
                            Data       = $parsed
                            Error      = $null
                        })
                    } catch {
                        $results.Add([PSCustomObject]@{
                            StatusCode = $statusCode
                            Data       = $null
                            Error      = "JSON parse error: $($_.Exception.Message)"
                        })
                    }
                } else {
                    $results.Add([PSCustomObject]@{
                        StatusCode = $statusCode
                        Data       = $null
                        Error      = "No JSON body found in batch response part"
                    })
                }
            }

            return $results
        } catch {
            $is429 = $_.Exception.Response.StatusCode -eq 429 -or $_.Exception.Message -like "*429*"
            # 429s always retry indefinitely — do not count against MaxAttempts
            if ($is429) { $attempts-- }
            $isTransientNetwork = $_.Exception.Message -like "*No such host is known*" -or $_.Exception.Message -like "*name or service not known*" -or $_.Exception.Message -like "*connection was forcibly closed*"
            if($attempts -ge $MaxAttempts) { throw $_ }
            $delay = 0
            if($is429) {
                try {
                    $retryAfter = $_.Exception.Response.Headers.GetValues("Retry-After")
                    if ($retryAfter -and $retryAfter.Count -gt 0 -and $retryAfter[0] -match '^\d+$') { $delay = [int]$retryAfter[0] }
                } catch {}
                if($delay -le 0) { $delay = [math]::Min(120, [math]::Pow(5, [math]::Max(1, $attempts))) }
            } elseif($isTransientNetwork) {
                $delay = [math]::Min(10, 2 * $attempts)
            } else {
                $delay = [math]::Pow(5, $attempts)
            }
            $null = Write-Log "[SharePoint] [WARNING] Batch request to '$SiteUrl' failed on attempt $attempts/$MaxAttempts, retrying in $($delay)s: $($_.Exception.Message)" "WARN"
            Start-Sleep -Seconds (1 + $delay)
        }
    }
}

#endregion

#region Main Script

try {
    $logDir = [System.IO.Path]::GetDirectoryName($global:octo.LogPath)
    if(!(Test-Path $logDir)){ New-Item -ItemType Directory -Path $logDir -Force | Out-Null }

    if($global:IsAzureAutomation) {
        Write-Output "NOTE: Azure Automation detected. Write-Host is not supported in this environment, using Write-Output for all log messages. All logs appear in the Output tab."
    }else{
        Start-Transcript -Path $global:octo.LogPath -Force
    }

    Write-Log "=== M365AutoLink Centralized v3.0 (Optimized) Started ===" "INFO"
    Write-Log "Optimizations: SP Batch ($BatchSize/batch) | Adaptive Throttle ($InitialParallelLimit-$MaxParallelLimit) | Shortcut Actions ($ShortcutActionParallelLimit threads)" "INFO"
    if($DryRun) { Write-Log "*** DRY RUN MODE — no changes will be made ***" "WARN" }

    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Web")

    $token = Get-AccessToken -resource $global:octo.graphUrl
    Write-Log "Authentication successful" "SUCCESS"

    $rootSite = New-GraphQuery -Uri "$($global:octo.graphUrl)/v1.0/sites/root" -Method GET
    $spHost = ([System.Uri]::new($rootSite.webUrl)).Host
    Write-Log "Tenant SharePoint URL: https://$spHost" "INFO"

    Write-Log "Exclusion patterns:" "INFO"
    foreach($pattern in $excludedSitesByWildcard){ Write-Log "  - $pattern" "INFO" }
    Write-Log "Inclusion patterns:" "INFO"
    foreach($pattern in $includedSitesByWildcard){ Write-Log "  - $pattern" "INFO" }

    $phaseTimings = [ordered]@{}
    $scriptStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    # ============================================================
    # PHASE 0: Get target users
    # ============================================================
    $phaseStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Log "--- Phase 0: User enumeration ---" "INFO"

    $targetUserList = @()
    switch($TargetMode) {
        "Group" {
            if(-not $TargetGroupId) { throw "TargetGroupId must be specified when TargetMode is 'Group'" }
            Write-Log "Getting members of group '$TargetGroupId' (transitive)..." "INFO"
            $members = New-GraphQuery -Uri "$($global:octo.graphUrl)/v1.0/groups/$TargetGroupId/transitiveMembers?`$select=id,userPrincipalName,displayName" -Method GET
            $targetUserList = @($members | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.user' })
        }
        "UserList" {
            if($TargetUsers.Count -eq 0) { throw "TargetUsers must contain at least one UPN when TargetMode is 'UserList'" }
            Write-Log "Resolving $($TargetUsers.Count) specified users..." "INFO"
            foreach($upn in $TargetUsers) {
                try {
                    $user = New-GraphQuery -Uri "$($global:octo.graphUrl)/v1.0/users/$([System.Web.HttpUtility]::UrlEncode($upn))?`$select=id,userPrincipalName,displayName" -Method GET
                    $targetUserList += $user
                } catch {
                    Write-Log "Could not find user '$upn': $($_.Exception.Message)" "ERROR"
                }
            }
        }
        "All" {
            Write-Log "Getting all enabled member users..." "INFO"
            $targetUserList = @(New-GraphQuery -Uri "$($global:octo.graphUrl)/v1.0/users?`$filter=accountEnabled eq true and userType eq 'Member'&`$select=id,userPrincipalName,displayName&`$top=999" -Method GET)
        }
        default { throw "Invalid TargetMode '$TargetMode'. Must be 'Group', 'UserList', or 'All'" }
    }

    Write-Log "Found $($targetUserList.Count) target users" "SUCCESS"
    if($targetUserList.Count -eq 0) { throw "No target users found for TargetMode '$TargetMode'" }

    $phaseTimings['Phase 0: User enumeration'] = $phaseStopwatch.Elapsed

    # ============================================================
    # PHASE 1: Pre-fetch all sites and enumerate document libraries
    # ============================================================
    $phaseStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Log "--- Phase 1: Site enumeration & library discovery ---" "INFO"

    Write-Log "Fetching all tenant sites..." "INFO"
    $allTenantSites = @(New-GraphQuery -Uri "$($global:octo.graphUrl)/v1.0/sites/getAllSites?`$select=id,displayName,webUrl&`$top=1000" -Method GET)
    Write-Log "Found $($allTenantSites.Count) sites in tenant" "SUCCESS"

    # Pre-filter sites locally
    $filteredSites = [System.Collections.Generic.List[object]]::new()
    foreach($site in $allTenantSites) {
        if($null -eq $site.webUrl) { continue }

        $isExcluded = $false
        foreach($pattern in $excludedSitesByWildcard){
            $wildcardPattern = "^" + [regex]::Escape($pattern) -replace "\\\*",".*"
            if($site.webUrl -match $wildcardPattern){
                $isExcluded = $true
                break
            }
        }
        if($isExcluded) { continue }

        $isIncluded = $false
        foreach($pattern in $includedSitesByWildcard){
            $wildcardPattern = "^" + [regex]::Escape($pattern) -replace "\\\*",".*"
            if($site.webUrl -match $wildcardPattern){
                $isIncluded = $true
                break
            }
        }
        if(-not $isIncluded) { continue }

        $filteredSites.Add($site)
    }
    Write-Log "Pre-filtered to $($filteredSites.Count) sites after inclusion/exclusion patterns" "INFO"

    if($MaxSitesToProcessForTesting -gt 0 -and $filteredSites.Count -gt $MaxSitesToProcessForTesting) {
        $limitedSites = [System.Collections.Generic.List[object]]::new()
        for($i = 0; $i -lt $MaxSitesToProcessForTesting; $i++) {
            $limitedSites.Add($filteredSites[$i])
        }
        $filteredSites = $limitedSites
        Write-Log "Test limiter active: restricting run to first $MaxSitesToProcessForTesting filtered sites" "WARN"
    }

    # Enumerate document libraries in parallel using adaptive throttle
    Write-Log "Enumerating document libraries (adaptive throttle: $InitialParallelLimit initial, $MaxParallelLimit max)..." "INFO"

    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $iss.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new('Get-AccessToken', (Get-Command Get-AccessToken).Definition))
    $iss.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new('New-GraphQuery', (Get-Command New-GraphQuery).Definition))
    $iss.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new('Write-Log', (Get-Command Write-Log).Definition))
    $iss.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new('ClientId', $ClientId, ''))
    $iss.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new('TenantId', $TenantId, ''))
    $iss.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new('CertificateThumbprint', $CertificateThumbprint, ''))
    $iss.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new('CertificatePath', $CertificatePath, ''))
    $iss.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new('CertificatePassword', $CertificatePassword, ''))

    $pool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $MaxParallelLimit, $iss, $Host)
    $pool.Open()

    try {

    $graphUrlValue = $global:octo.graphUrl
    $idpUrlValue = $global:octo.idpUrl
    $sharepointUrlValue = $global:octo.sharepointUrl

    # Pre-fetch tokens so runspaces don't each make their own token requests
    $graphToken = Get-AccessToken -resource $global:octo.graphUrl
    $spToken = Get-AccessToken -resource $global:octo.sharepointUrl
    $preSeededTokens = @{
        $global:octo.graphUrl = @{ accessToken = $graphToken; validFrom = Get-Date }
        $global:octo.sharepointUrl = @{ accessToken = $spToken; validFrom = Get-Date }
    }

    $processSiteBlock = {
        param([hashtable]$siteInfo, [string]$graphUrl, [string]$idpUrl, [string]$sharepointUrl, [int]$maxFC, [int]$minFC, [bool]$onlyConnected, [hashtable]$seedTokens)

        $global:octo = @{
            graphUrl       = $graphUrl
            idpUrl         = $idpUrl
            sharepointUrl  = $sharepointUrl
            LCCachedTokens = $seedTokens
            RunspaceLogBuffer = [System.Collections.Generic.List[hashtable]]::new()
        }
        [void][System.Reflection.Assembly]::LoadWithPartialName("System.Web")

        $result = [PSCustomObject]@{
            Libraries   = [System.Collections.Generic.List[hashtable]]::new()
            GroupId     = $null
            LogMessages = [System.Collections.Generic.List[hashtable]]::new()
            Succeeded   = $false
        }

        try {
            $siteDetails = New-GraphQuery -Uri "$($siteInfo.webUrl)/_api/site" -Method "GET" -MaxAttempts 1
            if($siteDetails.WriteLocked -or $siteDetails.ReadOnly){
                $result.LogMessages.Add(@{ Message = "Site '$($siteInfo.webUrl)' is locked/read-only, skipping..."; Level = "WARN" })
                return $result
            }

            if($onlyConnected -and (!$siteDetails.GroupId -or $siteDetails.GroupId -eq "00000000-0000-0000-0000-000000000000")){
                $result.LogMessages.Add(@{ Message = "Site '$($siteInfo.webUrl)' is not group-connected, skipping..."; Level = "INFO" })
                return $result
            }

            # Capture the associated M365 group ID (for group pre-filter optimization)
            if($siteDetails.GroupId -and $siteDetails.GroupId -ne "00000000-0000-0000-0000-000000000000") {
                $result.GroupId = $siteDetails.GroupId
            }

            $lists = @((New-GraphQuery -Uri "$graphUrl/v1.0/sites/$($siteInfo.id)/lists" -Method "GET" -MaxAttempts 3) | Where-Object { $_.list.template -eq "documentLibrary" -and !$_.list.hidden })

            foreach($list in $lists) {
                $listMetaData = New-GraphQuery -Uri "$($siteInfo.webUrl)/_api/lists/GetById('$($list.id)')" -Method GET
                if($listMetaData.Hidden) { continue }

                if($listMetaData.ItemCount -gt $maxFC){
                    $result.LogMessages.Add(@{ Message = "$($siteInfo.webUrl) - '$($list.displayName)' exceeds $maxFC files, skipping..."; Level = "WARN" })
                    continue
                }
                if($listMetaData.ItemCount -lt $minFC){
                    $result.LogMessages.Add(@{ Message = "$($siteInfo.webUrl) - '$($list.displayName)' below $minFC files, skipping..."; Level = "WARN" })
                    continue
                }

                $libKey = "$($siteInfo.webUrl)|$($list.id)"

                $result.Libraries.Add(@{
                    libKey          = $libKey
                    siteId          = $siteInfo.id
                    siteDisplayName = $siteInfo.displayName
                    siteWebUrl      = $siteInfo.webUrl
                    siteGroupId     = $result.GroupId
                    listId          = $list.id
                    listDisplayName = $list.displayName
                    shortcutInfo    = @{
                        siteId           = $siteInfo.id.Split(',')[1]
                        siteUrl          = $siteInfo.webUrl
                        webId            = $siteInfo.id.Split(',')[2]
                        listId           = $list.id
                        listItemUniqueId = "root"
                    }
                })
            }
            $result.Succeeded = $true
        } catch {
            $result.LogMessages.Add(@{ Message = "Failed to process site '$($siteInfo.webUrl)': $($_.Exception.Message)"; Level = "WARN" })
        }

        if($global:octo.RunspaceLogBuffer -and $global:octo.RunspaceLogBuffer.Count -gt 0) {
            foreach($entry in $global:octo.RunspaceLogBuffer) {
                $result.LogMessages.Add($entry)
            }
        }

        return $result
    }

    # Adaptive throttle: use a semaphore to control concurrency within the max pool
    $activeSemaphore = [System.Threading.SemaphoreSlim]::new($InitialParallelLimit, $MaxParallelLimit)
    $currentLimit = $InitialParallelLimit
    $throttleHitCount = 0
    $graphThrottleHitCount = 0
    $sharePointThrottleHitCount = 0
    $completedSinceAdjust = 0
    $phase1Total429s = 0
    $phase1TotalApiCalls = 0
    $lastMetricsLog = 0

    # Result collections — declared before submission loop for interleaved processing
    $allSiteLibraries = [System.Collections.Generic.List[hashtable]]::new()
    $filteredSiteCount = 0
    $completedCount = 0
    $totalJobs = $filteredSites.Count

    $jobs = [System.Collections.Generic.List[hashtable]]::new()
    foreach($site in $filteredSites) {
        # Wait for a free slot, processing completed jobs while waiting (prevents deadlock)
        while(-not $activeSemaphore.Wait(200)) {
            for($i = $jobs.Count - 1; $i -ge 0; $i--) {
                if($jobs[$i].Handle.IsCompleted) {
                    $completedCount++
                    $completedSinceAdjust++
                    if($activeSemaphore.CurrentCount -lt $MaxParallelLimit) { [void]$activeSemaphore.Release() }
                    Write-Progress -Id 0 -Activity "Phase 1: Enumerating sites" -Status "$completedCount/$totalJobs sites | concurrent: $currentLimit | 429s: $phase1Total429s | libs: $($allSiteLibraries.Count)" -PercentComplete ([math]::Min(100, [math]::Round(($completedCount / $totalJobs) * 100)))
                    try {
                        $jobResult = $jobs[$i].PowerShell.EndInvoke($jobs[$i].Handle)
                        if($jobResult) {
                            foreach($r in $jobResult) {
                                if($r.LogMessages) {
                                    foreach($log in $r.LogMessages) { Write-Log "  $($log.Message)" $log.Level }
                                    foreach($log in $r.LogMessages) {
                                        if($log.Message -match '\[Graph\].*429') { $throttleHitCount++; $graphThrottleHitCount++; $phase1Total429s++ }
                                        elseif($log.Message -match '\[SharePoint\].*429') { $throttleHitCount++; $sharePointThrottleHitCount++; $phase1Total429s++ }
                                    }
                                }
                                if($r.Libraries -and $r.Libraries.Count -gt 0) {
                                    $allSiteLibraries.AddRange([hashtable[]]$r.Libraries)
                                }
                                if($r.Succeeded) { $filteredSiteCount++ }
                            }
                        }
                        if($jobs[$i].PowerShell.Streams.Error.Count -gt 0) {
                            foreach($err in $jobs[$i].PowerShell.Streams.Error) {
                                Write-Log "  Runspace error for '$($jobs[$i].SiteUrl)': $($err.Exception.Message)" "WARN"
                                if($err.Exception.Message -like "*429*") { $throttleHitCount++; $phase1Total429s++ }
                            }
                        }
                        # Count 429 retries from buffered log messages returned by the runspace
                    } catch {
                        Write-Log "  Error collecting result for '$($jobs[$i].SiteUrl)': $($_.Exception.Message)" "WARN"
                    }
                    $jobs[$i].PowerShell.Dispose()
                    $jobs.RemoveAt($i)
                }
            }
        }
        $ps = [powershell]::Create()
        $ps.RunspacePool = $pool
        [void]$ps.AddScript($processSiteBlock)
        [void]$ps.AddParameter('siteInfo', @{ id = $site.id; displayName = $site.displayName; webUrl = $site.webUrl })
        [void]$ps.AddParameter('graphUrl', $graphUrlValue)
        [void]$ps.AddParameter('idpUrl', $idpUrlValue)
        [void]$ps.AddParameter('sharepointUrl', $sharepointUrlValue)
        [void]$ps.AddParameter('maxFC', $maxFileCount)
        [void]$ps.AddParameter('minFC', $minFileCount)
        [void]$ps.AddParameter('onlyConnected', $onlyConnectedSites)
        [void]$ps.AddParameter('seedTokens', $preSeededTokens)
        $jobs.Add(@{
            PowerShell = $ps
            Handle     = $ps.BeginInvoke()
            SiteUrl    = $site.webUrl
        })
    }

    # Collect remaining site enumeration results
    $totalJobs = $completedCount + $jobs.Count # Recalculate to account for already-processed jobs

    while($jobs.Count -gt 0) {
        for($i = $jobs.Count - 1; $i -ge 0; $i--) {
            if($jobs[$i].Handle.IsCompleted) {
                $completedCount++
                $completedSinceAdjust++
                if($activeSemaphore.CurrentCount -lt $MaxParallelLimit) { [void]$activeSemaphore.Release() }
                Write-Progress -Id 0 -Activity "Phase 1: Enumerating sites" -Status "$completedCount/$totalJobs sites | concurrent: $currentLimit | 429s: $phase1Total429s | libs: $($allSiteLibraries.Count)" -PercentComplete ([math]::Min(100, [math]::Round(($completedCount / $totalJobs) * 100)))
                try {
                    $jobResult = $jobs[$i].PowerShell.EndInvoke($jobs[$i].Handle)
                    if($jobResult) {
                        foreach($r in $jobResult) {
                            if($r.LogMessages) {
                                foreach($log in $r.LogMessages) { Write-Log "  $($log.Message)" $log.Level }
                                foreach($log in $r.LogMessages) {
                                    if($log.Message -match '\[Graph\].*429') { $throttleHitCount++; $graphThrottleHitCount++; $phase1Total429s++ }
                                    elseif($log.Message -match '\[SharePoint\].*429') { $throttleHitCount++; $sharePointThrottleHitCount++; $phase1Total429s++ }
                                }
                            }
                            if($r.Libraries -and $r.Libraries.Count -gt 0) {
                                $allSiteLibraries.AddRange([hashtable[]]$r.Libraries)
                            }
                            if($r.Succeeded) { $filteredSiteCount++ }
                        }
                    }
                    # Check for 429s in error stream to feed adaptive throttle
                    if($jobs[$i].PowerShell.Streams.Error.Count -gt 0) {
                        foreach($err in $jobs[$i].PowerShell.Streams.Error) {
                            Write-Log "  Runspace error for '$($jobs[$i].SiteUrl)': $($err.Exception.Message)" "WARN"
                            if($err.Exception.Message -like "*429*") { $throttleHitCount++; $phase1Total429s++ }
                        }
                    }
                    # Buffered logs already accounted for above.
                } catch {
                    Write-Log "  Error collecting result for '$($jobs[$i].SiteUrl)': $($_.Exception.Message)" "WARN"
                }
                $jobs[$i].PowerShell.Dispose()
                $jobs.RemoveAt($i)

                # Adaptive throttle adjustment every 50 completed jobs
                if($completedSinceAdjust -ge 50) {
                    $graphThrottleRate = $graphThrottleHitCount / $completedSinceAdjust
                    $sharePointThrottleRate = $sharePointThrottleHitCount / $completedSinceAdjust
                    $throttleRate = ($graphThrottleRate + $sharePointThrottleRate)
                    if(($graphThrottleRate -gt 0.1 -or $sharePointThrottleRate -gt 0.1) -and $currentLimit -gt 2) {
                        $reductionFactor = if($graphThrottleRate -gt 0.1 -and $sharePointThrottleRate -gt 0.1) { 0.5 } else { 0.75 }
                        $newLimit = [math]::Max(2, [math]::Floor($currentLimit * $reductionFactor))
                        $reduction = $currentLimit - $newLimit
                        for($r = 0; $r -lt $reduction; $r++) {
                            [void]$activeSemaphore.Wait(0)
                        }
                        $currentLimit = $newLimit
                        Write-Log "  Adaptive throttle: reducing to $currentLimit concurrent (Graph 429: $([math]::Round($graphThrottleRate * 100))%, SharePoint 429: $([math]::Round($sharePointThrottleRate * 100))%)" "WARN"
                    } elseif($throttleRate -eq 0 -and $currentLimit -lt $MaxParallelLimit) {
                        $increase = [math]::Min(2, $MaxParallelLimit - $currentLimit)
                        $currentLimit += $increase
                        if($increase -gt 0 -and $activeSemaphore.CurrentCount -lt $MaxParallelLimit) {
                            $safeRelease = [math]::Min($increase, $MaxParallelLimit - $activeSemaphore.CurrentCount)
                            if($safeRelease -gt 0) { [void]$activeSemaphore.Release($safeRelease) }
                        }
                        Write-Log "  Adaptive throttle: increasing to $currentLimit concurrent (no 429s)" "INFO"
                    }
                    $throttleHitCount = 0
                    $graphThrottleHitCount = 0
                    $sharePointThrottleHitCount = 0
                    $completedSinceAdjust = 0
                }
            }
        }
        if($jobs.Count -gt 0) { Start-Sleep -Milliseconds 100 }
    }

    Write-Progress -Id 0 -Activity "Phase 1: Enumerating sites" -Completed

    } finally {
        if($pool) { $pool.Close(); $pool.Dispose() }
        if($activeSemaphore) { $activeSemaphore.Dispose() }
    }

    # Deduplicate libraries (Graph pagination or runspace collection can produce duplicates)
    $beforeDedup = $allSiteLibraries.Count
    $seenLibKeys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $dedupedLibs = [System.Collections.Generic.List[hashtable]]::new()
    foreach($lib in $allSiteLibraries) {
        $libKey = "$($lib.siteWebUrl)|$($lib.listId)"
        if($seenLibKeys.Add($libKey)) { $dedupedLibs.Add($lib) }
    }
    $allSiteLibraries = $dedupedLibs
    if($allSiteLibraries.Count -lt $beforeDedup) {
        Write-Log "Deduplication removed $($beforeDedup - $allSiteLibraries.Count) duplicate library entries" "WARN"
    }

    Write-Log "Pre-cached $($allSiteLibraries.Count) document libraries across $filteredSiteCount sites" "SUCCESS"
    Write-Log "Phase 1 throttling: Graph 429s=$graphThrottleHitCount | SharePoint 429s=$sharePointThrottleHitCount" "INFO"
    Write-Log "Minimum permission level for shortcuts: $MinimumPermissionLevel" "INFO"

    $phaseTimings['Phase 1: Site & library discovery'] = $phaseStopwatch.Elapsed

    # ============================================================
    # PHASE 2: Site-centric permission matrix via SP $batch (P1+P2)
    # ============================================================
    $phaseStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Log "--- Phase 2: Building permission matrix (site-centric, batched) ---" "INFO"

    # Group libraries by site for batching
    $siteLibGroups = @{}
    foreach($lib in $allSiteLibraries) {
        if(-not $siteLibGroups.ContainsKey($lib.siteWebUrl)) {
            $siteLibGroups[$lib.siteWebUrl] = [System.Collections.Generic.List[hashtable]]::new()
        }
        $siteLibGroups[$lib.siteWebUrl].Add($lib)
    }

    Write-Log "Grouped $($allSiteLibraries.Count) libraries into $($siteLibGroups.Count) sites for batched permission checks" "INFO"

    # Build permission matrix: userUPN → list of libraries they have access to
    $permissionMatrix = @{} # "userUPN" → List of library hashtables (shortcut info)
    foreach($u in $targetUserList) {
        $permissionMatrix[$u.userPrincipalName] = [System.Collections.Generic.List[hashtable]]::new()
    }

    $targetUserByUpn = @{}
    $targetUserById = @{}
    foreach($u in $targetUserList) {
        if(-not [string]::IsNullOrWhiteSpace($u.userPrincipalName)) {
            $targetUserByUpn[$u.userPrincipalName.ToLowerInvariant()] = $u.userPrincipalName
        }
        if(-not [string]::IsNullOrWhiteSpace($u.id)) {
            $targetUserById[$u.id.ToLowerInvariant()] = $u.userPrincipalName
        }
    }

    $preCheckStats = @{
        SiteScopeHits    = 0
        ListScopeHits    = 0
        PreAllowedPairs  = 0
        FallbackPairs    = 0
        CacheHits        = 0
        CacheMisses      = 0
        RecursionCutoffs = 0
    }

    if($EnableRecursivePermissionPreCheck) {
        Write-Log "Recursive pre-check enabled: site/list owner-member expansion + Entra group caching" "INFO"
    } else {
        Write-Log "Recursive pre-check disabled: using full getUserEffectivePermissions matrix path" "INFO"
    }

    # Adaptive throttle for batch requests (reset for Phase 2)
    $batchThrottleHits = 0
    $batchCompletedSinceAdjust = 0
    $batchCurrentLimit = $InitialParallelLimit

    # Create runspace pool for batch permission checks
    $batchIss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $batchIss.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new('Get-AccessToken', (Get-Command Get-AccessToken).Definition))
    $batchIss.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new('New-GraphQuery', (Get-Command New-GraphQuery).Definition))
    $batchIss.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new('Write-Log', (Get-Command Write-Log).Definition))
    $batchIss.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new('Invoke-SharePointBatch', (Get-Command Invoke-SharePointBatch).Definition))
    $batchIss.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new('ClientId', $ClientId, ''))
    $batchIss.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new('TenantId', $TenantId, ''))
    $batchIss.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new('CertificateThumbprint', $CertificateThumbprint, ''))
    $batchIss.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new('CertificatePath', $CertificatePath, ''))
    $batchIss.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new('CertificatePassword', $CertificatePassword, ''))
    $batchPool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $MaxParallelLimit, $batchIss, $Host)
    $batchPool.Open()
    $batchSemaphore = [System.Threading.SemaphoreSlim]::new($InitialParallelLimit, $MaxParallelLimit)

    try {

    # Pre-fetch tokens for Phase 2 runspaces
    $graphToken2 = Get-AccessToken -resource $global:octo.graphUrl
    $spToken2 = Get-AccessToken -resource $global:octo.sharepointUrl
    $preSeededTokens2 = @{
        $global:octo.graphUrl = @{ accessToken = $graphToken2; validFrom = Get-Date }
        $global:octo.sharepointUrl = @{ accessToken = $spToken2; validFrom = Get-Date }
    }

    # Script block: for one site, check permissions for a chunk of (user, library) pairs using SP $batch
    $batchPermCheckBlock = {
        param(
            [string]$siteWebUrl,
            [array]$checkPairs,     # Array of @{ userUPN = "..."; listId = "..."; libKey = "..." }
            [int]$batchSize,
            [string]$graphUrl,
            [string]$idpUrl,
            [string]$sharepointUrl,
            [string]$minPermLevel,
            [hashtable]$seedTokens
        )

        $global:octo = @{
            graphUrl       = $graphUrl
            idpUrl         = $idpUrl
            sharepointUrl  = $sharepointUrl
            LCCachedTokens = $seedTokens
            RunspaceLogBuffer = [System.Collections.Generic.List[hashtable]]::new()
        }
        [void][System.Reflection.Assembly]::LoadWithPartialName("System.Web")

        $results = [System.Collections.Generic.List[hashtable]]::new()
        $logMessages = [System.Collections.Generic.List[hashtable]]::new()
        $throttleHits = 0

        # Process in batches of $batchSize
        for($bi = 0; $bi -lt $checkPairs.Count; $bi += $batchSize) {
            $chunk = $checkPairs[$bi..([math]::Min($bi + $batchSize - 1, $checkPairs.Count - 1))]

            $subRequests = [System.Collections.Generic.List[object]]::new()
            foreach($pair in $chunk) {
                $encodedUPN = "i%3A0%23.f%7Cmembership%7C$($pair.userUPN)"
                $subRequests.Add(@{
                    RelativeUrl = "_api/web/lists/GetById('$($pair.listId)')/getUserEffectivePermissions(@u)?@u='$encodedUPN'"
                    UserUPN     = $pair.userUPN
                    LibKey      = $pair.libKey
                    ListId      = $pair.listId
                })
            }

            try {
                $batchResults = Invoke-SharePointBatch -SiteUrl $siteWebUrl -SubRequests $subRequests.ToArray() -MaxAttempts 3

                # Match results back to sub-requests by index
                $resultIndex = 0
                foreach($sub in $subRequests) {
                    if($resultIndex -lt $batchResults.Count) {
                        $batchResult = $batchResults[$resultIndex]
                        $hasAccess = $false

                        if($batchResult.StatusCode -eq 200 -and $batchResult.Data) {
                            # Unwrap SP response — format varies by odata mode:
                            #   flat:    {"High":"48","Low":"1011030767"}
                            #   wrapped: {"GetUserEffectivePermissions":{"High":"48","Low":"1011030767"}}
                            #   verbose: {"d":{"GetUserEffectivePermissions":{"High":"48","Low":"1011030767"}}}
                            $permData = $batchResult.Data
                            if($null -ne $permData.d) { $permData = $permData.d }
                            if($null -ne $permData.GetUserEffectivePermissions) { $permData = $permData.GetUserEffectivePermissions }

                            $lowVal = $permData.Low
                            $highVal = $permData.High
                            if($null -ne $lowVal) {
                                $hasView = ([long]$lowVal -band 0x1) -ne 0
                                $hasEdit = ([long]$lowVal -band 0x4) -ne 0

                                if($minPermLevel -eq "Edit") {
                                    $hasAccess = $hasEdit
                                } else {
                                    $hasAccess = $hasView
                                }
                            }
                        } elseif($batchResult.StatusCode -in (401, 403, 404)) {
                            $hasAccess = $false
                        } elseif($batchResult.StatusCode -eq 429) {
                            $throttleHits++
                            $logMessages.Add(@{ Message = "429 throttled for user '$($sub.UserUPN)' on list '$($sub.ListId)' at '$siteWebUrl' - retrying individually"; Level = "WARN" })

                            # Retry this specific check immediately (bounded by New-GraphQuery retry behavior)
                            try {
                                $encodedUPN = "i%3A0%23.f%7Cmembership%7C$($sub.UserUPN)"
                                $effectivePermissions = New-GraphQuery -Uri "$siteWebUrl/_api/web/lists/GetById('$($sub.ListId)')/getUserEffectivePermissions(@u)?@u='$encodedUPN'" -Method GET -MaxAttempts 3
                                $permData = $effectivePermissions
                                if($null -ne $permData.d) { $permData = $permData.d }
                                if($null -ne $permData.GetUserEffectivePermissions) { $permData = $permData.GetUserEffectivePermissions }
                                if($null -ne $permData.Low) {
                                    $hasView = ([long]$permData.Low -band 0x1) -ne 0
                                    $hasEdit = ([long]$permData.Low -band 0x4) -ne 0
                                    $hasAccess = if($minPermLevel -eq "Edit") { $hasEdit } else { $hasView }
                                } else {
                                    $hasAccess = $false
                                }
                            } catch {
                                $logMessages.Add(@{ Message = "Individual retry after 429 failed for '$($sub.UserUPN)' on '$($sub.ListId)': $($_.Exception.Message)"; Level = "WARN" })
                                $hasAccess = $false
                            }
                        }

                        $results.Add(@{
                            UserUPN   = $sub.UserUPN
                            LibKey    = $sub.LibKey
                            HasAccess = $hasAccess
                        })
                    }
                    $resultIndex++
                }
            } catch {
                $logMessages.Add(@{ Message = "Batch failed for '$siteWebUrl': $($_.Exception.Message)"; Level = "WARN" })
                if($_.Exception.Message -like "*429*") { $throttleHits++ }

                # Fallback: try individual requests for this chunk
                foreach($pair in $chunk) {
                    try {
                        $encodedUPN = "i%3A0%23.f%7Cmembership%7C$($pair.userUPN)"
                        $effectivePermissions = New-GraphQuery -Uri "$siteWebUrl/_api/web/lists/GetById('$($pair.listId)')/getUserEffectivePermissions(@u)?@u='$encodedUPN'" -Method GET -MaxAttempts 1
                        # Unwrap SP response (same as batch path)
                        $permData = $effectivePermissions
                        if($null -ne $permData.d) { $permData = $permData.d }
                        if($null -ne $permData.GetUserEffectivePermissions) { $permData = $permData.GetUserEffectivePermissions }
                        $hasAccess = $false
                        if($null -ne $permData.Low) {
                            $hasView = ([long]$permData.Low -band 0x1) -ne 0
                            $hasEdit = ([long]$permData.Low -band 0x4) -ne 0
                            $hasAccess = if($minPermLevel -eq "Edit") { $hasEdit } else { $hasView }
                        }

                        $results.Add(@{
                            UserUPN   = $pair.userUPN
                            LibKey    = $pair.libKey
                            HasAccess = $hasAccess
                        })
                    } catch {
                        if($_.Exception.Message -notlike "*401*" -and $_.Exception.Message -notlike "*403*" -and $_.Exception.Message -notlike "*404*") {
                            $logMessages.Add(@{ Message = "Individual fallback failed for '$($pair.userUPN)' on '$($pair.listId)': $($_.Exception.Message)"; Level = "WARN" })
                        }
                        $results.Add(@{
                            UserUPN   = $pair.userUPN
                            LibKey    = $pair.libKey
                            HasAccess = $false
                        })
                    }
                }
            }

            # Micro-sleep between batches to avoid burst throttling
            if($bi + $batchSize -lt $checkPairs.Count) {
                Start-Sleep -Milliseconds 100
            }
        }

        return [PSCustomObject]@{
            Results      = $results
            LogMessages  = $logMessages
            ThrottleHits = $throttleHits
        }
    }

    # Now process site by site: for each site, determine which (user, library) pairs need checking
    $siteIndex = 0
    $siteTotal = $siteLibGroups.Count
    $totalBatchChecks = 0
    $phase2Total429s = 0
    $phase2CompletedJobs = 0

    $batchJobs = [System.Collections.Generic.List[hashtable]]::new()

    foreach($siteWebUrl in $siteLibGroups.Keys) {
        $siteIndex++
        $siteLibs = $siteLibGroups[$siteWebUrl]
        $sitePreAllowedPairCount = 0
        $siteFallbackPairCount = 0

        Write-Progress -Id 0 -Activity "Phase 2: Checking permissions" -Status "$siteIndex/$siteTotal sites | concurrent: $batchCurrentLimit | batch: $BatchSize | 429s: $phase2Total429s | checks: $totalBatchChecks" -PercentComplete ([math]::Min(100, [math]::Round(($siteIndex / $siteTotal) * 100)))

        # Build list of (user, library) pairs that need an actual API check.
        # Optional pre-check can short-circuit known access from recursive owner/member expansion.
        $checkPairs = [System.Collections.Generic.List[hashtable]]::new()
        $sitePreAllowedUsers = Ensure-StringSet -Set (New-CaseInsensitiveStringSet)
        $siteAllowAllUsers = $false

        if($EnableRecursivePermissionPreCheck) {
            $siteScope = Get-ScopeAllowedTargetUsers -SiteWebUrl $siteWebUrl -ScopeType Web -ListId $null -MinimumPermissionLevel $MinimumPermissionLevel -TargetUserByUpn $targetUserByUpn -TargetUserById $targetUserById -EntraGroupCache $global:octo.EntraGroupTargetMemberCache -PreCheckStats $preCheckStats -IncludeEveryoneClaims $PreCheckIncludeEveryoneClaims -MaxDepth $PreCheckMaxRecursionDepth -MaxExpandedPrincipals $PreCheckMaxExpandedPrincipals
            if($siteScope) {
                $sitePreAllowedUsers = Ensure-StringSet -Set $sitePreAllowedUsers
                if($siteScope.AllowAllUsers) { $siteAllowAllUsers = $true }
                foreach($upn in @($siteScope.AllowedUsers)) { [void]$sitePreAllowedUsers.Add($upn) }
                $siteScopeResolvedCount = @($siteScope.AllowedUsers).Count
                if($siteScope.AllowAllUsers -or $siteScopeResolvedCount -gt 0) { $preCheckStats.SiteScopeHits++ }
                if($PreCheckVerboseDiagnostics) {
                    Write-Log "Pre-check source '$siteWebUrl' [site-roles]: allowAll=$($siteScope.AllowAllUsers), resolvedUsers=$siteScopeResolvedCount" "INFO"
                }
            }

            $siteGroupId = $null
            foreach($sLib in $siteLibs) {
                if(-not [string]::IsNullOrWhiteSpace($sLib.siteGroupId)) {
                    $siteGroupId = $sLib.siteGroupId
                    break
                }
            }

            if(-not [string]::IsNullOrWhiteSpace($siteGroupId)) {
                $sitePreAllowedUsers = Ensure-StringSet -Set $sitePreAllowedUsers
                $groupMembers = Resolve-EntraGroupTargetUsers -GroupId $siteGroupId -TargetUserByUpn $targetUserByUpn -TargetUserById $targetUserById -EntraGroupCache $global:octo.EntraGroupTargetMemberCache -PreCheckStats $preCheckStats
                foreach($upn in $groupMembers) { [void]$sitePreAllowedUsers.Add($upn) }
                if($groupMembers.Count -gt 0) { $preCheckStats.SiteScopeHits++ }
                if($PreCheckVerboseDiagnostics) {
                    Write-Log "Pre-check source '$siteWebUrl' [site-group-members]: groupId=$siteGroupId, resolvedUsers=$($groupMembers.Count)" "INFO"
                }

                try {
                    $groupOwners = @(New-GraphQuery -Uri "$($global:octo.graphUrl)/v1.0/groups/$siteGroupId/owners?`$select=id,userPrincipalName" -Method GET -MaxAttempts 2)
                    $ownerResolved = 0
                    foreach($owner in $groupOwners) {
                        $ownerUpn = $null
                        try { $ownerUpn = [string]$owner.userPrincipalName } catch {}
                        $ownerId = $null
                        try { $ownerId = [string]$owner.id } catch {}
                        if(Add-ResolvedTargetUserToSet -CandidateUpn $ownerUpn -CandidateId $ownerId -TargetUserByUpn $targetUserByUpn -TargetUserById $targetUserById -OutputSet $sitePreAllowedUsers) { $ownerResolved++ }

                        $ownerType = ""
                        try { $ownerType = [string]$owner.'@odata.type' } catch {}
                        if($ownerType -eq '#microsoft.graph.group' -and -not [string]::IsNullOrWhiteSpace($ownerId)) {
                            $nestedOwnerGroupMembers = Resolve-EntraGroupTargetUsers -GroupId $ownerId -TargetUserByUpn $targetUserByUpn -TargetUserById $targetUserById -EntraGroupCache $global:octo.EntraGroupTargetMemberCache -PreCheckStats $preCheckStats
                            foreach($upn in $nestedOwnerGroupMembers) { [void]$sitePreAllowedUsers.Add($upn) }
                            $ownerResolved += $nestedOwnerGroupMembers.Count
                        }
                    }
                    if($PreCheckVerboseDiagnostics) {
                        Write-Log "Pre-check source '$siteWebUrl' [site-group-owners]: ownersProcessed=$(@($groupOwners).Count), resolvedUsers=$ownerResolved" "INFO"
                    }
                } catch {
                    Write-Log "Recursive pre-check: failed to read owners for site group '$siteGroupId': $($_.Exception.Message)" "WARN"
                }
            }
        }

        foreach($lib in $siteLibs) {
            $libKey = if(-not [string]::IsNullOrWhiteSpace($lib.libKey)) { $lib.libKey } else { "$($lib.siteWebUrl)|$($lib.listId)" }

            $libPreAllowedUsers = New-CaseInsensitiveStringSet
            $sitePreAllowedUsers = Ensure-StringSet -Set $sitePreAllowedUsers
            $libPreAllowedUsers = Ensure-StringSet -Set $libPreAllowedUsers
            foreach($upn in $sitePreAllowedUsers) { [void]$libPreAllowedUsers.Add($upn) }
            $libAllowAllUsers = $siteAllowAllUsers

            if($EnableRecursivePermissionPreCheck -and -not $siteAllowAllUsers) {
                $listScope = Get-ScopeAllowedTargetUsers -SiteWebUrl $siteWebUrl -ScopeType List -ListId $lib.listId -MinimumPermissionLevel $MinimumPermissionLevel -TargetUserByUpn $targetUserByUpn -TargetUserById $targetUserById -EntraGroupCache $global:octo.EntraGroupTargetMemberCache -PreCheckStats $preCheckStats -IncludeEveryoneClaims $PreCheckIncludeEveryoneClaims -MaxDepth $PreCheckMaxRecursionDepth -MaxExpandedPrincipals $PreCheckMaxExpandedPrincipals
                if($listScope) {
                    $libPreAllowedUsers = Ensure-StringSet -Set $libPreAllowedUsers
                    if($listScope.AllowAllUsers) { $libAllowAllUsers = $true }
                    foreach($upn in @($listScope.AllowedUsers)) { [void]$libPreAllowedUsers.Add($upn) }
                    $listScopeResolvedCount = @($listScope.AllowedUsers).Count
                    if($listScope.AllowAllUsers -or $listScopeResolvedCount -gt 0) { $preCheckStats.ListScopeHits++ }
                    if($PreCheckVerboseDiagnostics) {
                        Write-Log "Pre-check source '$siteWebUrl' [list-roles:$($lib.listId)]: allowAll=$($listScope.AllowAllUsers), resolvedUsers=$listScopeResolvedCount" "INFO"
                    }
                }
            }

            foreach($user in $targetUserList) {
                $userUPN = $user.userPrincipalName
                $isPreAllowedForLib = $false

                if($EnableRecursivePermissionPreCheck) {
                    if($libAllowAllUsers) {
                        $isPreAllowedForLib = $true
                    } elseif(
                        -not [string]::IsNullOrWhiteSpace($userUPN) -and
                        $null -ne $libPreAllowedUsers -and
                        $libPreAllowedUsers -is [System.Collections.Generic.HashSet[string]]
                    ) {
                        $isPreAllowedForLib = $libPreAllowedUsers.Contains($userUPN)
                    }
                }

                if($isPreAllowedForLib) {
                    $permissionMatrix[$userUPN].Add(@{
                        shortCut = $lib.shortcutInfo
                        siteName = $lib.siteDisplayName
                        listName = $lib.listDisplayName
                    })
                    $preCheckStats.PreAllowedPairs++
                    $sitePreAllowedPairCount++
                    continue
                }

                $checkPairs.Add(@{
                    userUPN = $userUPN
                    listId  = $lib.listId
                    libKey  = $libKey
                    libInfo = $lib
                })
                if($EnableRecursivePermissionPreCheck) {
                    $preCheckStats.FallbackPairs++
                    $siteFallbackPairCount++
                }
            }
        }

        if($EnableRecursivePermissionPreCheck -and $PreCheckVerboseDiagnostics) {
            Write-Log "Pre-check diagnostics '$siteWebUrl': pre-allowed=$sitePreAllowedPairCount, fallback=$siteFallbackPairCount, siteAllowAll=$siteAllowAllUsers" "INFO"
        }

        if($checkPairs.Count -eq 0) { continue }

        $totalBatchChecks += $checkPairs.Count

        # Submit batch permission check job for this site
        # Wait for a free slot, processing completed jobs while waiting (prevents deadlock)
        while(-not $batchSemaphore.Wait(200)) {
            for($ji = $batchJobs.Count - 1; $ji -ge 0; $ji--) {
                if($batchJobs[$ji].Handle.IsCompleted) {
                    $batchCompletedSinceAdjust++
                    if($batchSemaphore.CurrentCount -lt $MaxParallelLimit) { [void]$batchSemaphore.Release() }
                    try {
                        $batchJobResult = $batchJobs[$ji].PowerShell.EndInvoke($batchJobs[$ji].Handle)
                        if($batchJobResult) {
                            foreach($bjr in $batchJobResult) {
                                if($bjr.LogMessages) {
                                    foreach($log in $bjr.LogMessages) { Write-Log "  $($log.Message)" $log.Level }
                                }
                                if($bjr.ThrottleHits) { $batchThrottleHits += $bjr.ThrottleHits; $phase2Total429s += $bjr.ThrottleHits }
                                if($bjr.Results) {
                                    foreach($pr in $bjr.Results) {
                                        if($pr.HasAccess) {
                                            $matchingPair = $batchJobs[$ji].CheckPairs | Where-Object { $_.libKey -eq $pr.LibKey } | Select-Object -First 1
                                            if($matchingPair -and $matchingPair.libInfo) {
                                                $permissionMatrix[$pr.UserUPN].Add(@{
                                                    shortCut = $matchingPair.libInfo.shortcutInfo
                                                    siteName = $matchingPair.libInfo.siteDisplayName
                                                    listName = $matchingPair.libInfo.listDisplayName
                                                })
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        if($batchJobs[$ji].PowerShell.Streams.Error.Count -gt 0) {
                            foreach($err in $batchJobs[$ji].PowerShell.Streams.Error) {
                                Write-Log "  Batch runspace error: $($err.Exception.Message)" "WARN"
                            }
                        }
                        # Count 429 retries from Information stream (Write-Host output from Invoke-SharePointBatch retries)
                        foreach($info in $batchJobs[$ji].PowerShell.Streams.Information) {
                            if($info.MessageData -like "*429*") { $batchThrottleHits++; $phase2Total429s++ }
                        }
                    } catch {
                        Write-Log "  Error collecting batch result for '$($batchJobs[$ji].SiteUrl)': $($_.Exception.Message)" "WARN"
                    }
                    $batchJobs[$ji].PowerShell.Dispose()
                    $batchJobs.RemoveAt($ji)

                    $phase2CompletedJobs++
                    Write-Progress -Id 0 -Activity "Phase 2: Checking permissions" -Status "$phase2CompletedJobs/$siteTotal sites | concurrent: $batchCurrentLimit | batch: $BatchSize | 429s: $phase2Total429s | checks: $totalBatchChecks" -PercentComplete ([math]::Min(100, [math]::Round(($phase2CompletedJobs / [math]::Max(1,$siteTotal)) * 100)))
                }
            }
        }
        $ps = [powershell]::Create()
        $ps.RunspacePool = $batchPool
        [void]$ps.AddScript($batchPermCheckBlock)
        [void]$ps.AddParameter('siteWebUrl', $siteWebUrl)
        [void]$ps.AddParameter('checkPairs', $checkPairs.ToArray())
        [void]$ps.AddParameter('batchSize', $BatchSize)
        [void]$ps.AddParameter('graphUrl', $graphUrlValue)
        [void]$ps.AddParameter('idpUrl', $idpUrlValue)
        [void]$ps.AddParameter('sharepointUrl', $sharepointUrlValue)
        [void]$ps.AddParameter('minPermLevel', $MinimumPermissionLevel)
        [void]$ps.AddParameter('seedTokens', $preSeededTokens2)
        $batchJobs.Add(@{
            PowerShell = $ps
            Handle     = $ps.BeginInvoke()
            SiteUrl    = $siteWebUrl
            CheckPairs = $checkPairs
        })

        # Eagerly collect any completed jobs to free up semaphore slots
        for($ji = $batchJobs.Count - 1; $ji -ge 0; $ji--) {
            if($batchJobs[$ji].Handle.IsCompleted) {
                $batchCompletedSinceAdjust++
                if($batchSemaphore.CurrentCount -lt $MaxParallelLimit) { [void]$batchSemaphore.Release() }
                try {
                    $batchJobResult = $batchJobs[$ji].PowerShell.EndInvoke($batchJobs[$ji].Handle)
                    if($batchJobResult) {
                        foreach($bjr in $batchJobResult) {
                            if($bjr.LogMessages) {
                                foreach($log in $bjr.LogMessages) { Write-Log "  $($log.Message)" $log.Level }
                            }
                            if($bjr.ThrottleHits) { $batchThrottleHits += $bjr.ThrottleHits; $phase2Total429s += $bjr.ThrottleHits }
                            if($bjr.Results) {
                                foreach($pr in $bjr.Results) {
                                    if($pr.HasAccess) {
                                        # Find the lib info from checkPairs
                                        $matchingPair = $batchJobs[$ji].CheckPairs | Where-Object { $_.libKey -eq $pr.LibKey } | Select-Object -First 1
                                        if($matchingPair -and $matchingPair.libInfo) {
                                            $permissionMatrix[$pr.UserUPN].Add(@{
                                                shortCut = $matchingPair.libInfo.shortcutInfo
                                                siteName = $matchingPair.libInfo.siteDisplayName
                                                listName = $matchingPair.libInfo.listDisplayName
                                            })
                                        }
                                    }
                                }
                            }
                        }
                    }
                    if($batchJobs[$ji].PowerShell.Streams.Error.Count -gt 0) {
                        foreach($err in $batchJobs[$ji].PowerShell.Streams.Error) {
                            Write-Log "  Batch runspace error: $($err.Exception.Message)" "WARN"
                        }
                    }
                    # Count 429 retries from Information stream (Write-Host output from Invoke-SharePointBatch retries)
                    foreach($info in $batchJobs[$ji].PowerShell.Streams.Information) {
                        if($info.MessageData -like "*[SharePoint]*429*") { $batchThrottleHits++; $phase2Total429s++ }
                        elseif($info.MessageData -like "*429*") { $batchThrottleHits++; $phase2Total429s++ }
                    }
                } catch {
                    Write-Log "  Error collecting batch result for '$($batchJobs[$ji].SiteUrl)': $($_.Exception.Message)" "WARN"
                }
                $batchJobs[$ji].PowerShell.Dispose()
                $batchJobs.RemoveAt($ji)

                $phase2CompletedJobs++
                Write-Progress -Id 0 -Activity "Phase 2: Checking permissions" -Status "$phase2CompletedJobs/$siteTotal sites | concurrent: $batchCurrentLimit | batch: $BatchSize | 429s: $phase2Total429s | checks: $totalBatchChecks" -PercentComplete ([math]::Min(100, [math]::Round(($phase2CompletedJobs / [math]::Max(1,$siteTotal)) * 100)))

                # Adaptive throttle adjustment
                if($batchCompletedSinceAdjust -ge 20) {
                    $batchThrottleRate = $batchThrottleHits / $batchCompletedSinceAdjust
                    if($batchThrottleRate -gt 0.1 -and $batchCurrentLimit -gt 2) {
                        $newLimit = [math]::Max(2, [math]::Floor($batchCurrentLimit * 0.5))
                        $reduction = $batchCurrentLimit - $newLimit
                        for($r = 0; $r -lt $reduction; $r++) { [void]$batchSemaphore.Wait(0) }
                        $batchCurrentLimit = $newLimit
                        Write-Log "  Adaptive batch throttle: reducing to $batchCurrentLimit (429 rate: $([math]::Round($batchThrottleRate * 100))%)" "WARN"
                    } elseif($batchThrottleRate -eq 0 -and $batchCurrentLimit -lt $MaxParallelLimit) {
                        $increase = [math]::Min(2, $MaxParallelLimit - $batchCurrentLimit)
                        $batchCurrentLimit += $increase
                        if($increase -gt 0 -and $batchSemaphore.CurrentCount -lt $MaxParallelLimit) {
                            $safeRelease = [math]::Min($increase, $MaxParallelLimit - $batchSemaphore.CurrentCount)
                            if($safeRelease -gt 0) { [void]$batchSemaphore.Release($safeRelease) }
                        }
                        Write-Log "  Adaptive batch throttle: increasing to $batchCurrentLimit (no 429s)" "INFO"
                    }
                    $batchThrottleHits = 0
                    $batchCompletedSinceAdjust = 0
                }
            }
        }
    }

    # Collect remaining batch jobs
    Write-Log "Collecting remaining batch permission results..." "INFO"
    while($batchJobs.Count -gt 0) {
        for($ji = $batchJobs.Count - 1; $ji -ge 0; $ji--) {
            if($batchJobs[$ji].Handle.IsCompleted) {
                if($batchSemaphore.CurrentCount -lt $MaxParallelLimit) { [void]$batchSemaphore.Release() }
                try {
                    $batchJobResult = $batchJobs[$ji].PowerShell.EndInvoke($batchJobs[$ji].Handle)
                    if($batchJobResult) {
                        foreach($bjr in $batchJobResult) {
                            if($bjr.LogMessages) {
                                foreach($log in $bjr.LogMessages) { Write-Log "  $($log.Message)" $log.Level }
                            }
                            if($bjr.ThrottleHits) { $batchThrottleHits += $bjr.ThrottleHits; $phase2Total429s += $bjr.ThrottleHits }
                            if($bjr.Results) {
                                foreach($pr in $bjr.Results) {
                                    if($pr.HasAccess) {
                                        $matchingPair = $batchJobs[$ji].CheckPairs | Where-Object { $_.libKey -eq $pr.LibKey } | Select-Object -First 1
                                        if($matchingPair -and $matchingPair.libInfo) {
                                            $permissionMatrix[$pr.UserUPN].Add(@{
                                                shortCut = $matchingPair.libInfo.shortcutInfo
                                                siteName = $matchingPair.libInfo.siteDisplayName
                                                listName = $matchingPair.libInfo.listDisplayName
                                            })
                                        }
                                    }
                                }
                            }
                        }
                    }
                    # Count 429 retries from Information stream (Write-Host output from Invoke-SharePointBatch retries)
                    foreach($info in $batchJobs[$ji].PowerShell.Streams.Information) {
                        if($info.MessageData -like "*[SharePoint]*429*") { $batchThrottleHits++; $phase2Total429s++ }
                        elseif($info.MessageData -like "*429*") { $batchThrottleHits++; $phase2Total429s++ }
                    }
                } catch {
                    Write-Log "  Error collecting batch result for '$($batchJobs[$ji].SiteUrl)': $($_.Exception.Message)" "WARN"
                }
                $batchJobs[$ji].PowerShell.Dispose()
                $batchJobs.RemoveAt($ji)

                $phase2CompletedJobs++
                Write-Progress -Id 0 -Activity "Phase 2: Checking permissions" -Status "$phase2CompletedJobs/$siteTotal sites | concurrent: $batchCurrentLimit | batch: $BatchSize | 429s: $phase2Total429s | checks: $totalBatchChecks" -PercentComplete ([math]::Min(100, [math]::Round(($phase2CompletedJobs / [math]::Max(1,$siteTotal)) * 100)))
            }
        }
        if($batchJobs.Count -gt 0) { Start-Sleep -Milliseconds 100 }
    }

    Write-Progress -Id 0 -Activity "Phase 2: Checking permissions" -Completed

    } finally {
        if($batchPool) { $batchPool.Close(); $batchPool.Dispose() }
        if($batchSemaphore) { $batchSemaphore.Dispose() }
    }

    $totalPairs = $targetUserList.Count * $allSiteLibraries.Count
    Write-Log "Permission matrix complete:" "SUCCESS"
    Write-Log "  Total user×library pairs: $totalPairs" "INFO"
    Write-Log "  Actual API checks (batched): $totalBatchChecks" "INFO"
    if($EnableRecursivePermissionPreCheck) {
        Write-Log "  Pre-allowed pairs (site/list/group recursion): $($preCheckStats.PreAllowedPairs)" "INFO"
        Write-Log "  Fallback pairs (effective permission API): $($preCheckStats.FallbackPairs)" "INFO"
        Write-Log "  Scope hits - site: $($preCheckStats.SiteScopeHits), list: $($preCheckStats.ListScopeHits)" "INFO"
        Write-Log "  Entra group cache - hits: $($preCheckStats.CacheHits), misses: $($preCheckStats.CacheMisses), entries: $($global:octo.EntraGroupTargetMemberCache.Count)" "INFO"
        Write-Log "  Recursion cutoffs (depth/node guards): $($preCheckStats.RecursionCutoffs)" "INFO"
    }
    $phaseTimings['Phase 2: Permission matrix'] = $phaseStopwatch.Elapsed

    # ============================================================
    # PHASE 3: Per-user shortcut creation/deletion
    # ============================================================
    $phaseStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Log "--- Phase 3: Applying shortcuts per user ---" "INFO"

    $totalStats = @{
        UsersProcessed   = 0
        UsersSkipped     = 0
        UsersFailed      = 0
        ShortcutsCreated = 0
        ShortcutsRenamed = 0
        ShortcutsSkipped = 0
        ShortcutsDeleted = 0
        Warnings         = 0
        Errors           = 0
    }
    $existingShortcutConflicts = [System.Collections.Generic.List[hashtable]]::new()

    $userIndex = 0
    $userTotal = $targetUserList.Count
    foreach($targetUser in $targetUserList) {
        $userIndex++
        $userId = $targetUser.id
        $userUPN = $targetUser.userPrincipalName

        Write-Progress -Id 0 -Activity "Phase 3: Applying shortcuts" -Status "User $userIndex / $userTotal — $userUPN" -PercentComplete ([math]::Min(100, [math]::Round(($userIndex / $userTotal) * 100)))
        Write-Log "========================================" "INFO"
        Write-Log "Processing user: $userUPN ($userIndex/$userTotal)" "INFO"
        Write-Log "========================================" "INFO"

        try {
            # Check if user has a OneDrive provisioned
            try {
                $userDrive = New-GraphQuery -Uri "$($global:octo.graphUrl)/v1.0/users/$userId/drive" -Method GET -MaxAttempts 3
            } catch {
                Write-Log "  OneDrive not provisioned for this user, skipping..." "WARN"
                $totalStats.UsersSkipped++
                continue
            }

            # Retrieve the desired shortcuts from the permission matrix (already computed in Phase 2)
            # Deduplicate by target (siteId+webId+listId) in case the same library was added multiple times
            $rawDesired = @($permissionMatrix[$userUPN])
            $seenTargets = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $desiredShortcuts = @($rawDesired | Where-Object {
                $targetKey = "$($_.shortCut.siteId)|$($_.shortCut.webId)|$($_.shortCut.listId)"
                $seenTargets.Add($targetKey)
            })
            if($desiredShortcuts.Count -lt $rawDesired.Count) {
                Write-Log "  Deduplicated desired shortcuts: $($rawDesired.Count) → $($desiredShortcuts.Count)" "WARN"
            }
            Write-Log "  User has access to $($desiredShortcuts.Count) document libraries" "SUCCESS"

            # Check if target folder exists, create if not
            Write-Log "  Checking for '$FolderName' folder in OneDrive..." "INFO"
            $targetFolder = $null

            try {
                $targetFolder = New-GraphQuery -Uri "$($global:octo.graphUrl)/v1.0/users/$userId/drive/root:/$($FolderName)?`$expand=listItem" -Method "GET"
                Write-Log "  Folder '$FolderName' already exists" "INFO"
            } catch {
                if($DryRun) {
                    Write-Log "  [DRY RUN] Would create folder '$FolderName' — skipping user (folder needed for processing)" "WARN"
                    $totalStats.UsersSkipped++
                    continue
                }
                Write-Log "  Creating folder '$FolderName'..." "INFO"
                $folderBody = @{
                    name = $FolderName
                    folder = @{}
                    "@microsoft.graph.conflictBehavior" = "rename"
                } | ConvertTo-Json -Depth 3
                $null = New-GraphQuery -Uri "$($global:octo.graphUrl)/v1.0/users/$userId/drive/root/children" -Method POST -Body $folderBody
                Write-Log "  Folder created successfully" "SUCCESS"
                $targetFolder = New-GraphQuery -Uri "$($global:octo.graphUrl)/v1.0/users/$userId/drive/root:/$($FolderName)?`$expand=listItem" -Method "GET"
            }

            if(-not $targetFolder -or -not $targetFolder.listItem -or [string]::IsNullOrWhiteSpace($targetFolder.listItem.webUrl)) {
                Write-Log "  Could not resolve OneDrive folder webUrl metadata, skipping user..." "WARN"
                $totalStats.UsersSkipped++
                continue
            }

            # Determine root OneDrive URL
            $urlParts = $targetFolder.listItem.webUrl -split "/personal/"
            if($urlParts.Count -lt 2) {
                Write-Log "  Unexpected OneDrive URL format '$($targetFolder.listItem.webUrl)', skipping user..." "WARN"
                $totalStats.UsersSkipped++
                continue
            }
            $rootUrl = $urlParts[0]
            $relativeParts = $urlParts[1].Split('/')
            if($relativeParts.Count -lt 2 -or [string]::IsNullOrWhiteSpace($relativeParts[0]) -or [string]::IsNullOrWhiteSpace($relativeParts[1])) {
                Write-Log "  Could not parse OneDrive personal path components from '$($targetFolder.listItem.webUrl)', skipping user..." "WARN"
                $totalStats.UsersSkipped++
                continue
            }
            $userComponent = $relativeParts[0]
            $libraryName = $relativeParts[1]

            if(-not $targetFolder.parentReference -or [string]::IsNullOrWhiteSpace($targetFolder.parentReference.siteId)) {
                Write-Log "  Missing parentReference.siteId for OneDrive folder, skipping user..." "WARN"
                $totalStats.UsersSkipped++
                continue
            }

            $docLibrary = (New-GraphQuery -Uri "$($global:octo.graphUrl)/v1.0/sites/$($targetFolder.parentReference.siteId)/lists" -Method "GET") | Where-Object { $_.list.template -eq "mySiteDocumentLibrary" -and !$_.list.hidden}
            $docLibrary = @($docLibrary) | Select-Object -First 1
            if(-not $docLibrary -or [string]::IsNullOrWhiteSpace($docLibrary.id)) {
                Write-Log "  Could not resolve personal document library metadata, skipping user..." "WARN"
                $totalStats.UsersSkipped++
                continue
            }

            $graphTokenPhase3 = Get-AccessToken -resource $global:octo.graphUrl
            $spTokenPhase3 = Get-AccessToken -resource $global:octo.sharepointUrl
            $preSeededTokensPhase3 = @{
                $global:octo.graphUrl = @{ accessToken = $graphTokenPhase3; validFrom = Get-Date }
                $global:octo.sharepointUrl = @{ accessToken = $spTokenPhase3; validFrom = Get-Date }
            }

            $currentShortCuts = @()

            # Retrieve current shortcuts
            Write-Log "  Getting target info for all current shortcuts..." "INFO"
            $folderContents = New-GraphQuery -Uri "$rootUrl/personal/$userComponent/_api/web/GetFolderByServerRelativeUrl('/personal/$userComponent/$libraryName/$FolderName')/Files?`$top=5000&`$format=json&`$expand=listItem" -Method GET

            # Clean up unexpected folders
            if(-not $DryRun) {
                New-GraphQuery -Uri "$rootUrl/personal/$userComponent/_api/web/GetFolderByServerRelativeUrl('/personal/$userComponent/$libraryName/$FolderName')/Folders?`$top=5000&`$format=json&`$expand=listItem" -Method GET | ForEach-Object {
                    if($_.UniqueId){
                        New-GraphQuery -Uri "$rootUrl/personal/$userComponent/_api/web/GetFolderById('$($_.UniqueId)')/DeleteObject()" -Method POST
                        Write-Log "  Deleted unexpected folder '$($_.Name)'" "WARN"
                    }
                }
            }

            if(@($folderContents).Count -gt 0) {
                $shortcutsWithIds = @($folderContents | Where-Object { $_.UniqueId })
                for($si = 0; $si -lt $shortcutsWithIds.Count; $si += $BatchSize) {
                    $chunk = @($shortcutsWithIds[$si..([math]::Min($si + $BatchSize - 1, $shortcutsWithIds.Count - 1))])
                    $metaRequests = [System.Collections.Generic.List[object]]::new()
                    foreach($sc in $chunk) {
                        $metaRequests.Add(@{
                            ShortcutId  = [string]$sc.UniqueId
                            RelativeUrl = "_api/web/lists('$($docLibrary.id)')/GetItemByUniqueId('$($sc.UniqueId)')?`$expand=FieldValuesAsText"
                        })
                    }

                    try {
                        $metaBatchResults = Invoke-SharePointBatch -SiteUrl "$rootUrl/personal/$userComponent" -SubRequests $metaRequests.ToArray() -MaxAttempts 3
                        for($ri = 0; $ri -lt $metaRequests.Count; $ri++) {
                            $req = $metaRequests[$ri]
                            $batchResult = if($ri -lt $metaBatchResults.Count) { $metaBatchResults[$ri] } else { $null }
                            if($batchResult -and $batchResult.StatusCode -eq 200 -and $batchResult.Data) {
                                $metaData = $batchResult.Data
                                if($null -ne $metaData.d) { $metaData = $metaData.d }
                                $fv = $metaData.FieldValuesAsText
                                if($fv) {
                                    $currentShortCuts += @{
                                        "ID" = $req.ShortcutId
                                        "Name" = $fv.FileLeafRef
                                        "targetSiteId" = $fv.A2ODRemoteItemSiteId
                                        "targetWebId" = $fv.A2ODRemoteItemWebId
                                        "targetListId" = $fv.A2ODRemoteItemListId
                                        "targetItemUniqueId" = $fv.A2ODRemoteItemUniqueId
                                    }
                                    continue
                                }
                            }

                            # Fallback for failed/empty batch entries
                            try {
                                $shortCutMetaData = New-GraphQuery -Uri "$rootUrl/personal/$userComponent/_api/web/lists('$($docLibrary.id)')/GetItemByUniqueId('$($req.ShortcutId)')?`$expand=FieldValuesAsText" -Method GET -MaxAttempts 3
                                $currentShortCuts += @{
                                    "ID" = $req.ShortcutId
                                    "Name" = $shortCutMetaData.FieldValuesAsText.FileLeafRef
                                    "targetSiteId" = $shortCutMetaData.FieldValuesAsText.A2ODRemoteItemSiteId
                                    "targetWebId" = $shortCutMetaData.FieldValuesAsText.A2ODRemoteItemWebId
                                    "targetListId" = $shortCutMetaData.FieldValuesAsText.A2ODRemoteItemListId
                                    "targetItemUniqueId" = $shortCutMetaData.FieldValuesAsText.A2ODRemoteItemUniqueId
                                }
                            } catch {
                                Write-Log "  Failed to resolve shortcut metadata for item '$($req.ShortcutId)': $($_.Exception.Message)" "WARN"
                            }
                        }
                    } catch {
                        Write-Log "  Shortcut metadata batch failed, falling back to individual calls: $($_.Exception.Message)" "WARN"
                        foreach($req in $metaRequests) {
                            try {
                                $shortCutMetaData = New-GraphQuery -Uri "$rootUrl/personal/$userComponent/_api/web/lists('$($docLibrary.id)')/GetItemByUniqueId('$($req.ShortcutId)')?`$expand=FieldValuesAsText" -Method GET -MaxAttempts 3
                                $currentShortCuts += @{
                                    "ID" = $req.ShortcutId
                                    "Name" = $shortCutMetaData.FieldValuesAsText.FileLeafRef
                                    "targetSiteId" = $shortCutMetaData.FieldValuesAsText.A2ODRemoteItemSiteId
                                    "targetWebId" = $shortCutMetaData.FieldValuesAsText.A2ODRemoteItemWebId
                                    "targetListId" = $shortCutMetaData.FieldValuesAsText.A2ODRemoteItemListId
                                    "targetItemUniqueId" = $shortCutMetaData.FieldValuesAsText.A2ODRemoteItemUniqueId
                                }
                            } catch {
                                Write-Log "  Failed to resolve shortcut metadata for item '$($req.ShortcutId)': $($_.Exception.Message)" "WARN"
                            }
                        }
                    }
                }
            }

            # Deduplicate by ID in case the SP REST endpoint returned duplicate entries
            $beforeDedup = $currentShortCuts.Count
            $seenIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $currentShortCuts = @($currentShortCuts | Where-Object { $_.ID -and $seenIds.Add($_.ID) })
            if($currentShortCuts.Count -lt $beforeDedup) {
                Write-Log "  Deduplicated current shortcuts: $beforeDedup → $($currentShortCuts.Count)" "WARN"
            }

            Write-Log "  Currently has $($currentShortCuts.count) shortcuts" "INFO"

            $successCount = 0
            $skipCount = 0
            $warnCount = 0
            $errorCount = 0
            $renameCount = 0

            # Build target maps once for faster matching
            $existingTargetKeys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach($existing in $currentShortCuts) {
                $null = $existingTargetKeys.Add("$($existing.targetSiteId)|$($existing.targetWebId)|$($existing.targetListId)")
            }

            $desiredTargetKeys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            foreach($desired in $desiredShortcuts) {
                $null = $desiredTargetKeys.Add("$($desired.shortcut.siteId)|$($desired.shortcut.webId)|$($desired.shortcut.listId)")
            }

            $shortcutsToCreate = @($desiredShortcuts | Where-Object {
                -not $existingTargetKeys.Contains("$($_.shortcut.siteId)|$($_.shortcut.webId)|$($_.shortcut.listId)")
            })
            $skipCount = [math]::Max(0, $desiredShortcuts.Count - $shortcutsToCreate.Count)

            $shortcutsToDelete = @($currentShortCuts | Where-Object {
                -not $desiredTargetKeys.Contains("$($_.targetSiteId)|$($_.targetWebId)|$($_.targetListId)")
            })

            # Safety guard: never allow a full wipe of this folder in one run.
            # This protects against upstream metadata regressions that could mark all existing shortcuts as obsolete.
            if($shortcutsToDelete.Count -gt 0 -and $shortcutsToDelete.Count -eq $currentShortCuts.Count) {
                Write-Log "  Mass-delete protection triggered: calculated deletion of all $($currentShortCuts.Count) current shortcuts. Skipping deletion for this user." "WARN"
                $warnCount++
                $shortcutsToDelete = @()
            }

            # Create missing shortcuts
            if($DryRun) {
                foreach($desiredShortcut in $shortcutsToCreate) {
                    Write-Log "    [DRY RUN] Would create shortcut for '$($desiredShortcut.shortcut.siteUrl)'" "INFO"
                    $successCount++
                }
            } elseif($shortcutsToCreate.Count -gt 0) {
                $createIss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
                $createIss.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new('Get-AccessToken', (Get-Command Get-AccessToken).Definition))
                $createIss.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new('New-GraphQuery', (Get-Command New-GraphQuery).Definition))
                $createIss.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new('Write-Log', (Get-Command Write-Log).Definition))
                $createIss.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new('Get-CleanedShortcutName', (Get-Command Get-CleanedShortcutName).Definition))
                $createIss.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new('ClientId', $ClientId, ''))
                $createIss.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new('TenantId', $TenantId, ''))
                $createIss.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new('CertificateThumbprint', $CertificateThumbprint, ''))
                $createIss.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new('CertificatePath', $CertificatePath, ''))
                $createIss.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new('CertificatePassword', $CertificatePassword, ''))
                $createIss.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new('linkNameReplacements', $linkNameReplacements, ''))
                $createPool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, [math]::Max(1, $ShortcutActionParallelLimit), $createIss, $Host)
                $createPool.Open()

                try {

                $createBlock = {
                    param([hashtable]$desiredShortcut, [string]$userId, [string]$targetFolderId, [string]$folderName, [string]$graphUrl, [string]$idpUrl, [string]$sharepointUrl, [hashtable]$seedTokens, [int]$maxAttempts)

                    $global:octo = @{ graphUrl = $graphUrl; idpUrl = $idpUrl; sharepointUrl = $sharepointUrl; LCCachedTokens = $seedTokens }
                    $global:octo.RunspaceLogBuffer = [System.Collections.Generic.List[hashtable]]::new()
                    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Web")

                    try {
                        $shortcutBody = @{
                            name = $($desiredShortcut.listName)
                            remoteItem = @{ sharepointIds = $desiredShortcut.shortcut }
                            "@microsoft.graph.conflictBehavior" = "rename"
                        } | ConvertTo-Json -Depth 5

                        $newShortCut = New-GraphQuery -MaxAttempts $maxAttempts -Uri "$graphUrl/v1.0/users/$userId/drive/root/children" -Method POST -Body $shortcutBody

                        if($newShortCut.id -and $targetFolderId) {
                            $moveBody = @{ parentReference = @{ id = $targetFolderId } } | ConvertTo-Json -Depth 3
                            $newShortCut = New-GraphQuery -MaxAttempts $maxAttempts -Uri "$graphUrl/v1.0/users/$userId/drive/items/$($newShortCut.id)" -Method PATCH -Body $moveBody
                        }

                        $renamed = $false
                        if($newShortCut.id) {
                            $cleanName = Get-CleanedShortcutName -Name $newShortCut.name
                            if($cleanName -and $cleanName -ne $newShortCut.name) {
                                $renameSucceeded = $false
                                for($i = 0; $i -lt 6; $i++) {
                                    $candidateName = if($i -eq 0) { $cleanName } else { "$cleanName`_$i" }
                                    try {
                                        $renameBody = @{ name = $candidateName } | ConvertTo-Json
                                        $null = New-GraphQuery -MaxAttempts $maxAttempts -Uri "$graphUrl/v1.0/users/$userId/drive/items/$($newShortCut.id)" -Method PATCH -Body $renameBody
                                        $renameSucceeded = $true
                                        $renamed = ($candidateName -ne $newShortCut.name)
                                        break
                                    } catch {
                                        $renameMsg = $_.Exception.Message
                                        $isNameConflict = $renameMsg -like '*already exists*' -or $renameMsg -like '*nameAlreadyExists*' -or $renameMsg -like '*Conflict*'
                                        if(-not $isNameConflict -or $i -ge 5) {
                                            throw
                                        }
                                    }
                                }
                                if(-not $renameSucceeded) {
                                    return [PSCustomObject]@{ Status = 'Warning'; Url = $desiredShortcut.shortcut.siteUrl; Message = "Rename attempt did not complete for '$($desiredShortcut.shortcut.siteUrl)'"; Renamed = $false }
                                }
                            }
                        }

                        return [PSCustomObject]@{ Status = 'Created'; Url = $desiredShortcut.shortcut.siteUrl; Message = "Created shortcut for '$($desiredShortcut.shortcut.siteUrl)'"; Renamed = $renamed; LogMessages = @($global:octo.RunspaceLogBuffer) }
                    } catch {
                        $msg = $_.Exception.Message
                        if($msg -like '*descendant shortcut exists*' -or $msg -like '*shortcut already exists*') {
                            return [PSCustomObject]@{ Status = 'Conflict'; Url = $desiredShortcut.shortcut.siteUrl; Message = $msg; Renamed = $false; LogMessages = @($global:octo.RunspaceLogBuffer) }
                        }
                        return [PSCustomObject]@{ Status = 'Error'; Url = $desiredShortcut.shortcut.siteUrl; Message = $msg; Renamed = $false; LogMessages = @($global:octo.RunspaceLogBuffer) }
                    }
                }

                $createJobs = [System.Collections.Generic.List[hashtable]]::new()
                foreach($desiredShortcut in $shortcutsToCreate) {
                    $ps = [powershell]::Create()
                    $ps.RunspacePool = $createPool
                    [void]$ps.AddScript($createBlock)
                    [void]$ps.AddParameter('desiredShortcut', $desiredShortcut)
                    [void]$ps.AddParameter('userId', $userId)
                    [void]$ps.AddParameter('targetFolderId', $targetFolder.id)
                    [void]$ps.AddParameter('folderName', $FolderName)
                    [void]$ps.AddParameter('graphUrl', $global:octo.graphUrl)
                    [void]$ps.AddParameter('idpUrl', $global:octo.idpUrl)
                    [void]$ps.AddParameter('sharepointUrl', $global:octo.sharepointUrl)
                    [void]$ps.AddParameter('seedTokens', $preSeededTokensPhase3)
                    [void]$ps.AddParameter('maxAttempts', $GraphMutationMaxAttempts)
                    $createJobs.Add(@{ PowerShell = $ps; Handle = $ps.BeginInvoke() })
                }

                $createCompleted = 0
                while($createJobs.Count -gt 0) {
                    for($ci = $createJobs.Count - 1; $ci -ge 0; $ci--) {
                        if($createJobs[$ci].Handle.IsCompleted) {
                            $createCompleted++
                            Write-Progress -Id 1 -ParentId 0 -Activity "Creating shortcuts" -Status "$createCompleted / $($shortcutsToCreate.Count)" -PercentComplete ([math]::Min(100, [math]::Round(($createCompleted / [math]::Max(1,$shortcutsToCreate.Count)) * 100)))
                            try {
                                $createResults = $createJobs[$ci].PowerShell.EndInvoke($createJobs[$ci].Handle)
                                foreach($result in $createResults) {
                                    if($result.LogMessages) {
                                        foreach($log in $result.LogMessages) { Write-Log "    $($log.Message)" $log.Level }
                                    }
                                    switch($result.Status) {
                                        'Created' {
                                            $successCount++
                                            if($result.Renamed) { $renameCount++ }
                                            Write-Log "    $($result.Message)" "SUCCESS"
                                        }
                                        'Conflict' {
                                            $warnCount++
                                            Write-Log "    Existing shortcut conflict for '$($result.Url)': $($result.Message)" "WARN"
                                            $existingShortcutConflicts.Add(@{ User = $userUPN; Url = $result.Url })
                                        }
                                        'Warning' {
                                            $warnCount++
                                            Write-Log "    $($result.Message)" "WARN"
                                        }
                                        default {
                                            $errorCount++
                                            Write-Log "    Failed to create shortcut for '$($result.Url)': $($result.Message)" "ERROR"
                                        }
                                    }
                                }
                            } catch {
                                $errorCount++
                                Write-Log "    Failed to collect create result: $($_.Exception.Message)" "ERROR"
                            }
                            $createJobs[$ci].PowerShell.Dispose()
                            $createJobs.RemoveAt($ci)
                        }
                    }
                    if($createJobs.Count -gt 0) { Start-Sleep -Milliseconds 100 }
                }

                Write-Progress -Id 1 -ParentId 0 -Activity "Creating shortcuts" -Completed

                } finally {
                    if($createPool) { $createPool.Close(); $createPool.Dispose() }
                }
            }

            # Rename existing shortcuts if link name cleanup patterns apply
            if($linkNameReplacements.Count -gt 0) {
                Write-Log "  Checking existing shortcuts for name cleanup..." "INFO"
                foreach($existing in $currentShortCuts) {
                    if(-not $existing.Name) { continue }
                    $cleanedName = Get-CleanedShortcutName -Name $existing.Name
                    if($cleanedName -ne $existing.Name -and $currentShortCuts.Name -notcontains $cleanedName) {
                        if($DryRun) {
                            Write-Log "    [DRY RUN] Would rename '$($existing.Name)' → '$cleanedName'" "INFO"
                            $renameCount++
                            continue
                        }
                        try {
                            $renameBody = @{ name = $cleanedName } | ConvertTo-Json
                            $Null = New-GraphQuery -MaxAttempts $GraphMutationMaxAttempts -Uri "$($global:octo.graphUrl)/v1.0/users/$userId/drive/items/$($existing.ID)" -Method PATCH -Body $renameBody
                            Write-Log "    Renamed '$($existing.Name)' → '$cleanedName'" "SUCCESS"
                            Start-Sleep -Milliseconds 200
                            $renameCount++
                        } catch {
                            Write-Log "    Failed to rename '$($existing.Name)': $($_.Exception.Message)" "WARN"
                        }
                    }
                }
            }

            Write-Progress -Id 1 -ParentId 0 -Activity "Creating shortcuts" -Completed

            # Delete shortcuts user should no longer have access to
            $deletedCount = 0
            if($DryRun) {
                foreach($existing in $shortcutsToDelete) {
                    Write-Log "    [DRY RUN] Would delete obsolete shortcut '$($existing.Name)'" "INFO"
                    $deletedCount++
                }
            } elseif($shortcutsToDelete.Count -gt 0) {
                $deleteIss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
                $deleteIss.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new('Get-AccessToken', (Get-Command Get-AccessToken).Definition))
                $deleteIss.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new('New-GraphQuery', (Get-Command New-GraphQuery).Definition))
                $deleteIss.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new('Write-Log', (Get-Command Write-Log).Definition))
                $deleteIss.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new('ClientId', $ClientId, ''))
                $deleteIss.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new('TenantId', $TenantId, ''))
                $deleteIss.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new('CertificateThumbprint', $CertificateThumbprint, ''))
                $deleteIss.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new('CertificatePath', $CertificatePath, ''))
                $deleteIss.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new('CertificatePassword', $CertificatePassword, ''))
                $deletePool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, [math]::Max(1, $ShortcutActionParallelLimit), $deleteIss, $Host)
                $deletePool.Open()

                try {

                $deleteBlock = {
                    param([hashtable]$existingShortcut, [string]$userId, [string]$graphUrl, [string]$idpUrl, [string]$sharepointUrl, [hashtable]$seedTokens, [int]$maxAttempts)

                    $global:octo = @{ graphUrl = $graphUrl; idpUrl = $idpUrl; sharepointUrl = $sharepointUrl; LCCachedTokens = $seedTokens }
                    $global:octo.RunspaceLogBuffer = [System.Collections.Generic.List[hashtable]]::new()
                    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Web")

                    try {
                        $null = New-GraphQuery -MaxAttempts $maxAttempts -Uri "$graphUrl/v1.0/users/$userId/drive/items/$($existingShortcut.ID)" -Method DELETE
                        return [PSCustomObject]@{ Status = 'Deleted'; Name = $existingShortcut.Name; Message = "Deleted obsolete shortcut '$($existingShortcut.Name)'"; LogMessages = @($global:octo.RunspaceLogBuffer) }
                    } catch {
                        return [PSCustomObject]@{ Status = 'Error'; Name = $existingShortcut.Name; Message = $_.Exception.Message; LogMessages = @($global:octo.RunspaceLogBuffer) }
                    }
                }

                $deleteJobs = [System.Collections.Generic.List[hashtable]]::new()
                foreach($existing in $shortcutsToDelete) {
                    $ps = [powershell]::Create()
                    $ps.RunspacePool = $deletePool
                    [void]$ps.AddScript($deleteBlock)
                    [void]$ps.AddParameter('existingShortcut', $existing)
                    [void]$ps.AddParameter('userId', $userId)
                    [void]$ps.AddParameter('graphUrl', $global:octo.graphUrl)
                    [void]$ps.AddParameter('idpUrl', $global:octo.idpUrl)
                    [void]$ps.AddParameter('sharepointUrl', $global:octo.sharepointUrl)
                    [void]$ps.AddParameter('seedTokens', $preSeededTokensPhase3)
                    [void]$ps.AddParameter('maxAttempts', $GraphMutationMaxAttempts)
                    $deleteJobs.Add(@{ PowerShell = $ps; Handle = $ps.BeginInvoke() })
                }

                $deleteCompleted = 0
                while($deleteJobs.Count -gt 0) {
                    for($di = $deleteJobs.Count - 1; $di -ge 0; $di--) {
                        if($deleteJobs[$di].Handle.IsCompleted) {
                            $deleteCompleted++
                            Write-Progress -Id 1 -ParentId 0 -Activity "Cleaning obsolete shortcuts" -Status "$deleteCompleted / $($shortcutsToDelete.Count)" -PercentComplete ([math]::Min(100, [math]::Round(($deleteCompleted / [math]::Max(1,$shortcutsToDelete.Count)) * 100)))
                            try {
                                $deleteResults = $deleteJobs[$di].PowerShell.EndInvoke($deleteJobs[$di].Handle)
                                foreach($result in $deleteResults) {
                                    if($result.LogMessages) {
                                        foreach($log in $result.LogMessages) { Write-Log "    $($log.Message)" $log.Level }
                                    }
                                    if($result.Status -eq 'Deleted') {
                                        $deletedCount++
                                        Write-Log "    $($result.Message)" "SUCCESS"
                                    } else {
                                        $errorCount++
                                        Write-Log "    Failed to delete '$($result.Name)': $($result.Message)" "ERROR"
                                    }
                                }
                            } catch {
                                $errorCount++
                                Write-Log "    Failed to collect delete result: $($_.Exception.Message)" "ERROR"
                            }
                            $deleteJobs[$di].PowerShell.Dispose()
                            $deleteJobs.RemoveAt($di)
                        }
                    }
                    if($deleteJobs.Count -gt 0) { Start-Sleep -Milliseconds 100 }
                }

                Write-Progress -Id 1 -ParentId 0 -Activity "Cleaning obsolete shortcuts" -Completed

                } finally {
                    if($deletePool) { $deletePool.Close(); $deletePool.Dispose() }
                }
            }

            Write-Log "  --- User Summary for $userUPN ---" "INFO"
            Write-Log "  Created: $successCount | Renamed: $renameCount | Skipped: $skipCount | Deleted: $deletedCount | Warnings: $warnCount | Errors: $errorCount" "INFO"

            $totalStats.UsersProcessed++
            $totalStats.ShortcutsCreated += $successCount
            $totalStats.ShortcutsRenamed += $renameCount
            $totalStats.ShortcutsSkipped += $skipCount
            $totalStats.ShortcutsDeleted += $deletedCount
            $totalStats.Warnings += $warnCount
            $totalStats.Errors += $errorCount

        } catch {
            Write-Log "  Failed to process user '$userUPN': $($_.Exception.Message)" "ERROR"
            Write-Log "  $($_.ScriptStackTrace)" "ERROR"
            $totalStats.UsersFailed++
        }
    }

    Write-Progress -Id 0 -Activity "Phase 3: Applying shortcuts" -Completed

    # Existing shortcut conflicts table
    if($existingShortcutConflicts.Count -gt 0) {
        Write-Log "" "INFO"
        Write-Log "=== Existing Shortcut Conflicts ($($existingShortcutConflicts.Count)) ===" "WARN"
        Write-Log ("{0,-45} {1}" -f "User", "URL") "INFO"
        Write-Log ("{0,-45} {1}" -f "----", "---") "INFO"
        foreach($conflict in $existingShortcutConflicts) {
            Write-Log ("{0,-45} {1}" -f $conflict.User, $conflict.Url) "WARN"
        }
        Write-Log "" "INFO"
    }

    # Global summary
    $modeLabel = if($DryRun) { " (DRY RUN)" } else { "" }
    Write-Log "=== Global Summary$modeLabel ===" "INFO"
    Write-Log "Users Processed: $($totalStats.UsersProcessed)" "SUCCESS"
    Write-Log "Users Skipped (no OneDrive): $($totalStats.UsersSkipped)" "WARN"
    Write-Log "Users Failed: $($totalStats.UsersFailed)" $(if($totalStats.UsersFailed -gt 0){"ERROR"}else{"SUCCESS"})
    Write-Log "Total Shortcuts Created: $($totalStats.ShortcutsCreated)" "SUCCESS"
    Write-Log "Total Shortcuts Renamed: $($totalStats.ShortcutsRenamed)" "SUCCESS"
    Write-Log "Total Shortcuts Skipped: $($totalStats.ShortcutsSkipped)" "INFO"
    Write-Log "Total Shortcuts Deleted: $($totalStats.ShortcutsDeleted)" "SUCCESS"
    Write-Log "Total Warnings: $($totalStats.Warnings)" $(if($totalStats.Warnings -gt 0){"WARN"}else{"SUCCESS"})
    Write-Log "Total Errors: $($totalStats.Errors)" $(if($totalStats.Errors -gt 0){"ERROR"}else{"SUCCESS"})
    Write-Log "" "INFO"
    Write-Log "=== Performance Summary ===" "INFO"
    Write-Log "Total user×library pairs: $totalPairs" "INFO"
    if($EnableRecursivePermissionPreCheck) {
        Write-Log "  Pre-allowed pairs (site/list/group recursion): $($preCheckStats.PreAllowedPairs)" "INFO"
        Write-Log "  Fallback pairs (effective permission API): $($preCheckStats.FallbackPairs)" "INFO"
        Write-Log "  Scope hits - site: $($preCheckStats.SiteScopeHits), list: $($preCheckStats.ListScopeHits)" "INFO"
        Write-Log "  Entra group cache - hits: $($preCheckStats.CacheHits), misses: $($preCheckStats.CacheMisses), entries: $($global:octo.EntraGroupTargetMemberCache.Count)" "INFO"
        Write-Log "  Recursion cutoffs (depth/node guards): $($preCheckStats.RecursionCutoffs)" "INFO"
    }    

    $phaseTimings['Phase 3: Applying shortcuts'] = $phaseStopwatch.Elapsed
    $scriptStopwatch.Stop()
    Write-Log "" "INFO"
    Write-Log "=== Timing Summary ===" "INFO"
    foreach($phase in $phaseTimings.GetEnumerator()) {
        $ts = $phase.Value
        $formatted = if($ts.TotalMinutes -ge 1) { '{0:0}m {1:0}s' -f [math]::Floor($ts.TotalMinutes), $ts.Seconds } else { '{0:0.0}s' -f $ts.TotalSeconds }
        Write-Log "  $($phase.Key): $formatted" "INFO"
    }
    $totalTs = $scriptStopwatch.Elapsed
    $totalFormatted = if($totalTs.TotalMinutes -ge 1) { '{0:0}m {1:0}s' -f [math]::Floor($totalTs.TotalMinutes), $totalTs.Seconds } else { '{0:0.0}s' -f $totalTs.TotalSeconds }
    Write-Log "  Total runtime: $totalFormatted" "SUCCESS"

    Write-Log "=== Script Completed ===" "SUCCESS"
    if(!$global:IsAzureAutomation) {
        Stop-Transcript
    }
} catch {
    Write-Log "Fatal error: $($_.Exception.Message)" "ERROR"
    Write-Log $_.ScriptStackTrace "ERROR"
    throw
}

#endregion
