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

# --- Performance tuning (Optimized version) ---

# Initial concurrency for site enumeration AND permission check runspaces
# The adaptive throttle will scale this up/down based on 429 rate
$InitialParallelLimit = 10

# Absolute maximum concurrent threads (adaptive throttle ceiling)
$MaxParallelLimit = 25

# SharePoint REST $batch size (max sub-requests per batch call, Microsoft limit is 20)
$BatchSize = 20

# Cross-run cache settings
# Set to $true to cache permission results between runs (dramatically speeds up subsequent runs, but only works on persistent hosts)
$EnableCache = $false

if($EnableCache){
    # Cache location (default: alongside script). Set to a custom path if desired.
    $CachePath = Join-Path (Split-Path $MyInvocation.MyCommand.Path -Parent) "M365AutoLink_PermissionCache.json"
}

# Maximum age of cached permission results in hours. Entries older than this are re-checked.
$CacheMaxAgeHours = 24

# Set to $true to ignore the cache completely and do a full scan (useful after major permission changes)
$ForceFullScan = $false

# Dry-run mode: when $true, no shortcuts are created, deleted, or renamed
$DryRun = $false

##########END CONFIGURATION#############################


#base vars
$global:octo = @{}
$global:octo.LCCachedTokens = @{}
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
                    if($_.Exception.Message -like "*404*" -or $_.Exception.Message -like "*Request_ResourceNotFound*" -or $_.Exception.Message -like "*Resource*does not exist*" -or $_.Exception.Message -like "*403*" -or $_.Exception.StatusCode -in (401,403,"Unauthorized",404,"NotFound")){
                        $nextUrl = $Null
                        throw $_
                    }
                    $is429 = $_.Exception.Response.StatusCode -eq 429 -or $_.Exception.Message -like "*429*"
                    # 429s always retry indefinitely — do not count against MaxAttempts
                    if ($is429) { $attempts-- }
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
                    if($delay -le 0 -and $isTransientNetwork){
                        $delay = [math]::Min(10, 2 * $attempts)
                    }
                    if($delay -le 0){
                        $delay = [math]::Pow(5, $attempts)
                    }
                    $null = Write-Log "[WARNING] Transient error on attempt $attempts/$MaxAttempts, retrying in $($delay)s: $($_.Exception.Message)" -ForegroundColor Yellow
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
                        if(($_.Exception -and $_.Exception.StatusCode -and $_.Exception.StatusCode -in (401,403,"Unauthorized",404,"NotFound")) -or ($_.Exception.Message -like "*404*" -or $_.Exception.Message -like "*403*" -or $_.Exception.Message -like "*Request_ResourceNotFound*" -or $_.Exception.Message -like "*Resource*does not exist*")){
                            $nextUrl = $Null
                            throw $_
                        }

                        $is429 = $_.Exception.Response.StatusCode -eq 429 -or $_.Exception.Message -like "*429*"
                        # 429s always retry indefinitely — do not count against MaxAttempts
                        if ($is429) { $attempts-- }
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
                        if($delay -le 0 -and $isTransientNetwork){
                            $delay = [math]::Min(10, 2 * $attempts)
                        }
                        if($delay -le 0){
                            $delay = [math]::Pow(5, $attempts)
                        }
                        $null = Write-Log "[WARNING] Transient error on attempt $attempts/$MaxAttempts, retrying in $($delay)s: $($_.Exception.Message)" -ForegroundColor Yellow
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

#region Cache Functions

function Load-PermissionCache {
    param([string]$Path, [int]$MaxAgeHours)
    $cache = @{}
    if(-not (Test-Path $Path)) { return $cache }
    try {
        $rawJson = Get-Content -Path $Path -Raw -ErrorAction Stop
        $rawData = $rawJson | ConvertFrom-Json -ErrorAction Stop
        $cutoff = (Get-Date).AddHours(-$MaxAgeHours)
        foreach($prop in $rawData.PSObject.Properties) {
            $entry = $prop.Value
            $checkedAt = [datetime]::Parse($entry.checkedAt)
            if($checkedAt -gt $cutoff) {
                $cache[$prop.Name] = @{
                    hasAccess = [bool]$entry.hasAccess
                    checkedAt = $checkedAt
                }
            }
        }
        return $cache
    } catch {
        $null = Write-Log "  Warning: Could not load permission cache from '$Path': $($_.Exception.Message)" "WARN"
        return @{}
    }
}

function Save-PermissionCache {
    param([hashtable]$Cache, [string]$Path)
    try {
        $exportObj = [ordered]@{}
        foreach($key in $Cache.Keys) {
            $exportObj[$key] = @{
                hasAccess = $Cache[$key].hasAccess
                checkedAt = $Cache[$key].checkedAt.ToString("o")
            }
        }
        $exportObj | ConvertTo-Json -Depth 3 -Compress | Set-Content -Path $Path -Force -ErrorAction Stop
    } catch {
        $null = Write-Log "  Warning: Could not save permission cache to '$Path': $($_.Exception.Message)" "WARN"
    }
}

function Get-CacheKey {
    param([string]$UserUPN, [string]$SiteWebUrl, [string]$ListId)
    return "$($UserUPN)|$($SiteWebUrl)|$($ListId)"
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

                # Fallback: if segment splitting didn't isolate JSON, search the raw part with regex
                if($null -eq $jsonBody) {
                    if($part -match '(\{[^{}]*("High"|"Low"|"GetUserEffectivePermissions"|"error"|"value")[^{}]*\})') {
                        $jsonBody = $Matches[1]
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
            $null = Write-Log "[WARNING] Batch request to '$SiteUrl' failed on attempt $attempts/$MaxAttempts, retrying in $($delay)s: $($_.Exception.Message)" "WARN"
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
    Write-Log "Optimizations: SP Batch ($BatchSize/batch) | Adaptive Throttle ($InitialParallelLimit-$MaxParallelLimit) | Cache ($($EnableCache ? 'ON' : 'OFF'), ${CacheMaxAgeHours}h)" "INFO"
    if($DryRun) { Write-Log "*** DRY RUN MODE — no changes will be made ***" "WARN" }
    if($ForceFullScan) { Write-Log "*** FORCE FULL SCAN — cache ignored ***" "WARN" }

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

                $result.Libraries.Add(@{
                    siteId          = $siteInfo.id
                    siteDisplayName = $siteInfo.displayName
                    siteWebUrl      = $siteInfo.webUrl
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

        return $result
    }

    # Adaptive throttle: use a semaphore to control concurrency within the max pool
    $activeSemaphore = [System.Threading.SemaphoreSlim]::new($InitialParallelLimit, $MaxParallelLimit)
    $currentLimit = $InitialParallelLimit
    $throttleHitCount = 0
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
                        # Count 429 retries from Information stream (Write-Host output from New-GraphQuery retries)
                        foreach($info in $jobs[$i].PowerShell.Streams.Information) {
                            if($info.MessageData -like "*429*") { $throttleHitCount++; $phase1Total429s++ }
                        }
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
                    # Count 429 retries from Information stream (Write-Host output from New-GraphQuery retries)
                    foreach($info in $jobs[$i].PowerShell.Streams.Information) {
                        if($info.MessageData -like "*429*") { $throttleHitCount++; $phase1Total429s++ }
                    }
                } catch {
                    Write-Log "  Error collecting result for '$($jobs[$i].SiteUrl)': $($_.Exception.Message)" "WARN"
                }
                $jobs[$i].PowerShell.Dispose()
                $jobs.RemoveAt($i)

                # Adaptive throttle adjustment every 50 completed jobs
                if($completedSinceAdjust -ge 50) {
                    $throttleRate = $throttleHitCount / $completedSinceAdjust
                    if($throttleRate -gt 0.1 -and $currentLimit -gt 2) {
                        $newLimit = [math]::Max(2, [math]::Floor($currentLimit * 0.5))
                        $reduction = $currentLimit - $newLimit
                        for($r = 0; $r -lt $reduction; $r++) {
                            [void]$activeSemaphore.Wait(0)
                        }
                        $currentLimit = $newLimit
                        Write-Log "  Adaptive throttle: reducing to $currentLimit concurrent (429 rate: $([math]::Round($throttleRate * 100))%)" "WARN"
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
                    $completedSinceAdjust = 0
                }
            }
        }
        if($jobs.Count -gt 0) { Start-Sleep -Milliseconds 100 }
    }

    $pool.Close()
    $pool.Dispose()
    $activeSemaphore.Dispose()

    Write-Progress -Id 0 -Activity "Phase 1: Enumerating sites" -Completed

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
    Write-Log "Minimum permission level for shortcuts: $MinimumPermissionLevel" "INFO"

    $phaseTimings['Phase 1: Site & library discovery'] = $phaseStopwatch.Elapsed

    # ============================================================
    # PHASE 2: Site-centric permission matrix via SP $batch (P1+P2)
    # ============================================================
    $phaseStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Log "--- Phase 2: Building permission matrix (site-centric, batched) ---" "INFO"

    # Load cache if enabled
    $permissionCache = @{}
    if($EnableCache -and -not $ForceFullScan) {
        $permissionCache = Load-PermissionCache -Path $CachePath -MaxAgeHours $CacheMaxAgeHours
        Write-Log "Loaded $($permissionCache.Count) cached permission entries (max age: ${CacheMaxAgeHours}h)" "INFO"
    }

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
                            # Don't record — will be retried on next run via cache miss
                            $logMessages.Add(@{ Message = "429 throttled for user '$($sub.UserUPN)' on list '$($sub.ListId)' at '$siteWebUrl'"; Level = "WARN" })
                            $resultIndex++
                            continue
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
    $totalCacheHits = 0
    $totalBatchChecks = 0
    $phase2Total429s = 0
    $phase2CompletedJobs = 0
    $lastPhase2MetricsLog = 0

    $batchJobs = [System.Collections.Generic.List[hashtable]]::new()

    foreach($siteWebUrl in $siteLibGroups.Keys) {
        $siteIndex++
        $siteLibs = $siteLibGroups[$siteWebUrl]

        Write-Progress -Id 0 -Activity "Phase 2: Checking permissions" -Status "$siteIndex/$siteTotal sites | concurrent: $batchCurrentLimit | batch: $BatchSize | 429s: $phase2Total429s | checks: $totalBatchChecks" -PercentComplete ([math]::Min(100, [math]::Round(($siteIndex / $siteTotal) * 100)))

        # Build list of (user, library) pairs that need an actual API check
        $checkPairs = [System.Collections.Generic.List[hashtable]]::new()

        foreach($lib in $siteLibs) {
            foreach($user in $targetUserList) {
                $cacheKey = Get-CacheKey -UserUPN $user.userPrincipalName -SiteWebUrl $siteWebUrl -ListId $lib.listId

                # Check cross-run cache first (P4)
                if($EnableCache -and -not $ForceFullScan -and $permissionCache.ContainsKey($cacheKey)) {
                    $totalCacheHits++
                    if($permissionCache[$cacheKey].hasAccess) {
                        $permissionMatrix[$user.userPrincipalName].Add(@{
                            shortCut = $lib.shortcutInfo
                            siteName = $lib.siteDisplayName
                            listName = $lib.listDisplayName
                        })
                    }
                    continue
                }

                # Need actual API check
                $checkPairs.Add(@{
                    userUPN = $user.userPrincipalName
                    listId  = $lib.listId
                    libKey  = $cacheKey
                    libInfo = $lib
                })
            }
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
                                        $permissionCache[$pr.LibKey] = @{ hasAccess = $pr.HasAccess; checkedAt = Get-Date }
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
                                    # Update cache
                                    $permissionCache[$pr.LibKey] = @{ hasAccess = $pr.HasAccess; checkedAt = Get-Date }
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
                        if($info.MessageData -like "*429*") { $batchThrottleHits++; $phase2Total429s++ }
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
                                    $permissionCache[$pr.LibKey] = @{ hasAccess = $pr.HasAccess; checkedAt = Get-Date }
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
        if($batchJobs.Count -gt 0) { Start-Sleep -Milliseconds 100 }
    }

    $batchPool.Close()
    $batchPool.Dispose()
    $batchSemaphore.Dispose()

    Write-Progress -Id 0 -Activity "Phase 2: Checking permissions" -Completed

    # Save updated cache
    if($EnableCache) {
        Save-PermissionCache -Cache $permissionCache -Path $CachePath
        Write-Log "Saved $($permissionCache.Count) permission entries to cache" "SUCCESS"
    }

    $totalPairs = $targetUserList.Count * $allSiteLibraries.Count
    Write-Log "Permission matrix complete:" "SUCCESS"
    Write-Log "  Total user×library pairs: $totalPairs" "INFO"
    Write-Log "  Cache hits (skipped): $totalCacheHits" "INFO"
    Write-Log "  Actual API checks (batched): $totalBatchChecks" "INFO"
    $savingsPercent = if($totalPairs -gt 0) { [math]::Round(($totalCacheHits / $totalPairs) * 100, 1) } else { 0 }
    Write-Log "  API calls saved: $savingsPercent%" "SUCCESS"

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

            # Determine root OneDrive URL
            $urlParts = $targetFolder.listItem.webUrl -split "/personal/"
            $rootUrl = $urlParts[0]
            $userComponent = $urlParts[1].Split('/')[0]
            $libraryName = $urlParts[1].Split('/')[1]

            $docLibrary = (New-GraphQuery -Uri "$($global:octo.graphUrl)/v1.0/sites/$($targetFolder.parentReference.siteId)/lists" -Method "GET") | Where-Object { $_.list.template -eq "mySiteDocumentLibrary" -and !$_.list.hidden}

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

            foreach($shortCut in $folderContents){
                $shortCutMetaData = (New-GraphQuery -Uri "$rootUrl/personal/$userComponent/_api/web/lists('$($docLibrary.id)')/GetItemByUniqueId('$($shortCut.UniqueId)')?`$expand=FieldValuesAsText" -Method GET -MaxAttempts 5)
                $currentShortCuts += @{
                    "ID" = $shortCut.uniqueId
                    "Name" = $shortCutMetaData.FieldValuesAsText.FileLeafRef
                    "targetSiteId" = $shortCutMetaData.FieldValuesAsText.A2ODRemoteItemSiteId
                    "targetWebId" = $shortCutMetaData.FieldValuesAsText.A2ODRemoteItemWebId
                    "targetListId" = $shortCutMetaData.FieldValuesAsText.A2ODRemoteItemListId
                    "targetItemUniqueId" = $shortCutMetaData.FieldValuesAsText.A2ODRemoteItemUniqueId
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

            # Create missing shortcuts
            $scIndex = 0
            $scTotal = $desiredShortcuts.Count
            foreach($desiredShortcut in $desiredShortcuts) {
                $scIndex++
                Write-Progress -Id 1 -ParentId 0 -Activity "Creating shortcuts" -Status "Shortcut $scIndex / $scTotal" -PercentComplete ([math]::Min(100, [math]::Max(1, [math]::Round(($scIndex / [math]::Max(1,$scTotal)) * 100))))

                # Check if shortcut already exists
                $exists = $false
                foreach($existing in $currentShortCuts) {
                    if ($existing.targetSiteId -eq $desiredShortcut.shortcut.siteId -and $existing.targetWebId -eq $desiredShortcut.shortcut.webId -and $existing.targetListId -eq $desiredShortcut.shortcut.listId) {
                        $exists = $true
                        break
                    }
                }

                if ($exists) {
                    $skipCount++
                    continue
                }

                if($DryRun) {
                    Write-Log "    [DRY RUN] Would create shortcut for '$($desiredShortcut.shortcut.siteUrl)'" "INFO"
                    $successCount++
                    continue
                }

                try {
                    $shortcutBody = @{
                        name = $($desiredShortcut.listName)
                        remoteItem = @{
                            sharepointIds = $desiredShortcut.shortcut
                        }
                        "@microsoft.graph.conflictBehavior" = "rename"
                    } | ConvertTo-Json -Depth 3

                    Write-Log "    Creating shortcut ($($desiredShortcut.shortcut.siteUrl))..." "INFO"
                    $newShortCut = $Null; $newShortCut = New-GraphQuery -MaxAttempts 1 -Uri "$($global:octo.graphUrl)/v1.0/users/$userId/drive/items/$($targetFolder.id)/children" -Method POST -Body $shortcutBody

                    $cleanName = Get-CleanedShortcutName -Name $newShortCut.name
                    $i = 1
                    while($currentShortCuts.Name -contains $cleanName){
                        $cleanName = "$($cleanName)_$($i)"
                        $i++
                    }
                    if($newShortCut.id -and $cleanName -ne $newShortCut.name -and $currentShortCuts.Name -notcontains $cleanName){
                        try {
                            $renameBody = @{ name = $cleanName } | ConvertTo-Json
                            $Null = New-GraphQuery -MaxAttempts 1 -Uri "$($global:octo.graphUrl)/v1.0/users/$userId/drive/items/$($newShortCut.id)" -Method PATCH -Body $renameBody
                            Write-Log "    Renamed '$($newShortCut.name)' → '$cleanName'" "INFO"
                        } catch {
                            Write-Log "    Failed to rename '$($newShortCut.name)': $($_.Exception.Message)" "WARN"
                        }
                    }
                    Start-Sleep -Milliseconds 200
                    Write-Log "    Created shortcut for '$($desiredShortcut.shortcut.siteUrl)'" "SUCCESS"
                    $successCount++
                }catch{
                    if($_.Exception.Message -like '*descendant shortcut exists*' -or $_.Exception.Message -like '*shortcut already exists*') {
                        Write-Log "    Existing shortcut conflict for '$($desiredShortcut.shortcut.siteUrl)': $($_.Exception.Message)" "WARN"
                        $existingShortcutConflicts.Add(@{ User = $userUPN; Url = $desiredShortcut.shortcut.siteUrl })
                        $warnCount++
                    } else {
                        Write-Log "    Failed to create shortcut for '$($desiredShortcut.shortcut.siteUrl)': $($_.Exception.Message)" "ERROR"
                        $errorCount++
                    }
                }
            }

            # Rename existing shortcuts if link name cleanup patterns apply
            $renameCount = 0
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
                            $Null = New-GraphQuery -MaxAttempts 1 -Uri "$($global:octo.graphUrl)/v1.0/users/$userId/drive/items/$($existing.ID)" -Method PATCH -Body $renameBody
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
            $delIndex = 0
            $delTotal = $currentShortCuts.Count
            foreach($existing in $currentShortCuts) {
                $delIndex++
                if($delTotal -gt 0) { Write-Progress -Id 1 -ParentId 0 -Activity "Cleaning obsolete shortcuts" -Status "$delIndex / $delTotal" -PercentComplete ([math]::Min(100, [math]::Round(($delIndex / $delTotal) * 100))) }
                $shouldExist = $false
                foreach($desired in $desiredShortcuts) {
                    if ($existing.targetSiteId -eq $desired.shortcut.siteId -and $existing.targetWebId -eq $desired.shortcut.webId -and $existing.targetListId -eq $desired.shortcut.listId) {
                        $shouldExist = $true
                        break
                    }
                }

                if (-not $shouldExist) {
                    if($DryRun) {
                        Write-Log "    [DRY RUN] Would delete obsolete shortcut '$($existing.Name)'" "INFO"
                        $deletedCount++
                        continue
                    }
                    try {
                        Write-Log "    Deleting obsolete shortcut '$($existing.Name)'..." "INFO"
                        New-GraphQuery -MaxAttempts 2 -Uri "$($global:octo.graphUrl)/v1.0/users/$userId/drive/items/$($existing.ID)" -Method DELETE
                        Start-Sleep -Milliseconds 200
                        $deletedCount++
                        Write-Log "    Deleted obsolete shortcut" "SUCCESS"
                    } catch {
                        Write-Log "    Failed to delete '$($existing.Name)': $($_.Exception.Message)" "ERROR"
                        $errorCount++
                    }
                }
            }

            if($delTotal -gt 0) { Write-Progress -Id 1 -ParentId 0 -Activity "Cleaning obsolete shortcuts" -Completed }

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
    Write-Log "Cache hits: $totalCacheHits | API checks: $totalBatchChecks" "INFO"
    Write-Log "API calls saved: $savingsPercent%" "SUCCESS"

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
