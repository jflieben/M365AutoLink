<#
.SYNOPSIS
    Automatically links all Microsoft Teams and SharePoint sites to targeted users' OneDrive using Managed Identity (no user impersonation)

.DESCRIPTION
    This script authenticates to Microsoft Graph using a Managed Identity (Azure VM, Automation Account,
    Azure Function App, etc.) and creates OneDrive shortcuts for each target user's Microsoft Teams and 
    SharePoint sites, under an "AutoLink" (configurable) folder.

    The script first enumerates ALL sites in the tenant and their document libraries, applying
    include/exclude wildcard filters, file count limits, and archived/locked checks.
    Then, for each target user, it checks the user's effective permissions on each document library
    using the SharePoint getUserEffectivePermissions REST API, and creates shortcuts only for
    libraries the user actually has access to.

    The $MinimumPermissionLevel setting controls whether "View" (read-only) access is sufficient,
    or whether "Edit" (contribute) access is required before a shortcut is created.

.REQUIREMENTS
    - PowerShell 5.x or 7.x
    - Azure Managed Identity (System or User-assigned) on the compute resource running the script
      OR an Entra ID App Registration with client_id and a certificate (thumbprint or PFX file)
    - Microsoft Graph API application permissions granted to the managed identity or app registration

.MANAGED IDENTITY SETUP
    Grant the following APPLICATION permissions to your Managed Identity's service principal:

    Microsoft Graph (AppId: 00000003-0000-0000-c000-000000000000):
    - Sites.Read.All          - Read SharePoint site information
    - Files.ReadWrite.All     - Create shortcuts in users' OneDrive
    - User.Read.All           - Read user profiles for target user enumeration

    SharePoint (AppId: 00000003-0000-0ff1-ce00-000000000000):
    - Sites.FullControl.All   - Access SharePoint REST APIs for site metadata

    Example PowerShell to grant Graph permissions to a Managed Identity:
        
        $MIObjectId = "<Your-Managed-Identity-Object-Id>"
        $GraphAppId = "00000003-0000-0000-c000-000000000000"
        
        Connect-MgGraph -Scopes "AppRoleAssignment.ReadWrite.All"
        $graphSp = Get-MgServicePrincipal -Filter "appId eq '$GraphAppId'"
        $permissions = @("Sites.Read.All","Files.ReadWrite.All","User.Read.All")
        foreach($perm in $permissions){
            $role = $graphSp.AppRoles | Where-Object { $_.Value -eq $perm }
            New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $MIObjectId ``
                -PrincipalId $MIObjectId -ResourceId $graphSp.Id -AppRoleId $role.Id
        }

        # Grant SharePoint permissions
        $SPAppId = "00000003-0000-0ff1-ce00-000000000000"
        $spSp = Get-MgServicePrincipal -Filter "appId eq '$SPAppId'"
        $spRole = $spSp.AppRoles | Where-Object { $_.Value -eq "Sites.FullControl.All" }
        New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $MIObjectId ``
            -PrincipalId $MIObjectId -ResourceId $spSp.Id -AppRoleId $spRole.Id

.AUTHENTICATION FLOW
    Authentication methods are tried in order (first success wins):
    1. Azure Functions / App Service identity endpoint ($env:IDENTITY_ENDPOINT)
    2. Azure VM Instance Metadata Service (IMDS) at 169.254.169.254
    3. Az PowerShell module (Connect-AzAccount -Identity)
    4. Client Certificate (client_id + certificate thumbprint or PFX) - if $ClientId, $TenantId, and certificate are configured

    Managed Identity is preferred. Client certificate auth is only used as a fallback when MI is not available
    (e.g. running from a local workstation or non-Azure environment).

.NOTES
    Author: Jose Lieben
    Version: 2.0
    Date: 2026-02-27
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
$TargetMode = "UserList"

# When TargetMode = "Group", specify the group Object ID
$TargetGroupId = ""

# When TargetMode = "UserList", specify an array of UPNs
$TargetUsers = @(
    # "user1@contoso.com"
    # "user2@contoso.com"
)

#excluded sites will not be added a link if below pattern occurs in the site's URL. Use a * to match 1 or more characters
#the default list is recommended
#e.g. https://contoso.sharepoint.com/sites/HR*" would exclude all sites where the name starts with HR"
$excludedSitesByWildcard = @(
    "*/sites/Streamvideo*"
    "*/portals/personal/*"
    "*/sites/AllCompany*"
    "*/personal/*"
    "*/contentstorage/*"
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
)

#below variables can be used to filter based on the number of existing files in the target location before creating a link
$maxFileCount = 300000
$minFileCount = 0

# Permission level required before creating a shortcut for a user:
# "View" = create shortcut when user has View (read) or higher permissions
# "Edit" = create shortcut only when user has Edit (contribute) or higher permissions (view-only users are skipped)
$MinimumPermissionLevel = "View"

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
        $cleanedName = $Name.Trim() # fallback to original if cleaning produces empty string
    }
    return $cleanedName
}

function Get-AccessToken {    
    Param(
        [Parameter(Mandatory=$true)]
        [string]$resource
    )   

    # Return cached token if still valid (25 minute cache, tokens typically last 60+ minutes)
    if($global:octo.LCCachedTokens.$resource -and 
       $global:octo.LCCachedTokens.$resource.accessToken -and 
       $global:octo.LCCachedTokens.$resource.validFrom -gt (Get-Date).AddMinutes(-25)){
        return $global:octo.LCCachedTokens.$resource.accessToken
    }

    $token = $null
    $encodedResource = [System.Web.HttpUtility]::UrlEncode($resource)

    # Method 1: Client Certificate (client_id + certificate) fallback
    if(-not $token -and $ClientId -and $TenantId -and ($CertificateThumbprint -or $CertificatePath)){
        try {
            # Load the certificate
            $cert = $null
            if($CertificateThumbprint){
                $cert = Get-ChildItem -Path "Cert:\CurrentUser\My\$CertificateThumbprint" -ErrorAction SilentlyContinue
                if(-not $cert){
                    $cert = Get-ChildItem -Path "Cert:\LocalMachine\My\$CertificateThumbprint" -ErrorAction SilentlyContinue
                }
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

            # Build JWT header
            $certHash = [System.Convert]::ToBase64String($cert.GetCertHash()) -replace '\+','-' -replace '/','_' -replace '=',''
            $jwtHeader = @{ alg = "RS256"; typ = "JWT"; x5t = $certHash } | ConvertTo-Json -Compress
            $jwtHeaderB64 = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($jwtHeader)) -replace '\+','-' -replace '/','_' -replace '=',''

            # Build JWT payload
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

            # Sign the JWT
            $dataToSign = [System.Text.Encoding]::UTF8.GetBytes("$jwtHeaderB64.$jwtPayloadB64")
            $rsaKey = $cert.PrivateKey
            if(-not $rsaKey){
                # .NET Core / PS 7+ path
                $rsaKey = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
            }
            $signature = $rsaKey.SignData($dataToSign, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
            $signatureB64 = [System.Convert]::ToBase64String($signature) -replace '\+','-' -replace '/','_' -replace '=',''

            $clientAssertion = "$jwtHeaderB64.$jwtPayloadB64.$signatureB64"

            # Request token using client_credentials with client_assertion
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
            Write-Verbose "Token acquired via client certificate"
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
            Write-Verbose "Token acquired via App Service/Functions managed identity"
        } catch {
            Write-Warning "App Service MI endpoint failed: $($_.Exception.Message)"
        }
    }

    # Method 3: Azure VM Instance Metadata Service (IMDS)
    if(-not $token){
        try {
            $response = Invoke-RestMethod -Uri "http://169.254.169.254/metadata/identity/oauth2/token?api-version=2018-02-01&resource=$encodedResource" `
                -Headers @{Metadata="true"} -Method GET -ErrorAction Stop -Verbose:$false
            $token = $response.access_token
            Write-Verbose "Token acquired via Azure VM IMDS"
        } catch {
            Write-Verbose "IMDS endpoint not available: $($_.Exception.Message)"
        }
    }

    # Method 4: Az PowerShell module fallback (e.g. Azure Automation)
    if(-not $token){
        try {
            $null = Connect-AzAccount -Identity -ErrorAction Stop
            $tokenResponse = Get-AzAccessToken -ResourceUrl $resource -ErrorAction Stop
            if($tokenResponse.Token -is [string]){
                $token = $tokenResponse.Token
            }else{
                # PS 7.x+ may return SecureString
                $token = $tokenResponse.Token | ConvertFrom-SecureString -AsPlainText
            }
            Write-Verbose "Token acquired via Az PowerShell module"
        } catch {
            Write-Verbose "Az PowerShell MI not available: $($_.Exception.Message)"
        }
    }

    if(-not $token){
        $methods = "IMDS, App Service MI, Az PowerShell MI"
        if($ClientId -and $TenantId -and ($CertificateThumbprint -or $CertificatePath)){ $methods += ", Client Certificate" }
        throw "Failed to acquire token for resource '$resource'. Tried: $methods. Ensure this script is running with a managed identity or valid client certificate configured."
    }

    # Cache the token
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
        Param(
            [string]$apiUrl
        )

        # Auto-detect the token resource from the URL being called
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
            Write-Log "Failed to acquire token for '$tokenResource': $_" -Level "ERROR"
            throw
        }
        $headers = @{
            "Authorization" = "Bearer $token"
        }
        $headers['Accept-Language'] = "en-US"

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
                    $Data = $Null; $Data = (Invoke-RestMethod -Uri $nextURL -Method $Method -Headers $headers -Body $Body -ContentType $ContentType -Verbose:$False -ErrorAction Stop -UserAgent "ISV|LiebenConsultancy|M365AutoLink|2.0")
                    $attempts = $MaxAttempts
                }catch {
                    if($_.Exception.Message -like "*404*" -or $_.Exception.Message -like "*Request_ResourceNotFound*" -or $_.Exception.Message -like "*Resource*does not exist*" -or $_.Exception.Message -like "*403*" -or $_.Exception.StatusCode -in (401,403,"Unauthorized",404,"NotFound")){
                        Write-Debug "Not retrying 404 or 401"
                        $nextUrl = $Null
                        throw $_
                    }  
                    if ($attempts -ge $MaxAttempts) { 
                        Throw $_
                    }

                    $delay = 0
                    if ($_.Exception.Response.StatusCode -eq 429){
                        try {
                            $retryAfter = $_.Exception.Response.Headers.GetValues("Retry-After")
                            if ($retryAfter -and $retryAfter.Count -gt 0) {
                                $retryAfterValue = $retryAfter[0]
                                if ($retryAfterValue -match '^\d+$') {
                                    $delay = [int]$retryAfterValue
                                }
                            }
                        }catch {}
                    }
                    if($delay -le 0){
                        $delay = [math]::Pow(5, $attempts)
                    }
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
                        $headers = get-resourceHeaders -apiUrl $nextURL
                        $Data = $Null; $Data = (Invoke-RestMethod -Uri $nextURL -Method $Method -Headers $headers -ContentType $ContentType -Verbose:$False -ErrorAction Stop -UserAgent "ISV|LiebenConsultancy|M365AutoLink|2.0")
                        $attempts = $MaxAttempts
                    }catch {                 
                        if(($_.Exception -and $_.Exception.StatusCode -and $_.Exception.StatusCode -in (401,403,"Unauthorized",404,"NotFound")) -or ($_.Exception.Message -like "*404*" -or $_.Exception.Message -like "*403*" -or $_.Exception.Message -like "*Request_ResourceNotFound*" -or $_.Exception.Message -like "*Resource*does not exist*")){
                            Write-Debug "Not retrying $($_.Exception.StatusCode)"
                            $nextUrl = $Null
                            throw $_
                        }
              
                        if ($attempts -ge $MaxAttempts) { 
                            $nextURL = $null
                            Throw $_
                        }
                       
                        $delay = 0
                        if ($_.Exception.Response.StatusCode -eq 429){
                            try {
                                $retryAfter = $_.Exception.Response.Headers.GetValues("Retry-After")
                                if ($retryAfter -and $retryAfter.Count -gt 0) {
                                    $retryAfterValue = $retryAfter[0]
                                    if ($retryAfterValue -match '^\d+$') {
                                        $delay = [int]$retryAfterValue
                                    }
                                }
                            }catch {}
                        }
                        if($delay -le 0){
                            $delay = [math]::Pow(5, $attempts)
                        }
                        Start-Sleep -Seconds (1 + $delay)
                    }
                }

                if($nextURL -match "sharepoint"){
                    if($Data -and $Data.PSObject.TypeNames -notcontains "System.Management.Automation.PSCustomObject"){
                        try {
                            $Data = $Data | ConvertFrom-Json -AsHashtable
                        } catch {
                            # Fallback for PS5.1 duplicate key issues (e.g. 'Id' and 'ID')
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


#endregion

#region Main Script

try {
    $logDir = [System.IO.Path]::GetDirectoryName($global:octo.LogPath)
    if(!(Test-Path $logDir)){ New-Item -ItemType Directory -Path $logDir -Force | Out-Null }

    Start-Transcript -Path $global:octo.LogPath -Force

    Write-Log "=== M365AutoLink Centralized (Managed Identity) Started ===" "INFO"

    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Web")

    # Pre-populate the token cache via managed identity
    $token = Get-AccessToken -resource $global:octo.graphUrl
    Write-Log "Managed identity authentication successful" "SUCCESS"

    # Discover tenant SharePoint URL from root site
    $rootSite = New-GraphQuery -Uri "$($global:octo.graphUrl)/v1.0/sites/root" -Method GET
    $spHost = ([System.Uri]::new($rootSite.webUrl)).Host
    Write-Log "Tenant SharePoint URL: https://$spHost" "INFO"

    Write-Log "Will apply the following exclusion patterns:" "INFO"
    foreach($pattern in $excludedSitesByWildcard){
        Write-Log "  - $pattern" "INFO"
    }

    Write-Log "Will apply the following inclusion patterns (if defined):" "INFO"
    foreach($pattern in $includedSitesByWildcard){
        Write-Log "  - $pattern" "INFO"
    }

    # Get target users based on TargetMode
    $targetUserList = @()
    switch($TargetMode) {
        "Group" {
            if(-not $TargetGroupId) { throw "TargetGroupId must be specified when TargetMode is 'Group'" }
            Write-Log "Getting members of group '$TargetGroupId'..." "INFO"
            $members = New-GraphQuery -Uri "$($global:octo.graphUrl)/v1.0/groups/$TargetGroupId/members?`$select=id,userPrincipalName,displayName" -Method GET
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
        default {
            throw "Invalid TargetMode '$TargetMode'. Must be 'Group', 'UserList', or 'All'"
        }
    }

    Write-Log "Found $($targetUserList.Count) target users" "SUCCESS"
    if($targetUserList.Count -eq 0) { throw "No target users found for TargetMode '$TargetMode'" }

    # Pre-fetch all sites and their document libraries (done once, then permissions checked per user)
    Write-Log "Pre-fetching all tenant sites..." "INFO"
    $allTenantSites = @(New-GraphQuery -Uri "$($global:octo.graphUrl)/v1.0/sites/getAllSites?`$select=id,displayName,webUrl&`$top=1000" -Method GET)
    Write-Log "Found $($allTenantSites.Count) sites in tenant" "SUCCESS"

    Write-Log "Filtering sites and enumerating document libraries..." "INFO"
    $allSiteLibraries = @()
    $filteredSiteCount = 0

    foreach($site in $allTenantSites) {
        if($null -eq $site.webUrl) { continue }

        # Apply exclusion patterns
        $isExcluded = $false
        foreach($pattern in $excludedSitesByWildcard){
            $wildcardPattern = $pattern -replace "\*",".*"
            if($site.webUrl -match $wildcardPattern){
                $isExcluded = $true
                break
            }
        }
        if($isExcluded) { continue }

        # Apply inclusion patterns
        $isIncluded = $false
        foreach($pattern in $includedSitesByWildcard){
            $wildcardPattern = $pattern -replace "\*",".*"
            if($site.webUrl -match $wildcardPattern){
                $isIncluded = $true
                break
            }
        }
        if(-not $isIncluded) { continue }

        try {
            # Check if site is archived or read-only
            $siteDetails = New-GraphQuery -Uri "$($site.webUrl)/_api/site" -Method "GET" -MaxAttempts 1
            if($siteDetails.WriteLocked -or $siteDetails.ReadOnly){
                Write-Log "  Site '$($site.webUrl)' is locked or read-only, skipping..." "WARN"
                continue
            }

            # Get document libraries for this site
            $lists = @((New-GraphQuery -Uri "$($global:octo.graphUrl)/v1.0/sites/$($site.id)/lists" -Method "GET" -MaxAttempts 3) | Where-Object { $_.list.template -eq "documentLibrary" -and !$_.list.hidden })

            foreach($list in $lists) {
                # Get metadata for file count filtering
                $listMetaData = New-GraphQuery -Uri "$($site.webUrl)/_api/lists/GetById('$($list.id)')" -Method GET
                if($listMetaData.Hidden) { continue }

                if($listMetaData.ItemCount -gt $maxFileCount){
                    Write-Log "  $($site.webUrl) - '$($list.displayName)' exceeds $maxFileCount files, skipping..." "WARN"
                    continue
                }
                if($listMetaData.ItemCount -lt $minFileCount){
                    Write-Log "  $($site.webUrl) - '$($list.displayName)' below $minFileCount files, skipping..." "WARN"
                    continue
                }

                $allSiteLibraries += @{
                    siteId          = $site.id
                    siteDisplayName = $site.displayName
                    siteWebUrl      = $site.webUrl
                    listId          = $list.id
                    listDisplayName = $list.displayName
                    shortcutInfo    = @{
                        siteId           = $site.id.Split(',')[1]
                        siteUrl          = $site.webUrl
                        webId            = $site.id.Split(',')[2]
                        listId           = $list.id
                        listItemUniqueId = "root"
                    }
                }
                Write-Log "  Added '$($list.displayName)' from site '$($site.webUrl)' for potential linking" "INFO"
            }
            $filteredSiteCount++
        } catch {
            Write-Log "  Failed to process site '$($site.webUrl)': $($_.Exception.Message)" "WARN"
        }
    }

    Write-Log "Pre-cached $($allSiteLibraries.Count) document libraries across $filteredSiteCount sites" "SUCCESS"
    Write-Log "Minimum permission level for shortcuts: $MinimumPermissionLevel" "INFO"

    # Global counters
    $totalStats = @{
        UsersProcessed   = 0
        UsersSkipped     = 0
        UsersFailed      = 0
        ShortcutsCreated = 0
        ShortcutsRenamed = 0
        ShortcutsSkipped = 0
        ShortcutsDeleted = 0
        Errors           = 0
    }

    foreach($targetUser in $targetUserList) {
        $userId = $targetUser.id
        $userUPN = $targetUser.userPrincipalName

        Write-Log "========================================" "INFO"
        Write-Log "Processing user: $userUPN" "INFO"
        Write-Log "========================================" "INFO"

        try {
            # Check if user has a OneDrive provisioned
            try {
                $userDrive = New-GraphQuery -Uri "$($global:octo.graphUrl)/v1.0/users/$userId/drive" -Method GET -MaxAttempts 1
            } catch {
                Write-Log "  OneDrive not provisioned for this user, skipping..." "WARN"
                $totalStats.UsersSkipped++
                continue
            }

            # Check if target folder exists, create if not
            Write-Log "  Checking for '$FolderName' folder in OneDrive..." "INFO"
            $targetFolder = $null
    
            try {
                $targetFolder = New-GraphQuery -Uri "$($global:octo.graphUrl)/v1.0/users/$userId/drive/root:/$($FolderName)?`$expand=listItem" -Method "GET"
                Write-Log "  Folder '$FolderName' already exists" "INFO"
            } catch {
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

            #determine root onedrive url by list item:
            $urlParts = $targetFolder.listItem.webUrl -split "/personal/"
            $rootUrl = $urlParts[0]
            $userComponent = $urlParts[1].Split('/')[0]
            $libraryName = $urlParts[1].Split('/')[1]

            $docLibrary = (New-GraphQuery -Uri "$($global:octo.graphUrl)/v1.0/sites/$($targetFolder.parentReference.siteId)/lists" -Method "GET") | Where-Object { $_.list.template -eq "mySiteDocumentLibrary" -and !$_.list.hidden}

            $currentShortCuts = @()

            #retrieve current shortcuts
            Write-Log "Getting target info for all current shortcuts...." "INFO" 
            $folderContents = New-GraphQuery -Uri "$rootUrl/personal/$userComponent/_api/web/GetFolderByServerRelativeUrl('/personal/$userComponent/$libraryName/$FolderName')/Files?`$top=5000&`$format=json&`$expand=listItem" -Method GET
            
            #sometimes, e.g. when a library is changed to sync-blocked, onedrive changes it to a folder. These should be wiped as they would only confuse the user
            New-GraphQuery -Uri "$rootUrl/personal/$userComponent/_api/web/GetFolderByServerRelativeUrl('/personal/$userComponent/$libraryName/$FolderName')/Folders?`$top=5000&`$format=json&`$expand=listItem" -Method GET | ForEach-Object {
                if($_.UniqueId){
                    New-GraphQuery -Uri "$rootUrl/personal/$userComponent/_api/web/GetFolderById('$($_.UniqueId)')/DeleteObject()" -Method POST
                    Write-Log "Found and deleted an unexpected folder where only links should exist. Name: $($_.Name)" "ERROR" 
                }
            }
            
            foreach($shortCut in $folderContents){
                $shortCutMetaData = (New-GraphQuery -Uri "$rootUrl/personal/$userComponent/_api/web/lists('$($docLibrary.id)')/GetItemByUniqueId('$($shortCut.UniqueId)')?`$expand=FieldValuesAsText" -Method GET -MaxAttempts 1)
                $currentShortCuts += @{
                    "ID" = $shortCut.uniqueId
                    "Name" = $shortCutMetaData.FieldValuesAsText.FileLeafRef
                    "targetSiteId" = $shortCutMetaData.FieldValuesAsText.A2ODRemoteItemSiteId
                    "targetWebId" = $shortCutMetaData.FieldValuesAsText.A2ODRemoteItemWebId
                    "targetListId" = $shortCutMetaData.FieldValuesAsText.A2ODRemoteItemListId
                    "targetItemUniqueId" = $shortCutMetaData.FieldValuesAsText.A2ODRemoteItemUniqueId
                }
            }

            Write-Log " Currently has $($currentShortCuts.count) shortcuts" "INFO" 

            $desiredShortcuts = @()
            $successCount = 0
            $skipCount = 0
            $errorCount = 0

            # Check user's effective permissions on each pre-cached document library
            Write-Log "  Checking permissions on $($allSiteLibraries.Count) document libraries..." "INFO"
            foreach($lib in $allSiteLibraries) {
                try {
                    $effectivePermissions = New-GraphQuery -Uri "$($lib.siteWebUrl)/_api/web/lists/GetById('$($lib.listId)')/getUserEffectivePermissions(@u)?@u='i%3A0%23.f%7Cmembership%7C$($targetUser.userPrincipalName)'" -Method GET -MaxAttempts 1

                    # SharePoint BasePermissions bit flags (Low word):
                    #   ViewListItems  = 0x1 (bit 0) - read/view access
                    #   EditListItems  = 0x4 (bit 2) - edit/contribute access
                    $hasView = ($effectivePermissions.Low -band 0x1) -ne 0
                    $hasEdit = ($effectivePermissions.Low -band 0x4) -ne 0

                    if($MinimumPermissionLevel -eq "Edit") {
                        if(-not $hasEdit) {
                            if($hasView) {
                                Write-Log "    User has only view access on '$($lib.listDisplayName)' ($($lib.siteWebUrl)), skipping (Edit required)..." "WARN"
                            }
                            continue
                        }
                    } else {
                        # "View" mode - view or higher
                        if(-not $hasView) {
                            continue
                        }
                    }

                    $desiredShortcuts += @{
                        shortCut = $lib.shortcutInfo
                        siteName = $lib.siteDisplayName
                        listName = $lib.listDisplayName
                    }
                } catch {
                    # 401/403/404 = user has no access, which is expected for most libraries
                    if($_.Exception.Message -like "*401*" -or $_.Exception.Message -like "*403*" -or $_.Exception.Message -like "*404*" -or $_.Exception.StatusCode -in (401,403,"Unauthorized",404,"NotFound")) {
                        continue
                    }
                    Write-Log "    Failed to check permissions on '$($lib.listDisplayName)' ($($lib.siteWebUrl)): $($_.Exception.Message)" "WARN"
                }
            }

            Write-Log "  User has access to $($desiredShortcuts.Count) document libraries" "SUCCESS"

            foreach($desiredShortcut in $desiredShortcuts) {
                # Check if shortcut already exists
                $exists = $false
                foreach($existing in $currentShortCuts) {
                    if ($existing.targetSiteId -eq $desiredShortcut.shortcut.siteId -and $existing.targetWebId -eq $desiredShortcut.shortcut.webId -and $existing.targetListId -eq $desiredShortcut.shortcut.listId) {
                        $exists = $true
                        break
                    }
                }

                if ($exists) {
                    Write-Log "    Shortcut already exists for '$($desiredShortcut.shortcut.siteUrl)', skipping..." "WARN"
                    $skipCount++    
                    continue
                }
                try {
                    # Create the shortcut
                    $shortcutBody = @{
                        name = $($desiredShortcut.listName)
                        remoteItem = @{
                            sharepointIds = $desiredShortcut.shortcut
                        }
                        "@microsoft.graph.conflictBehavior" = "rename"
                    } | ConvertTo-Json -Depth 3
            
                    Write-Log "    Creating shortcut for '$($desiredShortcut.listName)' ($($desiredShortcut.shortcut.siteUrl))..." "INFO"
                    $newShortCut = $Null; $newShortCut = New-GraphQuery -Uri "$($global:octo.graphUrl)/v1.0/users/$userId/drive/items/$($targetFolder.id)/children" -Method POST -Body $shortcutBody
            
                    # Rename the shortcut if link name cleanup patterns apply
                    $cleanName = Get-CleanedShortcutName -Name $newShortCut.name
                    if($newShortCut.id -and $cleanName -ne $newShortCut.name -and $currentShortCuts.Name -notcontains $cleanName){
                        try {
                            $renameBody = @{ name = $cleanName } | ConvertTo-Json
                            $Null = New-GraphQuery -Uri "$($global:octo.graphUrl)/v1.0/users/$userId/drive/items/$($newShortCut.id)" -Method PATCH -Body $renameBody
                            Write-Log "    Renamed shortcut from '$($newShortCut.name)' to '$($cleanName)'" "INFO"
                        } catch {
                            Write-Log "    Failed to rename shortcut '$($newShortCut.name)': $($_.Exception.Message)" "WARN"
                        }
                    }            
                    # Small delay to avoid throttling
                    Start-Sleep -Milliseconds 500
                    Write-Log "    Successfully created shortcut for '$($desiredShortcut.shortcut.siteUrl)'" "SUCCESS"
                    $successCount++
                }catch{
                    Write-Log "    Failed to create shortcut for '$($desiredShortcut.shortcut.siteUrl)': $($_.Exception.Message)" "ERROR"
                    $errorCount++
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
                        try {
                            $renameBody = @{ name = $cleanedName } | ConvertTo-Json
                            $Null = New-GraphQuery -Uri "$($global:octo.graphUrl)/v1.0/users/$userId/drive/items/$($existing.ID)" -Method PATCH -Body $renameBody
                            Write-Log "    Renamed '$($existing.Name)' to '$cleanedName'" "SUCCESS"
                            Start-Sleep -Milliseconds 500                    
                            $renameCount++
                        } catch {
                            Write-Log "    Failed to rename '$($existing.Name)': $($_.Exception.Message)" "WARN"
                        }
                    }
                }
            }    

            # Delete shortcuts user should no longer have access to
            $deletedCount = 0
            foreach($existing in $currentShortCuts) {
                $shouldExist = $false
                foreach($desired in $desiredShortcuts) {
                    if ($existing.targetSiteId -eq $desired.shortcut.siteId -and $existing.targetWebId -eq $desired.shortcut.webId -and $existing.targetListId -eq $desired.shortcut.listId) {
                        $shouldExist = $true
                        break
                    }
                }

                if (-not $shouldExist) {
                    try {
                        Write-Log "    Deleting obsolete shortcut '$($existing.Name)' (ID: $($existing.ID))..." "INFO"
                        New-GraphQuery -Uri "$($global:octo.graphUrl)/v1.0/users/$userId/drive/items/$($existing.ID)" -Method DELETE
                        Start-Sleep -Milliseconds 500
                        $deletedCount++
                        Write-Log "    Successfully deleted obsolete shortcut" "SUCCESS"
                    } catch {
                        Write-Log "    Failed to delete obsolete shortcut '$($existing.Name)': $($_.Exception.Message)" "ERROR"
                        $errorCount++
                    }
                }
            }

            # Per-user summary
            Write-Log "  --- User Summary for $userUPN ---" "INFO"
            Write-Log "  Created: $successCount | Renamed: $renameCount | Skipped: $skipCount | Deleted: $deletedCount | Errors: $errorCount" "INFO"

            $totalStats.UsersProcessed++
            $totalStats.ShortcutsCreated += $successCount
            $totalStats.ShortcutsRenamed += $renameCount
            $totalStats.ShortcutsSkipped += $skipCount
            $totalStats.ShortcutsDeleted += $deletedCount
            $totalStats.Errors += $errorCount

        } catch {
            Write-Log "  Failed to process user '$userUPN': $($_.Exception.Message)" "ERROR"
            Write-Log "  $($_.ScriptStackTrace)" "ERROR"
            $totalStats.UsersFailed++
        }
    }

    # Global summary
    Write-Log "=== Global Summary ===" "INFO"
    Write-Log "Users Processed: $($totalStats.UsersProcessed)" "SUCCESS"
    Write-Log "Users Skipped (no OneDrive/sites): $($totalStats.UsersSkipped)" "WARN"
    Write-Log "Users Failed: $($totalStats.UsersFailed)" $(if($totalStats.UsersFailed -gt 0){"ERROR"}else{"SUCCESS"})
    Write-Log "Total Shortcuts Created: $($totalStats.ShortcutsCreated)" "SUCCESS"
    Write-Log "Total Shortcuts Renamed: $($totalStats.ShortcutsRenamed)" "SUCCESS"
    Write-Log "Total Shortcuts Skipped: $($totalStats.ShortcutsSkipped)" "INFO"
    Write-Log "Total Shortcuts Deleted: $($totalStats.ShortcutsDeleted)" "SUCCESS"
    Write-Log "Total Errors: $($totalStats.Errors)" $(if($totalStats.Errors -gt 0){"ERROR"}else{"SUCCESS"})
    
    Write-Log "=== Script Completed ===" "SUCCESS"
    Stop-Transcript
} catch {
    Write-Log "Fatal error: $($_.Exception.Message)" "ERROR"
    Write-Log $_.ScriptStackTrace "ERROR"
    throw
}

#endregion
