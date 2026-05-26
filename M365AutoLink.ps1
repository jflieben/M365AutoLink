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
    Version: 1.1
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

##########END CONFIGURATION#############################

#base vars
$global:octo = @{}
$global:octo.LCRefreshToken = $Null
$global:octo.LCCachedTokens = @{}
$global:octo.LCClientId = $ClientID
$global:octo.TokenCachePath = "$env:APPDATA\M365AutoLink\RefreshToken.xml"
$global:octo.LogPath = "$env:APPDATA\M365AutoLink\lastRun.log"

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

function get-AccessToken{    
    Param(
        [Parameter(Mandatory=$true)]$resource,
        [Switch]$returnHeader
    )   

    # Try to load refresh token from disk (completely silent)
    if(!$global:octo.LCRefreshToken -and (Test-Path $global:octo.TokenCachePath)){
        try {
            $global:octo.LCRefreshToken = (Import-Clixml $global:octo.TokenCachePath).GetNetworkCredential().Password
            Write-Verbose "Loaded refresh token from local storage"
        } catch {
            Write-Warning "Failed to load cached token, proceeding to authentication..."
            Remove-Item $global:octo.TokenCachePath -ErrorAction SilentlyContinue
        }     
    }

    # Use cached refresh token (completely silent)
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

    # Browser-based authentication (sometimes interactive)
    if(!$global:octo.LCRefreshToken){
        $global:octo.LCRefreshToken = Get-BrowserAuthorizationCode
    }

    # Now get the access token for the requested resource
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
            $siteWebUrl = [string]$map.SPWebUrl
            if([string]::IsNullOrWhiteSpace($siteWebUrl)) {
                $siteWebUrl = [string]$map.SPSiteUrl
            }
            if([string]::IsNullOrWhiteSpace($siteWebUrl)) {
                $siteWebUrl = [string]$map.Path
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
                listPath = [string]$map.Path
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
    $serverRelativeListPath = $null

    if(-not [string]::IsNullOrWhiteSpace($listPath)) {
        try {
            $listPathUri = [System.Uri]::new($listPath)
            $serverRelativeListPath = $listPathUri.AbsolutePath
        } catch {
            $serverRelativeListPath = $null
        }
    }

    try {
        return New-GraphQuery -resource $global:octo.sharepointUrl -Uri "$PrimarySiteUrl/_api/lists/GetById('$listId')" -Method GET
    } catch {
        $isNotFound = $_.Exception.Message -like "*404*" -or $_.Exception.Message -like "*Not Found*"
        if(-not $isNotFound -or [string]::IsNullOrWhiteSpace($serverRelativeListPath)) {
            throw
        }
    }

    $candidateBaseUrls = [System.Collections.Generic.List[string]]::new()
    $candidateBaseUrls.Add($PrimarySiteUrl)
    if(-not [string]::IsNullOrWhiteSpace($siteCollectionUrl)) {
        $candidateBaseUrls.Add($siteCollectionUrl.TrimEnd('/'))
    }

    foreach($baseUrl in $candidateBaseUrls | Select-Object -Unique) {
        try {
            $escapedPath = $serverRelativeListPath.Replace("'", "''")
            return New-GraphQuery -resource $global:octo.sharepointUrl -Uri "$baseUrl/_api/web/GetList('$escapedPath')" -Method GET
        } catch {
            $isNotFound = $_.Exception.Message -like "*404*" -or $_.Exception.Message -like "*Not Found*"
            if(-not $isNotFound) {
                throw
            }
        }
    }

    throw "List metadata lookup failed for list '$listId' using both GetById and GetList fallbacks"
}


#endregion

#region Main Script

try {
    $logDir = [System.IO.Path]::GetDirectoryName($global:octo.LogPath)
    if(!(Test-Path $logDir)){ New-Item -ItemType Directory -Path $logDir -Force | Out-Null }

    Start-Transcript -Path $global:octo.LogPath -Force

    Write-Log "=== M365AutoLink v1.1 Started ===" "INFO"
    if($DryRun) { Write-Log "*** DRY RUN MODE — no changes will be made ***" "WARN" }

    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Web")

    # Pre populate the token cache
    $token = Get-AccessToken -resource $global:octo.graphUrl
    
    # Check if target folder exists, create if not
    Write-Log "Checking for '$FolderName' folder in OneDrive..." "INFO"
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
    $discoveredLibraries = @(Get-SharePointDocumentLibrariesFromSearch -SearchRootUrl $searchRootUrl)

    if(!$discoveredLibraries -or $discoveredLibraries.Count -eq 0) {
        Write-Log "No searchable document libraries found for this user" "WARN"
        Exit 0
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

    # Process each site
    $successCount = 0
    $skipCount = 0
    $errorCount = 0
    
    Write-Log "Evaluating discovered libraries against site and library rules..." "INFO"
    $siteEvaluationCache = @{}
    $seenLibraryKeys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    foreach($library in $discoveredLibraries) {
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

        try {
            $listMetaData = Get-ListMetadataWithFallback -Library $library -PrimarySiteUrl $siteUrl

            $listDisplayName = [string]$listMetaData.Title
            if([string]::IsNullOrWhiteSpace($listDisplayName)) {
                $listDisplayName = [string]$library.listName
            }

            if($listMetaData.Hidden){
                Write-Log "  $listDisplayName is hidden, skipping..." "WARN"
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
            }
        }catch{
            Write-Log "  Failed to evaluate library '$($library.listName)' on '$siteUrl': $($_.Exception.Message)" "ERROR"
            $errorCount++
            continue
        }
    }

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
            Write-Log "  Shortcut already exists for '$($desiredShortcut.shortcut.siteUrl)', skipping..." "WARN"
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
            $safeShortcutName = Get-SafeDriveItemName -Name $desiredShortcut.listName
            if($safeShortcutName -ne $desiredShortcut.listName) {
                Write-Log "  Sanitized shortcut name '$($desiredShortcut.listName)' to '$safeShortcutName'" "WARN"
            }

            $shortcutBody = @{
                name = $safeShortcutName
                remoteItem = @{
                    sharepointIds = $desiredShortcut.shortcut
                }
                "@microsoft.graph.conflictBehavior" = "rename"
            } | ConvertTo-Json -Depth 3
            
            Write-Log "  Creating shortcut..." "INFO"
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
            if($newShortCut.id -and $cleanName -ne $newShortCut.name -and $currentShortCuts.Name -notcontains $cleanName){
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
            $successCount++
        }catch{
            Write-Log "  Failed to create shortcut for '$($desiredShortcut.shortcut.siteUrl)': $($_.Exception.Message)" "ERROR"
            $errorCount++
        }
    }
    
    # Rename existing shortcuts if link name cleanup patterns apply
    $renameCount = 0
    if($linkNameReplacements.Count -gt 0) {
        Write-Log "Checking existing shortcuts for name cleanup..." "INFO"
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
    foreach($existing in $currentShortCuts) {
        $shouldExist = $false
        foreach($desired in $desiredShortcuts) {
            if ($existing.targetSiteId -eq $desired.shortcut.siteId -and $existing.targetWebId -eq $desired.shortcut.webId -and $existing.targetListId -eq $desired.shortcut.listId) {
                $shouldExist = $true
                break
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
    
    Write-Log "=== Script Completed ===" "SUCCESS"
    Stop-Transcript
} catch {
    Write-Log "Fatal error: $($_.Exception.Message)" "ERROR"
    Write-Log $_.ScriptStackTrace "ERROR"
    throw
}

#endregion
