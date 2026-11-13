#
# Test-ArchiveMailbox.ps1
#
# By David Barrett, Microsoft Ltd. Use at your own risk. No warranties are given.
#
<#
.SYNOPSIS
Tests listing and getting folders and messages in an Exchange Online archive.

.DESCRIPTION
Starts from Mailbox (user principal name or Entra user ID). Discovers the main
archive through GET /beta/users/{Mailbox}/settings/exchange and its
inPlaceArchiveMailboxId property, including the archive and tenant GUIDs.
Then uses /users/MBX:{archive-guid}@{tenant-guid}/mailFolders and messages, not
the primary mailbox's well-known "archive" folder or the import/export APIs.
Enumerates every folder page and child folder, gets each folder individually,
and lists and gets up to ItemsPerFolder messages per folder (default 10).
Auxiliary archives are reached through HTTP 308 Location headers, or an explicit
ErrorArchiveFolderMovedPermanently error containing a URL. Redirects and next
links must remain on HTTPS graph.microsoft.com in the same tenant's MBX mail API.
Redirect loops and repeated pagination links are errors. No writes or automatic
retries are performed against Graph.

The default mail API version is v1.0; use -ApiVersion beta to compare versions.
Discovery always uses beta, where inPlaceArchiveMailboxId is exposed.
This tests observed service behavior, not a guarantee of archive API support.
An autoexpanded archive with populated expanded folders is needed to exercise
auxiliary routing. A run with no redirects does not validate that scenario.

A unique OutputPath is created for each run. Requests.jsonl records UTC timing,
HTTP status, URLs, response headers, correlation IDs and response bodies, without
authorization headers. Results.jsonl records folders, messages and failures.
Treat these files as mailbox data. -TraceGraphCalls uses TraceInvokeRestMethod
and its standard script-name .trace file, separate from OutputPath.
A partial failure is retained and ends with a terminating error.
Empty folders have no individual message GET to test.

.PREREQUISITES
Register and consent an application for Microsoft Graph with User.Read.All for
mailbox discovery, plus Mail.ReadBasic.All (application) or Mail.ReadBasic
(delegated) for the selected mail fields. Mail.Read is also suitable for mail.
Delegated authentication uses the shared browser flow. Supply TenantId when
using application authentication. No Exchange Online PowerShell connection or
manually supplied archive GUID is required. The script does not discover credentials.

.EXAMPLE
if ($secretKey -isnot [System.Security.SecureString]) { throw 'secretKey must be a SecureString.' }
.\Test-ArchiveMailbox.ps1 -AppId $clientId -TenantId $tenantId `
    -AppSecretKey $secretKey `
    -Mailbox $Mailbox -ItemsPerFolder 10 -TraceGraphCalls

Run in the intended profile-loaded lab session. The shared authentication code
accepts String or SecureString secrets and converts only for the token request.
Confirm the lab tenant and archive are the intended target before running.

.EXAMPLE
.\Test-ArchiveMailbox.ps1 -AppId $clientId -TenantId $tenantId `
    -Mailbox $Mailbox -ApiVersion beta -ItemsPerFolder 5

.LINK
https://learn.microsoft.com/en-us/graph/handle-archive-mailbox-redirects
.LINK
https://learn.microsoft.com/en-us/graph/api/user-list-mailfolders
.LINK
https://learn.microsoft.com/en-us/graph/api/usersettings-list-exchange?view=graph-rest-beta
#>

[CmdletBinding()]
param (
<# Auth.ps1 %PARAMS_START% #>
    [Parameter(Mandatory=$False,HelpMessage="Application Id (obtained when registering the application in Azure AD).")]
    [ValidateNotNullOrEmpty()]
    [string]$AppId,

    [Parameter(Mandatory=$False,HelpMessage="Application secret key as a String or SecureString (obtained when registering the application in Azure AD). If not specified, delegate permissions are assumed.")]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({
        if ($_ -isnot [string] -and $_ -isnot [System.Security.SecureString])
        {
            throw "AppSecretKey must be a String or SecureString."
        }
        if ($_.Length -eq 0) { throw "AppSecretKey must not be empty." }
        return $true
    })]
    [object]$AppSecretKey,

    [Parameter(Mandatory=$False,HelpMessage="Redirect Url (specified when registering the application in Azure AD, use localhost).  Required for delegate permissions.")]
    [ValidateNotNullOrEmpty()]
    [string]$RedirectUrl = "http://localhost/code",

    [Parameter(Mandatory=$False,HelpMessage="Tenant Id.  Default is common, which requires the application to be registered for multi-tenant use (and consented to in target tenant).")]
    [ValidateNotNullOrEmpty()]
    [string]$TenantId = "common",

    [Parameter(Mandatory=$False,HelpMessage="User (me to access as authenticating user when using delegate permissions).")]
    [ValidateNotNullOrEmpty()]
    [string]$User = "me",

    [Parameter(Mandatory=$False,HelpMessage="Timeout for HTTP requests (default is 30 seconds, which is same as Microsoft Graph endpoint).  Only works for Powershell 7.4+ (prior to that, timeout is infinite).")]
    [ValidateRange(0,300)]
    [int]$ConnectionTimeout = 30,    
<# Auth.ps1 %PARAMS_END% #>

<# Logging.ps1 %PARAMS_START% #>
    [Parameter(Mandatory=$False,HelpMessage="If included, activity is logged to a file (same as script name with .log appended).")]	
    [switch]$LogToFile,
<# Logging.ps1 %PARAMS_END% #>

<# Graph.ps1 %PARAMS_START% #>
    [Parameter(Mandatory=$False,HelpMessage="If specified, all Graph calls will be logged to file (same name as script, ending in .trace).")]
    [switch]$TraceGraphCalls,
<# Graph.ps1 %PARAMS_END% #>

    [Parameter(Mandatory=$True,HelpMessage="Mailbox owner's user principal name or Entra user ID; its archive is discovered automatically.")]
    [ValidateNotNullOrEmpty()]
    [string]$Mailbox,

    [Parameter(Mandatory=$False,HelpMessage="Maximum number of messages to list and get per folder (default 10).")]
    [ValidateRange(1,1000)]
    [int]$ItemsPerFolder = 10,

    [Parameter(Mandatory=$False,HelpMessage="Graph API version to test (default v1.0).")]
    [ValidateSet("v1.0","beta")]
    [string]$ApiVersion = "v1.0",

    [Parameter(Mandatory=$False,HelpMessage="Include hidden folders when enumerating the archive hierarchy.")]
    [switch]$IncludeHiddenFolders,

    [Parameter(Mandatory=$False,HelpMessage="Maximum redirects for a single request (default 10).")]
    [ValidateRange(1,50)]
    [int]$MaxRedirects = 10,

    [Parameter(Mandatory=$False,HelpMessage="New directory for request and result evidence; must not already exist.")]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath = (Join-Path (Get-Location) ("Test-ArchiveMailbox-{0}-{1}" -f [DateTime]::UtcNow.ToString("yyyyMMddTHHmmssfffZ"), [Guid]::NewGuid().ToString("N")))
)

$script:ScriptVersion = "1.0.4"
if ([string]::IsNullOrWhiteSpace($AppId)) { throw "AppId is required." }
if ([string]::IsNullOrWhiteSpace($Mailbox)) { throw "Mailbox is required." }
if ($null -ne $AppSecretKey -and $TenantId -eq "common")
{
    throw "Supply TenantId for application authentication."
}
$script:archiveMailboxId = $null
$script:archiveTenant = $null
$script:archiveDiscoveryUrl = "https://graph.microsoft.com/beta/users/$([Uri]::EscapeDataString($Mailbox))/settings/exchange"
$User = $Mailbox
if ($ConnectionTimeout -le 0) { throw "ConnectionTimeout must be greater than zero for this test." }
$OutputPath = (New-Item -Path $OutputPath -ItemType Directory -ErrorAction Stop).FullName
$script:requestFile = Join-Path $OutputPath "Requests.jsonl"
$script:resultFile = Join-Path $OutputPath "Results.jsonl"

<# Logging.ps1 %FUNCTIONS_START% #>
if ($LogToFile) {
    $script:logFile = "$($MyInvocation.InvocationName).log"
}

Function UpdateDetailsWithCallingMethod([string]$Details)
{
    # Update the log message with details of the function that logged it
    $timeInfo = "$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToLongTimeString())"
    $callingFunction = (Get-PSCallStack)[2].Command # The function we are interested in will always be frame 2 on the stack
    
    # Only add function name if it's actually a function (not a script file)
    if (![String]::IsNullOrEmpty($callingFunction) -and $callingFunction -notlike "*.ps1")
    {
        return "$timeInfo [$callingFunction] $Details"
    }
    return "$timeInfo $Details"
}

Function LogToFile([string]$LogEntry)
{
	if ( [String]::IsNullOrEmpty($script:logFile) ) { return }
	$LogEntry | Out-File $script:logFile -Append
}

Function LogVerbose([string]$Details)
{
    $Details = UpdateDetailsWithCallingMethod( $Details )
    Write-Verbose $Details
    if ( $VerbosePreference -eq "SilentlyContinue") { return }
    LogToFile $Details
}

Function Log([string]$Details, [System.ConsoleColor]$ForegroundColor = "Gray")
{
    $Details = UpdateDetailsWithCallingMethod( $Details )
    Write-Host $Details -ForegroundColor $ForegroundColor
    LogToFile $Details
}
Log "$($MyInvocation.MyCommand.Name) version $($script:ScriptVersion) starting"
<# Logging.ps1 %FUNCTIONS_END% #>

<# Auth.ps1 %FUNCTIONS_START% #>
function GetAppSecretKeyValue([object]$Secret)
{
    if ($Secret -is [string] -and $Secret.Length -gt 0) { return $Secret }
    if ($Secret -isnot [System.Security.SecureString] -or $Secret.Length -eq 0)
    {
        throw "AppSecretKey must be a nonempty String or SecureString."
    }

    $secretPointer = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Secret)
    try
    {
        return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($secretPointer)
    }
    finally
    {
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($secretPointer)
    }
}

function GetAppAuthToken()
{
    # Acquire token for application permissions
    $body = @{grant_type="client_credentials";scope="https://graph.microsoft.com/.default";client_id=$AppId}
    try
    {
        # Convert only at the token request boundary; never replace the caller's secret.
        $body.client_secret = GetAppSecretKeyValue $AppSecretKey
        $script:oauth = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token -Body $body
        $script:token_expires = (Get-Date).AddSeconds($oauth.expires_in)
    }
    catch
    {
        Write-Host "Failed to obtain OAuth token" -ForegroundColor Red
        $Error[0]
        exit # Failed to obtain a token
    }
    finally
    {
        $body.Remove("client_secret")
    }
    $script:graphBaseUrl = "https://graph.microsoft.com/v1.0/users/$User/"
}

function RenewOAuthToken()
{
    # If using app flow, we just obtain a new token

    if ($null -ne $AppSecretKey)
    {
        GetAppAuthToken
        return
    }

    # If using delegate flow, we need to refresh the token
    $body = @{grant_type="refresh_token";scope="https://graph.microsoft.com/.default";client_id=$AppId;refresh_token=$oauth.refresh_token}
    try
    {
        $script:oauth = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token -Body $body
        $script:token_expires = (Get-Date).AddSeconds($oauth.expires_in)
    }
    catch
    {
        Log "Failed to renew OAuth token" Red
        exit # Failed to renew the token
    }
    Log "Successfully renewed OAuth token"
}

function UpdateHeaders()
{
    # Ensure the OAuth token is renewed if it is about to expire
    if ($script:token_expires -lt (Get-Date).AddMinutes(5))
    {
        RenewOAuthToken
    }

    # Prepare the request headers
    if ($null -eq $script:headers -or $script:headers.Count -eq 0)
    {
        $script:headers = @{
            'Authorization'  = "$($oauth.token_type) $($oauth.access_token)";
            'Content-Type'   = 'application/json';
        }    
    }
    $script:headers["Authorization"] = "$($oauth.token_type) $($oauth.access_token)"
}



## Authentication

if ($null -eq $AppSecretKey)
{
    # Acquire auth code (needed to request token)
    $authUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/authorize?client_id=$AppId&response_type=code&redirect_uri=$RedirectUrl&response_mode=query&scope=openid%20profile%20email%20offline_access%20"
    Start-Process $authUrl

    Write-Host "Please complete log-in via the web browser, and then paste the redirect URL (including auth code) here to continue" -ForegroundColor Green
    $authcode = Read-Host "Auth code"
    $codeStart = $authcode.IndexOf("?code=")
    if ($codeStart -gt 0)
    {
        $authcode = $authcode.Substring($codeStart+6)
    }
    $codeEnd = $authcode.IndexOf("&session_state=")
    if ($codeEnd -gt 0)
    {
        $authcode = $authcode.Substring(0, $codeEnd)
    }
    Write-Verbose "Using auth code: $authcode"

    # Acquire token (using the auth code)
    $body = @{grant_type="authorization_code";scope="https://graph.microsoft.com/.default";client_id=$AppId;code=$authcode;redirect_uri=$RedirectUrl}
    try
    {
        $oauth = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token -Body $body
        $script:token_expires = (Get-Date).AddSeconds($oauth.expires_in)
    }
    catch
    {
        Log "Failed to obtain OAuth token" Red
        $Error[0]
        exit # Failed to obtain a token
    }
    $graphBaseUrl = "https://graph.microsoft.com/v1.0/me/"
}
else
{
    GetAppAuthToken
}
Log "Successfully obtained OAuth token"
<# Auth.ps1 %FUNCTIONS_END% #>

<# Graph.ps1 %FUNCTIONS_START% #>

$script:traceFile = "$($MyInvocation.InvocationName).trace"
function GetGraphErrorResponse([System.Management.Automation.ErrorRecord]$ErrorRecord)
{
    $response = $ErrorRecord.Exception.Response
    if ($null -eq $response) { return $null }
    $responseHeaders = @{}
    if ($response.Headers -is [System.Net.WebHeaderCollection])
    {
        foreach ($name in $response.Headers.AllKeys) { $responseHeaders[$name] = $response.Headers[$name] }
    }
    else
    {
        foreach ($header in $response.Headers.GetEnumerator()) { $responseHeaders[$header.Key] = $header.Value -join ", " }
    }
    if ($null -ne $response.Content -and $null -ne $response.Content.Headers)
    {
        foreach ($header in $response.Content.Headers.GetEnumerator()) { $responseHeaders[$header.Key] = $header.Value -join ", " }
    }
    $content = $ErrorRecord.ErrorDetails.Message
    if ([string]::IsNullOrEmpty($content))
    {
        if ($null -ne $response.Content)
        {
            try { $content = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult() }
            catch { $content = "" }
        }
        elseif ($response -is [System.Net.HttpWebResponse])
        {
            try
            {
                $reader = [System.IO.StreamReader]::new($response.GetResponseStream())
                try { $content = $reader.ReadToEnd() }
                finally { $reader.Dispose() }
            }
            catch { $content = "" }
        }
    }
    $status = [int]$response.StatusCode
    $description = $response.ReasonPhrase
    if ([string]::IsNullOrEmpty($description)) { $description = $response.StatusDescription }
    $rawContent = "HTTP $status $description`r`n"
    foreach ($name in $responseHeaders.Keys) { $rawContent += "${name}: $($responseHeaders[$name])`r`n" }
    $rawContent += "`r`n$content"
    return [PSCustomObject]@{ StatusCode = $status; Headers = $responseHeaders; Content = $content; RawContent = $rawContent }
}

# Full responses and caller-managed redirects are opt-in; existing calls still return content.
function TraceInvokeRestMethod($method, $url, $body, [switch]$ReturnResponse, [switch]$NoRedirect)
{
    UpdateHeaders
    Write-Verbose "Calling URL: $url"
    if ($TraceGraphCalls)
    {
        Add-Content -Path $script:traceFile -Value "$method $url" -Encoding UTF8
    }
    try
    {
        $result = $null
        $invokeParameters = @{
            Method = $method
            Uri = $url
            Headers = $script:headers
            ContentType = "application/json"
            MaximumRedirection = 0  # Disable automatic redirects so we can handle them manually
            UseBasicParsing = $true
        }
        if ($PSVersionTable.PSVersion -ge [version]7.4)
        {
            # PowerShell 7.4+ supports ConnectionTimeoutSeconds
            $invokeParameters.ConnectionTimeoutSeconds = $ConnectionTimeout
        }
        elseif ($ReturnResponse)
        {
            $invokeParameters.TimeoutSec = $ConnectionTimeout
        }

        # Write request headers, excluding Authorization
        if ($TraceGraphCalls)
        {
            $invokeParameters.Headers.GetEnumerator() | Where-Object { $_.Key -ne "Authorization" } | ForEach-Object {
                "$($_.Key): $($_.Value)" | out-file $script:traceFile -Append -Encoding UTF8
            }
            "Date: $([datetime]::UtcNow.ToString('r'))" | out-file $script:traceFile -Append -Encoding UTF8
        }

        if ($method -eq "POST" -or $method -eq "PATCH")
        {            
            if ($TraceGraphCalls)
            {
                "`r`n$body" | out-file $script:traceFile -Append -Encoding UTF8
            }
            $invokeParameters.Body = $body
        }

        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $result = Invoke-WebRequest @invokeParameters -ErrorAction Stop
        $stopwatch.Stop()
        LogVerbose "Web request total time: $($stopwatch.Elapsed.TotalMilliseconds) ms"

        if ($TraceGraphCalls)
        {
            "" | out-file $script:traceFile -Append -Encoding UTF8
            $result.RawContent | out-file $script:traceFile -Append -Encoding UTF8
            "`r`n" | out-file $script:traceFile -Append -Encoding UTF8
        }
    }
    catch
    {
        if ($ReturnResponse -and ($NoRedirect -or !$_.Exception.Response -or $_.Exception.Response.StatusCode -ne 308))
        {
            $errorResponse = GetGraphErrorResponse $_
            if ($TraceGraphCalls)
            {
                "" | out-file $script:traceFile -Append -Encoding UTF8
                if ($null -ne $errorResponse) { $errorResponse.RawContent | out-file $script:traceFile -Append -Encoding UTF8 }
                else { $_ | out-file $script:traceFile -Append -Encoding UTF8 }
                "`r`n" | out-file $script:traceFile -Append -Encoding UTF8
            }
            if ($null -ne $errorResponse) { return $errorResponse }
            throw
        }

        # Check for HTTP 308 Permanent Redirect (archive mailbox redirect)
        if (!$NoRedirect -and $_.Exception.Response -and $_.Exception.Response.StatusCode -eq 308)
        {
            $redirectUrl = $_.Exception.Response.Headers['Location']
            if ($redirectUrl)
            {
                LogVerbose "Received HTTP 308 redirect to: $redirectUrl"
                
                if ($TraceGraphCalls)
                {
                    "" | out-file $script:traceFile -Append -Encoding UTF8
                    "HTTP 308 Permanent Redirect" | out-file $script:traceFile -Append -Encoding UTF8
                    "Location: $redirectUrl" | out-file $script:traceFile -Append -Encoding UTF8
                    "Following redirect..." | out-file $script:traceFile -Append -Encoding UTF8
                    "`r`n" | out-file $script:traceFile -Append -Encoding UTF8
                }
                
                # Follow the redirect by making a new request
                return TraceInvokeRestMethod $method $redirectUrl $body -ReturnResponse:$ReturnResponse
            }
        }
        
        if ($TraceGraphCalls)
        {
            # Check if there is a result object
            if ($result) {
                # If the result contains raw content, log it to the trace file
                if ($result.RawContent) {
                    "" | out-file $script:traceFile -Append -Encoding UTF8
                    $result.RawContent | out-file $script:traceFile -Append -Encoding UTF8
                    "`r`n" | out-file $script:traceFile -Append -Encoding UTF8
                }
                else {
                    # If raw content is not available, log the entire result object
                    $result | out-file $script:traceFile -Append -Encoding UTF8
                }                
            }
            else {
                # If no result object is available, check for exception response details
                if ($Error[0].Exception.Response)
                {
                    # Extract and log response headers if available
                    $headers = $Error[0].Exception.Response.Headers
                    if ($headers -and $headers.Key.Count -gt 0) {
                        for ($i=0; $i -lt $headers.Key.Count; $i++)
                        {
                            "$($headers.Key[$i]): $($headers.Value[$i])" | out-file $script:traceFile -Append -Encoding UTF8
                        }
                        "" | out-file $script:traceFile -Append -Encoding UTF8
                    }
                    else {
                        # Log a message if no response headers are returned
                        Log "No response headers returned" Red
                    }
                }
                else {
                    # Log a message if no further error details are available
                    Log "No further error details available" Red
                }                
            }
            # Log the error object to the trace file
            $Error[0] | out-file $script:traceFile -Append -Encoding UTF8
        }
        Log "$method failed to URL: $url" Red
        if ($Error[0].ErrorDetails.Message)
        {
            $errResponse = $null
            $errResponse = $Error[0].ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue
            if ($errResponse)
            {
                if ($errResponse.error.details.message)
                {
                    Log "Error: $($errResponse.error.details.message)" Red
                }            
            }
        }
        return $null
    }
    if ($ReturnResponse) { return $result }
    return $result.Content
}

function ResolveGraphRequestUrl([string]$Url, [string]$BaseUrl, [scriptblock]$ValidateUrl)
{
    $target = $null
    if (![Uri]::TryCreate($Url, [UriKind]::Absolute, [ref]$target))
    {
        if ([string]::IsNullOrEmpty($BaseUrl)) { throw "An absolute Graph URL is required." }
        $target = [Uri]::new([Uri]$BaseUrl, $Url)
    }
    if ($target.Scheme -ne "https" -or $target.Host -ne "graph.microsoft.com" -or
        !$target.IsDefaultPort -or $target.UserInfo -or $target.Fragment)
    {
        throw "Rejected an unsafe Graph redirect or pagination URL."
    }
    if ($null -ne $ValidateUrl) { $null = & $ValidateUrl $target.AbsoluteUri }
    return $target.AbsoluteUri
}

# RequestOptions enables strict JSON responses, bounded redirects and per-response observation.
# Supported keys: MaxRedirects, ValidateUrl, OnResponse, OnRedirect.
# ValidateUrl must throw to reject a URL.
function InvokeGraphRequest([string]$method, [string]$url, [object]$body, [hashtable]$RequestOptions)
{
    if ($null -eq $RequestOptions) { $RequestOptions = @{} }
    $redirectLimit = 10
    if ($RequestOptions.ContainsKey("MaxRedirects")) { $redirectLimit = [int]$RequestOptions.MaxRedirects }
    if ($redirectLimit -lt 0 -or $redirectLimit -gt 50) { throw "MaxRedirects must be between 0 and 50." }
    $urlToGet = ResolveGraphRequestUrl $url -ValidateUrl $RequestOptions.ValidateUrl
    $initialUrl = $urlToGet
    $visited = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    for ($hop = 0; $hop -le $redirectLimit; $hop++)
    {
        if (!$visited.Add($urlToGet)) { throw "Graph redirect loop at $urlToGet" }
        $clientRequestId = [Guid]::NewGuid().ToString()
        $started = [DateTime]::UtcNow
        $response = $null
        $responseHeaders = @{}
        $transportError = $null
        try
        {
            if ($null -eq $script:headers) { $script:headers = @{ "Content-Type" = "application/json" } }
            $script:headers["client-request-id"] = $clientRequestId
            $script:headers["return-client-request-id"] = "true"
            $response = TraceInvokeRestMethod $method $urlToGet $body -ReturnResponse -NoRedirect
            foreach ($name in $response.Headers.Keys) { $responseHeaders[$name] = $response.Headers[$name] -join ", " }
        }
        catch { $transportError = $_.Exception.Message }
        $record = [PSCustomObject]@{
            StartedUtc = $started.ToString("o")
            FinishedUtc = [DateTime]::UtcNow.ToString("o")
            Method = $method
            Url = $urlToGet
            ClientRequestId = $clientRequestId
            Status = $(if ($null -ne $response) { [int]$response.StatusCode } else { $null })
            Headers = $(if ($null -ne $response) { $responseHeaders } else { $null })
            Body = $(if ($null -ne $response) { [string]$response.Content } else { $null })
            TransportError = $transportError
        }
        if ($null -ne $RequestOptions.OnResponse) { $null = & $RequestOptions.OnResponse $record }
        if ($transportError) { throw "Transport failure for ${urlToGet}: $transportError" }
        $data = $null
        if ($record.Status -ne 308 -and ![string]::IsNullOrWhiteSpace($record.Body))
        {
            try { $data = ConvertFrom-Json -InputObject $record.Body -ErrorAction Stop }
            catch { throw "HTTP $($record.Status) returned invalid JSON at $urlToGet." }
        }
        $redirect = $null
        if ($record.Status -eq 308)
        {
            $redirect = $responseHeaders["Location"]
            if ([string]::IsNullOrWhiteSpace($redirect)) { throw "HTTP 308 without a Location header at $urlToGet" }
        }
        elseif ($data.error.code -eq "ErrorArchiveFolderMovedPermanently")
        {
            $redirect = $data.error.message
            if ([string]::IsNullOrWhiteSpace($redirect)) { throw "Archive-moved error without a redirect URL at $urlToGet" }
        }
        if ($redirect)
        {
            $nextUrl = ResolveGraphRequestUrl $redirect $urlToGet -ValidateUrl $RequestOptions.ValidateUrl
            if ($hop -ge $redirectLimit) { throw "Exceeded MaxRedirects ($redirectLimit) for $initialUrl" }
            if ($visited.Contains($nextUrl)) { throw "Graph redirect loop at $nextUrl" }
            if ($null -ne $RequestOptions.OnRedirect) { $null = & $RequestOptions.OnRedirect $urlToGet $nextUrl }
            LogVerbose "Following Graph redirect: $urlToGet -> $nextUrl"
            $urlToGet = $nextUrl
            continue
        }
        if ($record.Status -lt 200 -or $record.Status -ge 300 -or $null -ne $data.error -or
            ($record.Status -eq 200 -and $null -eq $data))
        {
            throw "$method failed: HTTP $($record.Status), code '$($data.error.code)', URL $urlToGet."
        }
        return [PSCustomObject]@{
            Data = $data; Url = $urlToGet; RequestUrl = $initialUrl
            Status = $record.Status; Headers = $responseHeaders; Body = $record.Body
        }
    }
}

function GET($url, [hashtable]$RequestOptions)
{
    if ($null -ne $RequestOptions) { return InvokeGraphRequest "GET" $url $null $RequestOptions }
    Write-Verbose "Calling URL: $url"
    try
    {
        $result = TraceInvokeRestMethod "GET" $url
    }
    catch
    {
        Write-Host "GET failed from URL: $url" -ForegroundColor Red
        return $null
    }
    return $result
}

function POST($url, $body, $contentType="application/json")
{
    Write-Verbose "Calling URL: $url"
    try
    {
        $result = TraceInvokeRestMethod "POST" $url $body $contentType
    }
    catch
    {
        Write-Host "POST failed to URL: $url" -ForegroundColor Red
        return $null
    }
    # Return result or empty string (HTTP 202 with no content is success)
    if ($null -eq $result) { return "" }
    return $result
}

function PATCH($url, $body)
{
    Write-Verbose "Calling URL: $url"
    try
    {
        $result = TraceInvokeRestMethod "PATCH" $url $body
    }
    catch
    {
        Write-Host "PATCH failed at URL: $url" -ForegroundColor Red
    }
    return $result
}

function DELETE($url)
{
    Write-Verbose "Calling URL: $url"
    try
    {
        $result = TraceInvokeRestMethod "DELETE" $url
    }
    catch
    {
        Write-Host "DELETE failed at URL: $url" -ForegroundColor Red
        return $null
    }
    # Return result or empty string (HTTP 204 with no content is success)
    if ($null -eq $result) { return "" }
    return $result
}

function rawGET($url)
{
    return TraceInvokeRestMethod "GET" $url $null
}

<# Graph.ps1 %FUNCTIONS_END% #>

<# Mailbox.ps1 %FUNCTIONS_START% #>
# RequestOptions returns routed response metadata; omitting it retains parsed JSON results.
function GetMailboxFolder($queryString, [string]$BaseUrl = $script:graphBaseUrl, [hashtable]$RequestOptions,
    [string]$FolderId, [switch]$ChildFolders)
{
    $url = $BaseUrl.TrimEnd('/') + "/mailFolders"
    if (![string]::IsNullOrEmpty($FolderId)) { $url += "/" + [Uri]::EscapeDataString($FolderId) }
    if ($ChildFolders)
    {
        if ([string]::IsNullOrEmpty($FolderId)) { throw "FolderId is required when listing child folders." }
        $url += "/childFolders"
    }
    $url += $queryString
    if ($null -ne $RequestOptions) { return GET $url -RequestOptions $RequestOptions }
    $response = GET $url
    if ($null -eq $response)
    {
        return $null
    }
    return ConvertFrom-Json ($response)
}

function GetFolderMessages($folderId, $queryString, [string]$BaseUrl = $script:graphBaseUrl, [hashtable]$RequestOptions)
{
    if ($null -ne $RequestOptions) { $folderId = [Uri]::EscapeDataString($folderId) }
    $url = $BaseUrl.TrimEnd('/') + "/mailFolders/$folderId/messages$queryString"
    if ($null -ne $RequestOptions) { return GET $url -RequestOptions $RequestOptions }
    $response = GET $url
    if ($null -eq $response)
    {
        return $null
    }
    return ConvertFrom-Json ($response)
}

function GetMessages($queryString)
{
    $url = $script:graphBaseUrl + "messages$queryString"
    $response = GET $url
    if ($null -eq $response)
    {
        return $null
    }
    return ConvertFrom-Json ($response)
}

function GetFolderMessageById($folderId, $messageId, $queryString, [string]$BaseUrl = $script:graphBaseUrl, [hashtable]$RequestOptions)
{
    if ($null -ne $RequestOptions)
    {
        $folderId = [Uri]::EscapeDataString($folderId)
        $messageId = [Uri]::EscapeDataString($messageId)
    }
    $url = $BaseUrl.TrimEnd('/') + "/mailFolders/$folderId/messages/$messageId$queryString"
    if ($null -ne $RequestOptions) { return GET $url -RequestOptions $RequestOptions }
    $response = GET $url
    if ($null -eq $response)
    {
        return $null
    }
    return ConvertFrom-Json ($response)
}

function GetFolderMessageByIdAttachments($folderId, $messageId)
{
    $url = $script:graphBaseUrl + "mailFolders/$folderId/messages/$messageId/attachments"
    $response = GET $url
    if ($null -eq $response)
    {
        return $null
    }
    return ConvertFrom-Json ($response)
}

function DeleteFolderMessageById($folderId, $messageId)
{
    $url = $script:graphBaseUrl + "mailFolders/$folderId/messages/$messageId"
    $response = DELETE $url
    if ($null -eq $response)
    {
        return $null
    }
    return $response
}

function CreateMailboxFolder($parentFolderId, $folderName)
{
    $url = $script:graphBaseUrl + "mailFolders/$parentFolderId/childFolders"
    $body = @{
        displayName = $folderName
    } | ConvertTo-Json
    $response = POST $url $body
    if ($null -eq $response)
    {
        return $null
    }
    return ConvertFrom-Json ($response)
}

function CopyMessageToFolder($messageId, $destinationFolderId)
{
    $url = $script:graphBaseUrl + "messages/$messageId/copy"
    $body = @{
        destinationId = $destinationFolderId
    } | ConvertTo-Json
    $response = POST $url $body
    if ($null -eq $response)
    {
        return $null
    }
    return ConvertFrom-Json ($response)
}

<# Mailbox.ps1 %FUNCTIONS_END% #>

function Resolve-ArchiveUrl([string]$Url, [string]$BaseUrl)
{
    $target = [Uri](ResolveGraphRequestUrl $Url $BaseUrl)
    if ($target.AbsoluteUri -ceq $script:archiveDiscoveryUrl) { return $target.AbsoluteUri }
    $path = [Uri]::UnescapeDataString($target.AbsolutePath)
    $tenantPattern = [Regex]::Escape([string]$script:archiveTenant)
    if ([string]::IsNullOrEmpty($script:archiveTenant) -or
        $path -notmatch "^/(v1\.0|beta)/users/MBX:[0-9a-fA-F-]{36}@$tenantPattern/mailFolders(?:/|\('(?:[^']|'')+'\)(?:/|$)|$)")
    {
        throw "URL must target the supplied mailbox's discovery endpoint or an MBX mailFolders route in the discovered archive tenant."
    }
    return $target.AbsoluteUri
}

function New-ArchiveRequestOptions
{
    return @{
        MaxRedirects = $MaxRedirects
        ValidateUrl = { param([string]$Url) Resolve-ArchiveUrl $Url }
        OnResponse = {
            param([object]$Record)
            $Record | ConvertTo-Json -Depth 20 -Compress | Add-Content -LiteralPath $script:requestFile -Encoding UTF8 -ErrorAction Stop
        }
        OnRedirect = {
            param([string]$From, [string]$To)
            $script:archiveRedirectCount++
            Log "Archive redirect: $From -> $To" Yellow
        }
    }
}

function Get-ArchiveCollection([scriptblock]$ReadFirstPage, [int]$Limit = [int]::MaxValue)
{
    if ($Limit -le 0) { return }
    $seenPages = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $count = 0
    $page = & $ReadFirstPage
    while ($null -ne $page -and $count -lt $Limit)
    {
        if (!$seenPages.Add($page.Url)) { throw "Repeated archive pagination link: $($page.Url)" }
        if ($page.RequestUrl) { $null = $seenPages.Add($page.RequestUrl) }
        if ($null -eq $page.Data.PSObject.Properties["value"] -or $null -eq $page.Data.value -or $page.Data.value -isnot [array])
        {
            throw "Expected a Graph collection at $($page.Url)"
        }
        foreach ($entry in $page.Data.value)
        {
            if ([string]::IsNullOrWhiteSpace($entry.id)) { throw "Collection entry without an ID at $($page.Url)" }
            $entry
            $count++
            if ($count -ge $Limit) { break }
        }
        $nextLink = $page.Data.'@odata.nextLink'
        if (!$nextLink -or $count -ge $Limit) { break }
        $nextUrl = Resolve-ArchiveUrl $nextLink $page.Url
        if ($seenPages.Contains($nextUrl)) { throw "Repeated archive pagination link: $nextUrl" }
        $page = GET $nextUrl -RequestOptions $script:archiveRequestOptions
    }
}

function Write-ArchiveResult([object]$Result)
{
    $Result | ConvertTo-Json -Depth 20 -Compress | Add-Content -LiteralPath $script:resultFile -Encoding UTF8 -ErrorAction Stop
    Write-Output $Result
}

function Write-ArchiveFailure([string]$Operation, [string]$Id, [string]$Message)
{
    $script:archiveFailureCount++
    Log "$Operation failed for ${Id}: $Message" Red
    Write-ArchiveResult ([PSCustomObject]@{ Kind = "Failure"; Operation = $Operation; Id = $Id; Error = $Message })
}

# Redirects can rewrite mailbox-scoped IDs; validate against the final URL.
function Assert-ArchiveResourceId([object]$Response, [string]$ResourceName)
{
    $path = ([Uri]$Response.Url).AbsolutePath
    $key = [Regex]::Match($path, "/$ResourceName(?:/(?<slash>[^/]+)|\('(?<odata>(?:[^']|'')+)'\))$")
    if (!$key.Success) { throw "Cannot determine the resource ID from the final routed GET URL." }
    if ($key.Groups["slash"].Success) { $expectedId = [Uri]::UnescapeDataString($key.Groups["slash"].Value) }
    else { $expectedId = [Uri]::UnescapeDataString($key.Groups["odata"].Value).Replace("''", "'") }
    if ([string]::IsNullOrWhiteSpace($Response.Data.id) -or $Response.Data.id -cne $expectedId)
    {
        throw "Resource GET returned a different or missing ID than its final routed URL."
    }
}

function Initialize-ArchiveMailbox
{
    $script:archiveMailboxId = $null
    $script:archiveTenant = $null
    $settings = (GET $script:archiveDiscoveryUrl -RequestOptions $script:archiveRequestOptions).Data
    if ($null -eq $settings.PSObject.Properties["inPlaceArchiveMailboxId"])
    {
        throw "Exchange settings did not return inPlaceArchiveMailboxId for $Mailbox; archive discovery is unavailable in this response."
    }
    $discoveredId = [string]$settings.inPlaceArchiveMailboxId
    if ([string]::IsNullOrWhiteSpace($discoveredId))
    {
        throw "Exchange settings returned no in-place archive mailbox ID for $Mailbox."
    }
    if ($discoveredId -notmatch '^MBX:(?<archive>[0-9a-fA-F-]{36})@(?<tenant>[0-9a-fA-F-]{36})$')
    {
        throw "Exchange settings returned an invalid archive mailbox ID for $Mailbox."
    }
    $archiveGuid = [Guid]::Empty
    $tenantGuid = [Guid]::Empty
    if (![Guid]::TryParse($Matches.archive, [ref]$archiveGuid) -or $archiveGuid -eq [Guid]::Empty -or
        ![Guid]::TryParse($Matches.tenant, [ref]$tenantGuid) -or $tenantGuid -eq [Guid]::Empty)
    {
        throw "Exchange settings returned invalid archive or tenant GUIDs for $Mailbox."
    }
    $configuredTenant = [Guid]::Empty
    if ([Guid]::TryParse($TenantId, [ref]$configuredTenant) -and $configuredTenant -ne $tenantGuid)
    {
        throw "Discovered archive tenant does not match TenantId."
    }
    if ($discoveredId -eq $settings.primaryMailboxId)
    {
        throw "Exchange settings returned the primary mailbox ID as the archive ID."
    }
    $script:archiveMailboxId = $discoveredId
    $script:archiveTenant = $tenantGuid.ToString()
    Write-ArchiveResult ([PSCustomObject]@{
        Kind = "ArchiveMailbox"; Mailbox = $Mailbox; ArchiveMailboxId = $discoveredId
        ArchiveGuid = $archiveGuid.ToString(); TenantId = $script:archiveTenant
        PrimaryMailboxId = $settings.primaryMailboxId
    })
}

function Test-ArchiveFolders
{
    if ([string]::IsNullOrEmpty($script:archiveMailboxId)) { throw "Discover the archive mailbox before listing folders." }
    $mailboxRequest = @{
        BaseUrl = "https://graph.microsoft.com/$ApiVersion/users/$script:archiveMailboxId/"
        RequestOptions = $script:archiveRequestOptions
    }
    $hiddenQuery = "?includeHiddenFolders=$($IncludeHiddenFolders.IsPresent.ToString().ToLowerInvariant())"
    $folderSelect = '&$select=id,displayName,parentFolderId,childFolderCount,totalItemCount'
    $messageQuery = '?$select=id,subject,receivedDateTime,parentFolderId&$top=' + $ItemsPerFolder
    $queue = [System.Collections.Generic.Queue[object]]::new()
    $seenFolders = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    # Enqueue as pages arrive so already discovered folders survive a later page failure.
    try
    {
        Get-ArchiveCollection { GetMailboxFolder "$hiddenQuery$folderSelect" @mailboxRequest } | ForEach-Object { $queue.Enqueue($_) }
    }
    catch { Write-ArchiveFailure "ListFolders" $script:archiveMailboxId $_.Exception.Message }
    while ($queue.Count -gt 0)
    {
        $folder = $queue.Dequeue()
        if (!$seenFolders.Add($folder.id)) { continue }
        $script:archiveFolderCount++
        Write-ArchiveResult ([PSCustomObject]@{
            Kind = "Folder"; Id = $folder.id; DisplayName = $folder.displayName
            ParentFolderId = $folder.parentFolderId; TotalItemCount = $folder.totalItemCount
        })
        try
        {
            $getFolder = GetMailboxFolder '?$select=id,displayName,parentFolderId,childFolderCount,totalItemCount' -FolderId $folder.id @mailboxRequest
            Assert-ArchiveResourceId $getFolder "mailFolders"
            $null = $seenFolders.Add($getFolder.Data.id)
        }
        catch { Write-ArchiveFailure "GetFolder" $folder.id $_.Exception.Message }
        # Start each operation through the main archive to independently test its routing.
        try
        {
            Get-ArchiveCollection { GetFolderMessages $folder.id $messageQuery @mailboxRequest } -Limit $ItemsPerFolder | ForEach-Object {
                $message = $_
                $script:archiveMessageCount++
                Write-ArchiveResult ([PSCustomObject]@{
                    Kind = "Message"; FolderId = $folder.id; Id = $message.id
                    Subject = $message.subject; ReceivedDateTime = $message.receivedDateTime
                })
                try
                {
                    $getMessage = GetFolderMessageById $folder.id $message.id '?$select=id,subject,receivedDateTime,parentFolderId' @mailboxRequest
                    Assert-ArchiveResourceId $getMessage "messages"
                }
                catch { Write-ArchiveFailure "GetMessage" $message.id $_.Exception.Message }
            }
        }
        catch { Write-ArchiveFailure "ListMessages" $folder.id $_.Exception.Message }
        try
        {
            Get-ArchiveCollection { GetMailboxFolder "$hiddenQuery$folderSelect" -FolderId $folder.id -ChildFolders @mailboxRequest } | ForEach-Object { $queue.Enqueue($_) }
        }
        catch { Write-ArchiveFailure "ListChildFolders" $folder.id $_.Exception.Message }
    }
}

# Main code
$script:archiveRedirectCount = 0
$script:archiveFailureCount = 0
$script:archiveFolderCount = 0
$script:archiveMessageCount = 0
$script:archiveRequestOptions = New-ArchiveRequestOptions
Log "Discovering the archive for $Mailbox; evidence: $OutputPath"
try { Initialize-ArchiveMailbox }
catch
{
    Write-ArchiveFailure "DiscoverArchive" $Mailbox $_.Exception.Message
    throw
}
Log "Testing discovered archive $script:archiveMailboxId ($ApiVersion)"
Test-ArchiveFolders
Write-ArchiveResult ([PSCustomObject]@{
    Kind = "Summary"; Folders = $script:archiveFolderCount; Messages = $script:archiveMessageCount
    Redirects = $script:archiveRedirectCount; Failures = $script:archiveFailureCount
    AuxiliaryRouting = $(if ($script:archiveRedirectCount -gt 0) { "Redirects observed; inspect Requests.jsonl for target mailboxes." } else { "Not exercised: no archive redirects observed." })
})
if ($script:archiveFailureCount -gt 0) { throw "$($script:archiveFailureCount) archive operation(s) failed; see $OutputPath" }
Log "Completed"
