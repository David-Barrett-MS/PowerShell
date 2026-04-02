#
# Get-ServiceCommunications.ps1
#
# By David Barrett, Microsoft Ltd. Use at your own risk.  No warranties are given.
#
#  DISCLAIMER:
# THIS CODE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.

<#
.SYNOPSIS
Retrieves service information and messages using Microsoft Graph.

.DESCRIPTION
Example script showing how to retrieve service health information and messages using the service communications API in Microsoft Graph.
https://learn.microsoft.com/en-us/graph/api/resources/service-communications-api-overview?view=graph-rest-1.0

.EXAMPLE
.\Get-ServiceCommunications.ps1 -AppId $appId -TenantId $tenantId -AppSecretKey $secretKey -ListHealthOverviews -ListIssues -ListMessages

#>


param (
<# Auth.ps1 %PARAMS_START% #>
    [Parameter(Mandatory=$False,HelpMessage="Application Id (obtained when registering the application in Azure AD).")]
    [ValidateNotNullOrEmpty()]
    [string]$AppId,

    [Parameter(Mandatory=$False,HelpMessage="Application secret key (obtained when registering the application in Azure AD).  If not specified, delegate permissions are assumed.")]
    [ValidateNotNullOrEmpty()]
    [string]$AppSecretKey,

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

    [Parameter(Mandatory=$False,HelpMessage="Retrieve service health overviews.")]
    [switch]$ListHealthOverviews,

    [Parameter(Mandatory=$False,HelpMessage="Retrieve service health issues.")]
    [switch]$ListIssues,

    [Parameter(Mandatory=$False,HelpMessage="Retrieve service health messages.")]
    [switch]$ListMessages
)

$script:ScriptVersion = "1.0.0"

<# Logging.ps1 %FUNCTIONS_START% #>
if ($LogToFile) {
    $script:logFile = "$($MyInvocation.InvocationName).log"
}

Function UpdateDetailsWithCallingMethod([string]$Details)
{
    # Update the log message with details of the function that logged it
    $timeInfo = "$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToLongTimeString())"
    $callingFunction = (Get-PSCallStack)[2].Command # The function we are interested in will always be frame 2 on the stack
    if (![String]::IsNullOrEmpty($callingFunction))
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
function GetAppAuthToken()
{
    # Acquire token for application permissions
    $body = @{grant_type="client_credentials";scope="https://graph.microsoft.com/.default";client_id=$AppId;client_secret=$AppSecretKey}
    try
    {
        $script:oauth = Invoke-RestMethod -Method Post -Uri https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token -Body $body
        $script:token_expires = (Get-Date).AddSeconds($oauth.expires_in)
    }
    catch
    {
        Write-Host "Failed to obtain OAuth token" -ForegroundColor Red
        $Error[0]
        exit # Failed to obtain a token
    }
    $script:graphBaseUrl = "https://graph.microsoft.com/v1.0/users/$User/"
}

function RenewOAuthToken()
{
    # If using app flow, we just obtain a new token

    if (![String]::IsNullOrEmpty($AppSecretKey))
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
    if ($null -eq $script:headers -or $script.headers.Count -eq 0)
    {
        $script:headers = @{
            'Authorization'  = "$($oauth.token_type) $($oauth.access_token)";
            'Content-Type'   = 'application/json';
        }    
    }
}



## Authentication

if ([String]::IsNullOrEmpty($AppSecretKey))
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
function TraceInvokeRestMethod($method, $url, $body)
{
    UpdateHeaders
    Write-Verbose "Calling URL: $url"
    if ($TraceGraphCalls)
    {
        Add-Content -Path $script:traceFile -Value "$method $url"
    }
    try
    {
        $result = $null
        $invokeParameters = @{
            Method = $method
            Uri = $url
            Headers = $script:headers
            ContentType = "application/json"
        }
        if ($PSVersionTable.PSVersion -ge [version]7.4)
        {
            # PowerShell 7.4+ supports ConnectionTimeoutSeconds
            $invokeParameters.ConnectionTimeoutSeconds = $ConnectionTimeout
        }

        # Write request headers, excluding Authorization
        if ($TraceGraphCalls)
        {
            $invokeParameters.Headers.GetEnumerator() | Where-Object { $_.Key -ne "Authorization" } | ForEach-Object {
                "$($_.Key): $($_.Value)" | out-file $script:traceFile -Append
            }
            "Date: $([datetime]::UtcNow.ToString('r'))" | out-file $script:traceFile -Append
        }

        if ($method -eq "POST" -or $method -eq "PATCH")
        {            
            if ($TraceGraphCalls)
            {
                "`r`n$body" | out-file $script:traceFile -Append
            }
            $invokeParameters.Body = $body
        }

        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $result = Invoke-WebRequest @invokeParameters -ErrorAction Stop
        $stopwatch.Stop()
        LogVerbose "Web request total time: $($stopwatch.Elapsed.TotalMilliseconds) ms"

        if ($TraceGraphCalls)
        {
            "" | out-file $script:traceFile -Append
            $result.RawContent | out-file $script:traceFile -Append
            "`r`n" | out-file $script:traceFile -Append
        }
    }
    catch
    {
        if ($TraceGraphCalls)
        {
            # Check if there is a result object
            if ($result) {
                # If the result contains raw content, log it to the trace file
                if ($result.RawContent) {
                    "" | out-file $script:traceFile -Append
                    $result.RawContent | out-file $script:traceFile -Append
                    "`r`n" | out-file $script:traceFile -Append
                }
                else {
                    # If raw content is not available, log the entire result object
                    $result | out-file $script:traceFile -Append
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
                            "$($headers.Key[$i]): $($headers.Value[$i])" | out-file $script:traceFile -Append
                        }
                        "" | out-file $script:traceFile -Append
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
            $Error[0] | out-file $script:traceFile -Append
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
    return $result.Content
}

function GET($url)
{
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
    }
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
    }
    return $result
}

function rawGET($url)
{
    return TraceInvokeRestMethod "GET" $url $null
}

<# Graph.ps1 %FUNCTIONS_END% #>

# Service Communications API functions

function ListHealthOverviews()
{
    $url = "https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/healthOverviews"
    return GET $url
}

function ListIssues()
{
    $url = "https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/issues"
    return GET $url
}

function ListMessages()
{
    $url = "https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/messages"
    return GET $url
}

# Main code

# Service announcement endpoints (ServiceHealth.Read.All):
# https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/healthOverviews
# https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/healthOverviews/{ServiceName}

# https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/issues
# https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/issues/{serviceHealthIssueId}

# Service messages endpoints (ServiceMessage.Read.All):
# https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/messages
# https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/messages/{serviceUpdateMessageId}
# https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/messages/{serviceUpdateMessageId}/attachments
#
# Delegated only:
# POST https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/messages/markRead
# POST https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/messages/markUnread
# Content-Type: application/json
#
# {
#   "messageIds": ["MC172851", "MC167983"]
# }

$assignedVariableInfo = @()
if ($ListHealthOverviews)
{
    $global:healthOverviews = (ListHealthOverviews | ConvertFrom-Json).value
    Write-Host "Health overviews:"
    $global:healthOverviews | Format-Table -Property service, status
    $assignedVariableInfo += "HealthOverviews: `$healthOverviews"
}
if ($ListIssues)
{
    $global:issues = (ListIssues | ConvertFrom-Json).value
    Write-Host "Issues:"
    $global:issues | Format-Table -Property id, title, status
    $assignedVariableInfo += "Issues: `$issues"
}
if ($ListMessages)
{
    $global:messages = (ListMessages | ConvertFrom-Json).value
    Write-Host "Messages:"
    $global:messages | Format-Table -Property id, title, status
    $assignedVariableInfo += "Messages: `$messages"
}

if ($assignedVariableInfo.Count -gt 0)
{
    Write-Host "Assigned the following variables:`r`n$($assignedVariableInfo -join "`r`n")" -ForegroundColor Green
}


Log "Completed"
