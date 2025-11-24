#
# Test-ForwardMIME.ps1
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
Tests retrieving and then forwarding mailbox message in MIME format using Graph.

.DESCRIPTION
Script to test retrieving and then forwarding mailbox message in MIME format, using either delegated or application permissions.

.EXAMPLE

Delegate permissions (requires Mail.ReadWrite):
.\Test-ForwardMIME.ps1 -AppId $clientId -TenantId $tenantId -ForwardingRecipient $Mailbox -UpdateMessageId -RemoveFromHeader

Application permissions (requires Mail.ReadWrite):
.\Test-ForwardMIME.ps1 -mailbox $Mailbox -AppId $clientId -TenantId $tenantId -AppSecretKey $secretKey -ForwardingRecipient
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

    [Parameter(Mandatory=$False,HelpMessage="Mailbox (me to access own mailbox when using delegate permissions).")]
    [ValidateNotNullOrEmpty()]
    [string]$Mailbox = "me",
<# Auth.ps1 %PARAMS_END% #>

    [Parameter(Mandatory=$False,HelpMessage="Timeout for HTTP requests (default is 30 seconds, which is same as Microsoft Graph endpoint).  Only works for Powershell 7.4+ (prior to that, timeout is infinite).")]
    [ValidateRange(0,300)]
    [int]$ConnectionTimeout = 30,

<# Logging.ps1 %PARAMS_START% #>
    [Parameter(Mandatory=$False,HelpMessage="If included, activity is logged to a file (same as script name with .log appended).")]	
    [switch]$LogToFile,
<# Logging.ps1 %PARAMS_END% #>

<# Graph.ps1 %PARAMS_START% #>
    [Parameter(Mandatory=$False,HelpMessage="If specified, all Graph calls will be logged to file (same name as script, ending in .trace).")]
    [switch]$TraceGraphCalls,
<# Graph.ps1 %PARAMS_END% #>

    [Parameter(Mandatory=$True,HelpMessage="Mailbox to which the message will be forwarded.")]
    [ValidateNotNullOrEmpty()]
    [string]$ForwardingRecipient = "",

    [Parameter(Mandatory=$False,HelpMessage="If provided, this MIME will be sent (as is) instead of a message being taken from the Inbox and forwarded.")]
    [ValidateNotNullOrEmpty()]
    [string]$MIMEContent,

    [Parameter(Mandatory=$False,HelpMessage="Any additional headers to add to the forwarded message (these are not expected to already exist as no checks are made to replace).")]
    [string]$AdditionalHeaders = "",

    [Parameter(Mandatory=$False,HelpMessage="If set, the internet message-id and related headers will be updated (this is expected when forwarding from a mail client).")]
    [switch]$UpdateMessageId,

    [Parameter(Mandatory=$False,HelpMessage="If set, the FROM header is removed from the forwarded MIME.")]
    [switch]$RemoveFROMHeader,

    [Parameter(Mandatory=$False,HelpMessage="If set, the CC header is not removed from the forwarded MIME (it is by default so as to only forward to the specified recipient).")]
    [switch]$DoNotRemoveCCHeader,

    [Parameter(Mandatory=$False,HelpMessage="If set, original and updated MIME will be saved to disk.")]
    [switch]$ExportMIME,

    [Parameter(Mandatory=$False,HelpMessage="If not null or empty, filters for this specific value for the extended property.")]
    [string]$FilterSpecificValue = ""
)

$script:ScriptVersion = "1.0.0"

<# Logging.ps1 %FUNCTIONS_START% #>
if ($LogToFile) {
    $script:logFile = "$($MyInvocation.InvocationName).log"
}

Function LogToFile([string]$LogEntry)
{
	if ( [String]::IsNullOrEmpty($script:logFile) ) { return }
	$LogEntry | Out-File $script:logFile -Append
}

Function Log([string]$Details, [System.ConsoleColor]$ForegroundColor = "Gray")
{
    $logEntry = "$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToLongTimeString())   $Details"
    Write-Host $logEntry -ForegroundColor $ForegroundColor
    LogToFile $logEntry
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
    $script:graphBaseUrl = "https://graph.microsoft.com/v1.0/users/$Mailbox/"
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
function TraceInvokeRestMethod($method, $url, $body, $contentType="application/json")
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
            ContentType = $contentType
        }
        if ($PSVersionTable.PSVersion -ge [version]7.4)
        {
            # PowerShell 7.4+ supports ConnectionTimeoutSeconds
            $invokeParameters.ConnectionTimeoutSeconds = $ConnectionTimeout
        }

        if ($method -eq "POST" -or $method -eq "PATCH")
        {            
            if ($TraceGraphCalls)
            {
                "`r`n$body" | out-file $script:traceFile -Append
            }
            $invokeParameters.Body = $body
        }
        $result = Invoke-WebRequest @invokeParameters -ErrorAction Stop
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
                    Log "Return status code: $($Error[0].Exception.Response.StatusCode)" Red
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
        Log "$method failed to URL: $url"Red
        return $null
    }
    return $result.Content
}

function GET($url)
{
    UpdateHeaders
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

function PATCH($url, $body)
{
    UpdateHeaders
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

function POST($url, $body, $contentType="application/json")
{
    UpdateHeaders
    Write-Verbose "Calling URL: $url"
    try
    {
        $result = TraceInvokeRestMethod "POST" $url $body $contentType
    }
    catch
    {
        Write-Host "POST failed at URL: $url" -ForegroundColor Red
    }
    return $result
}

function DELETE($url)
{
    UpdateHeaders
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

function GetMailboxFolder($queryString)
{
    $url = $script:graphBaseUrl + "mailFolders$queryString"
    $response = GET $url
    if ($null -eq $response)
    {
        return $null
    }
    return ConvertFrom-Json ($response)
}

function GetFolderMessages($folderId, $queryString)
{
    $url = $script:graphBaseUrl + "mailFolders/$folderId/messages$queryString"
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

function GetFolderMessageById($folderId, $messageId, $queryString)
{
    $url = $script:graphBaseUrl + "mailFolders/$folderId/messages/$messageId$queryString"
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

<# Graph.ps1 %FUNCTIONS_END% #>


# Attempt to get the Inbox folder first to ensure we can access the mailbox
$inbox = GetMailboxFolder("?`$filter=displayName eq 'INBOX'")
if ($null -eq $inbox)
{
    Log "Failed to obtain inbox folder" Red
    exit # Failed to obtain inbox folder
}
Log "Inbox folder id: $($inbox.value.id)"

if (-not [String]::IsNullOrEmpty($MIMEContent))
{
    Log "Using provided MIME content"
    $newMimeContent = $MIMEContent
}
else
{
    # Retrieve the MIME of the first message in the Inbox
    $messages = GetFolderMessages $inbox.value.id "?`$top=1&`$select=id,subject"
    if ($null -eq $messages -or $messages.value.Count -eq 0)
    {
        Log "No messages found in inbox" Red
        exit # No messages found
    }
    $firstMessage = $messages.value[0]
    Log "Retrieving MIME for message: $($firstMessage.subject) (id: $($firstMessage.id))"
    $mimeContent = rawGET($script:graphBaseUrl + "mailFolders/$($inbox.value.id)/messages/$($firstMessage.id)/`$value")
    if ($null -eq $mimeContent)
    {
        Log "Failed to retrieve MIME content for message" Red
        exit # Failed to retrieve MIME content
    }
    Log "Successfully retrieved MIME content for message"

    # Save the MIME content to a file
    $mimeFilePath = "$($MyInvocation.MyCommand.Name)_original.eml"
    $mimeContent | Out-File -FilePath $mimeFilePath -Encoding utf8
    Log "Saved original MIME content to file: $mimeFilePath"

    # Split the MIME content into headers and body
    $mimeLines = $mimeContent -split "`r`n"
    $activeHeader = ""
    $headerLines = @{}
    $bodyLines = @()
    $inHeaders = $true
    foreach ($line in $mimeLines)
    {
        if ($inHeaders)
        {
            # Header can be split across multiple lines, so we extract each one individually to ensure that we update correctly
            if ([String]::IsNullOrEmpty($line))
            {
                $inHeaders = $false
            }
            else
            {
                if ($line.StartsWith(" ") -or $line.StartsWith("`t"))
                {
                    # Continuation of previous header
                    $headerLines[$activeHeader] += "`r`n" + $line
                }
                else
                {
                    # New header
                    $colonIndex = $line.IndexOf(":")
                    if ($colonIndex -gt 0)
                    {
                        $activeHeader = $line.Substring(0, $colonIndex)
                        $headerLines[$activeHeader] = $line.Substring($colonIndex + 1).Trim()
                    }
                }
            }
        }
        else
        {
            $bodyLines += $line
        }
    }

    # Go through each header and update as needed
    $newMessageId = "<" + [Guid]::NewGuid().ToString() + "@example.com>"
    if ($UpdateMessageId)
    {
        if ($headerLines.ContainsKey("Message-ID"))
        {
            $originalMessageId = $headerLines["Message-ID"]
            Log "Updating Message-ID to $newMessageId"
            $headerLines["Message-ID"] = "$newMessageId"

            # For a forwarded message, we should also update the References header to include the original message ID
            if ($headerLines.ContainsKey("References"))
            {        
                Log "Adding $originalMessageId to References header"
                $headerLines["References"] += " $originalMessageId"
            }
            else
            {
                Log "Creating References header with $originalMessageId"
                $headerLines["References"] = "$originalMessageId"
            }
        }
    }
    if ($RemoveFROMHeader)
    {
        if ($headerLines.ContainsKey("From"))
        {
            Log "Removing FROM header"
            $headerLines.Remove("From")
        }
    }
    if (-not $DoNotRemoveCCHeader)
    {
        if ($headerLines.ContainsKey("Cc"))
        {
            Log "Removing CC header"
            $headerLines.Remove("Cc")
        }
    }

    if (![String]::IsNullOrEmpty($ForwardingRecipient))
    {
        Log "Setting To header to $ForwardingRecipient"
        $headerLines["To"] = $ForwardingRecipient
    }

    # Parse the headers back into a list of lines
    $newHeaderLines = @()
    foreach ($key in $headerLines.Keys)
    {
        $newHeaderLines += "$($key): $($headerLines[$key])"
    }

    # Put the new MIME content back together
    $newMimeContent = ($newHeaderLines + $AdditionalHeaders + "" + $bodyLines) -join "`r`n"
}
# Save the new MIME content to a file
$mimeFilePath = "$($MyInvocation.MyCommand.Name)_updated.eml"
$newMimeContent | Out-File -FilePath $mimeFilePath -Encoding utf8
Log "Saved original MIME content to file: $mimeFilePath"

# Now send the MIME content as a new message to the specified recipient
$sendUrl = $script:graphBaseUrl + "sendMail"

# We need to Base64 encode the MIME content
$encodedMime = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($newMimeContent))
$result = POST $sendUrl $encodedMime "text/plain"

Log "Completed"
