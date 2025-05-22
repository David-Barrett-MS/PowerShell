#
# Test-MessageProcessing.ps1
#
# By David Barrett, Microsoft Ltd. 2025. Use at your own risk.  No warranties are given.
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
Simulates actions taken by an app automatically processing messages in a mailbox.

.DESCRIPTION
Script used for testing various Graph calls interacting with messages in a mailbox.

.EXAMPLE

Delegate permissions (requires Mail.ReadWrite):
.\Test-MessageProcessing.ps1 -AppId $clientId -TenantId $tenantId -TestCount 1

Application permissions (requires Mail.ReadWrite):
.\Test-MessageProcessing.ps1 -mailbox $Mailbox -AppId $clientId -TenantId $tenantId -AppSecretKey $secretKey -AnalyseAttachments -RetrieveAttachmentsIndividually

#>


param (
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

    [Parameter(Mandatory=$False,HelpMessage="If specified, MIME for the message will be requested (using `$value Graph call).")]
    [switch]$RetrieveFullMIME,

    [Parameter(Mandatory=$False,HelpMessage="If specified, more detailed information about message attachments will be shown.")]
    [switch]$AnalyseAttachments,

    [Parameter(Mandatory=$False,HelpMessage="If specified, each attachment will be retrieved by specific Id (requires -AnalyseAttachments).")]
    [switch]$RetrieveAttachmentsIndividually,

    [Parameter(Mandatory=$False,HelpMessage="Timeout for HTTP requests (default is 30 seconds, which is same as Microsoft Graph endpoint).  Only works for Powershell 7.4+ (prior to that, timeout is infinite).")]
    [ValidateRange(0,300)]
    [int]$ConnectionTimeout = 30,

    [Parameter(Mandatory=$False,HelpMessage="If specified, output will be logged to file (using same name as this script, with .log appended).")]
    [switch]$LogToFile,

    [Parameter(Mandatory=$False,HelpMessage="If specified, all Graph calls will be logged to file (same name as script, ending in .trace).")]
    [ValidateNotNullOrEmpty()]
    [switch]$TraceGraphCalls,

    [Parameter(Mandatory=$False,HelpMessage="If specified, the message will not be marked as read after processing.")]
    [ValidateNotNullOrEmpty()]
    [switch]$DoNotMarkUnread
)

$script:ScriptVersion = "1.0.3"
if ($LogToFile) {
    $script:logFile = "$($MyInvocation.InvocationName).log"
}
if ($TraceGraphCalls) {
    $script:traceFile = "$($MyInvocation.InvocationName).trace"
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

function TraceInvokeRestMethod($method, $url, $body)
{
    UpdateHeaders
    Write-Verbose "Calling URL: $url"
    if ($TraceGraphCalls)
    {
        "$method $url" | out-file $script:traceFile -Append
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

        if ($method -eq "POST" -or $method -eq "PATCH")
        {
            $jsonBody = $body | ConvertTo-Json -Depth 10
            if ($TraceGraphCalls)
            {
                "`r`n$jsonBody" | out-file $script:traceFile -Append
            }
            $invokeParameters.Body = $jsonBody
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
                    # Extract and log response headers if available
                    $headers = $Error[0].Exception.Response.Headers
                    if ($headers -and $headers.Count -gt 0) {
                        for ($i=0; $i -lt $headers.Count; $i++)
                        {
                            "$($headers.Keys[$i]): $($headers[$i])" | out-file $script:traceFile -Append
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
    $result = TraceInvokeRestMethod "GET" $url $null
    if ($null -eq $result)
    {
        return $null
    }
    return ConvertFrom-Json $result
}

function rawGET($url)
{
    return TraceInvokeRestMethod "GET" $url $null
}

function PATCH($url, $body)
{
    $result = TraceInvokeRestMethod "PATCH" $url $body
    if ($null -eq $result)
    {
        return $null
    }
    return ConvertFrom-Json $result
}

function DELETE($url)
{
    $result = TraceInvokeRestMethod "DELETE" $url $null
    if ($null -eq $result)
    {
        return $null
    }
    return ConvertFrom-Json $result
}

function GetMailboxFolder($queryString)
{
    $url = $script:graphBaseUrl + "mailFolders/$queryString"
    return GET $url
}

function GetFolderMessages($folderId, $queryString)
{
    $url = $script:graphBaseUrl + "mailFolders/$folderId/messages/$queryString"
    return GET $url
}

function MarkMessageAsRead($messageId)
{
    $url = $script:graphBaseUrl + "messages/$messageId"
    $body = @{
        isRead = $true
    }
    return PATCH $url $body
}

function MarkFolderMessageAsRead($folderId, $messageId)
{
    $url = $script:graphBaseUrl + "mailFolders/$folderId/messages/$messageId"
    $body = @{
        isRead = $true
    }
    return PATCH $url $body
}

function GetFolderMessageById($folderId, $messageId, $queryString)
{
    $url = $script:graphBaseUrl + "mailFolders/$folderId/messages/$messageId/$queryString"
    return GET $url
}

function GetFolderMessageByIdAttachments($folderId, $messageId)
{
    $url = $script:graphBaseUrl + "mailFolders/$folderId/messages/$messageId/attachments"
    return GET $url
}

function DeleteFolderMessageById($folderId, $messageId)
{
    $url = $script:graphBaseUrl + "mailFolders/$folderId/messages/$messageId"
    return DELETE $url
}

function GetMessageById($messageId, $queryString)
{
    $url = $script:graphBaseUrl + "messages/$messageId/$queryString"
    return GET $url
}

function GetMessageMIMEById($messageId, $queryString)
{
    $url = $script:graphBaseUrl + "messages/$messageId/$queryString"
    return rawGET $url
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
        exit # Failed to obtain a token
    }
    $graphBaseUrl = "https://graph.microsoft.com/v1.0/me/"
}
else
{
    GetAppAuthToken
}
Log "Successfully obtained OAuth token"

# GET /mailFolders/?$filter=displayName+eq+'INBOX'
$inbox = GetMailboxFolder("?`$filter=displayName eq 'INBOX'")
if ($null -eq $inbox)
{
    Log "Failed to obtain inbox folder" Red
    exit # Failed to obtain inbox folder
}
$inboxId = $inbox.value.id
Log "Inbox folder id: $($inboxId)"

# GET /mailfolders/inbox/messages/?$filter=isRead+eq+false&$count=true&$top=20&$orderBy=receivedDateTime&select=internetMessageId,id
$unreadMessages = GetFolderMessages $inboxId "?`$filter=isRead eq false&`$count=true&`$top=20&`$orderBy=receivedDateTime&select=internetMessageId,id"
if ($null -eq $unreadMessages)
{
    Log "Failed to obtain unread messages" Red
    exit # Failed to obtain unread messages
}
if ($unreadMessages.value.Count -eq 0)
{
    Log "No unread messages to process" Yellow
    exit # No unread messages to process
}
Log "$($unreadMessages.value.Count) unread messages found"

$global:messageToProcess = $unreadMessages.value[0]

if ($null -eq $messageToProcess)
{
    Log "Failed to obtain message to process" Red
    exit # Failed to obtain message to process
}

if (!$DoNotMarkUnread)
# PATCH message (mark as read)
{
    $global:updatedMessage = MarkFolderMessageAsRead $inboxId $messageToProcess.id
    if ($null -eq $updatedMessage)
    {
        Log "Failed to mark message as read" Red
    }
    Log "Message marked as read: $($updatedMessage.subject)"
}

# GET /mailFolders/?$filter=displayName+eq+'INBOX'
$inbox = GetMailboxFolder("?`$filter=displayName eq 'INBOX'")
if ($null -eq $inbox)
{
    Log "Failed to obtain inbox folder" Red
    exit # Failed to obtain inbox folder
}
$inboxId = $inbox.value.id
Log "Inbox folder id: $($inboxId)"

# GET /mailfolders/inbox(by Id)/MessageID?$select=*
$global:message = GetFolderMessageById $inboxId $messageToProcess.id "?`$select=*"
if ($null -eq $message)
{
    Log "Failed to obtain message by id: $($messageToProcess.id)" -ForegroundColor Red
    exit # Failed to obtain message
}

# GET /mailfolders/inbox(by Id)/MessageID/attachments
$global:attachments = GetFolderMessageByIdAttachments $inboxId $messageToProcess.id
if ($null -eq $attachments)
{
    Log "Failed to obtain message attachments" Red
    exit # Failed to obtain message attachments
}
Log "$($attachments.value.Count) attachments found"
$global:individualAttachments = @()

if ($AnalyseAttachments)
{
    foreach ($attachment in $attachments.value)
    {
        Log
        Log "Attachment type: $($attachment.'@odata.type')"
        Log "Name: $($attachment.name)"
        Log "Id: $($attachment.id)"
        Log "Size: $($attachment.size)"
        Log "Content type: $($attachment.contentType)"
        Log "Content id: $($attachment.contentId)"
        Log "Content location: $($attachment.contentLocation)"
        if ($attachment.contentBytes) {
            Log "Content bytes length: $($attachment.contentBytes.length)"
        } else {
            if ($attachment.'@odata.type' -ne "#microsoft.graph.itemAttachment")
            {
                Log "Content bytes empty" Red
            } else {
                Log "Content bytes not available (expected for item attachment)"
            }
        }
        if ($RetrieveAttachmentsIndividually)
        {
            Log "Retrieving attachment by id: $($attachment.id)"
            if ($attachment.'@odata.type' -eq "#microsoft.graph.itemAttachment")
            {
                $attachmentData = GetFolderMessageById $inboxId $messageToProcess.id "attachments/$($attachment.id)/?`$expand=microsoft.graph.itemattachment/item"
                if ($null -ne $attachmentData)
                {
                    Log "Attached item size: $($attachmentData.size)"
                }
            } else {
                $attachmentData = GetFolderMessageById $inboxId $messageToProcess.id "attachments/$($attachment.id)"
            }
            if ($null -eq $attachmentData)
            {
                $attachmentData = GetMessageById $messageToProcess.id "attachments/$($attachment.id)"
                if ($null -ne $attachmentData)
                {
                    Log "Successfully obtained attachment by id (excluding folder path): $($attachment.id)" Yellow
                }
                else {
                    Log "Failed to obtain attachment by id: $($attachment.id)" Red
                }                
            } else {
                $global:individualAttachments += $attachmentData
            }
        }
    }
    Log ""
}

if ($RetrieveFullMIME)
{
    # GET /mailFolders/inbox(by Id)/MessageID/$value
    Log "Retrieving full MIME for message id: $($messageToProcess.id)"
    $global:fullMIME = GetMessageMIMEById $messageToProcess.id "`$value"
    if ($null -eq $fullMIME)
    {
        Log "Failed to obtain full MIME" Red
    }
    Log "Full MIME length: $($global:fullMIME.length)"
}

# DELETE /mailfolders/inbox(by Id)/MessageID
#DeleteFolderMessageById $inboxId $messageToProcess.id

Write-Host
Log "Completed"
Write-Host "`r`n`$messageToProcess" -NoNewline -ForegroundColor Green
Write-Host " contains the message that was processed"
Write-Host "`$updatedMessage" -NoNewline -ForegroundColor Green
Write-Host " contains the message after marked as read"
Write-Host "`$message" -NoNewline -ForegroundColor Green
Write-Host " contains the message as retrieved again after second Inbox query"
Write-Host "`$attachments" -NoNewline -ForegroundColor Green
Write-Host " contains the attachments for the message that was processed"
if ($RetrieveAttachmentsIndividually)
{
    Write-Host "`$individualAttachments" -NoNewline -ForegroundColor Green
    Write-Host " contains the attachments retrieved individually"
}
if ($RetrieveFullMIME)
{
    Write-Host "`$fullMIME" -NoNewline -ForegroundColor Green
    Write-Host " contains the full MIME for the message that was processed"
}