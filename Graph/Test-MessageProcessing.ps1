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

$script:ScriptVersion = "1.0.2"
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
        Write-Host "Failed to renew OAuth token" -ForegroundColor Red
        exit # Failed to renew the token
    }
    Write-Host "Successfully renewed OAuth token"
}

function UpdateHeaders()
{
    # Prepare the request headers

    if ($null -eq $script:headers -or $script.headers.Count -eq 0)
    {
        $script:headers = @{
            'Authorization'  = "$($oauth.token_type) $($oauth.access_token)";
            'Content-Type'   = 'application/json';
        }    
    }

    if ($script:token_expires -lt (Get-Date).AddMinutes(5))
    {
        RenewOAuthToken
    }

}

function TraceInvokeRestMethod($method, $url, $body)
{
    UpdateHeaders
    Write-Verbose "Calling URL: $url"
    if ($TraceGraphCalls)
    {
        "$method $url" | out-file $script:traceFile -Append
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
        if ($PSVersionTable.PSVersion.Major -ge 7 -and $PSVersionTable.PSVersion.Minor -ge 4)
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
        $result = Invoke-WebRequest @invokeParameters
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
            if ($result) {
                if ($result.RawContent) {
                    "" | out-file $script:traceFile -Append
                    $result.RawContent | out-file $script:traceFile -Append
                    "`r`n" | out-file $script:traceFile -Append
                }
                else {
                    $result | out-file $script:traceFile -Append
                }                
            }
            else {
                if ($Error[0].Exception.Response)
                {
                    $headers = $Error[0].Exception.Response.Headers
                    if ($headers -and $headers.Count -gt 0) {
                        for ($i=0; $i -lt $headers.Count; $i++)
                        {
                            "$($headers.Keys[$i]): $($headers[$i])" | out-file $script:traceFile -Append
                        }
                        "" | out-file $script:traceFile -Append
                    }
                    else {
                        Write-Host "No response headers returned" -ForegroundColor Red
                    }
                }
                else {
                    Write-Host "No further error details available" -ForegroundColor Red
                }                
            }
            $Error[0] | out-file $script:traceFile -Append
        }
        Write-Host "$method failed to URL: $url" -ForegroundColor Red
        exit
    }
    return $result.Content
}

function GET($url)
{
    $result = TraceInvokeRestMethod "GET" $url $null
    return ConvertFrom-Json $result
}

function rawGET($url)
{
    return TraceInvokeRestMethod "GET" $url $null
}

function PATCH($url, $body)
{
    $result = TraceInvokeRestMethod "PATCH" $url $body
    return ConvertFrom-Json $result
}

function DELETE($url)
{
    $result = TraceInvokeRestMethod "DELETE" $url $null
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
        Write-Host "Failed to obtain OAuth token" -ForegroundColor Red
        exit # Failed to obtain a token
    }
    $graphBaseUrl = "https://graph.microsoft.com/v1.0/me/"
}
else
{
    GetAppAuthToken
}
Write-Host "Successfully obtained OAuth token"

# GET /mailFolders/?$filter=displayName+eq+'INBOX'
$inbox = GetMailboxFolder("?`$filter=displayName eq 'INBOX'")
if ($null -eq $inbox)
{
    Write-Host "Failed to obtain inbox folder" -ForegroundColor Red
    exit # Failed to obtain inbox folder
}
$inboxId = $inbox.value.id
Write-Host "Inbox folder id: $($inboxId)"

# GET /mailfolders/inbox/messages/?$filter=isRead+eq+false&$count=true&$top=20&$orderBy=receivedDateTime&select=internetMessageId,id
$unreadMessages = GetFolderMessages $inboxId "?`$filter=isRead eq false&`$count=true&`$top=20&`$orderBy=receivedDateTime&select=internetMessageId,id"
if ($null -eq $unreadMessages)
{
    Write-Host "Failed to obtain unread messages" -ForegroundColor Red
    exit # Failed to obtain unread messages
}
if ($unreadMessages.value.Count -eq 0)
{
    Write-Host "No unread messages to process" -ForegroundColor Yellow
    exit # No unread messages to process
}
Write-Host "$($unreadMessages.value.Count) unread messages found"

$global:messageToProcess = $unreadMessages.value[0]

# PATCH message (mark as read)
if ($null -eq $messageToProcess)
{
    Write-Host "Failed to obtain message to process" -ForegroundColor Red
    exit # Failed to obtain message to process
}
if (!$DoNotMarkUnread)
{
    $global:updatedMessage = MarkMessageAsRead $messageToProcess.id
    if ($null -eq $updatedMessage)
    {
        Write-Host "Failed to update message" -ForegroundColor Red
        exit # Failed to update message
    }
    Write-Host "Message marked as read: $($updatedMessage.subject)"
}

# GET /mailFolders/?$filter=displayName+eq+'INBOX'
$inbox = GetMailboxFolder("?`$filter=displayName eq 'INBOX'")
if ($null -eq $inbox)
{
    Write-Host "Failed to obtain inbox folder" -ForegroundColor Red
    exit # Failed to obtain inbox folder
}
$inboxId = $inbox.value.id
Write-Host "Inbox folder id: $($inboxId)"

# GET /mailfolders/inbox(by Id)/MessageID?$select=*
$message = GetFolderMessageById $inboxId $messageToProcess.id "?`$select=*"
if ($null -eq $message)
{
    Write-Host "Failed to obtain message by id: $($messageToProcess.id)" -ForegroundColor Red
    exit # Failed to obtain message
}

# GET /mailfolders/inbox(by Id)/MessageID/attachments
$global:attachments = GetFolderMessageByIdAttachments $inboxId $messageToProcess.id
if ($null -eq $attachments)
{
    Write-Host "Failed to obtain message attachments" -ForegroundColor Red
    exit # Failed to obtain message attachments
}
Write-Host "$($attachments.value.Count) attachments found"
$global:individualAttachments = @()

if ($AnalyseAttachments)
{
    foreach ($attachment in $attachments.value)
    {
        Write-Host
        Write-Host "Attachment type: $($attachment.'@odata.type')"
        Write-Host "Name: $($attachment.name)"
        Write-Host "Id: $($attachment.id)"
        Write-Host "Size: $($attachment.size)"
        Write-Host "Content type: $($attachment.contentType)"
        Write-Host "Content id: $($attachment.contentId)"
        Write-Host "Content location: $($attachment.contentLocation)"
        if ($attachment.contentBytes) {
            Write-Host "Content bytes length: $($attachment.contentBytes.length)"
        } else {
            if ($attachment.'@odata.type' -ne "#microsoft.graph.itemAttachment")
            {
                Write-Host "Content bytes empty" -ForegroundColor Red
            } else {
                Write-Host "Content bytes not available (expected for item attachment)"
            }
        }
        if ($RetrieveAttachmentsIndividually)
        {
            Write-Host "Retrieving attachment by id: $($attachment.id)"
            if ($attachment.'@odata.type' -eq "#microsoft.graph.itemAttachment")
            {
                $attachmentData = GetFolderMessageById $inboxId $messageToProcess.id "attachments/$($attachment.id)/?`$expand=microsoft.graph.itemattachment/item"
                if ($null -ne $attachmentData)
                {
                    Write-Host "Attached item size: $($attachmentData.size)"
                }
            } else {
                $attachmentData = GetFolderMessageById $inboxId $messageToProcess.id "attachments/$($attachment.id)"
            }
            if ($null -eq $attachmentData)
            {
                $attachmentData = GetMessageById $messageToProcess.id "attachments/$($attachment.id)"
                if ($null -ne $attachmentData)
                {
                    Write-Host "Successfully obtained attachment by id (excluding folder path): $($attachment.id)" -ForegroundColor Yellow
                }
                else {
                    Write-Host "Failed to obtain attachment by id: $($attachment.id)" -ForegroundColor Red
                }                
            } else {
                $global:individualAttachments += $attachmentData
            }
        }
    }
    Write-Host
}

if ($RetrieveFullMIME)
{
    # GET /mailFolders/inbox(by Id)/MessageID/$value
    Write-Host "Retrieving full MIME for message id: $($messageToProcess.id)"
    $global:fullMIME = GetMessageMIMEById $messageToProcess.id "`$value"
    if ($null -eq $fullMIME)
    {
        Write-Host "Failed to obtain full MIME" -ForegroundColor Red
    }
    Write-Host "Full MIME length: $($global:fullMIME.length)"
}

# DELETE /mailfolders/inbox(by Id)/MessageID
#DeleteFolderMessageById $inboxId $messageToProcess.id

Write-Host "`r`nCompleted`r`n"
Write-Host "`$messageToProcess" -NoNewline -ForegroundColor Green
Write-Host " contains the message that was processed"
Write-Host "`$updatedMessage" -NoNewline -ForegroundColor Green
Write-Host " contains the message after marked as read"
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