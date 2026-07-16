#
# Test-DeleteMessages.ps1
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
Tests message copy and delete operations via Microsoft Graph API, measuring latency and performance.

.DESCRIPTION
This script performs a test of message copy and delete operations in a user's mailbox:
1. Creates a test folder with subfolders in the Inbox
2. Copies a specified number of messages from Inbox to the subfolders (cycling through source messages if needed)
3. Optionally pauses before deletion
4. Deletes all copied messages
5. Cleans up test folders
6. Reports detailed timing statistics (min, max, average latency) for both copy and delete operations

The script can be helpful for performance testing, latency measurement, and verifying message operations via Graph API.

.EXAMPLE
.\Test-DeleteMessages.ps1  -AppId $clientId -AppSecretKey $secretKey -TenantId $tenantId -User $Mailbox -MessageCount 100 -PauseBeforeDelete -TraceGraphCalls

Copies 100 messages from Inbox to test subfolders, pauses for confirmation, then deletes all copied messages and displays performance statistics.

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

<# Mailbox.ps1 %PARAMS_START% #>
    [Parameter(Mandatory=$False,HelpMessage="Mailbox (me to access own mailbox when using delegate permissions).")]
    [ValidateNotNullOrEmpty()]
    [string]$Mailbox = "me",
<# Mailbox.ps1 %PARAMS_END% #>

    [Parameter(Mandatory=$True,HelpMessage="Number of messages to copy and delete in the test.")]
    [ValidateRange(1,10000)]
    [int]$MessageCount,

    [Parameter(Mandatory=$False,HelpMessage="If specified, script will pause before deleting messages.")]
    [switch]$PauseBeforeDelete,

    [Parameter(Mandatory=$False,HelpMessage="Maximum number of concurrent operations (default is 4).")]
    [ValidateRange(1,20)]
    [int]$MaxConcurrency = 4
)

$script:ScriptVersion = "1.0.1"

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
            MaximumRedirection = 0  # Disable automatic redirects so we can handle them manually
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
        # Check for HTTP 308 Permanent Redirect (archive mailbox redirect)
        if ($_.Exception.Response -and $_.Exception.Response.StatusCode -eq 308)
        {
            $redirectUrl = $_.Exception.Response.Headers['Location']
            if ($redirectUrl)
            {
                LogVerbose "Received HTTP 308 redirect to: $redirectUrl"
                
                if ($TraceGraphCalls)
                {
                    "" | out-file $script:traceFile -Append
                    "HTTP 308 Permanent Redirect" | out-file $script:traceFile -Append
                    "Location: $redirectUrl" | out-file $script:traceFile -Append
                    "Following redirect..." | out-file $script:traceFile -Append
                    "`r`n" | out-file $script:traceFile -Append
                }
                
                # Follow the redirect by making a new request
                return TraceInvokeRestMethod $method $redirectUrl $body
            }
        }
        
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
    }
    return $result
}

function rawGET($url)
{
    return TraceInvokeRestMethod "GET" $url $null
}

<# Graph.ps1 %FUNCTIONS_END% #>

<# Mailbox.ps1 %FUNCTIONS_START% #>
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


# Main code

# Statistics tracking (use thread-safe collections for parallel processing)
$script:copyLatencies = [System.Collections.Concurrent.ConcurrentBag[double]]::new()
$script:deleteLatencies = [System.Collections.Concurrent.ConcurrentBag[double]]::new()
$script:copiedMessages = [System.Collections.Concurrent.ConcurrentBag[hashtable]]::new()

# Step 1: Get Inbox folder
Log "Getting Inbox folder..." Cyan
$inbox = GetMailboxFolder "?`$filter=displayName eq 'Inbox'"
if ($null -eq $inbox -or $inbox.value.Count -eq 0)
{
    Log "Failed to retrieve Inbox folder" Red
    exit
}
$inboxId = $inbox.value[0].id
Log "Inbox folder ID: $inboxId" Green

# Step 2: Create test folder structure
Log "Creating test folder structure..." Cyan
$testFolderName = "TestDeleteMessages_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$testFolder = CreateMailboxFolder $inboxId $testFolderName
$stopwatch.Stop()
if ($null -eq $testFolder)
{
    Log "Failed to create test folder" Red
    exit
}
Log "Created test folder '$testFolderName' (ID: $($testFolder.id)) in $($stopwatch.Elapsed.TotalMilliseconds) ms" Green

# Create subfolders
$subfolderNames = @("Subfolder1", "Subfolder2", "Subfolder3")
$subfolders = @()
foreach ($subfolderName in $subfolderNames)
{
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $subfolder = CreateMailboxFolder $testFolder.id $subfolderName
    $stopwatch.Stop()
    if ($null -eq $subfolder)
    {
        Log "Failed to create subfolder '$subfolderName'" Red
        exit
    }
    $subfolders += $subfolder
    Log "Created subfolder '$subfolderName' (ID: $($subfolder.id)) in $($stopwatch.Elapsed.TotalMilliseconds) ms" Green
}

# Step 3: Get messages from Inbox
Log "Retrieving messages from Inbox..." Cyan
$messages = GetFolderMessages $inboxId "?`$top=999&`$select=id,subject"
if ($null -eq $messages -or $messages.value.Count -eq 0)
{
    Log "No messages found in Inbox. Please ensure there are messages to copy." Red
    exit
}
$sourceMessages = $messages.value
Log "Found $($sourceMessages.Count) messages in Inbox" Green

# Step 4: Copy messages to subfolders
Log "Copying $MessageCount messages to subfolders (parallel with max concurrency: $MaxConcurrency)..." Cyan
$copyTotalStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# Create work items for parallel processing
$copyWorkItems = @()
for ($i = 0; $i -lt $MessageCount; $i++)
{
    $sourceMessage = $sourceMessages[$i % $sourceMessages.Count]
    $targetFolder = $subfolders[$i % $subfolders.Count]
    $copyWorkItems += @{
        SourceMessageId = $sourceMessage.id
        TargetFolderId = $targetFolder.id
        Subject = $sourceMessage.subject
        Index = $i
    }
}

# Process copies in parallel
$copyWorkItems | ForEach-Object -ThrottleLimit $MaxConcurrency -Parallel {
    $workItem = $_
    
    # Import necessary variables into parallel runspace
    $graphBaseUrl = $using:graphBaseUrl
    $oauth = $using:oauth
    $ConnectionTimeout = $using:ConnectionTimeout
    $copyLatencies = $using:copyLatencies
    $copiedMessages = $using:copiedMessages
    $MessageCount = $using:MessageCount
    
    # Inline simplified HTTP POST for message copy
    $url = $graphBaseUrl + "messages/$($workItem.SourceMessageId)/copy"
    $body = @{
        destinationId = $workItem.TargetFolderId
    } | ConvertTo-Json
    
    $headers = @{
        'Authorization'  = "$($oauth.token_type) $($oauth.access_token)"
        'Content-Type'   = 'application/json'
    }
    
    $invokeParameters = @{
        Method = 'POST'
        Uri = $url
        Headers = $headers
        Body = $body
        ContentType = 'application/json'
    }
    
    if ($PSVersionTable.PSVersion -ge [version]7.4)
    {
        $invokeParameters.ConnectionTimeoutSeconds = $ConnectionTimeout
    }
    
    # Copy the message
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    try
    {
        $result = Invoke-WebRequest @invokeParameters -ErrorAction Stop
        $stopwatch.Stop()
        $latency = $stopwatch.Elapsed.TotalMilliseconds
        
        if ($result.Content)
        {
            $copiedMessage = ConvertFrom-Json $result.Content
            $copyLatencies.Add($latency)
            $copiedMessages.Add(@{
                Id = $copiedMessage.id
                FolderId = $workItem.TargetFolderId
                Subject = $workItem.Subject
            })
            
            # Progress reporting
            $current = $copiedMessages.Count
            if ($current % 10 -eq 0)
            {
                Write-Host "Copied $current/$MessageCount messages (latest: $([math]::Round($latency, 2)) ms)" -ForegroundColor Gray
            }
        }
    }
    catch
    {
        $stopwatch.Stop()
        # Silently continue on errors in parallel processing
    }
}

$copyTotalStopwatch.Stop()
$messagesCopied = $script:copiedMessages.Count
Log "Successfully copied $messagesCopied messages in $([math]::Round($copyTotalStopwatch.Elapsed.TotalSeconds, 2)) seconds" Green
$copyLatenciesArray = $script:copyLatencies.ToArray()
Log "Copy latency - Min: $([math]::Round(($copyLatenciesArray | Measure-Object -Minimum).Minimum, 2)) ms, Max: $([math]::Round(($copyLatenciesArray | Measure-Object -Maximum).Maximum, 2)) ms, Avg: $([math]::Round(($copyLatenciesArray | Measure-Object -Average).Average, 2)) ms" Cyan

# Step 5: Optional pause before deletion
if ($PauseBeforeDelete)
{
    Log "Pausing before deletion. Press any key to continue..." Yellow
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    Log "Resuming..." Green
}

# Step 6: Delete copied messages
Log "Deleting copied messages (parallel with max concurrency: $MaxConcurrency)..." Cyan
$deleteTotalStopwatch = [System.Diagnostics.Stopwatch]::StartNew()

$script:copiedMessages.ToArray() | ForEach-Object -ThrottleLimit $MaxConcurrency -Parallel {
    $copiedMsg = $_
    
    # Import necessary variables into parallel runspace
    $graphBaseUrl = $using:graphBaseUrl
    $oauth = $using:oauth
    $ConnectionTimeout = $using:ConnectionTimeout
    $deleteLatencies = $using:deleteLatencies
    $messagesCopied = $using:messagesCopied
    
    # Inline simplified HTTP DELETE for message deletion
    $url = $graphBaseUrl + "mailFolders/$($copiedMsg.FolderId)/messages/$($copiedMsg.Id)"
    
    $headers = @{
        'Authorization'  = "$($oauth.token_type) $($oauth.access_token)"
        'Content-Type'   = 'application/json'
    }
    
    $invokeParameters = @{
        Method = 'DELETE'
        Uri = $url
        Headers = $headers
        ContentType = 'application/json'
    }
    
    if ($PSVersionTable.PSVersion -ge [version]7.4)
    {
        $invokeParameters.ConnectionTimeoutSeconds = $ConnectionTimeout
    }
    
    # Delete the message
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    try
    {
        $result = Invoke-WebRequest @invokeParameters -ErrorAction Stop
        $stopwatch.Stop()
        $latency = $stopwatch.Elapsed.TotalMilliseconds
        
        $deleteLatencies.Add($latency)
        
        # Progress reporting
        $current = $deleteLatencies.Count
        if ($current % 10 -eq 0)
        {
            Write-Host "Deleted $current/$messagesCopied messages (latest: $([math]::Round($latency, 2)) ms)" -ForegroundColor Gray
        }
    }
    catch
    {
        $stopwatch.Stop()
        # Silently continue on errors in parallel processing
    }
}

$deleteTotalStopwatch.Stop()
$messagesDeleted = $script:deleteLatencies.Count
Log "Successfully deleted $messagesDeleted messages in $([math]::Round($deleteTotalStopwatch.Elapsed.TotalSeconds, 2)) seconds" Green
$deleteLatenciesArray = $script:deleteLatencies.ToArray()
Log "Delete latency - Min: $([math]::Round(($deleteLatenciesArray | Measure-Object -Minimum).Minimum, 2)) ms, Max: $([math]::Round(($deleteLatenciesArray | Measure-Object -Maximum).Maximum, 2)) ms, Avg: $([math]::Round(($deleteLatenciesArray | Measure-Object -Average).Average, 2)) ms" Cyan

# Step 7: Clean up test folders
Log "Cleaning up test folders..." Cyan
foreach ($subfolder in $subfolders)
{
    $url = $script:graphBaseUrl + "mailFolders/$($subfolder.id)"
    try
    {
        $null = DELETE $url
        Log "Deleted subfolder: $($subfolder.displayName)" Gray
    }
    catch
    {
        Log "Failed to delete subfolder: $($subfolder.displayName)" Red
    }
}

# Delete parent test folder
$url = $script:graphBaseUrl + "mailFolders/$($testFolder.id)"
try
{
    $null = DELETE $url
    Log "Deleted test folder: $testFolderName" Green
}
catch
{
    Log "Failed to delete test folder: $testFolderName" Red
}

# Display summary statistics
$copyLatenciesArray = $script:copyLatencies.ToArray()
$deleteLatenciesArray = $script:deleteLatencies.ToArray()

Log "`n========== TEST SUMMARY ==========" Cyan
Log "Max Concurrency: $MaxConcurrency" White
Log "Messages copied: $messagesCopied" White
Log "Messages deleted: $messagesDeleted" White
Log "`nCopy Operations:" White
Log "  Total time: $([math]::Round($copyTotalStopwatch.Elapsed.TotalSeconds, 2)) seconds" White
Log "  Min latency: $([math]::Round(($copyLatenciesArray | Measure-Object -Minimum).Minimum, 2)) ms" White
Log "  Max latency: $([math]::Round(($copyLatenciesArray | Measure-Object -Maximum).Maximum, 2)) ms" White
Log "  Avg latency: $([math]::Round(($copyLatenciesArray | Measure-Object -Average).Average, 2)) ms" White
Log "`nDelete Operations:" White
Log "  Total time: $([math]::Round($deleteTotalStopwatch.Elapsed.TotalSeconds, 2)) seconds" White
Log "  Min latency: $([math]::Round(($deleteLatenciesArray | Measure-Object -Minimum).Minimum, 2)) ms" White
Log "  Max latency: $([math]::Round(($deleteLatenciesArray | Measure-Object -Maximum).Maximum, 2)) ms" White
Log "  Avg latency: $([math]::Round(($deleteLatenciesArray | Measure-Object -Average).Average, 2)) ms" White
Log "==================================`n" Cyan

Log "Completed"
