#
# ImportExport-Messages.ps1
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
Sample script demonstrating how to export/import messages to/from an Exchange Online mailbox using Microsoft Graph.

.DESCRIPTION
This script demonstrates how to use Microsoft Graph to:
1. Export messages from a mailbox to the file system
2. Copy messages between mailboxes  
3. Import messages from the file system to a mailbox

https://learn.microsoft.com/en-us/graph/mailbox-import-export-concept-overview

When exporting, messages are saved as JSON files with metadata (including original message ID, subject, received date/time, etc.) in a separate metadata file.
When importing, the script scans all folders under the ImportPath and imports messages to the corresponding folders in the target mailbox.

.EXAMPLE

# Export messages from Inbox and Sent Items folders of a mailbox to the file system
.\ImportExport-Messages.ps1 -AppId $clientId -TenantId $tenantId -AppSecretKey $secretKey -SourceMailbox $SourceMailbox -Folders @("Inbox","SentItems") -ExportPath "C:\temp\mbxexport"

# Export messages from Inbox folder of a mailbox to the file system, including subfolders and creating folders in target if they do not exist
.\ImportExport-Messages.ps1 -AppId $clientId -TenantId $tenantId -AppSecretKey $secretKey -SourceMailbox $SourceMailbox -TargetMailbox $TargetMailbox -Folders @("Inbox") -CreateFolders -IncludeSubfolders -ExportPath "C:\temp\mbxexport"

# Import messages from the file system to a target mailbox
.\ImportExport-Messages.ps1 -AppId $clientId -TenantId $tenantId -AppSecretKey $secretKey -TargetMailbox $TargetMailbox -ImportPath "C:\temp\mbxexport" -CreateFolders
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

    [Parameter(Mandatory=$False,HelpMessage="Source mailbox (required for export mode).")]
    [ValidateNotNullOrEmpty()]
    [string]$SourceMailbox = "",

    [Parameter(Mandatory=$False,HelpMessage="Target mailbox (required if importing items to another mailbox or from file system).")]
    [ValidateNotNullOrEmpty()]
    [string]$TargetMailbox = "",

    [Parameter(Mandatory=$False,HelpMessage="If specified, folders will be created in the target mailbox if they do not exist.")]
    [ValidateNotNullOrEmpty()]
    [switch]$CreateFolders,

    [Parameter(Mandatory=$False,HelpMessage="If specified, subfolders will be included.")]
    [ValidateNotNullOrEmpty()]
    [switch]$IncludeSubfolders,

    [Parameter(Mandatory=$False,HelpMessage="Export path (only required if exporting items to the file system).")]
    [ValidateNotNullOrEmpty()]
    [string]$ExportPath = "",

    [Parameter(Mandatory=$False,HelpMessage="Import path (only required if importing items from the file system).")]
    [ValidateNotNullOrEmpty()]
    [string]$ImportPath = "",

    [Parameter(Mandatory=$False,HelpMessage="Folders to process (required for export mode, ignored for import mode).")]
    [ValidateNotNullOrEmpty()]
    $Folders = @("Inbox")

)

$script:ScriptVersion = "1.0.1"
$scriptStartTime = [DateTime]::Now

# Parameter validation
if (![String]::IsNullOrEmpty($ImportPath))
{
    # Import mode - require TargetMailbox
    if ([String]::IsNullOrEmpty($TargetMailbox))
    {
        Write-Host "Error: TargetMailbox is required when ImportPath is specified" -ForegroundColor Red
        exit
    }
}
else
{
    # Export mode - require SourceMailbox
    if ([String]::IsNullOrEmpty($SourceMailbox))
    {
        Write-Host "Error: SourceMailbox is required for export mode" -ForegroundColor Red
        exit
    }
}

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
            $errResponse = $Error[0].ErrorDetails.Message | ConvertFrom-Json
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

# Well-known folder names from https://learn.microsoft.com/en-us/graph/api/resources/mailfolder?view=graph-rest-1.0
$script:wellKnownFolderNames = @(
    "archive",
    "clutter",
    "conflicts",
    "conversationhistory",
    "deleteditems",
    "drafts",
    "inbox",
    "junkemail",
    "localfailures",
    "msgfolderroot",
    "outbox",
    "recoverableitemsdeletions",
    "scheduled",
    "searchfolders",
    "sentitems",
    "serverfailures",
    "syncissues"
)

# Helper function to get the folder cache for the current mailbox
function GetCurrentFolderCache()
{
    # Extract mailbox identifier from current graphBaseUrl
    # Format: https://graph.microsoft.com/v1.0/users/{mailbox}/
    #      or https://graph.microsoft.com/v1.0/me/
    $mailboxKey = "unknown"
    
    if ($script:graphBaseUrl -match 'users/([^/]+)/')
    {
        $mailboxKey = $matches[1]
    }
    elseif ($script:graphBaseUrl -match '/me/')
    {
        $mailboxKey = "me"
    }
    
    # Get or create cache for this mailbox
    if (-not $script:folderCaches.ContainsKey($mailboxKey))
    {
        $script:folderCaches[$mailboxKey] = @{}
        LogVerbose "Created new folder cache for mailbox: $mailboxKey"
    }
    
    return $script:folderCaches[$mailboxKey]
}

# Function to find a folder by path (e.g., "\inbox\subfolder1")
function FindFolder($folderPath)
{
    LogVerbose "Searching for folder path: $folderPath"
    
    # Normalize the path (remove leading/trailing backslashes, convert to lowercase)
    $folderPath = $folderPath.Trim('\').ToLower()
    
    # Get the cache for the current mailbox
    $cache = GetCurrentFolderCache
    
    # Check cache first
    if ($cache.ContainsKey($folderPath))
    {
        LogVerbose "Found folder in cache: $folderPath"
        return $cache[$folderPath]
    }
    
    # Split the path into individual folder names
    $pathParts = $folderPath -split '\\'
    
    $currentFolderId = $null
    $currentPath = ""
    
    for ($i = 0; $i -lt $pathParts.Length; $i++)
    {
        $folderName = $pathParts[$i]
        
        # Build the current path being processed
        if ($i -eq 0)
        {
            $currentPath = $folderName
        }
        else
        {
            $currentPath += "\$folderName"
        }
        
        # Check if this path is already cached
        if ($cache.ContainsKey($currentPath))
        {
            $currentFolderId = $cache[$currentPath].id
            LogVerbose "Found cached folder: $currentPath (ID: $currentFolderId)"
            continue
        }
        
        # Process the current folder level
        if ($i -eq 0)
        {
            # This is the root level - check if it's a well-known folder name
            if ($script:wellKnownFolderNames -contains $folderName)
            {
                LogVerbose "Accessing well-known folder: $folderName"
                $url = "$($script:graphBaseUrl)mailFolders/$folderName"
                $result = GET $url
                
                if ($result)
                {
                    $folder = $result | ConvertFrom-Json
                    $currentFolderId = $folder.id
                    $cache[$currentPath] = $folder
                    Log "Found well-known folder '$folderName' with ID: $currentFolderId"
                }
                else
                {
                    Log "Failed to access well-known folder '$folderName'" Red
                    return $null
                }
            }
            else
            {
                Log "Invalid root folder name '$folderName'. Must be one of the well-known folder names: $($script:wellKnownFolderNames -join ', ')" Red
                return $null
            }
        }
        else
        {
            # This is a subfolder - search for it under the current parent folder
            LogVerbose "Searching for subfolder '$folderName' under parent folder ID: $currentFolderId"
            $url = "$($script:graphBaseUrl)mailFolders/$currentFolderId/childFolders?`$filter=displayName eq '$folderName'"
            $result = GET $url
            
            if ($result)
            {
                $folders = ($result | ConvertFrom-Json).value
                if ($folders.Count -gt 0)
                {
                    $folder = $folders[0]
                    $currentFolderId = $folder.id
                    $cache[$currentPath] = $folder
                    LogVerbose "Found subfolder '$folderName' with ID: $currentFolderId"
                    
                    if ($folders.Count -gt 1)
                    {
                        Log "Warning: Multiple subfolders found with name '$folderName', using first match"
                    }
                }
                else
                {
                    Log "Subfolder '$folderName' not found under parent folder" Red
                    return $null
                }
            }
            else
            {
                Log "Failed to search for subfolder '$folderName'" Red
                return $null
            }
        }
    }
    
    # Return the final folder object from cache
    if ($cache.ContainsKey($folderPath))
    {
        Log "Successfully resolved folder path: $folderPath"
        return $cache[$folderPath]
    }
    
    Log "Folder path '$folderPath' could not be resolved" Red
    return $null
}

# Function to create a folder
function CreateFolder($parentFolderId, $folderName)
{
    LogVerbose "Creating folder '$folderName' under parent folder ID: $parentFolderId"
    
    # Build the request body
    $folderRequest = @{
        displayName = $folderName
    } | ConvertTo-Json
    
    # Create the folder
    $url = "$($script:graphBaseUrl)mailFolders/$parentFolderId/childFolders"
    $result = POST $url $folderRequest
    
    if ($result)
    {
        $folder = $result | ConvertFrom-Json
        Log "Created folder '$folderName' with ID: $($folder.id)" Green
        return $folder
    }
    else
    {
        Log "Failed to create folder '$folderName'" Red
        return $null
    }
}

# Function to find or create a folder by path
function FindOrCreateFolder($folderPath)
{
    LogVerbose "Finding or creating folder path: $folderPath"
    
    # Keep the original path for folder creation (preserves case)
    $originalPath = $folderPath.Trim('\\')
    
    # Normalize the path for cache lookups (lowercase)
    $normalizedPath = $originalPath.ToLower()
    
    # Get the cache for the current mailbox
    $cache = GetCurrentFolderCache
    
    # Check cache first (using normalized path)
    if ($cache.ContainsKey($normalizedPath))
    {
        LogVerbose "Found folder in cache: $normalizedPath"
        return $cache[$normalizedPath]
    }
    
    # Split the paths into individual folder names
    $pathParts = $normalizedPath -split '\\'
    $originalPathParts = $originalPath -split '\\'
    
    $currentFolderId = $null
    $currentPath = ""
    $currentOriginalPath = ""
    
    for ($i = 0; $i -lt $pathParts.Length; $i++)
    {
        $folderName = $pathParts[$i]
        $originalFolderName = $originalPathParts[$i]
        
        # Build the current path being processed
        if ($i -eq 0)
        {
            $currentPath = $folderName
            $currentOriginalPath = $originalFolderName
        }
        else
        {
            $currentPath += "\$folderName"
            $currentOriginalPath += "\$originalFolderName"
        }
        
        # Check if this path is already cached (using normalized path)
        if ($cache.ContainsKey($currentPath))
        {
            $currentFolderId = $cache[$currentPath].id
            LogVerbose "Found cached folder: $currentPath (ID: $currentFolderId)"
            continue
        }
        
        # Process the current folder level
        if ($i -eq 0)
        {
            # This is the root level - check if it's a well-known folder name
            if ($script:wellKnownFolderNames -contains $folderName)
            {
                LogVerbose "Accessing well-known folder: $folderName"
                $url = "$($script:graphBaseUrl)mailFolders/$folderName"
                $result = GET $url
                
                if ($result)
                {
                    $folder = $result | ConvertFrom-Json
                    $currentFolderId = $folder.id
                    $cache[$currentPath] = $folder
                    LogVerbose "Found well-known folder '$folderName' with ID: $currentFolderId"
                }
                else
                {
                    Log "Failed to access well-known folder '$folderName'" Red
                    return $null
                }
            }
            else
            {
                Log "Invalid root folder name '$folderName'. Must be one of the well-known folder names: $($script:wellKnownFolderNames -join ', ')" Red
                return $null
            }
        }
        else
        {
            # This is a subfolder - search for it under the current parent folder
            LogVerbose "Searching for subfolder '$folderName' under parent folder ID: $currentFolderId"
            $url = "$($script:graphBaseUrl)mailFolders/$currentFolderId/childFolders?`$filter=displayName eq '$folderName'"
            $result = GET $url
            
            if ($result)
            {
                $folders = ($result | ConvertFrom-Json).value
                if ($folders.Count -gt 0)
                {
                    $folder = $folders[0]
                    $currentFolderId = $folder.id
                    $cache[$currentPath] = $folder
                    LogVerbose "Found subfolder '$folderName' with ID: $currentFolderId"
                    
                    if ($folders.Count -gt 1)
                    {
                        Log "Warning: Multiple subfolders found with name '$folderName', using first match" Yellow
                    }
                }
                else
                {
                    # Subfolder not found, create it if CreateFolders is set
                    if ($CreateFolders)
                    {
                        $folder = CreateFolder $currentFolderId $originalFolderName
                        if ($folder)
                        {
                            $currentFolderId = $folder.id
                            $cache[$currentPath] = $folder
                        }
                        else
                        {
                            Log "Failed to create subfolder '$originalFolderName'" Red
                            return $null
                        }
                    }
                    else
                    {
                        Log "Subfolder '$folderName' not found under parent folder" Red
                        return $null
                    }
                }
            }
            else
            {
                Log "Failed to search for subfolder '$folderName'" Red
                return $null
            }
        }
    }
    
    # Return the final folder object from cache (using normalized path)
    if ($cache.ContainsKey($normalizedPath))
    {
        Log "Successfully resolved folder path: $originalPath"
        return $cache[$normalizedPath]
    }
    
    Log "Folder path '$originalPath' could not be resolved" Red
    return $null
}

# Function to get items from a folder using the export API (beta endpoint)
function ListItemsForExport($folderId, $folderName)
{
    LogVerbose "Retrieving items for export from folder: $folderName"
    
    $items = @()
    # Use the correct mailbox export API endpoint
    # GET /admin/exchange/mailboxes/{mailboxId}/folders/{mailboxFolderId}/items
    # Only request the minimum required fields: id, receivedDateTime, and subject
    $url = "https://graph.microsoft.com/beta/admin/exchange/mailboxes/$SourceMailbox/folders/$folderId/items"
    
    do
    {
        $result = GET $url
        if ($result)
        {
            $data = $result | ConvertFrom-Json
            if ($data.value)
            {
                $items += $data.value
                Log "Retrieved $($data.value.Count) items (total: $($items.Count))"
            }
            $url = $data.'@odata.nextLink'
        }
        else
        {
            $url = $null
        }
    } while ($url)
    
    Log "Total items retrieved from '$folderName': $($items.Count)"
    return $items
}

# Function to export an item using the beta mailbox export API
function ExportItem($folderId, $itemId)
{
    LogVerbose "Exporting item: $itemId from folder: $folderId"
    
    # Stage 1: Get the item metadata
    # GET /admin/exchange/mailboxes/{mailboxId}/folders/{mailboxFolderId}/items/{mailboxItemId}
    $metadataUrl = "https://graph.microsoft.com/beta/admin/exchange/mailboxes/$SourceMailbox/folders/$folderId/items/$itemId"
    $metadata = rawGET $metadataUrl
    
    if (!$metadata)
    {
        Log "Failed to retrieve item metadata for $itemId" Red
        return $null
    }
    
    LogVerbose "Retrieved item metadata for $itemId"
    
    # Stage 2: Export the item data using the exportItems API
    # POST /admin/exchange/mailboxes/{mailboxId}/exportItems
    $exportUrl = "https://graph.microsoft.com/beta/admin/exchange/mailboxes/$SourceMailbox/exportItems"
    
    # Build the request body with the item IDs to export
    # The API expects ItemIds as an array of item IDs
    $exportRequest = @{
        ItemIds = @($itemId)
    } | ConvertTo-Json -Depth 10
    
    $exportData = POST $exportUrl $exportRequest
    
    if (!$exportData)
    {
        Log "Failed to export item data for $itemId" Red
        return $null
    }
    
    LogVerbose "Successfully exported item $itemId"
    
    # Return both metadata and export data
    return @{
        metadata = $metadata
        exportData = $exportData
    }
}

# Function to save exported item (metadata and data) to files
function SaveExportedItemToFiles($exportResult, $messageId, $folderPath, $subject, $receivedDateTime)
{
    if ([String]::IsNullOrEmpty($ExportPath))
    {
        LogVerbose "ExportPath not specified, skipping file save"
        return $false
    }
    
    if (!$exportResult -or !$exportResult.metadata -or !$exportResult.exportData)
    {
        Log "No export data to save for message $messageId" Yellow
        return $false
    }
    
    try
    {
        # Normalize the folder path (remove leading backslash)
        $normalizedFolderPath = $folderPath.Trim('\')
        
        # Create the target directory path
        $targetDirectory = Join-Path -Path $ExportPath -ChildPath $normalizedFolderPath
        
        # Create directory if it doesn't exist
        if (-not (Test-Path -Path $targetDirectory))
        {
            New-Item -ItemType Directory -Path $targetDirectory -Force | Out-Null
            LogVerbose "Created directory: $targetDirectory"
        }
        
        # Save metadata as JSON file with .metadata.json extension
        $metadataFileName = "$messageId.metadata.json"
        $metadataFilePath = Join-Path -Path $targetDirectory -ChildPath $metadataFileName
        [System.IO.File]::WriteAllText($metadataFilePath, $exportResult.metadata)
        LogVerbose "Saved metadata to: $metadataFilePath"
        
        # Save export data as JSON file
        $dataFileName = "$messageId.json"
        $dataFilePath = Join-Path -Path $targetDirectory -ChildPath $dataFileName
        [System.IO.File]::WriteAllText($dataFilePath, $exportResult.exportData)
        LogVerbose "Saved data to: $dataFilePath"
        
        Log "Saved item files: $messageId.metadata.json and $messageId.json (Subject: $subject, Received: $receivedDateTime)"
        return $true
    }
    catch
    {
        Log "Failed to save item $messageId to files: $($_.Exception.Message)" Red
        return $false
    }
}

# Function to read exported item (metadata and data) from files
function ReadExportedItemFromFiles($dataFilePath)
{
    try
    {
        # Check if the data file exists
        if (-not (Test-Path -Path $dataFilePath))
        {
            Log "Data file not found: $dataFilePath" Red
            return $null
        }
        
        # Metadata file path (same name with .metadata.json extension)
        $messageId = [System.IO.Path]::GetFileNameWithoutExtension($dataFilePath)
        $directory = [System.IO.Path]::GetDirectoryName($dataFilePath)
        $metadataFilePath = Join-Path -Path $directory -ChildPath "$messageId.metadata.json"
        
        # Read the data file
        $exportData = [System.IO.File]::ReadAllText($dataFilePath)
        LogVerbose "Read export data from: $dataFilePath"
        
        # Read metadata file (optional - for logging purposes)
        $metadata = $null
        if (Test-Path -Path $metadataFilePath)
        {
            $metadata = [System.IO.File]::ReadAllText($metadataFilePath)
            LogVerbose "Read metadata from: $metadataFilePath"
        }
        else
        {
            LogVerbose "Metadata file not found: $metadataFilePath"
        }
        
        # Return both metadata and export data
        return @{
            metadata = $metadata
            exportData = $exportData
            messageId = $messageId
        }
    }
    catch
    {
        Log "Failed to read exported item from files: $($_.Exception.Message)" Red
        return $null
    }
}

# Function to import an item to a target mailbox
function ImportItemToMailbox($targetMailbox, $folderId, $exportData)
{
    LogVerbose "Importing item to mailbox: $targetMailbox, folder: $folderId"
    
    # Step 1: Create an import session to get the upload URL
    # POST /admin/exchange/mailboxes/{mailboxId}/createImportSession
    $createSessionUrl = "https://graph.microsoft.com/beta/admin/exchange/mailboxes/$targetMailbox/createImportSession"
    
    try
    {
        # Build the request body for creating the import session
        $sessionRequest = @{
            folderId = $folderId
        } | ConvertTo-Json -Depth 10
        
        LogVerbose "Creating import session for folder: $folderId"
        $sessionResult = POST $createSessionUrl $sessionRequest
        
        if (!$sessionResult)
        {
            Log "Failed to create import session" Red
            return $false
        }
        
        # Parse the session response to get the import URL
        $sessionData = $sessionResult | ConvertFrom-Json
        $uploadUrl = $sessionData.importUrl
        
        if ([String]::IsNullOrEmpty($uploadUrl))
        {
            Log "No import URL returned from import session" Red
            return $false
        }
        
        LogVerbose "Import session created, import URL: $uploadUrl"
        
        # Step 2: Upload the exported data to the import URL
        LogVerbose "Uploading exported data to import session"
        
        # Parse the export data to extract the message data
        $exportDataObj = $exportData | ConvertFrom-Json
        
        # Log the structure for debugging
        LogVerbose "Export data structure: $($exportDataObj | ConvertTo-Json -Depth 2 -Compress)"
        
        # The export API returns: {"@odata.context": "...", "value": [{"itemId": "...", "changeKey": "...", "data": "..."}]}
        # The import API expects: {"FolderId": "...", "Mode": "create", "Data": "..."}
        
        if ($exportDataObj.value -and $exportDataObj.value.Count -gt 0)
        {
            $messageData = $exportDataObj.value[0].data
            
            if ([String]::IsNullOrEmpty($messageData))
            {
                Log "No message data found in export response" Red
                return $false
            }
            
            LogVerbose "Extracted message data (length: $($messageData.Length) characters)"
            
            # Construct the import request body per API documentation
            $importBody = @{
                FolderId = $folderId
                Mode = "create"
                Data = $messageData
            } | ConvertTo-Json -Depth 10
        }
        else
        {
            Log "Export data does not contain expected 'value' array" Red
            return $false
        }
        
        # The import URL is pre-authenticated (token in URL), so we should not include auth headers
        # Use minimal headers for the POST request
        $uploadHeaders = @{
            'Content-Type' = 'application/json'
        }
        
        $invokeParameters = @{
            Method = "POST"
            Uri = $uploadUrl
            Headers = $uploadHeaders
            Body = $importBody
            ContentType = "application/json"
        }
        
        if ($PSVersionTable.PSVersion -ge [version]7.4)
        {
            $invokeParameters.ConnectionTimeoutSeconds = $ConnectionTimeout
        }
        
        $uploadResult = Invoke-WebRequest @invokeParameters -ErrorAction Stop
        
        if ($uploadResult.StatusCode -eq 200 -or $uploadResult.StatusCode -eq 201)
        {
            Log "Successfully imported item to mailbox $targetMailbox"
            return $true
        }
        else
        {
            Log "Failed to upload data to import session (Status: $($uploadResult.StatusCode))" Red
            return $false
        }
    }
    catch
    {
        Log "Failed to import item: $($_.Exception.Message)" Red
        if ($_.ErrorDetails.Message)
        {
            LogVerbose "Error response: $($_.ErrorDetails.Message)"
        }
        return $false
    }
}

# Function to get child folders from a folder
function GetChildFolders($folderId, $folderPath)
{
    LogVerbose "Getting child folders from: $folderPath"
    
    $childFolders = @()
    $url = "$($script:graphBaseUrl)mailFolders/$folderId/childFolders"
    
    do
    {
        $result = GET $url
        if ($result)
        {
            $data = $result | ConvertFrom-Json
            if ($data.value)
            {
                foreach ($folder in $data.value)
                {
                    # Build the full path for the child folder
                    $childPath = if ($folderPath) { "$folderPath\$($folder.displayName)" } else { $folder.displayName }
                    
                    # Add folder info with full path
                    $childFolders += @{
                        id = $folder.id
                        displayName = $folder.displayName
                        path = $childPath
                    }
                }
            }
            
            # Check for paging
            if ($data.'@odata.nextLink')
            {
                $url = $data.'@odata.nextLink'
            }
            else
            {
                $url = $null
            }
        }
        else
        {
            $url = $null
        }
    } while ($url)
    
    LogVerbose "Found $($childFolders.Count) child folder(s)"
    return $childFolders
}

# Function to process a folder and optionally its subfolders
function ProcessFolderAndSubfolders($folderPath)
{
    Log "Processing folder: $folderPath"
    
    $folder = FindFolder $folderPath
    if (!$folder)
    {
        Log "Skipping folder '$folderPath' - not found" Yellow
        return
    }
    
    # Process messages in this folder
    $messages = ListItemsForExport $folder.id $folderPath
    
    if ($messages.Count -gt 0)
    {
        Log "Found $($messages.Count) message(s) in '$folderPath'"

        # Validate target mailbox folder once per folder (if importing)
        $targetFolder = $null
        if (![String]::IsNullOrEmpty($TargetMailbox))
        {
            LogVerbose "Validating target folder in mailbox: $TargetMailbox"
            
            # Temporarily update graphBaseUrl to target mailbox
            # (GetCurrentFolderCache will automatically use the target mailbox cache)
            $savedGraphBaseUrl = $script:graphBaseUrl
            $script:graphBaseUrl = "https://graph.microsoft.com/v1.0/users/$TargetMailbox/"
            
            # Use FindOrCreateFolder if CreateFolders is set, otherwise use FindFolder
            if ($CreateFolders)
            {
                $targetFolder = FindOrCreateFolder $folderPath
            }
            else
            {
                $targetFolder = FindFolder $folderPath
            }
            
            # Restore original graphBaseUrl
            $script:graphBaseUrl = $savedGraphBaseUrl
            
            if ($targetFolder)
            {
                Log "Target folder '$folderPath' validated in mailbox $TargetMailbox (ID: $($targetFolder.id))"
            }
            else
            {
                Log "Target folder '$folderPath' not found in mailbox $TargetMailbox, import will be skipped" Yellow
            }
        }

        # Process messages for export
        $exportedCount = 0
        $exportFailedCount = 0
        $importSuccessCount = 0
        $importFailedCount = 0
        $saveSuccessCount = 0
        $saveFailedCount = 0
        
        foreach ($message in $messages)
        {
            LogVerbose "Processing message: $($message.subject) (Received: $($message.receivedDateTime))"
            
            # Export the message content (two-stage process: metadata + data)
            $exportResult = ExportItem $folder.id $message.id
            
            if ($exportResult)
            {
                $exportedCount++
                $script:statistics.ExportSuccess++
                
                # Save to file system if ExportPath is specified
                if (![String]::IsNullOrEmpty($ExportPath))
                {
                    $saved = SaveExportedItemToFiles $exportResult $message.id $folderPath $message.subject $message.receivedDateTime
                    if ($saved)
                    {
                        $saveSuccessCount++
                        $script:statistics.SaveSuccess++
                    }
                    else
                    {
                        $saveFailedCount++
                        $script:statistics.SaveFailed++
                    }
                }

                # Import to target mailbox if folder validation succeeded
                if ($targetFolder)
                {
                    $imported = ImportItemToMailbox $TargetMailbox $targetFolder.id $exportResult.exportData
                    if ($imported)
                    {
                        $importSuccessCount++
                        $script:statistics.ImportSuccess++
                        LogVerbose "Successfully imported message $($message.id) to target mailbox"
                    }
                    else
                    {
                        $importFailedCount++
                        $script:statistics.ImportFailed++
                        Log "Failed to import message $($message.id) to target mailbox" Red
                    }
                }
            }
            else
            {
                $exportFailedCount++
                $script:statistics.ExportFailed++
                Log "Failed to export message $($message.id): $($message.subject)" Red
            }
        }
        
        # Update folders processed count
        $script:statistics.FoldersProcessed++
        
        # Build summary message for this folder
        $summaryParts = @()
        $summaryParts += "Exports: $exportedCount succeeded, $exportFailedCount failed"
        
        if (![String]::IsNullOrEmpty($ExportPath))
        {
            $summaryParts += "Saves: $saveSuccessCount succeeded, $saveFailedCount failed"
        }
        
        if (![String]::IsNullOrEmpty($TargetMailbox))
        {
            $summaryParts += "Imports: $importSuccessCount succeeded, $importFailedCount failed"
        }
        
        $summary = $summaryParts -join " | "
        Log "Complete for '$folderPath': $summary" $(if ($exportFailedCount -eq 0 -and $saveFailedCount -eq 0 -and $importFailedCount -eq 0) { "Green" } else { "Yellow" })
    }
    else
    {
        Log "No messages found in folder '$folderPath'" Yellow
    }
    
    # Process subfolders if IncludeSubfolders is specified
    if ($IncludeSubfolders)
    {
        $childFolders = GetChildFolders $folder.id $folderPath
        
        if ($childFolders.Count -gt 0)
        {
            LogVerbose "Processing $($childFolders.Count) subfolder(s) of '$folderPath'"
            
            foreach ($childFolder in $childFolders)
            {
                # Recursively process each child folder
                ProcessFolderAndSubfolders $childFolder.path
            }
        }
        else
        {
            LogVerbose "No subfolders found in '$folderPath'"
        }
    }
}

# Function to import all exported items from the file system
function ImportFromFileSystem()
{
    if ([String]::IsNullOrEmpty($ImportPath))
    {
        Log "ImportPath not specified" Red
        return
    }
    
    if ([String]::IsNullOrEmpty($TargetMailbox))
    {
        Log "TargetMailbox must be specified when importing from file system" Red
        return
    }
    
    if (-not (Test-Path -Path $ImportPath))
    {
        Log "ImportPath does not exist: $ImportPath" Red
        return
    }
    
    Log "Importing messages from: $ImportPath to mailbox: $TargetMailbox"
    
    # Update graphBaseUrl to target mailbox
    $script:graphBaseUrl = "https://graph.microsoft.com/v1.0/users/$TargetMailbox/"
    
    # Find all .json files (excluding .metadata.json files)
    $dataFiles = Get-ChildItem -Path $ImportPath -Filter "*.json" -Recurse | Where-Object { $_.Name -notlike "*.metadata.json" }
    
    if ($dataFiles.Count -eq 0)
    {
        Log "No exported message files found in $ImportPath" Yellow
        return
    }
    
    Log "Found $($dataFiles.Count) message file(s) to import"
    
    # Group files by folder path
    $folderGroups = @{}
    foreach ($file in $dataFiles)
    {
        # Get the relative path from ImportPath
        $relativePath = $file.DirectoryName.Substring($ImportPath.Length).TrimStart('\', '/')
        
        # Normalize path (use backslash as separator)
        $folderPath = $relativePath -replace '/', '\'
        
        if ([String]::IsNullOrEmpty($folderPath))
        {
            Log "Skipping file at root level (no folder path): $($file.Name)" Yellow
            continue
        }
        
        if (-not $folderGroups.ContainsKey($folderPath))
        {
            $folderGroups[$folderPath] = @()
        }
        
        $folderGroups[$folderPath] += $file
    }
    
    Log "Messages organized into $($folderGroups.Count) folder(s)"
    
    # Process each folder
    foreach ($folderPath in $folderGroups.Keys)
    {
        Log "Processing folder: $folderPath"
        
        # Find or create the target folder
        $targetFolder = $null
        if ($CreateFolders)
        {
            $targetFolder = FindOrCreateFolder $folderPath
        }
        else
        {
            $targetFolder = FindFolder $folderPath
        }
        
        if (-not $targetFolder)
        {
            Log "Target folder '$folderPath' not found in mailbox $TargetMailbox, skipping $($folderGroups[$folderPath].Count) message(s)" Yellow
            $script:statistics.ImportFailed += $folderGroups[$folderPath].Count
            continue
        }
        
        Log "Target folder '$folderPath' validated (ID: $($targetFolder.id))"
        
        # Import each message file in this folder
        $importSuccessCount = 0
        $importFailedCount = 0
        
        foreach ($file in $folderGroups[$folderPath])
        {
            LogVerbose "Importing file: $($file.FullName)"
            
            # Read the exported item from files
            $exportedItem = ReadExportedItemFromFiles $file.FullName
            
            if (-not $exportedItem)
            {
                Log "Failed to read exported item from: $($file.FullName)" Red
                $importFailedCount++
                $script:statistics.ImportFailed++
                continue
            }
            
            # Extract subject and received date from metadata (if available)
            $subject = "Unknown"
            $receivedDateTime = "Unknown"
            if ($exportedItem.metadata)
            {
                try
                {
                    $metadataObj = $exportedItem.metadata | ConvertFrom-Json
                    if ($metadataObj.subject)
                    {
                        $subject = $metadataObj.subject
                    }
                    if ($metadataObj.receivedDateTime)
                    {
                        $receivedDateTime = $metadataObj.receivedDateTime
                    }
                }
                catch
                {
                    LogVerbose "Failed to parse metadata: $($_.Exception.Message)"
                }
            }
            
            LogVerbose "Importing message: $subject (Received: $receivedDateTime)"
            
            # Import the item to the target mailbox
            $imported = ImportItemToMailbox $TargetMailbox $targetFolder.id $exportedItem.exportData
            
            if ($imported)
            {
                $importSuccessCount++
                $script:statistics.ImportSuccess++
                Log "Imported: $($file.Name) (Subject: $subject)"
            }
            else
            {
                $importFailedCount++
                $script:statistics.ImportFailed++
                Log "Failed to import: $($file.Name) (Subject: $subject)" Red
            }
        }
        
        # Update folders processed count
        $script:statistics.FoldersProcessed++
        
        Log "Complete for '$folderPath': Imports: $importSuccessCount succeeded, $importFailedCount failed" $(if ($importFailedCount -eq 0) { "Green" } else { "Yellow" })
    }
}

################
# Main code
################

# Cache for folder paths and their IDs (per mailbox)
$script:folderCaches = @{}

# Global statistics tracking
$script:statistics = @{
    ExportSuccess = 0
    ExportFailed = 0
    ImportSuccess = 0
    ImportFailed = 0
    SaveSuccess = 0
    SaveFailed = 0
    FoldersProcessed = 0
}

# Check if we're importing from file system or exporting
if (![String]::IsNullOrEmpty($ImportPath))
{
    # Import mode - read from file system and import to target mailbox
    Log "Operating in IMPORT mode (from file system)"
    ImportFromFileSystem
}
else
{
    # Export mode - export from source mailbox
    Log "Operating in EXPORT mode (from source mailbox)"
    
    # Update the base Graph URL to use the source mailbox
    if (![String]::IsNullOrEmpty($SourceMailbox))
    {
        if ($SourceMailbox -eq "me" -and [String]::IsNullOrEmpty($AppSecretKey))
        {
            # Using delegate permissions with "me"
            $script:graphBaseUrl = "https://graph.microsoft.com/v1.0/me/"
        }
        else
        {
            # Using specific mailbox (email address or user ID)
            $script:graphBaseUrl = "https://graph.microsoft.com/v1.0/users/$SourceMailbox/"
        }
        Log "Using source mailbox: $SourceMailbox"
    }
    
    # Process each folder specified in $Folders parameter
    foreach ($folderName in $Folders)
    {
        ProcessFolderAndSubfolders $folderName
    }
}

# Log final statistics
Log ""
Log "========================================" Cyan
Log "         OPERATION SUMMARY" Cyan
Log "========================================" Cyan
Log "Folders processed: $($script:statistics.FoldersProcessed)"
Log ""
Log "Export Operations:"
Log "  Successful: $($script:statistics.ExportSuccess)" $(if ($script:statistics.ExportSuccess -gt 0) { "Green" } else { "Gray" })
Log "  Failed:     $($script:statistics.ExportFailed)" $(if ($script:statistics.ExportFailed -gt 0) { "Red" } else { "Gray" })

if (![String]::IsNullOrEmpty($ExportPath))
{
    Log ""
    Log "File System Save Operations:"
    Log "  Successful: $($script:statistics.SaveSuccess)" $(if ($script:statistics.SaveSuccess -gt 0) { "Green" } else { "Gray" })
    Log "  Failed:     $($script:statistics.SaveFailed)" $(if ($script:statistics.SaveFailed -gt 0) { "Red" } else { "Gray" })
}

if (![String]::IsNullOrEmpty($TargetMailbox))
{
    Log ""
    Log "Import Operations:"
    Log "  Successful: $($script:statistics.ImportSuccess)" $(if ($script:statistics.ImportSuccess -gt 0) { "Green" } else { "Gray" })
    Log "  Failed:     $($script:statistics.ImportFailed)" $(if ($script:statistics.ImportFailed -gt 0) { "Red" } else { "Gray" })
}

Log "========================================" Cyan

# Determine overall status
$totalFailures = $script:statistics.ExportFailed + $script:statistics.SaveFailed + $script:statistics.ImportFailed
if ($totalFailures -eq 0)
{
    Log "Completed successfully - no failures" Green
}
else
{
    Log "Completed with $totalFailures total failure(s)" Yellow
}

Log "Script finished in $([DateTime]::Now.Subtract($scriptStartTime).ToString())" Green