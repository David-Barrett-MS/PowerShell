#
# Rename-RoomMailboxes.ps1
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
Processes room mailbox calendar events by removing and re-adding the room to organizer's events.

.DESCRIPTION
This script processes calendar events for specified room mailboxes occurring on or after a given date.
For each event, the script locates the organizer's copy of the event and performs the following actions:
1. Removes the room mailbox from the event's attendees list
2. Saves the updated event
3. Re-adds the room mailbox to the event's attendees list
4. Saves the event again

This can be useful for refreshing room bookings or forcing synchronization of room calendar entries.

The script uses application permissions and requires access to all calendars in the organization.
Events are identified across mailboxes using the iCalUId property to ensure the correct event is updated.

.PREREQUISITES
Entra ID App Registration:
- Register an application in the Entra ID (Azure AD) portal
- Create a client secret for the application
- Configure the following API permissions:
  * Microsoft Graph > Application Permissions > Calendars.ReadWrite
  * Microsoft Graph > Application Permissions > User.Read.All (if needed to resolve user information)
- Grant admin consent for the permissions in your tenant
- Note the Application (client) ID, Directory (tenant) ID, and client secret value

IMPORTANT: This script requires application permissions (client credential flow) and will not work with delegate permissions.
The AppSecretKey parameter must be provided to authenticate using application permissions.

.EXAMPLE
.\Rename-RoomMailboxes.ps1 -AppId "12345678-1234-1234-1234-123456789012" -AppSecretKey "your-secret-key" -TenantId "contoso.onmicrosoft.com" -RoomMailboxes @("confroom1@contoso.com") -StartDate "2026-06-24"

Processes all calendar events for confroom1@contoso.com starting from June 24, 2026, removing and re-adding the room to each organizer's event.

.EXAMPLE
.\Rename-RoomMailboxes.ps1 -AppId "12345678-1234-1234-1234-123456789012" -AppSecretKey "your-secret-key" -TenantId "contoso.onmicrosoft.com" -RoomMailboxes @("room1@contoso.com", "room2@contoso.com", "room3@contoso.com") -StartDate "2026-07-01T09:00:00" -LogToFile

Processes multiple room mailboxes starting from July 1, 2026 at 9:00 AM, with all activity logged to a file.

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

    [Parameter(Mandatory=$True,HelpMessage="Array of room mailbox SMTP addresses to process.")]
    [ValidateNotNullOrEmpty()]
    [string[]]$RoomMailboxes,

    [Parameter(Mandatory=$True,HelpMessage="Process events occurring on or after this date (format: yyyy-MM-dd or yyyy-MM-ddTHH:mm:ss).")]
    [ValidateNotNullOrEmpty()]
    [string]$StartDate
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

<# Mailbox.ps1 %FUNCTIONS_END% #>


# Calendar-specific functions

function GetCalendarEvents($mailbox, $queryString)
{
    # Get calendar events for a specific mailbox
    $url = "https://graph.microsoft.com/v1.0/users/$mailbox/calendar/events$queryString"
    $response = GET $url
    if ($null -eq $response)
    {
        return $null
    }
    return ConvertFrom-Json ($response)
}

function GetEventById($mailbox, $eventId, $queryString)
{
    # Get a specific event by ID
    $url = "https://graph.microsoft.com/v1.0/users/$mailbox/events/$eventId$queryString"
    $response = GET $url
    if ($null -eq $response)
    {
        return $null
    }
    return ConvertFrom-Json ($response)
}

function UpdateEvent($mailbox, $eventId, $body)
{
    # Update an event using PATCH
    $url = "https://graph.microsoft.com/v1.0/users/$mailbox/events/$eventId"
    $response = PATCH $url $body
    if ($null -eq $response)
    {
        return $null
    }
    return ConvertFrom-Json ($response)
}

function FindEventByICalUId($mailbox, $iCalUId)
{
    # Find an event in a mailbox by its iCalUId
    $queryString = "?`$filter=iCalUId eq '$iCalUId'&`$select=id,subject,start,end,organizer,attendees,iCalUId"
    $response = GetCalendarEvents $mailbox $queryString
    if ($null -eq $response -or $null -eq $response.value -or $response.value.Count -eq 0)
    {
        return $null
    }
    return $response.value[0]
}

function RemoveAttendee($attendeesArray, $emailAddress)
{
    # Remove an attendee from the attendees array
    $newAttendees = @()
    foreach ($attendee in $attendeesArray)
    {
        if ($attendee.emailAddress.address -ne $emailAddress)
        {
            $newAttendees += $attendee
        }
    }
    return $newAttendees
}

function AddAttendee($attendeesArray, $emailAddress, $attendeeType = "required")
{
    # Add an attendee to the attendees array
    $newAttendee = @{
        emailAddress = @{
            address = $emailAddress
        }
        type = $attendeeType
    }
    $attendeesArray += $newAttendee
    return $attendeesArray
}

# Main code

# Verify that application permissions are being used (not delegate)
if ([String]::IsNullOrEmpty($AppSecretKey))
{
    Log "ERROR: This script requires application permissions (client credential flow)." Red
    Log "You must provide the -AppSecretKey parameter to use application permissions." Red
    Log "Delegate permissions are not supported for this script." Red
    exit
}

Log "Using application permissions (client credential flow)" Green

# Validate the start date
try
{
    $startDateTime = [DateTime]::Parse($StartDate)
    Log "Processing events from: $($startDateTime.ToString('yyyy-MM-dd HH:mm:ss'))"
}
catch
{
    Log "Invalid date format: $StartDate. Please use yyyy-MM-dd or yyyy-MM-ddTHH:mm:ss format." Red
    exit
}

# Convert to ISO 8601 format for Graph API
$startDateFilter = $startDateTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

# Process each room mailbox
$totalRooms = $RoomMailboxes.Count
$currentRoom = 0

foreach ($roomMailbox in $RoomMailboxes)
{
    $currentRoom++
    Log "`n[$currentRoom/$totalRooms] Processing room mailbox: $roomMailbox" Cyan
    
    # Get calendar events for the room mailbox that start on or after the specified date
    $filter = "`$filter=start/dateTime ge '$startDateFilter'"
    $select = "`$select=id,subject,start,end,organizer,attendees,iCalUId"
    $queryString = "?$filter&$select&`$top=999"
    
    $eventsResponse = GetCalendarEvents $roomMailbox $queryString
    
    if ($null -eq $eventsResponse -or $null -eq $eventsResponse.value)
    {
        Log "  No events found or unable to retrieve events for $roomMailbox" Yellow
        continue
    }
    
    $events = $eventsResponse.value
    $totalEvents = $events.Count
    Log "  Found $totalEvents event(s) to process"
    
    # Process pagination if there are more events
    while ($null -ne $eventsResponse.'@odata.nextLink')
    {
        $eventsResponse = GET $eventsResponse.'@odata.nextLink'
        if ($null -ne $eventsResponse)
        {
            $eventsResponse = ConvertFrom-Json $eventsResponse
            $events += $eventsResponse.value
            $totalEvents = $events.Count
            Log "  Retrieved additional events, total now: $totalEvents"
        }
    }
    
    $currentEvent = 0
    foreach ($event in $events)
    {
        $currentEvent++
        Log "  [$currentEvent/$totalEvents] Processing event: $($event.subject)" Gray
        
        # Get the organizer email address
        $organizerEmail = $event.organizer.emailAddress.address
        if ([String]::IsNullOrEmpty($organizerEmail))
        {
            Log "    Skipping: No organizer found" Yellow
            continue
        }
        
        Log "    Organizer: $organizerEmail"
        
        # Find the same event in the organizer's calendar using iCalUId
        $iCalUId = $event.iCalUId
        if ([String]::IsNullOrEmpty($iCalUId))
        {
            Log "    Skipping: No iCalUId found for event" Yellow
            continue
        }
        
        Log "    Finding event in organizer's calendar (iCalUId: $iCalUId)"
        $organizerEvent = FindEventByICalUId $organizerEmail $iCalUId
        
        if ($null -eq $organizerEvent)
        {
            Log "    Skipping: Could not find event in organizer's calendar" Yellow
            continue
        }
        
        Log "    Found event in organizer's calendar (ID: $($organizerEvent.id))"
        
        # Check if the room is actually in the attendees list
        $roomIsAttendee = $false
        foreach ($attendee in $organizerEvent.attendees)
        {
            if ($attendee.emailAddress.address -eq $roomMailbox)
            {
                $roomIsAttendee = $true
                break
            }
        }
        
        if (-not $roomIsAttendee)
        {
            Log "    Skipping: Room $roomMailbox is not an attendee of this event" Yellow
            continue
        }
        
        # Remove the room mailbox from attendees
        Log "    Removing room from attendees list"
        $updatedAttendees = RemoveAttendee $organizerEvent.attendees $roomMailbox
        
        $updateBody = @{
            attendees = $updatedAttendees
        } | ConvertTo-Json -Depth 10
        
        $updateResult = UpdateEvent $organizerEmail $organizerEvent.id $updateBody
        
        if ($null -eq $updateResult)
        {
            Log "    ERROR: Failed to remove room from event" Red
            continue
        }
        
        Log "    Successfully removed room from event" Green
        
        # Add the room mailbox back to attendees
        Log "    Adding room back to attendees list"
        $readdedAttendees = AddAttendee $updatedAttendees $roomMailbox "required"
        
        $readdBody = @{
            attendees = $readdedAttendees
        } | ConvertTo-Json -Depth 10
        
        $readdResult = UpdateEvent $organizerEmail $organizerEvent.id $readdBody
        
        if ($null -eq $readdResult)
        {
            Log "    ERROR: Failed to add room back to event" Red
            continue
        }
        
        Log "    Successfully added room back to event" Green
    }
}

Log "`nCompleted processing all room mailboxes" Green
