#
# Test-RenameRoomMailboxes.ps1
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
Creates test calendar events for testing Rename-RoomMailboxes.ps1.

.DESCRIPTION
This script creates calendar events for specified organisers with room mailboxes as locations.
For each organiser, the script creates:
- Single-instance meetings (specified by -MeetingsPerOrganiser parameter)
- Recurring meetings (specified by -RecurringMeetingsPerOrganiser parameter, default: 8)

Single-instance meetings:
- Scheduled starting from tomorrow at 9 AM, one day apart
- Room mailboxes cycle through the provided list

Recurring meetings:
- Start at 1 PM the day after tomorrow
- Recurrence patterns cycle: daily (7 occurrences), weekly (3 months), monthly (12 months)
- Every other recurring meeting includes 1-2 cancelled occurrences
- One in three recurring meetings has a time-modified occurrence (moved forward 1 hour)
- One in four recurring meetings has a location-modified occurrence (different room)

This is useful for creating test data to validate the Rename-RoomMailboxes.ps1 script functionality,
especially for testing recurring event handling.

.PREREQUISITES
Entra ID App Registration with the following permissions:
- Microsoft Graph > Application Permissions > Calendars.ReadWrite
- Microsoft Graph > Application Permissions > User.Read.All
- Grant admin consent for the permissions in your tenant

IMPORTANT: This script requires application permissions (client credential flow).
The AppSecretKey parameter must be provided.

.EXAMPLE
$organisers = @("user1@contoso.com", "user2@contoso.com")
$rooms = @("room1@contoso.com", "room2@contoso.com")
.\Test-RenameRoomMailboxes.ps1 -AppId $clientId -AppSecretKey $secretKey -TenantId $tenantId -Organisers $organisers -RoomMailboxes $rooms -MeetingsPerOrganiser 5

Creates 5 single-instance test meetings for each organiser (10 total), plus 8 recurring meetings per organiser (16 total), cycling through the room mailboxes as locations.

.EXAMPLE
.\Test-RenameRoomMailboxes.ps1 -AppId $clientId -AppSecretKey $secretKey -TenantId $tenantId -Organisers @("user@contoso.com") -RoomMailboxes @("room@contoso.com") -MeetingsPerOrganiser 10 -RecurringMeetingsPerOrganiser 12 -LogToFile

Creates 10 single-instance meetings and 12 recurring meetings for a single organiser with logging enabled.

.EXAMPLE
.\Test-RenameRoomMailboxes.ps1 -AppId $clientId -AppSecretKey $secretKey -TenantId $tenantId -Organisers @("user@contoso.com") -RoomMailboxes @("room@contoso.com") -MeetingsPerOrganiser 5 -RecurringMeetingsPerOrganiser 0

Creates only single-instance meetings, no recurring meetings.

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

    [Parameter(Mandatory=$True,HelpMessage="Array of organiser email addresses who will create the test meetings.")]
    [ValidateNotNullOrEmpty()]
    [string[]]$Organisers,

    [Parameter(Mandatory=$True,HelpMessage="Array of room mailbox SMTP addresses to use as meeting locations.")]
    [ValidateNotNullOrEmpty()]
    [string[]]$RoomMailboxes,

    [Parameter(Mandatory=$False,HelpMessage="Number of test meetings to create for each organiser.")]
    [ValidateRange(1,100)]
    [int]$MeetingsPerOrganiser = 10,

    [Parameter(Mandatory=$False,HelpMessage="Number of recurring test meetings to create for each organiser (default: 8).")]
    [ValidateRange(0,50)]
    [int]$RecurringMeetingsPerOrganiser = 8,

    [Parameter(Mandatory=$False,HelpMessage="Clear all existing calendar events from room mailboxes and organisers before creating new test meetings.")]
    [switch]$Clear,

    [Parameter(Mandatory=$False,HelpMessage="Clear all existing calendar events from room mailboxes and organisers, then exit without creating new meetings.")]
    [switch]$ClearOnly
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


# Calendar-specific functions

function GetRoomDisplayName($roomMailbox)
{
    # Get the display name of a room mailbox from Graph
    $url = "https://graph.microsoft.com/v1.0/users/$roomMailbox`?`$select=displayName"
    $response = GET $url
    
    if ($null -eq $response)
    {
        # If we can't get the display name, fall back to email address
        LogVerbose "Could not retrieve display name for $roomMailbox, using email address"
        return $roomMailbox.Split('@')[0]
    }
    
    $userObject = ConvertFrom-Json $response
    return $userObject.displayName
}

function ClearCalendarEvents($mailbox)
{
    # Delete all calendar events for a mailbox
    Log "  Retrieving events for $mailbox..." Gray
    
    $url = "https://graph.microsoft.com/v1.0/users/$mailbox/calendar/events"
    $response = GET $url
    
    if ($null -eq $response)
    {
        Log "    Could not retrieve events" Yellow
        return 0
    }
    
    $eventsResponse = ConvertFrom-Json $response
    $events = $eventsResponse.value
    $totalEvents = $events.Count
    
    # Handle pagination
    while ($null -ne $eventsResponse.'@odata.nextLink')
    {
        $response = GET $eventsResponse.'@odata.nextLink'
        if ($null -ne $response)
        {
            $eventsResponse = ConvertFrom-Json $response
            $events += $eventsResponse.value
            $totalEvents = $events.Count
        }
    }
    
    if ($totalEvents -eq 0)
    {
        Log "    No events to delete" Gray
        return 0
    }
    
    Log "    Found $totalEvents event(s), deleting..." Gray
    
    $deletedCount = 0
    $failedCount = 0
    foreach ($event in $events)
    {
        $deleteUrl = "https://graph.microsoft.com/v1.0/users/$mailbox/events/$($event.id)"
        
        # Store error count before DELETE
        $errorCountBefore = $Error.Count
        $result = DELETE $deleteUrl
        
        # If no new errors were added, deletion succeeded
        # (DELETE returns empty string for HTTP 204 success)
        if ($Error.Count -eq $errorCountBefore)
        {
            $deletedCount++
        }
        else
        {
            $failedCount++
        }
    }
    
    if ($failedCount -gt 0)
    {
        Log "    Deleted $deletedCount of $totalEvents event(s), $failedCount failed" Yellow
    }
    else
    {
        Log "    Deleted $deletedCount of $totalEvents event(s)" Green
    }
    return $deletedCount
}

function CreateCalendarEvent($organizerEmail, $subject, $startDateTime, $endDateTime, $locationEmail, $locationDisplayName, $attendeesArray)
{
    # Create a calendar event for an organiser
    $url = "https://graph.microsoft.com/v1.0/users/$organizerEmail/events"
    
    # Build location object with actual display name
    $location = @{
        displayName = $locationDisplayName
        locationType = "conferenceRoom"
        uniqueId = $locationEmail
        uniqueIdType = "directory"
    }
    
    # Build attendees array
    $attendees = @()
    
    # Add the room mailbox as a resource attendee first
    $attendees += @{
        emailAddress = @{
            address = $locationEmail
            name = $locationDisplayName
        }
        type = "resource"
    }
    
    # Add other attendees
    foreach ($attendeeEmail in $attendeesArray)
    {
        $attendees += @{
            emailAddress = @{
                address = $attendeeEmail
                name = $attendeeEmail
            }
            type = "required"
        }
    }
    
    # Build event body
    $eventBody = @{
        subject = $subject
        start = @{
            dateTime = $startDateTime.ToString("yyyy-MM-ddTHH:mm:ss")
            timeZone = "UTC"
        }
        end = @{
            dateTime = $endDateTime.ToString("yyyy-MM-ddTHH:mm:ss")
            timeZone = "UTC"
        }
        location = $location
        attendees = $attendees
        isOnlineMeeting = $false
    } | ConvertTo-Json -Depth 10
    
    $response = POST $url $eventBody
    
    if ($null -eq $response)
    {
        return $null
    }
    
    return ConvertFrom-Json $response
}

function CreateRecurringCalendarEvent($organizerEmail, $subject, $startDateTime, $endDateTime, $locationEmail, $locationDisplayName, $attendeesArray, $recurrencePattern, $recurrenceRange)
{
    # Create a recurring calendar event for an organiser
    $url = "https://graph.microsoft.com/v1.0/users/$organizerEmail/events"
    
    # Build location object with actual display name
    $location = @{
        displayName = $locationDisplayName
        locationType = "conferenceRoom"
        uniqueId = $locationEmail
        uniqueIdType = "directory"
    }
    
    # Build attendees array
    $attendees = @()
    
    # Add the room mailbox as a resource attendee first
    $attendees += @{
        emailAddress = @{
            address = $locationEmail
            name = $locationDisplayName
        }
        type = "resource"
    }
    
    # Add other attendees
    foreach ($attendeeEmail in $attendeesArray)
    {
        $attendees += @{
            emailAddress = @{
                address = $attendeeEmail
                name = $attendeeEmail
            }
            type = "required"
        }
    }
    
    # Build recurrence object
    $recurrence = @{
        pattern = $recurrencePattern
        range = $recurrenceRange
    }
    
    # Build event body
    $eventBody = @{
        subject = $subject
        start = @{
            dateTime = $startDateTime.ToString("yyyy-MM-ddTHH:mm:ss")
            timeZone = "UTC"
        }
        end = @{
            dateTime = $endDateTime.ToString("yyyy-MM-ddTHH:mm:ss")
            timeZone = "UTC"
        }
        location = $location
        attendees = $attendees
        recurrence = $recurrence
        isOnlineMeeting = $false
    } | ConvertTo-Json -Depth 10
    
    $response = POST $url $eventBody
    
    if ($null -eq $response)
    {
        return $null
    }
    
    return ConvertFrom-Json $response
}

function CancelOccurrence($organizerEmail, $seriesMasterId, $occurrenceDate)
{
    # Cancel a specific occurrence of a recurring event
    # First, get the instances to find the occurrence ID
    # Ensure we're working with UTC times
    $occurrenceDateUtc = $occurrenceDate.ToUniversalTime()
    $startDateTime = $occurrenceDateUtc.AddDays(-1).ToString('yyyy-MM-ddTHH:mm:ss')
    $endDateTime = $occurrenceDateUtc.AddDays(1).ToString('yyyy-MM-ddTHH:mm:ss')
    $instancesUrl = "https://graph.microsoft.com/v1.0/users/$organizerEmail/events/$seriesMasterId/instances?startDateTime=$startDateTime&endDateTime=$endDateTime"
    
    LogVerbose "      Querying instances between $startDateTime and $endDateTime"
    $response = GET $instancesUrl
    
    if ($null -eq $response)
    {
        Log "      Failed to get instances for cancellation (window: $startDateTime to $endDateTime)" Red
        return $false
    }
    
    $instances = ConvertFrom-Json $response
    if ($null -eq $instances.value -or $instances.value.Count -eq 0)
    {
        Log "      No occurrence found to cancel" Yellow
        return $false
    }
    
    # Find the occurrence that matches our target date/time
    $targetOccurrence = $null
    $targetDateStr = $occurrenceDateUtc.ToString('yyyy-MM-ddTHH:mm:ss')
    foreach ($instance in $instances.value)
    {
        $instanceStart = if ($instance.start.dateTime -is [DateTime]) {
            $instance.start.dateTime
        } else {
            # Parse and specify UTC kind to prevent local time conversion
            $parsed = [DateTime]::Parse($instance.start.dateTime, [System.Globalization.CultureInfo]::InvariantCulture)
            [DateTime]::SpecifyKind($parsed, [DateTimeKind]::Utc)
        }
        $instanceStartStr = $instanceStart.ToString('yyyy-MM-ddTHH:mm:ss')
        
        if ($instanceStartStr -eq $targetDateStr)
        {
            $targetOccurrence = $instance
            break
        }
    }
    
    if ($null -eq $targetOccurrence)
    {
        Log "      No occurrence found matching $targetDateStr" Yellow
        return $false
    }
    
    # Cancel the occurrence by sending a cancellation
    $occurrenceId = $targetOccurrence.id
    $cancelUrl = "https://graph.microsoft.com/v1.0/users/$organizerEmail/events/$occurrenceId/cancel"
    
    $cancelBody = @{
        comment = "Meeting cancelled"
    } | ConvertTo-Json
    
    $result = POST $cancelUrl $cancelBody
    
    if ($null -eq $result)
    {
        return $false
    }
    
    return $true
}

function ModifyOccurrenceTime($organizerEmail, $seriesMasterId, $occurrenceDate, $hourOffset)
{
    # Modify the time of a specific occurrence (creates an exception)
    # First, get the instances to find the occurrence ID
    # Ensure we're working with UTC times
    $occurrenceDateUtc = $occurrenceDate.ToUniversalTime()
    $startDateTime = $occurrenceDateUtc.AddDays(-1).ToString('yyyy-MM-ddTHH:mm:ss')
    $endDateTime = $occurrenceDateUtc.AddDays(1).ToString('yyyy-MM-ddTHH:mm:ss')
    $instancesUrl = "https://graph.microsoft.com/v1.0/users/$organizerEmail/events/$seriesMasterId/instances?startDateTime=$startDateTime&endDateTime=$endDateTime"
    
    LogVerbose "      Querying instances between $startDateTime and $endDateTime"
    $response = GET $instancesUrl
    
    if ($null -eq $response)
    {
        Log "      Failed to get instances for time modification (window: $startDateTime to $endDateTime)" Red
        return $false
    }
    
    $instances = ConvertFrom-Json $response
    if ($null -eq $instances.value -or $instances.value.Count -eq 0)
    {
        Log "      No occurrence found to modify" Yellow
        return $false
    }
    
    # Find the occurrence that matches our target date/time
    $targetOccurrence = $null
    $targetDateStr = $occurrenceDateUtc.ToString('yyyy-MM-ddTHH:mm:ss')
    foreach ($instance in $instances.value)
    {
        $instanceStart = if ($instance.start.dateTime -is [DateTime]) {
            $instance.start.dateTime
        } else {
            # Parse and specify UTC kind to prevent local time conversion
            $parsed = [DateTime]::Parse($instance.start.dateTime, [System.Globalization.CultureInfo]::InvariantCulture)
            [DateTime]::SpecifyKind($parsed, [DateTimeKind]::Utc)
        }
        $instanceStartStr = $instanceStart.ToString('yyyy-MM-ddTHH:mm:ss')
        
        if ($instanceStartStr -eq $targetDateStr)
        {
            $targetOccurrence = $instance
            break
        }
    }
    
    if ($null -eq $targetOccurrence)
    {
        Log "      No occurrence found matching $targetDateStr" Yellow
        return $false
    }
    
    # Update the occurrence time
    $occurrenceId = $targetOccurrence.id
    $occurrence = $targetOccurrence
    
    # Parse DateTime - handle both string and DateTime objects
    $startTime = if ($occurrence.start.dateTime -is [DateTime]) {
        $occurrence.start.dateTime
    } else {
        $parsed = [DateTime]::Parse($occurrence.start.dateTime, [System.Globalization.CultureInfo]::InvariantCulture)
        [DateTime]::SpecifyKind($parsed, [DateTimeKind]::Utc)
    }
    
    $endTime = if ($occurrence.end.dateTime -is [DateTime]) {
        $occurrence.end.dateTime
    } else {
        $parsed = [DateTime]::Parse($occurrence.end.dateTime, [System.Globalization.CultureInfo]::InvariantCulture)
        [DateTime]::SpecifyKind($parsed, [DateTimeKind]::Utc)
    }
    
    $newStart = $startTime.AddHours($hourOffset)
    $newEnd = $endTime.AddHours($hourOffset)
    
    $updateUrl = "https://graph.microsoft.com/v1.0/users/$organizerEmail/events/$occurrenceId"
    $updateBody = @{
        start = @{
            dateTime = $newStart.ToString("yyyy-MM-ddTHH:mm:ss")
            timeZone = $occurrence.start.timeZone
        }
        end = @{
            dateTime = $newEnd.ToString("yyyy-MM-ddTHH:mm:ss")
            timeZone = $occurrence.end.timeZone
        }
    } | ConvertTo-Json -Depth 10
    
    $result = PATCH $updateUrl $updateBody
    
    if ($null -eq $result)
    {
        return $false
    }
    
    return $true
}

function ModifyOccurrenceLocation($organizerEmail, $seriesMasterId, $occurrenceDate, $newLocationEmail, $newLocationDisplayName)
{
    # Modify the location of a specific occurrence (creates an exception)
    # First, get the instances to find the occurrence ID
    # Ensure we're working with UTC times
    $occurrenceDateUtc = $occurrenceDate.ToUniversalTime()
    $startDateTime = $occurrenceDateUtc.AddDays(-1).ToString('yyyy-MM-ddTHH:mm:ss')
    $endDateTime = $occurrenceDateUtc.AddDays(1).ToString('yyyy-MM-ddTHH:mm:ss')
    $instancesUrl = "https://graph.microsoft.com/v1.0/users/$organizerEmail/events/$seriesMasterId/instances?startDateTime=$startDateTime&endDateTime=$endDateTime"
    
    LogVerbose "      Querying instances between $startDateTime and $endDateTime"
    $response = GET $instancesUrl
    
    if ($null -eq $response)
    {
        Log "      Failed to get instances for location modification (window: $startDateTime to $endDateTime)" Red
        return $false
    }
    
    $instances = ConvertFrom-Json $response
    if ($null -eq $instances.value -or $instances.value.Count -eq 0)
    {
        Log "      No occurrence found to modify" Yellow
        return $false
    }
    
    # Find the occurrence that matches our target date/time
    $targetOccurrence = $null
    $targetDateStr = $occurrenceDateUtc.ToString('yyyy-MM-ddTHH:mm:ss')
    foreach ($instance in $instances.value)
    {
        $instanceStart = if ($instance.start.dateTime -is [DateTime]) {
            $instance.start.dateTime
        } else {
            # Parse and specify UTC kind to prevent local time conversion
            $parsed = [DateTime]::Parse($instance.start.dateTime, [System.Globalization.CultureInfo]::InvariantCulture)
            [DateTime]::SpecifyKind($parsed, [DateTimeKind]::Utc)
        }
        $instanceStartStr = $instanceStart.ToString('yyyy-MM-ddTHH:mm:ss')
        
        if ($instanceStartStr -eq $targetDateStr)
        {
            $targetOccurrence = $instance
            break
        }
    }
    
    if ($null -eq $targetOccurrence)
    {
        Log "      No occurrence found matching $targetDateStr" Yellow
        return $false
    }
    
    # Update the occurrence location
    $occurrenceId = $targetOccurrence.id
    $occurrence = $targetOccurrence
    
    # Build new location object
    $newLocation = @{
        displayName = $newLocationDisplayName
        locationType = "conferenceRoom"
        uniqueId = $newLocationEmail
        uniqueIdType = "directory"
    }
    
    # Update attendees to include new room instead of old room
    $updatedAttendees = @()
    foreach ($attendee in $occurrence.attendees)
    {
        # Skip the old room (it was a resource)
        if ($attendee.type -eq "resource")
        {
            continue
        }
        $updatedAttendees += $attendee
    }
    
    # Add the new room as a resource attendee
    $updatedAttendees += @{
        emailAddress = @{
            address = $newLocationEmail
            name = $newLocationDisplayName
        }
        type = "resource"
    }
    
    $updateUrl = "https://graph.microsoft.com/v1.0/users/$organizerEmail/events/$occurrenceId"
    $updateBody = @{
        location = $newLocation
        attendees = $updatedAttendees
    } | ConvertTo-Json -Depth 10
    
    $result = PATCH $updateUrl $updateBody
    
    if ($null -eq $result)
    {
        return $false
    }
    
    return $true
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

# Validate inputs
if ($Organisers.Count -eq 0)
{
    Log "ERROR: At least one organiser must be specified." Red
    exit
}

if ($RoomMailboxes.Count -eq 0)
{
    Log "ERROR: At least one room mailbox must be specified." Red
    exit
}

Log "Configuration:" Cyan
Log "  Organisers: $($Organisers.Count)"
Log "  Room Mailboxes: $($RoomMailboxes.Count)"
if (-not $ClearOnly)
{
    Log "  Meetings per Organiser: $MeetingsPerOrganiser"
    Log "  Recurring Meetings per Organiser: $RecurringMeetingsPerOrganiser"
    Log "  Total meetings to create: $($Organisers.Count * ($MeetingsPerOrganiser + $RecurringMeetingsPerOrganiser))"
}
Log "  Clear existing events: $($Clear -or $ClearOnly)"
if ($ClearOnly)
{
    Log "  Clear only mode: Will exit after clearing events" Yellow
}
Log ""

# Clear existing calendar events if requested
if ($Clear -or $ClearOnly)
{
    Log "Clearing existing calendar events..." Cyan
    
    # Clear organiser calendars first (this should cancel room events)
    Log "Clearing organiser calendars:" Cyan
    $totalOrganiserEventsDeleted = 0
    foreach ($organiser in $Organisers)
    {
        $deleted = ClearCalendarEvents $organiser
        $totalOrganiserEventsDeleted += $deleted
    }
    Log "  Total organiser events deleted: $totalOrganiserEventsDeleted" Green
    Log ""
    
    # Clear room mailboxes (to clean up any remaining events)
    Log "Clearing room mailbox calendars:" Cyan
    $totalRoomEventsDeleted = 0
    foreach ($room in $RoomMailboxes)
    {
        $deleted = ClearCalendarEvents $room
        $totalRoomEventsDeleted += $deleted
    }
    Log "  Total room events deleted: $totalRoomEventsDeleted" Green
    Log ""
    
    # If ClearOnly mode, exit now
    if ($ClearOnly)
    {
        Log "ClearOnly mode: Exiting after clearing events" Green
        exit
    }
}

# Get display names for all room mailboxes
Log "Retrieving room display names..." Cyan
$roomDisplayNames = @{}
foreach ($room in $RoomMailboxes)
{
    $displayName = GetRoomDisplayName $room
    $roomDisplayNames[$room] = $displayName
    Log "  $room -> $displayName" Gray
}
Log ""

# Start creating meetings from tomorrow (UTC)
$startDate = [DateTime]::SpecifyKind([DateTime]::Today.AddDays(1).AddHours(9), [DateTimeKind]::Utc)
$currentRoomIndex = 0
$totalMeetingsCreated = 0
$totalMeetingsFailed = 0
$globalMeetingNumber = 0 # Track meeting number across all organisers

foreach ($organiser in $Organisers)
{
    Log "Processing organiser: $organiser" Cyan
    
    # Build attendees list (all other organisers)
    $attendeesList = $Organisers | Where-Object { $_ -ne $organiser }
    
    for ($i = 1; $i -le $MeetingsPerOrganiser; $i++)
    {
        # Select room mailbox (cycle through the list)
        $roomMailbox = $RoomMailboxes[$currentRoomIndex]
        $currentRoomIndex = ($currentRoomIndex + 1) % $RoomMailboxes.Count
        
        # Calculate meeting times (1 hour meetings, one day apart)
        # Use global counter to prevent conflicts across organisers
        $meetingStart = $startDate.AddDays($globalMeetingNumber)
        $meetingEnd = $meetingStart.AddHours(1)
        $globalMeetingNumber++
        
        # Get room display name
        $roomDisplayName = $roomDisplayNames[$roomMailbox]
        $subject = "Meeting $i in $roomDisplayName"
        
        Log "  [$i/$MeetingsPerOrganiser] Creating: '$subject'" Gray
        Log "    Room: $roomMailbox"
        Log "    Time: $($meetingStart.ToString('yyyy-MM-dd HH:mm')) - $($meetingEnd.ToString('HH:mm')) UTC"
        Log "    Attendees: $($attendeesList.Count)"
        
        # Create the event
        $result = CreateCalendarEvent $organiser $subject $meetingStart $meetingEnd $roomMailbox $roomDisplayName $attendeesList
        
        if ($null -ne $result)
        {
            Log "    SUCCESS: Event created (ID: $($result.id))" Green
            $totalMeetingsCreated++
        }
        else
        {
            Log "    FAILED: Could not create event" Red
            $totalMeetingsFailed++
        }
    }
    
    Log ""
}

# Create recurring meetings
if ($RecurringMeetingsPerOrganiser -gt 0)
{
    Log "Creating recurring meetings..." Cyan
    Log ""
    
    # Start recurring meetings at 1 PM the day after tomorrow (UTC)
    $recurringStartDate = [DateTime]::SpecifyKind([DateTime]::Today.AddDays(2).AddHours(13), [DateTimeKind]::Utc)
    $totalRecurringCreated = 0
    $totalRecurringFailed = 0
    
    # Recurrence pattern cycle: daily, weekly, monthly
    $recurrencePatterns = @("daily", "weekly", "monthly")
    $patternIndex = 0
    
    foreach ($organiser in $Organisers)
    {
        Log "Processing recurring meetings for organiser: $organiser" Cyan
        
        # Build attendees list (all other organisers)
        $attendeesList = $Organisers | Where-Object { $_ -ne $organiser }
        
        for ($i = 1; $i -le $RecurringMeetingsPerOrganiser; $i++)
        {
            # Select room mailbox (cycle through the list)
            $roomMailbox = $RoomMailboxes[$currentRoomIndex]
            $currentRoomIndex = ($currentRoomIndex + 1) % $RoomMailboxes.Count
            
            # Get the recurrence pattern for this meeting
            $patternType = $recurrencePatterns[$patternIndex]
            $patternIndex = ($patternIndex + 1) % $recurrencePatterns.Length
            
            # Calculate meeting times (1 hour meetings, offset by meeting number)
            $meetingStart = $recurringStartDate.AddHours($i * 2)
            $meetingEnd = $meetingStart.AddHours(1)
            
            # Get room display name
            $roomDisplayName = $roomDisplayNames[$roomMailbox]
            $subject = "Recurring Meeting $i ($patternType) in $roomDisplayName"
            
            Log "  [$i/$RecurringMeetingsPerOrganiser] Creating: '$subject'" Gray
            Log "    Room: $roomMailbox"
            Log "    Pattern: $patternType"
            Log "    Start: $($meetingStart.ToString('yyyy-MM-dd HH:mm')) UTC"
            
            # Build recurrence pattern
            $recurrencePattern = @{}
            $recurrenceRange = @{
                startDate = $meetingStart.ToString("yyyy-MM-dd")
            }
            
            switch ($patternType)
            {
                "daily" {
                    $recurrencePattern = @{
                        type = "daily"
                        interval = 1
                    }
                    $recurrenceRange.type = "numbered"
                    $recurrenceRange.numberOfOccurrences = 7
                    Log "    Occurrences: 7"
                }
                "weekly" {
                    $recurrencePattern = @{
                        type = "weekly"
                        interval = 1
                        daysOfWeek = @($meetingStart.DayOfWeek.ToString().ToLower())
                    }
                    $recurrenceRange.type = "endDate"
                    $recurrenceRange.endDate = $meetingStart.AddMonths(3).ToString("yyyy-MM-dd")
                    Log "    Duration: 3 months (ends: $($recurrenceRange.endDate))"
                }
                "monthly" {
                    $recurrencePattern = @{
                        type = "absoluteMonthly"
                        interval = 1
                        dayOfMonth = $meetingStart.Day
                    }
                    $recurrenceRange.type = "endDate"
                    $recurrenceRange.endDate = $meetingStart.AddMonths(12).ToString("yyyy-MM-dd")
                    Log "    Duration: 12 months (ends: $($recurrenceRange.endDate))"
                }
            }
            
            # Create the recurring event
            $result = CreateRecurringCalendarEvent $organiser $subject $meetingStart $meetingEnd $roomMailbox $roomDisplayName $attendeesList $recurrencePattern $recurrenceRange
            
            if ($null -ne $result)
            {
                Log "    SUCCESS: Recurring event created (ID: $($result.id))" Green
                $totalRecurringCreated++
                
                # Add exceptions/cancellations based on rules
                $addedExceptions = $false
                
                # Every other recurring meeting should have cancellations
                if ($i % 2 -eq 0)
                {
                    Log "    Adding cancellations..." Gray
                    
                    # Calculate which occurrences to cancel (pick 1st and 3rd for variety)
                    $cancellations = @()
                    
                    switch ($patternType)
                    {
                        "daily" {
                            # Cancel 1st and 3rd occurrence
                            $cancellations += $meetingStart.AddDays(0) # 1st
                            if (7 -ge 3) { $cancellations += $meetingStart.AddDays(2) } # 3rd
                        }
                        "weekly" {
                            # Cancel 1st and 3rd week
                            $cancellations += $meetingStart.AddDays(0) # 1st week
                            $cancellations += $meetingStart.AddDays(14) # 3rd week
                        }
                        "monthly" {
                            # Cancel 1st and 3rd month
                            $cancellations += $meetingStart.AddMonths(0) # 1st month
                            $cancellations += $meetingStart.AddMonths(2) # 3rd month
                        }
                    }
                    
                    foreach ($cancelDate in $cancellations)
                    {
                        Log "      Cancelling occurrence on $($cancelDate.ToUniversalTime().ToString('yyyy-MM-dd HH:mm')) UTC..."
                        $cancelled = CancelOccurrence $organiser $result.id $cancelDate
                        if ($cancelled)
                        {
                            Log "        Cancelled" Green
                            $addedExceptions = $true
                        }
                        else
                        {
                            Log "        Failed to cancel" Red
                        }
                        #Start-Sleep -Seconds 1
                    }
                }
                
                # One in three should have time modified occurrence
                if ($i % 3 -eq 0)
                {
                    Log "    Adding time modification exception..." Gray
                    
                    # Modify the 2nd occurrence (move forward 1 hour)
                    $modifyDate = $null
                    switch ($patternType)
                    {
                        "daily" { $modifyDate = $meetingStart.AddDays(1) }
                        "weekly" { $modifyDate = $meetingStart.AddDays(7) }
                        "monthly" { $modifyDate = $meetingStart.AddMonths(1) }
                    }
                    
                    if ($null -ne $modifyDate)
                    {
                        Log "      Modifying occurrence on $($modifyDate.ToUniversalTime().ToString('yyyy-MM-dd HH:mm')) UTC (move forward 1 hour)..."
                        $modified = ModifyOccurrenceTime $organiser $result.id $modifyDate 1
                        if ($modified)
                        {
                            Log "        Modified" Green
                            $addedExceptions = $true
                        }
                        else
                        {
                            Log "        Failed to modify" Red
                        }
                        #Start-Sleep -Seconds 1
                    }
                }
                
                # One in four should have location modified occurrence
                if ($i % 4 -eq 0)
                {
                    Log "    Adding location modification exception..." Gray
                    
                    # Modify the 2nd occurrence to a different room
                    # Pick a different room from the list
                    $alternateRoomIndex = ($currentRoomIndex + 1) % $RoomMailboxes.Count
                    $alternateRoom = $RoomMailboxes[$alternateRoomIndex]
                    $alternateRoomDisplayName = $roomDisplayNames[$alternateRoom]
                    
                    $modifyDate = $null
                    switch ($patternType)
                    {
                        "daily" { $modifyDate = $meetingStart.AddDays(1) }
                        "weekly" { $modifyDate = $meetingStart.AddDays(7) }
                        "monthly" { $modifyDate = $meetingStart.AddMonths(1) }
                    }
                    
                    if ($null -ne $modifyDate)
                    {
                        Log "      Modifying occurrence on $($modifyDate.ToUniversalTime().ToString('yyyy-MM-dd HH:mm')) UTC to room $alternateRoom..."
                        $modified = ModifyOccurrenceLocation $organiser $result.id $modifyDate $alternateRoom $alternateRoomDisplayName
                        if ($modified)
                        {
                            Log "        Modified" Green
                            $addedExceptions = $true
                        }
                        else
                        {
                            Log "        Failed to modify" Red
                        }
                        #Start-Sleep -Seconds 1
                    }
                }
            }
            else
            {
                Log "    FAILED: Could not create recurring event" Red
                $totalRecurringFailed++
            }
        }
        
        Log ""
    }
    
    Log "Recurring meetings summary:" Cyan
    Log "  Total recurring meetings created: $totalRecurringCreated" Green
    if ($totalRecurringFailed -gt 0)
    {
        Log "  Total recurring meetings failed: $totalRecurringFailed" Red
    }
    Log ""
}

Log "Summary:" Cyan
Log "  Total single-instance meetings created: $totalMeetingsCreated" Green
if ($totalMeetingsFailed -gt 0)
{
    Log "  Total single-instance meetings failed: $totalMeetingsFailed" Red
}
if ($RecurringMeetingsPerOrganiser -gt 0)
{
    Log "  Total recurring meetings created: $totalRecurringCreated" Green
    if ($totalRecurringFailed -gt 0)
    {
        Log "  Total recurring meetings failed: $totalRecurringFailed" Red
    }
}
Log ""

# Suggest a good start date for testing Rename-RoomMailboxes.ps1
# Use 3 days from now to test scenario where recurring series start before cutoff but have occurrences after
$suggestedStartDate = [DateTime]::Today.AddDays(3)
Log "Suggested start date for Rename-RoomMailboxes.ps1:" Cyan
Log "  -StartDate $($suggestedStartDate.ToString('yyyy-MM-dd'))" Green
Log ""
Log "  Why this date ($($suggestedStartDate.ToString('yyyy-MM-dd')))?:" Cyan
Log "    - Tests the room rename date scenario" Gray
Log "    - Single-instance meetings start tomorrow (before cutoff) - will NOT be processed" Gray
Log "    - Recurring series start day after tomorrow (before cutoff)" Gray
Log "    - BUT recurring series have occurrences on/after this cutoff - WILL be processed" Gray
Log "    - Script automatically looks back 1 year to find relevant recurring series" Gray
Log ""
Log "  Test coverage with this date:" Gray
Log "    - Single-instance meetings: Only those on/after cutoff processed" Gray
Log "    - Daily recurring series: Split at cutoff (occurrences before stay with old room)" Gray
Log "    - Weekly/monthly recurring: Entire future series gets new room name" Gray
Log "    - Cancelled/modified occurrences: Properly handled in split series" Gray
Log ""
Log "Completed" Green
