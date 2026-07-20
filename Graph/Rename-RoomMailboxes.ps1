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
1. Removes the room mailbox from the organizer's event attendees list
2. Saves the updated event and waits for the event to be removed from the room's calendar
3. Re-adds the room mailbox to the organizer's event attendees list (preserving the room's display name)
4. Saves the event again
5. Allows the room mailbox to auto-process the meeting invitation

This can be useful for refreshing room bookings or forcing synchronization of room calendar entries.

The script uses application permissions and requires access to all calendars in the organization.
Events are identified across mailboxes using the iCalUId property to ensure the correct event is updated.

IMPORTANT: Room mailboxes must be configured to automatically accept meeting invitations. The script
relies on the room's auto-accept processing to handle the re-added invitation after removal.

.PREREQUISITES
Entra ID App Registration:
- Register an application in the Entra ID (Azure AD) portal
- Create a client secret for the application
- Configure the following API permissions:
  * Microsoft Graph > Application Permissions > Calendars.ReadWrite (required)
  * Microsoft Graph > Application Permissions > User.Read.All (optional, if needed to resolve user information)
- Grant admin consent for the permissions in your tenant
- Note the Application (client) ID, Directory (tenant) ID, and client secret value

IMPORTANT: This script requires application permissions (client credential flow) and will not work with delegate permissions.
The AppSecretKey parameter must be provided to authenticate using application permissions.

.EXAMPLE
$roomMailboxes = @("confroom1@contoso.com")
$startDate = [DateTime]::Today.AddDays(7)
.\Rename-RoomMailboxes.ps1 -AppId $clientId -AppSecretKey $secretKey -TenantId $tenantId -RoomMailboxes $roomMailboxes -StartDate $startDate

Processes all calendar events for confroom1@contoso.com starting 7 days from now. The script will delete each event from the room's calendar, remove and re-add the room to the organizer's event, then programmatically accept the meeting on behalf of the room.

.EXAMPLE
$roomMailboxes = @("room1@contoso.com", "room2@contoso.com", "room3@contoso.com")
$startDate = [DateTime]::Today.AddDays(7)
.\Rename-RoomMailboxes.ps1 -AppId $clientId -AppSecretKey $secretKey -TenantId $tenantId -RoomMailboxes $roomMailboxes -StartDate $startDate -LogToFile

Processes multiple room mailboxes starting 7 days from now, with all activity logged to a file.
Room mailboxes will auto-process the refreshed meeting invitations.

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

<# Calendar.ps1 %PARAMS_START% #>
<# Calendar.ps1 %PARAMS_END% #>

    [Parameter(Mandatory=$True,HelpMessage="Array of room mailbox SMTP addresses to process.")]
    [ValidateNotNullOrEmpty()]
    [string[]]$RoomMailboxes,

    [Parameter(Mandatory=$True,HelpMessage="Date of the room name change. Single-instance events on/after this date are processed. Recurring series are processed if they have occurrences on/after this date (even if the series started earlier). Format: yyyy-MM-dd or yyyy-MM-ddTHH:mm:ss")]
    [ValidateNotNullOrEmpty()]
    [string]$StartDate
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

<# Calendar.ps1 %FUNCTIONS_START% #>
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
    # Ensure input is treated as an array (PowerShell can unwrap single-element arrays)
    $attendees = @($attendeesArray)
    $newAttendees = @()
    foreach ($attendee in $attendees)
    {
        if ($attendee.emailAddress.address -ne $emailAddress)
        {
            $newAttendees += $attendee
        }
    }
    return $newAttendees
}

function GetMailboxDisplayName($mailbox)
{
    # Get the current display name for a mailbox (with caching)
    # Check cache first
    if ($script:displayNameCache.ContainsKey($mailbox))
    {
        return $script:displayNameCache[$mailbox]
    }
    
    # Not in cache, retrieve from Graph API
    $url = "https://graph.microsoft.com/v1.0/users/$mailbox`?`$select=displayName"
    $response = GET $url
    
    $displayName = $mailbox  # Default to email address
    
    if ($null -ne $response)
    {
        $user = ConvertFrom-Json $response
        if (![String]::IsNullOrEmpty($user.displayName))
        {
            $displayName = $user.displayName
        }
    }
    
    # Store in cache
    $script:displayNameCache[$mailbox] = $displayName
    
    return $displayName
}

function AddAttendee($attendeesArray, $emailAddress, $attendeeType = "required", $displayName = $null)
{
    # Add an attendee to the attendees array with optional display name
    # Ensure input is treated as an array (PowerShell can unwrap single-element arrays)
    $attendees = @($attendeesArray)
    
    $emailAddressObj = @{
        address = $emailAddress
    }
    
    # Include display name if provided
    if (![String]::IsNullOrEmpty($displayName))
    {
        $emailAddressObj.name = $displayName
    }
    
    $newAttendee = @{
        emailAddress = $emailAddressObj
        type = $attendeeType
    }
    $attendees += $newAttendee
    return $attendees
}

function DeleteEventFromCalendar($mailbox, $iCalUId, $eventId = $null)
{
    # Delete an event from a mailbox's calendar
    # If eventId is provided, use it directly; otherwise search by iCalUId
    
    if ([String]::IsNullOrEmpty($eventId))
    {
        # Find the event by iCalUId if no event ID was provided
        $calendarEvent = FindEventByICalUId $mailbox $iCalUId
        
        if ($null -eq $calendarEvent)
        {
            return $false
        }
        
        $eventId = $calendarEvent.id
    }
    
    $url = "https://graph.microsoft.com/v1.0/users/$mailbox/events/$eventId"
    $response = DELETE $url
    
    # DELETE returns true/false indicating success
    return $response
}

function CreateEvent($mailbox, $body)
{
    # Create a new event in a mailbox's calendar
    $url = "https://graph.microsoft.com/v1.0/users/$mailbox/events"
    $response = POST $url $body
    if ($null -eq $response)
    {
        return $null
    }
    return ConvertFrom-Json ($response)
}

function GetEventWithInstances($mailbox, $eventId, $eventStartDate, $eventRecurrence)
{
    # Get a recurring event with all its instances (occurrences and exceptions)
    $queryString = "?`$select=id,subject,start,end,organizer,attendees,iCalUId,type,recurrence,seriesMasterId,location,body,bodyPreview,importance,sensitivity,showAs,isReminderOn,reminderMinutesBeforeStart,responseRequested,allowNewTimeProposals,isOnlineMeeting,onlineMeetingProvider,onlineMeetingUrl,categories"
    $event = GetEventById $mailbox $eventId $queryString
    
    if ($null -eq $event -or $event.type -ne "seriesMaster")
    {
        return $event
    }
    
    # Calculate the time window for instances query
    # Use series start date as the beginning
    $instancesStart = $eventStartDate.ToString("yyyy-MM-ddTHH:mm:ss")
    
    # Calculate end date based on recurrence pattern
    $instancesEnd = $eventStartDate.AddYears(2)  # Default: 2 years from start
    
    if ($null -ne $eventRecurrence -and $null -ne $eventRecurrence.range)
    {
        if ($eventRecurrence.range.type -eq "endDate" -and ![String]::IsNullOrEmpty($eventRecurrence.range.endDate))
        {
            # Use the series end date plus a buffer
            $seriesEndDate = [DateTime]::Parse($eventRecurrence.range.endDate, [System.Globalization.CultureInfo]::InvariantCulture)
            $seriesEndDate = [DateTime]::SpecifyKind($seriesEndDate, [DateTimeKind]::Utc)
            $instancesEnd = $seriesEndDate.AddDays(1)  # Add 1 day buffer
        }
        elseif ($eventRecurrence.range.type -eq "numbered" -and $null -ne $eventRecurrence.range.numberOfOccurrences)
        {
            # Estimate end date based on number of occurrences and pattern
            $occurrences = [int]$eventRecurrence.range.numberOfOccurrences
            
            # Calculate estimated duration based on recurrence pattern
            $pattern = $eventRecurrence.pattern
            switch ($pattern.type)
            {
                "daily" { $instancesEnd = $eventStartDate.AddDays($occurrences * $pattern.interval + 7) }
                "weekly" { $instancesEnd = $eventStartDate.AddDays(($occurrences * $pattern.interval * 7) + 7) }
                "monthly" { $instancesEnd = $eventStartDate.AddMonths($occurrences * $pattern.interval + 1) }
                "yearly" { $instancesEnd = $eventStartDate.AddYears($occurrences * $pattern.interval + 1) }
                default { $instancesEnd = $eventStartDate.AddYears(2) }
            }
        }
    }
    
    # Get instances (occurrences and exceptions) for the series with time window
    $instancesUrl = "https://graph.microsoft.com/v1.0/users/$mailbox/events/$eventId/instances?startDateTime=$instancesStart&endDateTime=$($instancesEnd.ToString('yyyy-MM-ddTHH:mm:ss'))&`$select=id,subject,start,end,organizer,attendees,iCalUId,type,seriesMasterId,location,body,bodyPreview,importance,sensitivity,showAs,isReminderOn,reminderMinutesBeforeStart,responseRequested,allowNewTimeProposals,isOnlineMeeting,onlineMeetingProvider,onlineMeetingUrl,categories"
    $instancesResponse = GET $instancesUrl
    
    if ($null -ne $instancesResponse)
    {
        $instances = ConvertFrom-Json $instancesResponse
        $event | Add-Member -MemberType NoteProperty -Name "instances" -Value $instances.value -Force
        
        # Handle pagination if there are many instances
        while ($null -ne $instances.'@odata.nextLink')
        {
            $instancesResponse = GET $instances.'@odata.nextLink'
            if ($null -ne $instancesResponse)
            {
                $instances = ConvertFrom-Json $instancesResponse
                $event.instances += $instances.value
            }
        }
    }
    
    return $event
}

function UpdateRecurrenceEndDate($recurrencePattern, $endDate)
{
    # Update a recurrence pattern to end on a specific date
    if ($null -eq $recurrencePattern)
    {
        return $null
    }
    
    # Create a copy of the recurrence pattern
    $updatedRecurrence = $recurrencePattern | ConvertTo-Json -Depth 10 | ConvertFrom-Json
    
    # Update the range to end on the specified date
    if ($null -eq $updatedRecurrence.range)
    {
        $updatedRecurrence.range = @{}
    }
    
    $updatedRecurrence.range.type = "endDate"
    $updatedRecurrence.range.endDate = $endDate.ToString("yyyy-MM-dd")
    
    # Remove numberOfOccurrences if it exists (can't have both endDate and numberOfOccurrences)
    if ($updatedRecurrence.range.PSObject.Properties.Name -contains "numberOfOccurrences")
    {
        $updatedRecurrence.range.PSObject.Properties.Remove("numberOfOccurrences")
    }
    
    return $updatedRecurrence
}

function UpdateRecurrenceStartDate($recurrencePattern, $startDate)
{
    # Update a recurrence pattern to start on a specific date
    if ($null -eq $recurrencePattern)
    {
        return $null
    }
    
    # Create a copy of the recurrence pattern
    $updatedRecurrence = $recurrencePattern | ConvertTo-Json -Depth 10 | ConvertFrom-Json
    
    # Update the range to start on the specified date
    if ($null -eq $updatedRecurrence.range)
    {
        $updatedRecurrence.range = @{}
    }
    
    $updatedRecurrence.range.startDate = $startDate.ToString("yyyy-MM-dd")
    
    return $updatedRecurrence
}

<# Calendar.ps1 %FUNCTIONS_END% #>

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
    # Specify as UTC to prevent local timezone conversion
    $startDateTime = [DateTime]::SpecifyKind($startDateTime, [DateTimeKind]::Utc)
    Log "Processing events from: $($startDateTime.ToString('yyyy-MM-dd HH:mm:ss'))"
}
catch
{
    Log "Invalid date format: $StartDate. Please use yyyy-MM-dd or yyyy-MM-ddTHH:mm:ss format." Red
    exit
}

# Convert to ISO 8601 format for Graph API (already UTC, no conversion needed)
$startDateFilter = $startDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ")

# For recurring events, we need to look back to catch series masters that start before
# the cutoff but have occurrences after. Use a 1-year lookback window.
$lookbackStartDate = $startDateTime.AddYears(-1)
$lookbackFilter = $lookbackStartDate.ToString("yyyy-MM-ddTHH:mm:ssZ")

# Initialize cache for mailbox display names
$script:displayNameCache = @{}

# Process each room mailbox
$totalRooms = $RoomMailboxes.Count
$currentRoom = 0

foreach ($roomMailbox in $RoomMailboxes)
{
    $currentRoom++
    Log "`n[$currentRoom/$totalRooms] Processing room mailbox: $roomMailbox" Cyan
    
    # Get the current display name for the room mailbox
    Log "  Retrieving current display name for room mailbox..."
    $roomDisplayName = GetMailboxDisplayName $roomMailbox
    Log "  Room display name: $roomDisplayName" Gray
    
    # Get calendar events for the room mailbox
    # Use lookback date to catch recurring series that start before cutoff but have occurrences after
    $filter = "`$filter=start/dateTime ge '$lookbackFilter'"
    $select = "`$select=id,subject,start,end,organizer,attendees,iCalUId,type,recurrence,seriesMasterId"
    $queryString = "?$filter&$select&`$top=999"
    
    $eventsResponse = GetCalendarEvents $roomMailbox $queryString
    
    if ($null -eq $eventsResponse -or $null -eq $eventsResponse.value)
    {
        Log "  Unable to retrieve events for $roomMailbox" Yellow
        continue
    }
    
    $events = $eventsResponse.value
    
    # Process pagination if there are more events
    while ($null -ne $eventsResponse.'@odata.nextLink')
    {
        $eventsResponse = GET $eventsResponse.'@odata.nextLink'
        if ($null -ne $eventsResponse)
        {
            $eventsResponse = ConvertFrom-Json $eventsResponse
            $events += $eventsResponse.value
        }
    }
    
    # Filter events to only those relevant to our cutoff date:
    # - Single-instance events starting on/after cutoff
    # - Recurring series (will be checked for occurrences later)
    $filteredEvents = @()
    foreach ($event in $events)
    {
        if ($event.type -eq "seriesMaster")
        {
            # Include all recurring series - we'll check occurrences during processing
            $filteredEvents += $event
        }
        elseif ($event.type -eq "singleInstance")
        {
            # Only include single-instance events starting on/after cutoff
            $eventStartDate = [DateTime]::Parse($event.start.dateTime, [System.Globalization.CultureInfo]::InvariantCulture)
            $eventStartDate = [DateTime]::SpecifyKind($eventStartDate, [DateTimeKind]::Utc)
            if ($eventStartDate -ge $startDateTime)
            {
                $filteredEvents += $event
            }
        }
        # Skip exception and occurrence types - they're part of a series
    }
    
    $totalEvents = $filteredEvents.Count
    
    if ($totalEvents -eq 0)
    {
        Log "  No events found on/after $($startDateTime.ToString('yyyy-MM-dd'))" Yellow
        continue
    }
    
    Log "  Found $totalEvents event(s) to process"
    
    $currentEvent = 0
    foreach ($event in $events)
    {
        $currentEvent++
        Log "  [$currentEvent/$totalEvents] Processing event: $($event.subject)" Gray
        
        # Skip occurrence and exception events - these will be handled through the series master
        if ($event.type -eq "occurrence" -or $event.type -eq "exception")
        {
            Log "    Skipping: This is an occurrence/exception event (handled through series master)" Yellow
            continue
        }
        
        # Get the organizer email address
        $organizerEmail = $event.organizer.emailAddress.address
        if ([String]::IsNullOrEmpty($organizerEmail))
        {
            Log "    Skipping: No organizer found" Yellow
            continue
        }
        
        Log "    Organizer: $organizerEmail"
        Log "    Event type: $($event.type)"
        
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
        
        # Check if this is a recurring event (seriesMaster) that starts before the cutoff date
        if ($event.type -eq "seriesMaster")
        {
            Log "    This is a recurring event series" Cyan
            
            # Parse the event start date (Graph API returns UTC strings)
            $eventStartDate = [DateTime]::Parse($event.start.dateTime, [System.Globalization.CultureInfo]::InvariantCulture)
            $eventStartDate = [DateTime]::SpecifyKind($eventStartDate, [DateTimeKind]::Utc)
            
            # Check if the series starts before the cutoff date
            if ($eventStartDate -lt $startDateTime)
            {
                Log "    Series starts before cutoff date ($($eventStartDate.ToString('yyyy-MM-dd')))" Cyan
                
                # Check if the series has ended before the cutoff date
                if ($null -ne $event.recurrence -and $null -ne $event.recurrence.range)
                {
                    if ($event.recurrence.range.type -eq "endDate" -and $null -ne $event.recurrence.range.endDate)
                    {
                        $seriesEndDate = [DateTime]::Parse($event.recurrence.range.endDate, [System.Globalization.CultureInfo]::InvariantCulture)
                        $seriesEndDate = [DateTime]::SpecifyKind($seriesEndDate, [DateTimeKind]::Utc)
                        
                        if ($seriesEndDate -lt $startDateTime)
                        {
                            Log "    Series ended before cutoff date ($($seriesEndDate.ToString('yyyy-MM-dd'))) - skipping" Yellow
                            continue
                        }
                    }
                }
                
                Log "    Series has occurrences on/after cutoff - will split the series" Cyan
                
                # Get the full event details with instances
                Log "    Retrieving full event details and instances..."
                $fullOrganizerEvent = GetEventWithInstances $organizerEmail $organizerEvent.id $eventStartDate $event.recurrence
                
                if ($null -eq $fullOrganizerEvent)
                {
                    Log "    ERROR: Failed to retrieve full event details" Red
                    continue
                }
                
                # Filter instances to find exceptions and cancellations after the cutoff date
                # Also count occurrences before and after to determine if splitting is needed
                $exceptionsAfterCutoff = @()
                $cancellationsAfterCutoff = @()
                $occurrencesBeforeCutoff = 0
                $occurrencesAfterCutoff = 0
                
                if ($null -ne $fullOrganizerEvent.instances)
                {
                    foreach ($instance in $fullOrganizerEvent.instances)
                    {
                        # Parse instance start date (Graph API returns UTC strings)
                        $instanceStartDate = [DateTime]::Parse($instance.start.dateTime, [System.Globalization.CultureInfo]::InvariantCulture)
                        $instanceStartDate = [DateTime]::SpecifyKind($instanceStartDate, [DateTimeKind]::Utc)
                        
                        if ($instanceStartDate -ge $startDateTime)
                        {
                            $occurrencesAfterCutoff++
                            
                            if ($instance.type -eq "exception")
                            {
                                $exceptionsAfterCutoff += $instance
                            }
                            elseif ($instance.isCancelled -eq $true)
                            {
                                $cancellationsAfterCutoff += $instance
                            }
                        }
                        else
                        {
                            $occurrencesBeforeCutoff++
                        }
                    }
                }
                
                Log "    Found $occurrencesBeforeCutoff occurrence(s) before cutoff and $occurrencesAfterCutoff after cutoff"
                Log "    Found $($exceptionsAfterCutoff.Count) exception(s) and $($cancellationsAfterCutoff.Count) cancellation(s) after cutoff date"
                
                # If there are no occurrences before the cutoff, don't split - just process normally
                if ($occurrencesBeforeCutoff -eq 0)
                {
                    Log "    No occurrences before cutoff - series will be processed as a whole (not split)" Cyan
                    # Fall through to standard processing below
                }
                else
                {
                    # Step 1: Update the original series to end on the day before the cutoff date
                    Log "    Updating original series to end on $($startDateTime.AddDays(-1).ToString('yyyy-MM-dd'))..."
                
                $updatedRecurrence = UpdateRecurrenceEndDate $fullOrganizerEvent.recurrence $startDateTime.AddDays(-1)
                
                $updateOriginalBody = @{
                    recurrence = $updatedRecurrence
                } | ConvertTo-Json -Depth 10 -Compress:$false
                
                $updateOriginalResult = UpdateEvent $organizerEmail $fullOrganizerEvent.id $updateOriginalBody
                
                if ($null -eq $updateOriginalResult)
                {
                    Log "    ERROR: Failed to update original series end date" Red
                    continue
                }
                
                Log "    Successfully updated original series end date" Green
                
                # Step 2: Create a new series for the remainder with updated room details
                Log "    Creating new series for remainder starting from $($startDateTime.ToString('yyyy-MM-dd'))..."
                
                # Remove the old room and add the new one
                $newAttendees = RemoveAttendee $fullOrganizerEvent.attendees $roomMailbox
                $newAttendees = AddAttendee $newAttendees $roomMailbox "required" $roomDisplayName
                
                # Update recurrence to start from the cutoff date
                $newRecurrence = UpdateRecurrenceStartDate $fullOrganizerEvent.recurrence $startDateTime
                
                # Copy the range end date from the original if it exists (and is valid)
                $hasValidEndDate = $false
                if ($null -ne $fullOrganizerEvent.recurrence.range.endDate)
                {
                    $endDateStr = $fullOrganizerEvent.recurrence.range.endDate
                    # Check if it's a valid date (not the default 0001-01-01)
                    if ($endDateStr -ne "0001-01-01" -and ![String]::IsNullOrEmpty($endDateStr))
                    {
                        $newRecurrence.range.endDate = $endDateStr
                        $newRecurrence.range.type = "endDate"
                        $hasValidEndDate = $true
                        
                        # Remove numberOfOccurrences if present
                        if ($newRecurrence.range.PSObject.Properties.Name -contains "numberOfOccurrences")
                        {
                            $newRecurrence.range.PSObject.Properties.Remove("numberOfOccurrences")
                        }
                    }
                }
                
                if (-not $hasValidEndDate)
                {
                    if ($null -ne $fullOrganizerEvent.recurrence.range.numberOfOccurrences)
                    {
                        # Keep the numberOfOccurrences (but note this may not be accurate after the split)
                        $newRecurrence.range.numberOfOccurrences = $fullOrganizerEvent.recurrence.range.numberOfOccurrences
                        $newRecurrence.range.type = "numbered"
                        
                        # Remove endDate if present
                        if ($newRecurrence.range.PSObject.Properties.Name -contains "endDate")
                        {
                            $newRecurrence.range.PSObject.Properties.Remove("endDate")
                        }
                    }
                    else
                    {
                        $newRecurrence.range.type = "noEnd"
                        
                        # Remove both endDate and numberOfOccurrences if present
                        if ($newRecurrence.range.PSObject.Properties.Name -contains "endDate")
                        {
                            $newRecurrence.range.PSObject.Properties.Remove("endDate")
                        }
                        if ($newRecurrence.range.PSObject.Properties.Name -contains "numberOfOccurrences")
                        {
                            $newRecurrence.range.PSObject.Properties.Remove("numberOfOccurrences")
                        }
                    }
                }
                
                # Update location to the room
                $newLocation = @{
                    displayName = $roomDisplayName
                    locationType = "conferenceRoom"
                    uniqueId = $roomMailbox
                    uniqueIdType = "directory"
                }
                
                # Calculate the new start and end times
                # The new series should start at the same time of day as the original, but on the cutoff date
                $originalStart = [DateTime]::Parse($fullOrganizerEvent.start.dateTime, [System.Globalization.CultureInfo]::InvariantCulture)
                $originalStart = [DateTime]::SpecifyKind($originalStart, [DateTimeKind]::Utc)
                $originalEnd = [DateTime]::Parse($fullOrganizerEvent.end.dateTime, [System.Globalization.CultureInfo]::InvariantCulture)
                $originalEnd = [DateTime]::SpecifyKind($originalEnd, [DateTimeKind]::Utc)
                
                # Keep the time component from the original start, but use the cutoff date
                $newSeriesStart = [DateTime]::new($startDateTime.Year, $startDateTime.Month, $startDateTime.Day, 
                                                   $originalStart.Hour, $originalStart.Minute, $originalStart.Second, 
                                                   [DateTimeKind]::Utc)
                
                # Calculate duration and apply to new start
                $duration = $originalEnd - $originalStart
                $newSeriesEnd = $newSeriesStart.Add($duration)
                
                # Prepare the new event body
                $newEventBody = @{
                    subject = $fullOrganizerEvent.subject
                    body = $fullOrganizerEvent.body
                    start = @{
                        dateTime = $newSeriesStart.ToString("yyyy-MM-ddTHH:mm:ss")
                        timeZone = $fullOrganizerEvent.start.timeZone
                    }
                    end = @{
                        dateTime = $newSeriesEnd.ToString("yyyy-MM-ddTHH:mm:ss")
                        timeZone = $fullOrganizerEvent.end.timeZone
                    }
                    location = $newLocation
                    attendees = @($newAttendees)
                    recurrence = $newRecurrence
                    importance = $fullOrganizerEvent.importance
                    sensitivity = $fullOrganizerEvent.sensitivity
                    showAs = $fullOrganizerEvent.showAs
                    isReminderOn = $fullOrganizerEvent.isReminderOn
                    reminderMinutesBeforeStart = $fullOrganizerEvent.reminderMinutesBeforeStart
                    responseRequested = $fullOrganizerEvent.responseRequested
                    allowNewTimeProposals = $fullOrganizerEvent.allowNewTimeProposals
                }
                
                # Include online meeting details if present
                if ($fullOrganizerEvent.isOnlineMeeting -eq $true)
                {
                    $newEventBody.isOnlineMeeting = $true
                    if (![String]::IsNullOrEmpty($fullOrganizerEvent.onlineMeetingProvider))
                    {
                        $newEventBody.onlineMeetingProvider = $fullOrganizerEvent.onlineMeetingProvider
                    }
                }
                
                # Include categories if present
                if ($null -ne $fullOrganizerEvent.categories -and $fullOrganizerEvent.categories.Count -gt 0)
                {
                    $newEventBody.categories = $fullOrganizerEvent.categories
                }
                
                $newEventBodyJson = $newEventBody | ConvertTo-Json -Depth 10 -Compress:$false
                
                $newEvent = CreateEvent $organizerEmail $newEventBodyJson
                
                if ($null -eq $newEvent)
                {
                    Log "    ERROR: Failed to create new series" Red
                    continue
                }
                
                Log "    Successfully created new series (ID: $($newEvent.id))" Green
                
                # Step 3: Handle exceptions after the cutoff date
                if ($exceptionsAfterCutoff.Count -gt 0)
                {
                    Log "    Handling $($exceptionsAfterCutoff.Count) exception occurrence(s)..."
                    
                    # For each exception, we need to update it in the new series
                    # This is complex because we need to find the corresponding occurrence in the new series
                    # For now, we'll log a warning that manual intervention may be needed
                    Log "    Warning: Exception occurrences detected. These may need manual review." Yellow
                    
                    foreach ($exception in $exceptionsAfterCutoff)
                    {
                        Log "      Exception on $($exception.start.dateTime): $($exception.subject)" Yellow
                    }
                }
                
                Log "    Successfully processed recurring event series (split)" Green
                continue
                } # End of split logic (else block for occurrencesBeforeCutoff > 0)
                
                # If we reach here, series starts before cutoff but has no occurrences before cutoff
                # Fall through to standard processing
            }
            else
            {
                Log "    Series starts on or after cutoff date - will use standard processing" Gray
            }
        }
        
        # Standard processing for single instance events or series masters that start on/after the cutoff
        # Remove the room mailbox from attendees and send update
        Log "    Removing room from organizer's event attendees list"
        $updatedAttendees = RemoveAttendee $organizerEvent.attendees $roomMailbox
        
        $updateBody = @{
            attendees = @($updatedAttendees)
        } | ConvertTo-Json -Depth 10 -Compress:$false
        
        $updateResult = UpdateEvent $organizerEmail $organizerEvent.id $updateBody
        
        if ($null -eq $updateResult)
        {
            Log "    ERROR: Failed to remove room from event" Red
            continue
        }
        
        Log "    Successfully removed room from organizer's event" Green
        
        # Wait for the event to be deleted from the room's calendar
        Log "    Waiting for event to be removed from room's calendar..."
        $maxWaitAttempts = 12  # Wait up to 60 seconds (12 * 5 seconds)
        $eventRemovedFromRoom = $false
        
        for ($waitAttempt = 1; $waitAttempt -le $maxWaitAttempts; $waitAttempt++)
        {
            Start-Sleep -Seconds 5
            
            # Check if the event still exists in the room's calendar
            $roomEventCheck = FindEventByICalUId $roomMailbox $iCalUId
            
            if ($null -eq $roomEventCheck)
            {
                Log "    Event successfully removed from room's calendar" Green
                $eventRemovedFromRoom = $true
                break
            }
            else
            {
                Log "    Waiting for event removal (attempt $waitAttempt/$maxWaitAttempts)..." Gray
            }
        }
        
        if (-not $eventRemovedFromRoom)
        {
            Log "    Warning: Event still exists in room's calendar after waiting" Yellow
            Log "    Continuing anyway - the room may process the update shortly" Yellow
        }
        
        # Add the room mailbox back to attendees with display name and set location
        Log "    Adding room back to organizer's event attendees list and setting location"
        $readdedAttendees = AddAttendee $updatedAttendees $roomMailbox "required" $roomDisplayName
        
        # Also update the location to point to the room mailbox
        $location = @{
            displayName = $roomDisplayName
            locationType = "conferenceRoom"
            uniqueId = $roomMailbox
            uniqueIdType = "directory"
        }
        
        $readdBody = @{
            attendees = @($readdedAttendees)
            location = $location
        } | ConvertTo-Json -Depth 10 -Compress:$false
        
        $readdResult = UpdateEvent $organizerEmail $organizerEvent.id $readdBody
        
        if ($null -eq $readdResult)
        {
            Log "    ERROR: Failed to add room back to event" Red
            continue
        }
        
        Log "    Successfully added room back to organizer's event" Green
        Log "    Room will auto-process the meeting invitation" Gray
    }
}

Log "`nCompleted processing all room mailboxes" Green
