#
# Test-FreeBusyLatency.ps1
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
Tests latency between calendar and FreeBusy (using getSchedule to query FreeBusy)

.DESCRIPTION
This script creates meetings in the test mailbox and then queries FreeBusy from the affected mailboxes until the meeting is present.

https://learn.microsoft.com/en-us/graph/api/user-post-events
https://learn.microsoft.com/en-us/graph/api/calendar-getschedule

.EXAMPLE

Delegate permissions (requires Calendars.ReadWrite):
.\Test-FreeBusyLatency.ps1 -AppId $clientId -TenantId $tenantId -Attendees $attendees -TestCount 1

Application permissions (requires Calendars.ReadWrite):
.\Test-FreeBusyLatency.ps1 -mailbox $Mailbox -AppId $clientId -TenantId $tenantId -AppSecretKey $secretKey -Attendees $attendees -StartDate $([DateTime]::Today.AddDays(8)) -TestCount 1

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

    [Parameter(Mandatory=$False,HelpMessage="Attendees to be added to the created meetings (will also be included in the getSchedule request).")]
    [ValidateNotNullOrEmpty()]
    [string[]]$Attendees,

    [Parameter(Mandatory=$False,HelpMessage="If specified, we'll test for the specific meeting being returned by FreeBusy.  This will only work if the organizer has permissions to the attendee's calendar.")]
    [switch]$OrganizerCanViewAttendeesCalendars,

    [Parameter(Mandatory=$False,HelpMessage="If specified, output will be logged to file (using same name as this script).")]
    [switch]$LogToFile,

    [Parameter(Mandatory=$False,HelpMessage="Number of times to run the test.")]
    [int]$TestCount = 10,

    [Parameter(Mandatory=$False,HelpMessage="Start of the time range to create appointments.")]
    [datetime]$StartDate = [DateTime]::Today.AddDays(8)
)

$script:ScriptVersion = "1.0.0"
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
    Write-Host "Successfully renewed OAuth token" -ForegroundColor Green
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

function CreateMeeting($subject, $start, $end)
{
    # Create and send a meeting, logging the creation time
    $meeting = @{
        subject = $subject
        start = @{
            dateTime = $start.ToUniversalTime().ToString("s")
            timeZone = "UTC"
        }
        end = @{
            dateTime = $end.ToUniversalTime().ToString("s")
            timeZone = "UTC"
        }
        attendees = @()
    }
    foreach ($attendee in $Attendees)
    {
        $meeting.attendees += @{
            emailAddress = @{
                address = $attendee
                name = $attendee
            }
            type = "required"
        }
    }

    $meetingRequestUri = "$($graphBaseUrl)events"
    try
    {
        $meetingJson = $meeting | ConvertTo-Json -Depth 5
        UpdateHeaders
        $meetingResponse = Invoke-RestMethod -Method Post -Uri $meetingRequestUri -Headers $script:headers -Body $meetingJson
        Log "Created meeting: $($meetingResponse.subject)"
        return $meetingResponse
    }
    catch
    {
        Log "ERROR: $meetingRequestUri" Red
        return $null
    }
}

function DeleteMeeting($MeetingId)
{
    # Delete the specified meeting
    $meetingRequestUri = "$($graphBaseUrl)events/$MeetingId"
    try
    {
        UpdateHeaders
        Invoke-RestMethod -Method Delete -Uri $meetingRequestUri -Headers $script:headers
        Log "Meeting deleted: $MeetingId"
        return $true
    }
    catch
    {
        Log "ERROR: Failed to delete meeting $MeetingId " Red        
    }
    return $false
}

function GetSchedule($start, $end)
{
    # Perform getSchedule request for the given time range (with $Mailbox and $Attendees)

    $getScheduleRequest = @(
        @{
            schedules = $Attendees
            startTime = @{
                dateTime = $start.ToUniversalTime().ToString("s")
                timeZone = "UTC"
            }
            endTime = @{
                dateTime = $end.ToUniversalTime().ToString("s")
                timeZone = "UTC"
            }
            availabilityViewInterval = "30"
        }
    )
    
    $getScheduleRequestUri = "$($graphBaseUrl)calendar/getSchedule"
    try
    {
        Write-Verbose "Requesting schedule: $getScheduleRequestUri"
        $getScheduleJson = $getScheduleRequest | ConvertTo-Json -Depth 5
        UpdateHeaders
        $getScheduleResponse = Invoke-RestMethod -Method Post -Uri $getScheduleRequestUri -Headers $script:headers -Body $getScheduleJson
        return $getScheduleResponse
    }
    catch
    {
        Log "ERROR: getSchedule failed" Red
        return $null
    }
}

function RetrieveCalendarView($start, $end)
{
    # Retrieve the calendar view

    # Build the query string
    $viewStart = "$($start.ToUniversalTime().ToString("s"))Z"
    $viewEnd = "$($end.ToUniversalTime().ToString("s"))Z"
    $calenderViewRequest = "$($graphUrl)?startDateTime=$viewStart&endDateTime=$viewEnd"
    if (-not [String]::IsNullOrEmpty($viewQuery))
    {
        $calenderViewRequest = "$calenderViewRequest&$viewQuery"
    }

    # Send the request and return any items found
    $calendarItems = $null
    try
    {
        Write-Verbose "Requesting calendar: $calenderViewRequest"
        if ($psversiontable.PSVersion.Major -gt 6)
        {
            $thisCalendarView = Invoke-WebRequest -Method Get -Uri $calenderViewRequest -Headers $script:headers -SkipHeaderValidation
        }
        else
        {
            $thisCalendarView = Invoke-WebRequest -Method Get -Uri $calenderViewRequest -Headers $script:headers
        }
        $calendarItems = $thisCalendarView.Content | ConvertFrom-Json
        Log "$($calendarItems.value.Count) item(s) returned: $calenderViewRequest"
        $global:successQueries += $calenderViewRequest
        return $calendarItems.value
    }
    catch
    {
        $global:errorQueries += $calenderViewRequest # Log the calls that generate an error
        Log "ERROR: $calenderViewRequest" Red
    }

    # We get here if an error occurred.  If we have an additional query, retry the request without the query
    if ([String]::IsNullOrEmpty($viewQuery))
    {
        return $null
    }

    $calenderViewRequest = "$($graphUrl)?startDateTime=$viewStart&endDateTime=$viewEnd"
    try
    {
        Write-Verbose "Requesting calendar: $calenderViewRequest"
        if ($psversiontable.PSVersion.Major -gt 6)
        {
            $thisCalendarView = Invoke-WebRequest -Method Get -Uri $calenderViewRequest -Headers $script:headers -SkipHeaderValidation
        }
        else
        {
            $thisCalendarView = Invoke-WebRequest -Method Get -Uri $calenderViewRequest -Headers $script:headers
        }
        $calendarItems = $thisCalendarView.Content | ConvertFrom-Json
        Log "$($calendarItems.value.Count) item(s) returned: $calenderViewRequest"
        $global:successQueries += $calenderViewRequest
        return $calendarItems.value
    }
    catch
    {
        $global:errorQueries += $calenderViewRequest # Log the calls that generate an error
        Log "ERROR: $calenderViewRequest" Red
    }
    return $null
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
Write-Host "Successfully obtained OAuth token" -ForegroundColor Green


# Create a number of meetings in the mailbox and then perform getSchedule repeatedly until the meeting is present in the FreeBusy response for all attendees.
# The latency includes both the delivery time for the meeting invitation and the time taken for the FreeBusy to update.

$global:latencies = @()

# Now create calendar items at 30 minute intervals
$meetingStartTime = $StartDate.AddHours(9)
$meetingEndTime = $meetingStartTime.AddMinutes(30)
$meetingBaseSubject = "Test Meeting $(New-Guid):"

for ($testIndex=1; $testIndex -le $TestCount; $testIndex++) {
    $meetingSubject = "$meetingBaseSubject $testIndex"
    Log "Creating meeting at $meetingStartTime"
    $meeting = CreateMeeting $meetingSubject $meetingStartTime $meetingEndTime
    $pingInterval = 1
    if ($null -ne $meeting -and ![String]::IsNullOrEmpty($meeting.Id))
    {
        $meetingCreatedTime = [DateTime]::Now
        Write-Progress -Activity "Waiting for FreeBusy to update" -Status "0 seconds"
        do {
            $freeBusy = GetSchedule $StartDate $StartDate.AddDays(1)
            if ($null -ne $freeBusy) {
                # We check that we have the meeting in each attendee's schedule
                $freeBusyUpdated = $true
                foreach ($attendee in $Attendees) {
                    $attendeeFreeBusy = $freeBusy.value | Where-Object { $_.scheduleId -eq $attendee }
                    if ($null -eq $attendeeFreeBusy) {
                        $freeBusyUpdated = $false
                        break
                    }

                    if ($OrganizerCanViewAttendeesCalendars) {
                        # Search for specific meeting by subject (this doesn't work if the organizer doesn't have permission to the attendee's calendar)
                        $meetingsMatchingSubject = $attendeeFreeBusy.scheduleItems | Where-Object { $_.subject -eq $meetingSubject }
                        if ($null -eq $meetingsMatchingSubject) {
                            $freeBusyUpdated = $false               
                            break
                        }
                    } else {
                        # When the FreeBusy is updated, the time slot of the newly created meeting will be marked as "tentative" or "busy"
                        $unavailable = $attendeeFreeBusy.scheduleItems | Where-Object { $_.status -eq "tentative" -or $_.status -eq "busy" }
                        $unavailableAtMeetingStart = $unavailable | Where-Object { $_.start.dateTime -eq $meetingStartTime }
                        if ($null -eq $unavailableAtMeetingStart) {
                            $freeBusyUpdated = $false               
                            break
                        }
                    }
                }
            } else {
                $freeBusyUpdated = $false
            }
            if (!$freeBusyUpdated) {
                Start-Sleep -Seconds $pingInterval
                $elapsedTimeInSeconds = [int]([DateTime]::Now.Subtract($meetingCreatedTime).TotalSeconds)
                if ($pingInterval -eq 1 -and $elapsedTimeInSeconds -gt 360) {
                    # After five minutes we'll only check every 5 seconds
                    $pingInterval = 5
                }
                Write-Progress -Activity "Waiting for FreeBusy to update" -Status "$($elapsedTimeInSeconds) second(s)"
            }            
        } while (!$freeBusyUpdated)
        Write-Progress -Activity "Waiting for FreeBusy to update" -Completed
        $freeBusyUpdatedTime = [DateTime]::Now
        $latency = $freeBusyUpdatedTime.Subtract($meetingCreatedTime)
        Log "FreeBusy latency: $latency"
        $global:latencies += $latency

        # Now delete the meeting
        DeleteMeeting $meeting.id | out-null

        $meetingStartTime = $meetingStartTime.AddMinutes(30)
        if ($meetingStartTime.Hour -ge 17) {
            $StartDate = $StartDate.AddDays(1)
            $meetingStartTime = $StartDate.AddHours(9)
        }
        $meetingEndTime = $meetingStartTime.AddMinutes(30)
    }
    else {
        if ($testIndex -eq 1) {
            Log "Failed to create the first meeting.  Exiting." Red
            exit
        }
        Log "Failed to create meeting: $meetingSubject" Red
    }
}

Log "Latencies available in `$latencies." Green
$latencies.TotalSeconds | Measure-Object -AllStats
Log "$($MyInvocation.MyCommand.Name) version $($script:ScriptVersion) finished."