#
# Calendar.ps1
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
Calendar helper function for Microsoft Graph API.  Intended to be used as part of the build process for Graph scripts in this repository.

.DESCRIPTION
Contains functions for accessing calendar folders and items via Microsoft Graph.

.EXAMPLE

#>

param (
<# Calendar.ps1 %PARAMS_START% #>
<# Calendar.ps1 %PARAMS_END% #>

    [Parameter(Mandatory=$False,HelpMessage="Not included in template.  Only included to prevent error in code editors.")]
    [switch]$Dummy
)


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

<# Calendar.ps1 %FUNCTIONS_END% #>