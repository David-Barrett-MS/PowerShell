#
# CreateEventAsDelegate.ps1
#
# By David Barrett, Microsoft Ltd. Use at your own risk.  No warranties are given.
#
#  DISCLAIMER:
# THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
# MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR
# A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL
# MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS,
# BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR INABILITY TO USE THE
# SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION
# OF LIABILITY FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.

# This script demonstrates how to create an appointment in a shared calendar using Outlook Interop.
# Shared mailbox must have permissions granted.  For example:
# Set-MailboxFolderPermission -Identity shared@domain.com:\Calendar -User delegate@domain.com -AccessRights Editor


param (
    [Parameter(Position=0,Mandatory=$True,HelpMessage="Specifies the SMTP address of the mailbox where the calendar event will be created.")]
    [ValidateNotNullOrEmpty()]
    [string]$SharedMailbox
)

$script:ScriptVersion = "1.0.0"

# Connect to Outlook
Add-Type -Assembly "Microsoft.Office.Interop.Outlook"
$OutlookApp = New-Object -ComObject "Outlook.Application"

# Obtain the shared calendar folder
$sharedMailboxRecipient = $OutlookApp.Session.CreateRecipient($SharedMailbox)
if (!$sharedMailboxRecipient.Resolve()) {
    Write-Host "Failed to resolve $SharedMailbox" -ForegroundColor Red
    exit
}
$sharedCalendar = $OutlookApp.Session.GetSharedDefaultFolder($sharedMailboxRecipient, 9) # olFolderCalendar
if ($null -eq $sharedCalendar)
{
    Write-Host "Failed to open shared calendar of $SharedMailbox" -ForegroundColor Red
    exit
}

function AddUserProperty($item, $name, $type, $value)
{
    Write-Host "Attempting to add user property $name of type $type with value $value"
    try {
        $prop = $item.UserProperties.Add($name, $type) # add to folder fields
        $prop.Value = $value
        Write-Host "Added property $name (with add to folder fields)" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Failed to add property $name (with add to folder fields)" -ForegroundColor Red
    }
    try {
        $prop = $item.UserProperties.Add($name, $type, $false) # do not add to folder fields
        $prop.Value = $value
        Write-Host "Added property $name (without add to folder fields)" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Failed to add property $name" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
    }
    return $false
}

# Create the draft message with extended property
$testAppointmentSubject = "Test Meeting created at $([DateTime]::Now)"
$startTime = [DateTime]::Now.AddHours(1)
Write-Host "Meeting start time: $startTime"
$testAppointment = $sharedCalendar.Items.Add(1) # olAppointmentItem
$testAppointment.Subject = $testAppointmentSubject
$testAppointment.Start = $startTime
$testAppointment.End = $testAppointment.Start.AddMinutes(30)
$testAppointment.Body = "This is a test appointment."

# Add two user properties.  Type 1 = string, 5 = date
if ((AddUserProperty $testAppointment "TestProperty1" 1 "TestValue1") -and (AddUserProperty $testAppointment  "TestProperty2" 5 $([DateTime]::Now.ToString()))) 
{
    $testAppointment.Save()
    Write-Host "Created appointment with user properties: $testAppointmentSubject"
}
else {
    Write-Host "Failed to add UserProperties, item not saved" -ForegroundColor Red
}

