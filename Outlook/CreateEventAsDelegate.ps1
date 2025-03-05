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
# Set the shared mailbox and run the script to create an appointment in the shared calendar.
$sharedMailbox = ""

# Connect to Outlook
Add-Type -Assembly "Microsoft.Office.Interop.Outlook"
$OutlookApp = New-Object -ComObject "Outlook.Application"

# Obtain the shared calendar folder
$sharedMailboxRecipient = $OutlookApp.Session.CreateRecipient($sharedMailbox)
if (!$sharedMailboxRecipient.Resolve()) {
    Write-Host "Failed to resolve shared mailbox recipient" -ForegroundColor Red
    exit
}
$sharedCalendar = $OutlookApp.Session.GetSharedDefaultFolder($sharedMailboxRecipient, 9) # olFolderCalendar
if ($null -eq $sharedCalendar)
{
    Write-Host "Failed to open shared calendar" -ForegroundColor Red
    exit
}

# Create the draft message with extended property
$testAppointmentSubject = "Test Meeting created at $([DateTime]::Now)"
$startTime = [DateTime]::Now.AddHours(1)
Write-Host "Meeting start time: $startTime"
try {
    $testAppointment = $sharedCalendar.Items.Add(1) # olAppointmentItem
    $testAppointment.Subject = $testAppointmentSubject
    $testAppointment.Start = $startTime
    $testAppointment.End = $testAppointment.Start.AddMinutes(30)
    $testAppointment.Body = "This is a test appointment."
    $testAppointment.Save()        
}
catch {
    Write-Host "Failed to create appointment" -ForegroundColor Red
    exit
}
Write-Host "Created appointment: $testAppointmentSubject"