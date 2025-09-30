#
# Test-SendOnBehalfOf.ps1
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

# This script Sends an email on behalf of another user

param (
	[Parameter(Mandatory=$False,HelpMessage="The account to send on behalf of.")]
	[ValidateNotNullOrEmpty()]
	[string]$SendOnBehalfOf,

	[Parameter(Mandatory=$False,HelpMessage="Subject of the message to be sent.")]
	[ValidateNotNullOrEmpty()]
	[string]$Subject = "Test Send On Behalf Of",

	[Parameter(Mandatory=$False,HelpMessage="Recipient for the message.")]
	[ValidateNotNullOrEmpty()]
	[string]$Recipient
)

# Connect to Outlook
Add-Type -Assembly "Microsoft.Office.Interop.Outlook"
$OutlookApp = New-Object -ComObject "Outlook.Application"

# Create the message
$message = $OutlookApp.CreateItem(0) # 0 = olMailItem
$message.Subject = $Subject
$message.Body = "This is a test message sent on behalf of $SendOnBehalfOf"
$message.To = $Recipient

$sendOnBehalfOfExchangeUser = $OutlookApp.Session.CreateRecipient($SendOnBehalfOf)
if ($sendOnBehalfOfExchangeUser.Resolve() -eq $false)
{
    Write-Host "Failed to resolve SendOnBehalfOf user: $SendOnBehalfOf" -ForegroundColor Red
    exit
}
Write-Host "Sending message on behalf of $($sendOnBehalfOfExchangeUser.Address) to $Recipient" -ForegroundColor Green
$message.SentOnBehalfOfName = $SendOnBehalfOf

# Send the message
$message.Send()
$OutlookApp.Quit() # This only quits Outlook if it was started by this script
