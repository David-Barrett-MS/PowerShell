#
# Enable-FolderAuditing.ps1
#
# By David Barrett, Microsoft Ltd. 2024. Use at your own risk.  No warranties are given.
#
#  DISCLAIMER:
# THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
# MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR
# A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL
# MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS,
# BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR INABILITY TO USE THE
# SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION
# OF LIABILITY FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.

# Enable auditing on Outlook add-in cache folder:
# .\Enable-FolderAuditing.ps1 -FolderPath "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef\" -EnableLocalPolicy

param (
    [Parameter(Position=0,Mandatory=$True,HelpMessage="Specifies the folder on which auditing (delete/write/modify) will be enabled.")]
    [ValidateNotNullOrEmpty()]
    [string]$FolderPath,

    [Parameter(Position=1,Mandatory=$False,HelpMessage="If specified, attempt update local computer policy to enable auditing.")]
    [switch]$EnableLocalPolicy
)

if ($EnableLocalPolicy) {
    # Check audit settings
    AuditPol.exe /set /subcategory:"File System" /success:enable /failure:enable | Tee-Object -Variable setAuditResult
    if (!($setAuditResult.StartsWith("The command was successfully executed"))) {
        Write-Host "Failed to update audit settings.  Is this an administrator PowerShell session?" -ForegroundColor Red
    }
}

$AuditRule = New-Object System.Security.AccessControl.FileSystemAuditRule('Everyone', 'Delete,DeleteSubdirectoriesAndFiles,Write,Modify', 'none', 'none', 'Success,Failure')
$Acl = Get-Acl -Path $FolderPath
$Acl.AddAuditRule($AuditRule)
Set-Acl -Path $FolderPath -AclObject $Acl

(Get-Acl $FolderPath -Audit).Audit