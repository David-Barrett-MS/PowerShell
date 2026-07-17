param(
    [Parameter(Mandatory=$False, HelpMessage="Optional file path to write output to (in addition to console)")]
    [string]$WriteToFile
)

function Write-Output-Dual {
    param(
        [string]$Message,
        [string]$ForegroundColor = "White"
    )
    
    Write-Host $Message -ForegroundColor $ForegroundColor
    if ($WriteToFile) {
        $Message | Out-File -FilePath $WriteToFile -Append -Encoding UTF8
    }
}

# Initialize output file if specified
if ($WriteToFile) {
    "" | Out-File -FilePath $WriteToFile -Encoding UTF8
}

$keys = @(
    "HKCU:\Software\Microsoft\Office\Outlook\Addins", # Install path
    "HKCU:\Software\Microsoft\Office\16.0\Outlook\Addins", # Data/telemetry
	"HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency",
	"HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList", # Not set by policy
	"HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency\NotificationReminderAddinData",
	"HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency\DisabledItems",
	"HKCU:\Software\Policies\Microsoft\Office\16.0\Outlook\Resiliency",
    "HKCU:\Software\Policies\Microsoft\Office\16.0\Outlook\Resiliency\addinlist") # Set when GPO always enables add-in

foreach ($key in $keys) {
    if (-not (Test-Path $key)) {
        Write-Output-Dual -Message "`n[$key] - missing" -ForegroundColor Red
        continue
    }
    Write-Output-Dual -Message "`n[$key]" -ForegroundColor Green
    
    # Display values
    $props = Get-ItemProperty -Path $key
    $values = $props.PSObject.Properties | Where-Object { $_.Name -notin 'PSPath','PSParentPath','PSChildName','PSDrive','PSProvider' }
    if ($values) {
        foreach ($prop in $values) {
            Write-Output-Dual -Message "  $($prop.Name) = $($prop.Value)"
        }
    }
    
    # Display direct subkeys
    $subkeys = Get-ChildItem -Path $key -ErrorAction SilentlyContinue
    if ($subkeys) {
        foreach ($subkey in $subkeys) {
            Write-Output-Dual -Message "`t[$($subkey.PSChildName)]" -ForegroundColor Green
        }
    }
}