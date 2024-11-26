
# Check we have an Exchange Online connection
if (!(Get-Command Get-App -ErrorAction SilentlyContinue)) {
    if (Get-Command Connect-ExchangeOnline -ErrorAction SilentlyContinue) {
        Connect-ExchangeOnline
    }
    if (!(Get-Command Get-App -ErrorAction SilentlyContinue)) {
        Write-Host "This script requires the Exchange Online Management module to be installed (and connection established). Please install it from the PowerShell Gallery using the command 'Install-Module ExchangeOnlineManagement'." -ForegroundColor Red
        exit
    }
}

$legacyAddins = Import-CSV "add-ins-using-exchange-tokens.csv"

Write-Host "Retrieving mailboxes..." -NoNewline
$mbxs = Get-EXOMailbox -Properties Identity -ResultSize Unlimited
Write-Host " complete"

# Check mailbox apps
$progressActivity = "Checking apps in mailboxes..."
$mbxsProcessed = 0
$global:appsFound = New-Object System.Collections.ArrayList

$mbxs | ForEach-Object {
    Get-App -Mailbox $_.identity | Select-Object -Property DisplayName, Enabled, AppVersion, AppId | ForEach-Object {
        $appIsLegacy = $legacyAddins | Where-Object ProductId -eq $_.AppId
        if ($appIsLegacy) {
            # App found in legacy list - collect the details
            $global:appsFound += $appIsLegacy
        }
    }
    $mbxsProcessed++
    Write-Progress -Activity $progressActivity -Status "$mbxsProcessed mailboxes processed (out of $($mbxs.Count))" -PercentComplete ($mbxsProcessed / $mbxs.Count * 100)
}
Write-Progress -Activity $progressActivity -Status "$mbxsProcessed mailboxes processed (out of $($mbxs.Count))" -Completed
if ($global:appsFound.Count -gt 0) {
    Write-Host "Apps found in mailboxes (stored in `$appsFound):"
    $global:appsFound
}
else {
    Write-Host "No affected apps found in mailboxes"
}

# Check organization apps
Write-Host "Checking organization apps..." -NoNewline
$global:orgAppsFound = New-Object System.Collections.ArrayList
Get-App -OrganizationApp | ForEach-Object {
    $appIsLegacy = $legacyAddins | Where-Object ProductId -eq $_.AppId
    if ($appIsLegacy) {
        # App found in legacy list - collect the details
        $global:orgAppsFound += $appIsLegacy
    }
}
Write-Host " complete"
if ($global:orgAppsFound.Count -gt 0) {
    Write-Host "Organization apps found (stored in `$orgAppsFound):"
    $global:orgAppsFound
}