<#
.SYNOPSIS
    Sends a test message through Exchange Online SMTP using OAuth.
.DESCRIPTION
    PowerShell translation of https://github.com/David-Barrett-MS/SMTPOAuthSample/blob/master/Program.cs.
#>
[CmdletBinding(DefaultParameterSetName = 'Delegated')]
param(
    [Parameter(Mandatory, Position = 0, ParameterSetName = 'Delegated')]
    [Parameter(Mandatory, Position = 0, ParameterSetName = 'Application')]
    [string]$TenantId,

    [Parameter(Mandatory, Position = 1, ParameterSetName = 'Delegated')]
    [Parameter(Mandatory, Position = 1, ParameterSetName = 'Application')]
    [string]$ClientId,

    [Parameter(Mandatory, Position = 2, ParameterSetName = 'Application')]
    [string]$ClientSecret,

    [Parameter(Mandatory, Position = 3, ParameterSetName = 'Application')]
    [string]$Mailbox,

    [Parameter(Position = 2, ParameterSetName = 'Delegated')]
    [Parameter(Position = 4, ParameterSetName = 'Application')]
    [string]$EmlFile,

    [string]$SmtpEndpoint = 'outlook.office365.com',
    [int]$Port = 587
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$asciiEncoding = [System.Text.Encoding]::ASCII

function Import-MsalAssembly {
    if ([Type]::GetType('Microsoft.Identity.Client.AuthenticationResult, Microsoft.Identity.Client', $false)) {
        return
    }

    try {
        Add-Type -AssemblyName 'Microsoft.Identity.Client' | Out-Null
    }
    catch {
        $localDll = Join-Path -Path $PSScriptRoot -ChildPath 'Microsoft.Identity.Client.dll'
        if (Test-Path -LiteralPath $localDll) {
            Add-Type -Path $localDll | Out-Null
        }
        else {
            throw 'Unable to load Microsoft.Identity.Client. Install the Microsoft.Identity.Client package or place Microsoft.Identity.Client.dll next to this script.'
        }
    }
}

function Get-DelegatedToken {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$ScopeHost
    )

    $scope = [string[]]("https://$ScopeHost/SMTP.Send")
    $builder = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($ClientId)
    $builder = $builder.WithAuthority([Microsoft.Identity.Client.AzureCloudInstance]::AzurePublic, $TenantId)
    $builder = $builder.WithRedirectUri('http://localhost')
    $pca = $builder.Build()

    Write-Host 'Requesting access token (user must log-in via browser)'
    return $pca.AcquireTokenInteractive($scope).ExecuteAsync().GetAwaiter().GetResult()
}

function Get-ApplicationToken {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$Secret,
        [string]$ScopeHost
    )

    $scope = [string[]]("https://$ScopeHost/.default")
    $builder = [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create($ClientId)
    $builder = $builder.WithAuthority([Microsoft.Identity.Client.AzureCloudInstance]::AzurePublic, $TenantId)
    $builder = $builder.WithClientSecret($Secret)
    $cca = $builder.Build()

    Write-Host 'Requesting access token (client credentials - no user interaction required)'
    return $cca.AcquireTokenForClient($scope).ExecuteAsync().GetAwaiter().GetResult()
}

function Get-XOAuthPayload {
    param(
        [string]$Mailbox,
        [string]$Token
    )

    $ctrlA = [char]1
    $login = "user=$Mailbox$ctrlA" + "auth=Bearer $Token$ctrlA$ctrlA"
    return [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($login))
}

function Write-SmtpStream {
    param(
        [System.IO.Stream]$Stream,
        [string]$Data
    )

    if (-not $Data.EndsWith("`r`n")) {
        $Data = "$Data`r`n"
    }

    $bytes = $asciiEncoding.GetBytes($Data)
    $Stream.Write($bytes, 0, $bytes.Length)
    $Stream.Flush()
    Write-Host $Data -NoNewline
}

function Read-SmtpStream {
    param([System.IO.Stream]$Stream)

    $buffer = New-Object byte[] 4096
    $bytesRead = $Stream.Read($buffer, 0, $buffer.Length)
    if ($bytesRead -le 0) {
        throw 'SMTP server closed the connection.'
    }

    $response = $asciiEncoding.GetString($buffer, 0, $bytesRead)
    Write-Host $response
    return $response
}

function Invoke-XOAuth2Logon {
    param(
        [System.Net.Security.SslStream]$Stream,
        [string]$Mailbox,
        [string]$Token
    )

    Write-SmtpStream -Stream $Stream -Data 'AUTH XOAUTH2'
    if (-not (Read-SmtpStream -Stream $Stream).StartsWith('334')) {
        throw 'Failed on AUTH XOAUTH2'
    }

    Write-SmtpStream -Stream $Stream -Data (Get-XOAuthPayload -Mailbox $Mailbox -Token $Token)
    if (-not (Read-SmtpStream -Stream $Stream).StartsWith('235')) {
        throw 'Log on failed'
    }
}

function Get-MessagePayload {
    param([string]$EmlFile)

    if ([string]::IsNullOrEmpty($EmlFile)) {
        return "Subject: OAuth SMTP Test Message`r`n`r`nThis is a test.`r`n.`r`n"
    }

    $content = Get-Content -LiteralPath $EmlFile -Raw
    if (-not $content.EndsWith("`r`n")) {
        $content = "$content`r`n"
    }

    return "$content.`r`n"
}

function Invoke-SmtpSession {
    param(
        [string]$Server,
        [int]$Port,
        [string]$Sender,
        [string]$Token,
        [string]$MessagePayload
    )

    $client = [System.Net.Sockets.TcpClient]::new()
    $networkStream = $null
    $sslStream = $null

    try {
        $client.Connect($Server, $Port)
        $networkStream = $client.GetStream()

        try {
            if (-not (Read-SmtpStream -Stream $networkStream).StartsWith('220')) {
                throw 'Unexpected welcome message'
            }

            Write-SmtpStream -Stream $networkStream -Data 'EHLO OAuthTest.app'
            if (-not (Read-SmtpStream -Stream $networkStream).StartsWith('250')) {
                throw 'Failed on EHLO'
            }

            Write-SmtpStream -Stream $networkStream -Data 'STARTTLS'
        }
        catch {
            try { Write-SmtpStream -Stream $networkStream -Data 'QUIT' } catch {}
            throw
        }

        if (-not (Read-SmtpStream -Stream $networkStream).StartsWith('220')) {
            throw 'Failed to start TLS'
        }

        $sslStream = [System.Net.Security.SslStream]::new($networkStream, $false)
        $sslStream.AuthenticateAsClient($Server)

        Write-SmtpStream -Stream $sslStream -Data 'EHLO'
        Read-SmtpStream -Stream $sslStream | Out-Null

        Invoke-XOAuth2Logon -Stream $sslStream -Mailbox $Sender -Token $Token

        Write-SmtpStream -Stream $sslStream -Data "MAIL FROM:<$Sender>"
        if (-not (Read-SmtpStream -Stream $sslStream).StartsWith('250')) {
            throw 'Failed at MAIL FROM'
        }

        Write-SmtpStream -Stream $sslStream -Data "RCPT TO:<$Sender>"
        if (-not (Read-SmtpStream -Stream $sslStream).StartsWith('250')) {
            throw 'Failed at RCPT TO'
        }

        Write-SmtpStream -Stream $sslStream -Data 'DATA'
        if (-not (Read-SmtpStream -Stream $sslStream).StartsWith('354')) {
            throw 'Failed at DATA'
        }

        Write-SmtpStream -Stream $sslStream -Data $MessagePayload
        if (-not (Read-SmtpStream -Stream $sslStream).StartsWith('250')) {
            throw 'Failed to send data'
        }

        Write-SmtpStream -Stream $sslStream -Data 'QUIT'
        Read-SmtpStream -Stream $sslStream | Out-Null
        Write-Host 'Closing connection'
    }
    catch [System.Net.Sockets.SocketException] {
        throw $_.Exception
    }
    finally {
        if ($sslStream) { $sslStream.Dispose() }
        if ($networkStream) { $networkStream.Dispose() }
        if ($client) { $client.Dispose() }
    }
}

Import-MsalAssembly

if ($EmlFile) {
    if (-not (Test-Path -LiteralPath $EmlFile)) {
        throw "Couldn't find email: $EmlFile"
    }
    $EmlFile = (Resolve-Path -LiteralPath $EmlFile).ProviderPath
}

try {
    if ($PSCmdlet.ParameterSetName -eq 'Application') {
        $authResult = Get-ApplicationToken -TenantId $TenantId -ClientId $ClientId -Secret $ClientSecret -ScopeHost $SmtpEndpoint
        if (-not $authResult.AccessToken) {
            throw 'No token received'
        }
        Write-Host 'Token received'
        $sender = $Mailbox
    }
    else {
        $authResult = Get-DelegatedToken -TenantId $TenantId -ClientId $ClientId -ScopeHost $SmtpEndpoint
        if (-not $authResult.AccessToken) {
            throw 'No token received'
        }
        if ($authResult.Account -and $authResult.Account.Username) {
            Write-Host ("Token received for {0}" -f $authResult.Account.Username)
            $sender = $authResult.Account.Username
        }
        else {
            Write-Host 'Token received'
            $sender = $Mailbox
        }
    }

    $payload = Get-MessagePayload -EmlFile $EmlFile
    Invoke-SmtpSession -Server $SmtpEndpoint -Port $Port -Sender $sender -Token $authResult.AccessToken -MessagePayload $payload
}
catch {
    Write-Error $_
    exit 1
}