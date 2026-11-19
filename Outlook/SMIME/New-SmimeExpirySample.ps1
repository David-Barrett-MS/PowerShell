#requires -Version 7.2
using namespace System.Security.Cryptography
using namespace System.Security.Cryptography.X509Certificates
using namespace System.Security.Cryptography.Pkcs

<#
.SYNOPSIS
Creates certificates and signed/encrypted S/MIME expiry lab messages.
.DESCRIPTION
Creates a new lab CA, valid/expired sender and recipient certificates, and eight
EML files. Exports all private keys as password-protected PFX files using the
same random password, deliberately logged in PfxPassword.txt for disposable lab
use only. Restricts the new output folder to the current Windows user and SYSTEM.
-TrustLabCA installs only the public CA in CurrentUser\Root on the lab machine.
Does not import user private keys, send mail, change system time, or use Outlook.
.PREREQUISITES
PowerShell 7.2 or later on Windows, with System.Security.Cryptography.Pkcs.
#>
[CmdletBinding(SupportsShouldProcess=$True, ConfirmImpact='None')]
param (
    [Parameter(Mandatory=$True, HelpMessage='Actual SMTP address of the disposable lab recipient configured in Outlook.')]
    [ValidateNotNullOrEmpty()]
    [ValidateScript({
        try {
            $address = [Net.Mail.MailAddress]::new($_)
            if ($_ -match '[^\x21-\x7e]' -or $address.Address -cne $_ -or $address.DisplayName) {
                throw 'Not a bare ASCII SMTP address.'
            }
            $True
        }
        catch { throw 'LabRecipient must be one bare ASCII SMTP address, without a display name or line breaks.' }
    })]
    [string]$LabRecipient,

    [Parameter(Mandatory=$False, HelpMessage='New private output directory; an existing directory is never overwritten.')]
    [ValidateNotNullOrEmpty()]
    [string]$OutputFolder = (Join-Path $PSScriptRoot ('SmimeSamples-{0}-{1}' -f [DateTime]::UtcNow.ToString('yyyyMMddTHHmmssZ'), [guid]::NewGuid().ToString('N'))),

    [Parameter(Mandatory=$False, HelpMessage='Trust the generated public CA in CurrentUser\Root on this disposable lab machine.')]
    [switch]$TrustLabCA,

    [Parameter(Mandatory=$True, HelpMessage='Acknowledges disposable lab use, exported private keys, and a plaintext password log.')]
    [switch]$LabOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if (-not $IsWindows) { throw 'This generator requires Windows for private output-folder permissions.' }
if (-not $LabOnly) { throw 'Only disposable lab use is permitted.' }
$provider = $drive = $null
$OutputFolder = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFolder, [ref]$provider, [ref]$drive)
if ($provider.Name -ne 'FileSystem') { throw 'OutputFolder must use the FileSystem provider.' }
if (Test-Path -LiteralPath $OutputFolder) { throw 'OutputFolder must not already exist.' }
if (-not $PSCmdlet.ShouldProcess($OutputFolder, 'Create private-key S/MIME lab fixtures and optionally trust their public CA')) {
    return
}
Add-Type -AssemblyName System.Security.Cryptography.Pkcs

function New-PrivateLabFolder {
    param (
        [Parameter(Mandatory=$True, HelpMessage='New folder to protect before writing key material.')]
        [string]$Path
    )
    $acl = [Security.AccessControl.DirectorySecurity]::new()
    $sid = [Security.Principal.WindowsIdentity]::GetCurrent().User
    $acl.SetOwner($sid)
    $acl.SetAccessRuleProtection($True, $False)
    $inherit = [Security.AccessControl.InheritanceFlags]::ContainerInherit -bor [Security.AccessControl.InheritanceFlags]::ObjectInherit
    foreach ($identity in @($sid, [Security.Principal.SecurityIdentifier]::new('S-1-5-18'))) {
        $acl.AddAccessRule([Security.AccessControl.FileSystemAccessRule]::new(
            $identity, [Security.AccessControl.FileSystemRights]::FullControl,
            $inherit, [Security.AccessControl.PropagationFlags]::None, [Security.AccessControl.AccessControlType]::Allow))
    }
    $null = [IO.FileSystemAclExtensions]::CreateDirectory($acl, $Path)
}

function New-LabLeaf {
    param (
        [Parameter(Mandatory=$True, HelpMessage='Email address placed in the RFC822 subject alternative name.')]
        [string]$Email,

        [Parameter(Mandatory=$True, HelpMessage='Synthetic certificate role.')]
        [ValidateSet('signer', 'recipient')]
        [string]$Role,

        [Parameter(Mandatory=$True, HelpMessage='Certificate expiry scenario.')]
        [ValidateSet('Valid', 'Expired')]
        [string]$Scenario
    )
    $key = [RSA]::Create(2048)
    $public = $null
    try {
        $request = [CertificateRequest]::new(
            "CN=S-MIME expiry lab $Role", $key, [HashAlgorithmName]::SHA256, [RSASignaturePadding]::Pkcs1)
        $request.CertificateExtensions.Add([X509BasicConstraintsExtension]::new($False, $False, 0, $True))
        $usage = [X509KeyUsageFlags]::DigitalSignature
        if ($Role -eq 'recipient') { $usage = $usage -bor [X509KeyUsageFlags]::KeyEncipherment }
        $request.CertificateExtensions.Add([X509KeyUsageExtension]::new($usage, $True))
        $eku = [OidCollection]::new()
        $null = $eku.Add([Oid]::new('1.3.6.1.5.5.7.3.4'))
        $request.CertificateExtensions.Add([X509EnhancedKeyUsageExtension]::new($eku, $False))
        $san = [SubjectAlternativeNameBuilder]::new()
        $san.AddEmailAddress($Email)
        $request.CertificateExtensions.Add($san.Build())
        $request.CertificateExtensions.Add([X509SubjectKeyIdentifierExtension]::new($request.PublicKey, $False))
        $notAfter = if ($Scenario -eq 'Expired') { $now.AddDays(-1) } else { $now.AddYears(1) }
        $serial = [RandomNumberGenerator]::GetBytes(16)
        $serial[0] = ($serial[0] -band 0x7f) -bor 1
        $public = $request.Create($root, $now.AddYears(-1), $notAfter, $serial)
        return [RSACertificateExtensions]::CopyWithPrivateKey($public, $key)
    }
    finally {
        if ($null -ne $public) { $public.Dispose() }
        $key.Dispose()
    }
}

function Test-LabCertificate {
    param (
        [Parameter(Mandatory=$True, HelpMessage='Certificate to check with an in-memory custom root.')]
        [X509Certificate2]$Certificate,

        [Parameter(Mandatory=$True, HelpMessage='Expected expiry scenario.')]
        [ValidateSet('Valid', 'Expired')]
        [string]$Scenario
    )
    $chain = [X509Chain]::new()
    try {
        $chain.ChainPolicy.TrustMode = [X509ChainTrustMode]::CustomRootTrust
        $null = $chain.ChainPolicy.CustomTrustStore.Add($root)
        $chain.ChainPolicy.RevocationMode = [X509RevocationMode]::NoCheck
        $chain.ChainPolicy.DisableCertificateDownloads = $True
        $chain.ChainPolicy.VerificationTime = $now.UtcDateTime
        $null = $chain.ChainPolicy.ApplicationPolicy.Add([Oid]::new('1.3.6.1.5.5.7.3.4'))
        $valid = $chain.Build($Certificate)
        $statuses = @($chain.ChainStatus | ForEach-Object { $_.Status.ToString() })
        if ($Scenario -eq 'Expired') {
            if ($valid -or $statuses.Count -ne 1 -or $statuses[0] -ne 'NotTimeValid') {
                throw "Expired certificate did not fail exclusively on time validity: $statuses"
            }
        }
        elseif (-not $valid) { throw "Valid control failed certificate validation: $statuses" }
        [pscustomobject]@{ ChainValidWithLabRoot = $valid; ChainStatus = $statuses }
    }
    finally { $chain.Dispose() }
}

function New-SignedContent {
    param (
        [Parameter(Mandatory=$True, HelpMessage='MIME entity to sign.')]
        [byte[]]$Content,

        [Parameter(Mandatory=$True, HelpMessage='Signing certificate including its in-memory private key.')]
        [X509Certificate2]$Certificate
    )
    $cms = [SignedCms]::new([ContentInfo]::new($Content), $False)
    $signer = [CmsSigner]::new($Certificate)
    $signer.IncludeOption = [X509IncludeOption]::EndCertOnly
    $signer.DigestAlgorithm = [Oid]::new('2.16.840.1.101.3.4.2.1')
    $null = $signer.Certificates.Add($root)
    $null = $signer.SignedAttributes.Add([Pkcs9SigningTime]::new($messageTime.UtcDateTime))
    $cms.ComputeSignature($signer, $True)
    return ,$cms.Encode()
}

function New-EncryptedContent {
    param (
        [Parameter(Mandatory=$True, HelpMessage='MIME entity to encrypt.')]
        [byte[]]$Content,

        [Parameter(Mandatory=$True, HelpMessage='Public recipient certificate, valid or expired.')]
        [X509Certificate2]$Certificate
    )
    $algorithm = [AlgorithmIdentifier]::new([Oid]::new('2.16.840.1.101.3.4.1.42'))
    $cms = [EnvelopedCms]::new([ContentInfo]::new($Content), $algorithm)
    $cms.Encrypt([CmsRecipient]::new([SubjectIdentifierType]::IssuerAndSerialNumber, $Certificate))
    return ,$cms.Encode()
}

function ConvertTo-PkcsMime {
    param (
        [Parameter(Mandatory=$True, HelpMessage='Encoded CMS signed or enveloped data.')]
        [byte[]]$Encoded,

        [Parameter(Mandatory=$True, HelpMessage='S/MIME MIME entity type.')]
        [ValidateSet('signed-data', 'enveloped-data')]
        [string]$SmimeType
    )
    $base64 = [Convert]::ToBase64String($Encoded, [Base64FormattingOptions]::InsertLineBreaks)
    return ,([Text.Encoding]::ASCII.GetBytes(
        "Content-Type: application/pkcs7-mime; smime-type=$SmimeType; name=`"smime.p7m`"`r`n" +
        "Content-Transfer-Encoding: base64`r`nContent-Disposition: attachment; filename=`"smime.p7m`"`r`n`r`n$base64`r`n"))
}

function Read-PkcsMime {
    param (
        [Parameter(Mandatory=$True, HelpMessage='Persisted EML or nested MIME entity to decode.')]
        [byte[]]$Content
    )
    $text = [Text.Encoding]::ASCII.GetString($Content)
    $separator = $text.IndexOf("`r`n`r`n", [StringComparison]::Ordinal)
    if ($separator -lt 0) { throw 'Missing MIME header/body separator.' }
    return ,([Convert]::FromBase64String($text.Substring($separator + 4)))
}

function Add-LabRootTrust {
    param (
        [Parameter(Mandatory=$True, HelpMessage='Public-only CA certificate to trust in the disposable current-user lab.')]
        [X509Certificate2]$PublicRoot
    )
    if ($PublicRoot.HasPrivateKey) { throw 'Refusing to install a CA private key in the root trust store.' }
    $store = [X509Store]::new([StoreName]::Root, [StoreLocation]::CurrentUser)
    try {
        $store.Open([OpenFlags]::ReadWrite)
        $store.Add($PublicRoot)
    }
    finally { $store.Dispose() }
}

$now = [DateTimeOffset]::UtcNow
$messageTime = $now.AddDays(-2)
$senderEmail = 'smime-lab@example.test'
$rootKey = $root = $publicRoot = $securePassword = $null
$password = $null
$certificates = @{}
$imported = @{}
$certificateResults = [Collections.Generic.List[object]]::new()
$messageResults = [Collections.Generic.List[object]]::new()
$files = [Collections.Generic.Dictionary[string,byte[]]]::new()
$rootTrusted = $False
try {
    New-PrivateLabFolder -Path $OutputFolder
    $password = [Convert]::ToBase64String([RandomNumberGenerator]::GetBytes(32))
    $securePassword = ConvertTo-SecureString -String $password -AsPlainText -Force
    [IO.File]::WriteAllText((Join-Path $OutputFolder 'PfxPassword.txt'), $password + "`r`n", [Text.Encoding]::ASCII)
    Write-Warning 'Disposable lab only: private-key PFXs and their plaintext password are stored together in the private output folder.'

    $rootKey = [RSA]::Create(2048)
    $rootRequest = [CertificateRequest]::new(
        'CN=S-MIME expiry disposable lab root', $rootKey, [HashAlgorithmName]::SHA256, [RSASignaturePadding]::Pkcs1)
    $rootRequest.CertificateExtensions.Add([X509BasicConstraintsExtension]::new($True, $True, 0, $True))
    $rootRequest.CertificateExtensions.Add([X509KeyUsageExtension]::new(
        ([X509KeyUsageFlags]::KeyCertSign -bor [X509KeyUsageFlags]::CrlSign), $True))
    $rootRequest.CertificateExtensions.Add([X509SubjectKeyIdentifierExtension]::new($rootRequest.PublicKey, $False))
    $root = $rootRequest.CreateSelfSigned($now.AddYears(-2), $now.AddYears(2))
    $certificates['LabRoot'] = $root
    foreach ($role in @('signer', 'recipient')) {
        $email = if ($role -eq 'recipient') { $LabRecipient } else { $senderEmail }
        foreach ($scenario in @('Valid', 'Expired')) {
            $name = "$scenario-$role"
            $certificate = New-LabLeaf -Email $email -Role $role -Scenario $scenario
            $certificates[$name] = $certificate
            $validation = Test-LabCertificate -Certificate $certificate -Scenario $scenario
            $certificateResults.Add([pscustomobject]@{
                Name = $name; Role = $role; Email = $email; Thumbprint = $certificate.Thumbprint
                NotBeforeUtc = $certificate.NotBefore.ToUniversalTime().ToString('o')
                NotAfterUtc = $certificate.NotAfter.ToUniversalTime().ToString('o')
                ChainValidWithLabRoot = $validation.ChainValidWithLabRoot
                ChainStatus = $validation.ChainStatus
            })
        }
    }
    foreach ($name in $certificates.Keys) {
        $certificate = $certificates[$name]
        $files.Add("$name.cer", $certificate.Export([X509ContentType]::Cert))
        $files.Add("$name.pfx", $certificate.Export([X509ContentType]::Pfx, $securePassword))
    }

    $plainMime = [Text.Encoding]::ASCII.GetBytes(
        "Content-Type: text/plain; charset=us-ascii`r`nContent-Transfer-Encoding: 7bit`r`n`r`nDisposable S/MIME expiry lab message. No customer data.`r`n")
    $signed = @{}
    foreach ($scenario in @('Valid', 'Expired')) {
        $signed[$scenario] = New-SignedContent -Content $plainMime -Certificate $certificates["$scenario-signer"]
        $messageResults.Add([pscustomobject]@{
            MessageFile = "$scenario-signed.eml"; Signer = "$scenario-signer"; Recipient = $null
            SmimeType = 'signed-data'; CmsSignatureVerified = $False; DecryptionVerified = $null
        })
        $messageResults.Add([pscustomobject]@{
            MessageFile = "$scenario-encrypted.eml"; Signer = $null; Recipient = "$scenario-recipient"
            SmimeType = 'enveloped-data'; CmsSignatureVerified = $null; DecryptionVerified = $False
        })
    }
    foreach ($signing in @('Valid', 'Expired')) {
        foreach ($encryption in @('Valid', 'Expired')) {
            $messageResults.Add([pscustomobject]@{
                MessageFile = "$signing-signed-$encryption-encrypted.eml"; Signer = "$signing-signer"; Recipient = "$encryption-recipient"
                SmimeType = 'enveloped-data'; CmsSignatureVerified = $False; DecryptionVerified = $False
            })
        }
    }
    foreach ($message in $messageResults) {
        $content = $plainMime
        if ($message.Signer) {
            $scenario = $message.Signer.Split('-')[0]
            $content = ConvertTo-PkcsMime -Encoded $signed[$scenario] -SmimeType 'signed-data'
        }
        if ($message.Recipient) {
            $encoded = New-EncryptedContent -Content $content -Certificate $certificates[$message.Recipient]
        }
        else { $encoded = $signed[$message.Signer.Split('-')[0]] }
        $entity = ConvertTo-PkcsMime -Encoded $encoded -SmimeType $message.SmimeType
        $headers = @(
            "From: S-MIME Lab <$senderEmail>"
            "To: S-MIME Lab Recipient <$LabRecipient>"
            ('Date: ' + $messageTime.UtcDateTime.ToString("ddd, dd MMM yyyy HH:mm:ss '+0000'", [Globalization.CultureInfo]::InvariantCulture))
            ('Message-ID: <{0}@example.test>' -f [guid]::NewGuid().ToString('N'))
            ('Subject: S-MIME lab - ' + $message.MessageFile.Replace('.eml', ''))
            'MIME-Version: 1.0'
        ) -join "`r`n"
        $files.Add($message.MessageFile, [Text.Encoding]::ASCII.GetBytes($headers + "`r`n" + [Text.Encoding]::ASCII.GetString($entity)))
    }
    foreach ($file in $files.GetEnumerator()) {
        [IO.File]::WriteAllBytes((Join-Path $OutputFolder $file.Key), $file.Value)
    }

    # Validate the persisted PFXs and messages, not just their in-memory inputs.
    foreach ($name in $certificates.Keys) {
        $certificate = [X509Certificate2]::new(
            [IO.File]::ReadAllBytes((Join-Path $OutputFolder "$name.pfx")), $securePassword, [X509KeyStorageFlags]::EphemeralKeySet)
        $imported[$name] = $certificate
        if (-not $certificate.HasPrivateKey -or $certificate.Thumbprint -ne $certificates[$name].Thumbprint) {
            throw "PFX private key or certificate mismatch: $name"
        }
    }
    foreach ($message in $messageResults) {
        $encoded = Read-PkcsMime -Content ([IO.File]::ReadAllBytes((Join-Path $OutputFolder $message.MessageFile)))
        if ($message.Recipient) {
            $envelope = [EnvelopedCms]::new()
            $envelope.Decode($encoded)
            $key = [RSACertificateExtensions]::GetRSAPrivateKey($imported[$message.Recipient])
            try { $envelope.Decrypt($envelope.RecipientInfos[0], $key) }
            finally { $key.Dispose() }
            $content = $envelope.ContentInfo.Content
            $message.DecryptionVerified = $True
            if ($message.Signer) { $encoded = Read-PkcsMime -Content $content }
        }
        if ($message.Signer) {
            $cms = [SignedCms]::new()
            $cms.Decode($encoded)
            $cms.CheckSignature($True)
            if ($cms.SignerInfos[0].Certificate.Thumbprint -ne $certificates[$message.Signer].Thumbprint) {
                throw 'Unexpected signer in saved message.'
            }
            $content = $cms.ContentInfo.Content
            $message.CmsSignatureVerified = $True
        }
        if ([Convert]::ToBase64String($content) -cne [Convert]::ToBase64String($plainMime)) {
            throw "Decrypted/verified content mismatch: $($message.MessageFile)"
        }
    }
    $publicRoot = [X509Certificate2]::new($root.Export([X509ContentType]::Cert))
    $manifest = [pscustomobject]@{
        GeneratedUtc = $now.ToString('o'); PowerShellVersion = $PSVersionTable.PSVersion.ToString()
        LabRecipient = $LabRecipient; Sender = $senderEmail
        PrivateKeysExported = $True; PasswordFile = 'PfxPassword.txt'
        CertificateStoresModified = $False; RootTrustedInCurrentUserRoot = $False; RootThumbprint = $root.Thumbprint
        ClaimedSigningTimeUtc = $messageTime.ToString('o'); TrustedTimestamp = $False
        Scope = 'Synthetic signed/encrypted fixtures; Outlook dialog reproduction not performed.'
        Certificates = $certificateResults.ToArray(); Messages = $messageResults.ToArray()
    }
    $manifestPath = Join-Path $OutputFolder 'Validation.json'
    $manifest | ConvertTo-Json -Depth 7 | Set-Content -LiteralPath $manifestPath -Encoding utf8 -Confirm:$False
    if ($TrustLabCA) {
        Add-LabRootTrust -PublicRoot $publicRoot
        $rootTrusted = $True
        $manifest.CertificateStoresModified = $True
        $manifest.RootTrustedInCurrentUserRoot = $True
        $manifest | ConvertTo-Json -Depth 7 | Set-Content -LiteralPath $manifestPath -Encoding utf8 -Confirm:$False
    }
    Write-Host "PFX password log: $(Join-Path $OutputFolder 'PfxPassword.txt')"
    [pscustomobject]@{
        OutputFolder = $OutputFolder; MessageCount = $messageResults.Count
        PasswordFile = Join-Path $OutputFolder 'PfxPassword.txt'
        RootThumbprint = $root.Thumbprint; RootTrusted = $rootTrusted
    }
}
finally {
    foreach ($certificate in $imported.Values) { $certificate.Dispose() }
    foreach ($certificate in $certificates.Values) { $certificate.Dispose() }
    if ($null -ne $publicRoot) { $publicRoot.Dispose() }
    if ($null -ne $rootKey) { $rootKey.Dispose() }
    if ($null -ne $securePassword) { $securePassword.Dispose() }
    $password = $null
}
