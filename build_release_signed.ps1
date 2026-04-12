param(
    [Parameter(Mandatory = $false)]
    [string]$Version = "1.0.12",

    [Parameter(Mandatory = $false)]
    [string]$PfxPath = ".\pidskr_cert.pfx",

    [Parameter(Mandatory = $false)]
    [System.Security.SecureString]$PfxPassword,

    [Parameter(Mandatory = $false)]
    [string]$StoreThumbprint,

    [Parameter(Mandatory = $false)]
    [string]$StoreSubject,

    [Parameter(Mandatory = $false)]
    [ValidateSet('CurrentUser', 'LocalMachine')]
    [string]$StoreScope = 'CurrentUser',

    [Parameter(Mandatory = $false)]
    [switch]$EnsureLocalTrust,

    [Parameter(Mandatory = $false)]
    [string]$TimestampServer = "http://timestamp.digicert.com"
)

$ErrorActionPreference = 'Stop'

$python = "C:/Users/ryoum/.local/bin/python3.14.exe"
& $python -m PyInstaller pnet_order_reply.spec --noconfirm

& .\sign_release.ps1 -PfxPath $PfxPath -PfxPassword $PfxPassword -StoreThumbprint $StoreThumbprint -StoreSubject $StoreSubject -StoreScope $StoreScope -EnsureLocalTrust:$EnsureLocalTrust -TimestampServer $TimestampServer
& .\verify_signature.ps1

$zipName = "pnet_order_reply_v$Version.zip"
$zipPath = Join-Path (Get-Location) $zipName
if (Test-Path $zipPath) {
    Remove-Item $zipPath -Force
}

Compress-Archive -Path .\dist\pnet_order_reply\* -DestinationPath $zipPath -Force
Write-Host "Created release package: $zipPath"
