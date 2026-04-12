param(
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
    [string[]]$TargetFiles = @(
        ".\build\pnet_order_reply\pnet_order_reply.exe",
        ".\dist\pnet_order_reply\pnet_order_reply.exe"
    ),

    [Parameter(Mandatory = $false)]
    [string]$TimestampServer = "http://timestamp.digicert.com"
)

$ErrorActionPreference = 'Stop'

function Get-StoreSigningCertificate {
    param(
        [string]$Thumbprint,
        [string]$Subject,
        [string]$Scope
    )

    $storePath = "Cert:\$Scope\My"
    $candidates = Get-ChildItem $storePath | Where-Object {
        $_.HasPrivateKey -and (
            $_.EnhancedKeyUsageList.FriendlyName -contains 'Code Signing' -or
            $_.EnhancedKeyUsageList.ObjectId -contains '1.3.6.1.5.5.7.3.3'
        )
    }

    if ($Thumbprint) {
        $normalizedThumbprint = ($Thumbprint -replace '\s', '').ToUpperInvariant()
        $candidates = $candidates | Where-Object { $_.Thumbprint.ToUpperInvariant() -eq $normalizedThumbprint }
    }

    if ($Subject) {
        $candidates = $candidates | Where-Object { $_.Subject -like "*$Subject*" }
    }

    $matching = @($candidates)
    if ($matching.Count -eq 0) {
        return $null
    }

    if ($matching.Count -gt 1) {
        throw "Multiple matching code-signing certificates found in $storePath. Narrow the selection with -StoreThumbprint or -StoreSubject."
    }

    return $matching[0]
}

function Import-PfxSigningCertificate {
    param(
        [string]$ResolvedPfxPath,
        [System.Security.SecureString]$ResolvedPassword,
        [string]$Scope
    )

    if (-not (Test-Path $ResolvedPfxPath)) {
        throw "PFX file not found: $ResolvedPfxPath"
    }

    if (-not $ResolvedPassword) {
        $securePassword = Read-Host "PFX password" -AsSecureString
    }
    else {
        $securePassword = $ResolvedPassword
    }

    $certificate = Import-PfxCertificate -FilePath $ResolvedPfxPath -CertStoreLocation "Cert:\$Scope\My" -Password $securePassword -Exportable
    if (-not $certificate.HasPrivateKey) {
        throw "Imported certificate does not have an accessible private key."
    }

    return $certificate
}

function Ensure-CertificateTrusted {
    param(
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
        [string]$Scope
    )

    $rootStorePath = "Cert:\$Scope\Root\$($Certificate.Thumbprint)"
    if (-not (Test-Path $rootStorePath)) {
        $rootStore = [System.Security.Cryptography.X509Certificates.X509Store]::new('Root', $Scope)
        $rootStore.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
        try {
            $rootStore.Add($Certificate)
        }
        finally {
            $rootStore.Close()
        }
    }

    $publisherStorePath = "Cert:\$Scope\TrustedPublisher\$($Certificate.Thumbprint)"
    if (-not (Test-Path $publisherStorePath)) {
        $publisherStore = [System.Security.Cryptography.X509Certificates.X509Store]::new('TrustedPublisher', $Scope)
        $publisherStore.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
        try {
            $publisherStore.Add($Certificate)
        }
        finally {
            $publisherStore.Close()
        }
    }
}

$importedFromPfx = $false
$certificate = Get-StoreSigningCertificate -Thumbprint $StoreThumbprint -Subject $StoreSubject -Scope $StoreScope
if (-not $certificate) {
    $resolvedPfxPath = (Resolve-Path $PfxPath).Path
    $certificate = Import-PfxSigningCertificate -ResolvedPfxPath $resolvedPfxPath -ResolvedPassword $PfxPassword -Scope $StoreScope
    $importedFromPfx = $true
}

if ($EnsureLocalTrust) {
    Ensure-CertificateTrusted -Certificate $certificate -Scope $StoreScope
}

try {
    foreach ($targetFile in $TargetFiles) {
        if (-not (Test-Path $targetFile)) {
            throw "Signing target not found: $targetFile"
        }

        $signature = Set-AuthenticodeSignature -FilePath $targetFile -Certificate $certificate -TimestampServer $TimestampServer
        if ($signature.Status -notin @('Valid', 'UnknownError')) {
            throw "Signing failed for $targetFile. Status: $($signature.Status). Message: $($signature.StatusMessage)"
        }

        Write-Host "Signed: $targetFile"
    }
}
finally {
    if ($importedFromPfx) {
        $storePath = "Cert:\$StoreScope\My\$($certificate.Thumbprint)"
        if (Test-Path $storePath) {
            Remove-Item $storePath -DeleteKey -Force
        }
    }
}
