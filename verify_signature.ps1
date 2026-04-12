param(
    [Parameter(Mandatory = $false)]
    [string[]]$TargetFiles = @(
        ".\build\pnet_order_reply\pnet_order_reply.exe",
        ".\dist\pnet_order_reply\pnet_order_reply.exe"
    )
)

$ErrorActionPreference = 'Stop'

$results = foreach ($targetFile in $TargetFiles) {
    if (-not (Test-Path $targetFile)) {
        [pscustomobject]@{
            Path = $targetFile
            Status = 'Missing'
            StatusMessage = 'File not found'
            Signer = $null
            Thumbprint = $null
        }
        continue
    }

    $signature = Get-AuthenticodeSignature $targetFile
    [pscustomobject]@{
        Path = $targetFile
        Status = $signature.Status
        StatusMessage = $signature.StatusMessage
        Signer = $signature.SignerCertificate.Subject
        Thumbprint = $signature.SignerCertificate.Thumbprint
    }
}

$results | Format-Table -AutoSize

$invalid = $results | Where-Object { $_.Status -ne 'Valid' }
if ($invalid) {
    throw "One or more files are not validly signed."
}
