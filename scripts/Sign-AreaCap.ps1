[CmdletBinding(DefaultParameterSetName = 'CertificateStore')]
param(
    [string]$FilePath = (Join-Path $PSScriptRoot '..\dist\areacap.exe'),
    [Parameter(Mandatory, ParameterSetName = 'Pfx')]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]$PfxPath,
    [Parameter(Mandatory, ParameterSetName = 'Pfx')]
    [SecureString]$PfxPassword,
    [Parameter(Mandatory, ParameterSetName = 'CertificateStore')]
    [ValidatePattern('^[A-Fa-f0-9]{40}$')]
    [string]$CertificateThumbprint,
    [string]$TimestampUrl = 'http://timestamp.digicert.com'
)

$ErrorActionPreference = 'Stop'
$resolvedFilePath = (Resolve-Path -LiteralPath $FilePath).Path

function Find-SignTool {
    $onPath = Get-Command signtool.exe -ErrorAction SilentlyContinue
    if ($onPath) { return $onPath.Source }

    $sdkDirectory = Join-Path ${env:ProgramFiles(x86)} 'Windows Kits\10\bin'
    if (Test-Path -LiteralPath $sdkDirectory) {
        $candidate = Get-ChildItem -LiteralPath $sdkDirectory -Recurse -Filter signtool.exe -File |
            Sort-Object FullName -Descending | Select-Object -First 1
        if ($candidate) { return $candidate.FullName }
    }

    throw 'signtool.exe was not found. Install the Windows SDK Signing Tools.'
}

$signTool = Find-SignTool
$signArguments = @('sign', '/fd', 'SHA256', '/tr', $TimestampUrl, '/td', 'SHA256')

if ($PSCmdlet.ParameterSetName -eq 'Pfx') {
    # The password is never placed on the command line.
    $certificate = Import-PfxCertificate -FilePath $PfxPath -CertStoreLocation 'Cert:\CurrentUser\My' -Password $PfxPassword
    try {
        $signArguments += @('/sha1', $certificate.Thumbprint)
        & $signTool @signArguments $resolvedFilePath
    }
    finally {
        Remove-Item -LiteralPath ("Cert:\CurrentUser\My\{0}" -f $certificate.Thumbprint) -Force -ErrorAction SilentlyContinue
    }
}
else {
    $signArguments += @('/sha1', $CertificateThumbprint)
    & $signTool @signArguments $resolvedFilePath
}

if ($LASTEXITCODE -ne 0) {
    throw "Signing failed (signtool exit code: $LASTEXITCODE)."
}

& $signTool verify '/pa' '/v' $resolvedFilePath
if ($LASTEXITCODE -ne 0) {
    throw "Signature verification failed (signtool exit code: $LASTEXITCODE)."
}

Write-Host "Signing and verification completed: $resolvedFilePath"
