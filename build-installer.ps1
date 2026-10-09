param(
    [string]$OfficeCertificateThumbprint,

    [string]$CargoTargetDirectory,

    [switch]$SkipOfficeAddinBuild,

    [switch]$SkipDesktopBuild
)

$ErrorActionPreference = 'Stop'
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

if ($CargoTargetDirectory) {
    $env:CARGO_TARGET_DIR = $CargoTargetDirectory
}

# Build and stage signed VSTO payloads before Tauri packages the NSIS installer.
$officePayloadDirectory = Join-Path $scriptRoot 'office-addin\dist'
if ($SkipOfficeAddinBuild) {
    if (-not (Test-Path $officePayloadDirectory)) {
        throw "Office add-in payload directory was not found: $officePayloadDirectory"
    }
}
else {
    if (-not $OfficeCertificateThumbprint) {
        throw 'OfficeCertificateThumbprint is required unless SkipOfficeAddinBuild is specified.'
    }
    & (Join-Path $scriptRoot 'office-addin\Build-OfficeAddins.ps1') -CertificateThumbprint $OfficeCertificateThumbprint
    if ($LASTEXITCODE -ne 0) {
        throw 'Office add-in staging failed.'
    }
}

Push-Location $scriptRoot
try {
    if ($SkipDesktopBuild) {
        cargo tauri bundle --bundles nsis --no-sign
    }
    else {
        cargo tauri build --bundles nsis
    }
    if ($LASTEXITCODE -ne 0) {
        throw 'NSIS installer build failed.'
    }
}
finally {
    Pop-Location
}
