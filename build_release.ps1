param(
    [string]$PythonExe = "python",
    [string]$ReleaseDir = ""
)

$ErrorActionPreference = "Stop"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$VenvDir = Join-Path $ProjectDir ".venv-build"
$VenvPython = Join-Path $VenvDir "Scripts\python.exe"
$PortableFileName = "PDF-Word-PPT-Converter-v0.5.0-Portable-x64.exe"
$LegacyPortableFileName = "PDF-Word-PPT批量转换工具-v0.5.0-便携版-x64.exe"

function Assert-Python312X64([string]$Executable, [string]$Description) {
    & $Executable -c "import struct, sys; raise SystemExit(0 if sys.version_info[:2] == (3, 12) and struct.calcsize('P') == 8 else 1)"
    if ($LASTEXITCODE -ne 0) {
        throw "$Description must be Python 3.12 x64: $Executable"
    }
}

if (-not $ReleaseDir) {
    $ReleaseDir = Join-Path $ProjectDir "release\v0.5.0"
}

Assert-Python312X64 $PythonExe "Build interpreter"
if (-not (Test-Path -LiteralPath $VenvPython)) {
    & $PythonExe -m venv $VenvDir
    if ($LASTEXITCODE -ne 0) {
        throw "Could not create build environment: $VenvDir"
    }
}
Assert-Python312X64 $VenvPython "Build environment"

& $VenvPython -m pip install --upgrade pip
if ($LASTEXITCODE -ne 0) { throw "Could not upgrade pip" }
& $VenvPython -m pip install -r (Join-Path $ProjectDir "requirements-build.txt")
if ($LASTEXITCODE -ne 0) { throw "Could not install build requirements" }

$BuildDir = Join-Path $ProjectDir "build"
$DistDir = Join-Path $ProjectDir "dist"
if (Test-Path -LiteralPath $BuildDir) {
    Remove-Item -LiteralPath $BuildDir -Recurse -Force
}
if (Test-Path -LiteralPath $DistDir) {
    Remove-Item -LiteralPath $DistDir -Recurse -Force
}

Push-Location $ProjectDir
try {
    & $VenvPython -m PyInstaller --noconfirm --clean "packaging\PDFWordConverter.spec"
    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller build failed with exit code $LASTEXITCODE"
    }
}
finally {
    Pop-Location
}

New-Item -ItemType Directory -Path $ReleaseDir -Force | Out-Null
$PortableSource = Join-Path $DistDir $PortableFileName
if (-not (Test-Path -LiteralPath $PortableSource -PathType Leaf)) {
    throw "Portable executable was not created: $PortableSource"
}

$PortableRelease = Join-Path $ReleaseDir $PortableFileName
$LegacyPortableRelease = Join-Path $ReleaseDir $LegacyPortableFileName
if ($LegacyPortableRelease -ne $PortableRelease -and (Test-Path -LiteralPath $LegacyPortableRelease)) {
    Remove-Item -LiteralPath $LegacyPortableRelease -Force
}

Copy-Item -LiteralPath $PortableSource -Destination $PortableRelease -Force
Copy-Item -LiteralPath (Join-Path $ProjectDir "packaging\README-使用说明.txt") -Destination $ReleaseDir -Force
Copy-Item -LiteralPath (Join-Path $ProjectDir "CHANGELOG.md") -Destination $ReleaseDir -Force
Copy-Item -LiteralPath (Join-Path $ProjectDir "THIRD_PARTY_NOTICES.md") -Destination $ReleaseDir -Force
Copy-Item -LiteralPath (Join-Path $ProjectDir "LICENSE") -Destination $ReleaseDir -Force

$ReleaseLicensesDir = Join-Path $ReleaseDir "licenses"
if (Test-Path -LiteralPath $ReleaseLicensesDir) {
    Remove-Item -LiteralPath $ReleaseLicensesDir -Recurse -Force
}
Copy-Item -LiteralPath (Join-Path $ProjectDir "packaging\licenses") -Destination $ReleaseLicensesDir -Recurse -Force

$LicenseArchive = Join-Path $ReleaseDir "THIRD-PARTY-LICENSES-v0.5.0.zip"
if (Test-Path -LiteralPath $LicenseArchive) {
    Remove-Item -LiteralPath $LicenseArchive -Force
}
$LicenseInputs = @(
    (Join-Path $ReleaseDir "LICENSE"),
    (Join-Path $ReleaseDir "THIRD_PARTY_NOTICES.md"),
    $ReleaseLicensesDir
)
Compress-Archive -LiteralPath $LicenseInputs -DestinationPath $LicenseArchive -CompressionLevel Optimal

$Hash = Get-FileHash -Algorithm SHA256 -LiteralPath $PortableRelease
$HashLine = "{0}  {1}" -f $Hash.Hash.ToLowerInvariant(), (Split-Path -Leaf $Hash.Path)
$Utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllLines(
    (Join-Path $ReleaseDir "SHA256SUMS.txt"),
    [string[]]@($HashLine),
    $Utf8NoBom
)

Write-Host "Portable release completed: $ReleaseDir"
