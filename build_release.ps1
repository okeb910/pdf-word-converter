param(
    [string]$PythonExe = "python",
    [string]$ReleaseDir = ""
)

$ErrorActionPreference = "Stop"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$VenvDir = Join-Path $ProjectDir ".venv-build"
$VenvPython = Join-Path $VenvDir "Scripts\python.exe"

if (-not (Test-Path -LiteralPath $VenvPython)) {
    & $PythonExe -m venv $VenvDir
}

& $VenvPython -m pip install --upgrade pip
& $VenvPython -m pip install -r (Join-Path $ProjectDir "requirements-build.txt")

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

if ($ReleaseDir) {
    New-Item -ItemType Directory -Path $ReleaseDir -Force | Out-Null
    $PortableSource = Join-Path $DistDir "PDFWordConverter-v0.4.1-Portable-x64.exe"
    $PortableRelease = Join-Path $ReleaseDir "PDF-Word-PPT批量转换工具-v0.4.1-便携版-x64.exe"
    Copy-Item -LiteralPath $PortableSource -Destination $PortableRelease -Force
    Copy-Item -LiteralPath (Join-Path $ProjectDir "packaging\README-使用说明.txt") -Destination $ReleaseDir -Force
    Copy-Item -LiteralPath (Join-Path $ProjectDir "THIRD_PARTY_NOTICES.md") -Destination $ReleaseDir -Force

    $Hash = Get-FileHash -Algorithm SHA256 -LiteralPath $PortableRelease
    $HashLine = "{0}  {1}" -f $Hash.Hash.ToLowerInvariant(), (Split-Path -Leaf $Hash.Path)
    Set-Content -LiteralPath (Join-Path $ReleaseDir "SHA256SUMS.txt") -Value $HashLine -Encoding utf8
}

Write-Host "Portable build completed: $DistDir"
