param(
    [string]$PythonExe = "python",
    [string]$InnoCompiler = "",
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

    if (-not $InnoCompiler) {
        $Candidates = @(
            "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe",
            "$env:ProgramFiles\Inno Setup 6\ISCC.exe",
            "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe"
        )
        $InnoCompiler = $Candidates | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1
    }
    if (-not $InnoCompiler) {
        throw "Inno Setup 6 compiler (ISCC.exe) was not found."
    }

    & $InnoCompiler "packaging\installer.iss"
    if ($LASTEXITCODE -ne 0) {
        throw "Inno Setup build failed with exit code $LASTEXITCODE"
    }
}
finally {
    Pop-Location
}

if ($ReleaseDir) {
    New-Item -ItemType Directory -Path $ReleaseDir -Force | Out-Null
    $PortableSource = Join-Path $DistDir "PDFWordConverter-v0.4.0-Portable-x64.exe"
    $SetupSource = Join-Path $DistDir "PDFWordConverter-v0.4.0-Setup-x64.exe"
    $PortableRelease = Join-Path $ReleaseDir "PDF-Word-PPT批量转换工具-v0.4.0-便携版-x64.exe"
    $SetupRelease = Join-Path $ReleaseDir "PDF-Word-PPT批量转换工具-v0.4.0-安装版-x64.exe"
    Copy-Item -LiteralPath $PortableSource -Destination $PortableRelease -Force
    Copy-Item -LiteralPath $SetupSource -Destination $SetupRelease -Force
    Copy-Item -LiteralPath (Join-Path $ProjectDir "packaging\README-使用说明.txt") -Destination $ReleaseDir -Force

    $Hashes = Get-FileHash -Algorithm SHA256 -LiteralPath $PortableRelease, $SetupRelease
    $HashLines = $Hashes | ForEach-Object { "{0}  {1}" -f $_.Hash.ToLowerInvariant(), (Split-Path -Leaf $_.Path) }
    Set-Content -LiteralPath (Join-Path $ReleaseDir "SHA256SUMS.txt") -Value $HashLines -Encoding utf8
}

Write-Host "Build completed: $DistDir"
