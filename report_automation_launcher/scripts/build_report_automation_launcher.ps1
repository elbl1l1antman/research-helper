param(
    [string]$OutputDir,
    [string]$SourcePath,
    [switch]$DevelopmentOnly
)

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent $ScriptDir
$RepositoryRoot = Split-Path -Parent $ProjectRoot

if (-not $OutputDir) {
    $OutputDir = Join-Path $ProjectRoot "bin"
}
$OutputDir = [System.IO.Path]::GetFullPath($OutputDir)

if (-not $SourcePath) {
    $SourcePath = Join-Path $ProjectRoot "src\ReportAutomationLauncher.cs"
}

if (-not (Test-Path -LiteralPath $SourcePath)) {
    throw "Launcher source not found: $SourcePath"
}

$BundleSourcePath = Join-Path (Split-Path -Parent $SourcePath) "StandaloneBundle.cs"
if (-not (Test-Path -LiteralPath $BundleSourcePath)) {
    throw "Standalone bundle source not found: $BundleSourcePath"
}

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

$cscCandidates = @(
    "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe",
    "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\csc.exe"
)

$csc = $cscCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (-not $csc) {
    throw "C# compiler not found. Install .NET Framework developer tools or the .NET SDK."
}

$exePath = Join-Path $OutputDir "ReportAutomationLauncher.exe"
$staging = Join-Path ([System.IO.Path]::GetTempPath()) ("ReportAutomationBuild-" + [Guid]::NewGuid().ToString("N"))
$staging = [System.IO.Path]::GetFullPath($staging)
New-Item -ItemType Directory -Path $staging | Out-Null
try {
    $temporaryExe = Join-Path $staging "ReportAutomationLauncher.exe"
    $compilerArgs = @(
        "/nologo", "/target:winexe", "/platform:x64", "/codepage:65001",
        "/out:$temporaryExe",
        "/reference:System.dll", "/reference:System.Core.dll",
        "/reference:System.Drawing.dll", "/reference:System.Web.Extensions.dll",
        "/reference:System.Windows.Forms.dll", "/reference:Microsoft.CSharp.dll",
        "/reference:System.IO.Compression.dll", "/reference:System.IO.Compression.FileSystem.dll"
    )
    if (-not $DevelopmentOnly) {
        $python = Join-Path $RepositoryRoot ".venv\Scripts\python.exe"
        if (-not (Test-Path -LiteralPath $python)) {
            throw "Installed build Python not found: $python"
        }
        $payload = Join-Path $staging "ReportAutomation.Payload.zip"
        & $python -E -s -B (Join-Path $ScriptDir "build_standalone_payload.py") --repository $RepositoryRoot --output $payload
        if ($LASTEXITCODE -ne 0) {
            throw "Standalone payload build/smoke failed with exit code $LASTEXITCODE"
        }
        $compilerArgs += "/resource:$payload,ReportAutomation.Payload.zip"
    }
    $compilerArgs += @($SourcePath, $BundleSourcePath)
    & $csc @compilerArgs
    if ($LASTEXITCODE -ne 0) {
        throw "Launcher build failed with exit code $LASTEXITCODE"
    }
    if ((Get-Item -LiteralPath $temporaryExe).Length -ge 100MB) {
        throw "Launcher exceeds the GitHub 100 MiB file limit; the existing EXE was not replaced."
    }
    $publishTemp = $exePath + ".tmp." + [Guid]::NewGuid().ToString("N")
    try {
        Copy-Item -LiteralPath $temporaryExe -Destination $publishTemp
        if ([System.IO.File]::Exists($exePath)) {
            [System.IO.File]::Replace($publishTemp, $exePath, [System.Management.Automation.Language.NullString]::Value, $true)
        }
        else {
            [System.IO.File]::Move($publishTemp, $exePath)
        }
    }
    finally {
        if ([System.IO.File]::Exists($publishTemp)) {
            [System.IO.File]::Delete($publishTemp)
        }
    }
}
finally {
    $tempRoot = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath()).TrimEnd([char[]]"\/") + [System.IO.Path]::DirectorySeparatorChar
    if (-not $staging.StartsWith($tempRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing cleanup outside the build temporary directory: $staging"
    }
    Remove-Item -LiteralPath $staging -Recurse -Force
}

Write-Output "Created launcher: $exePath"
