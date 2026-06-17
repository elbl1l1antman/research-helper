param(
    [string]$OutputDir,
    [string]$SourcePath
)

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent $ScriptDir

if (-not $OutputDir) {
    $OutputDir = Join-Path $ProjectRoot "bin"
}

if (-not $SourcePath) {
    $SourcePath = Join-Path $ProjectRoot "src\ReportAutomationLauncher.cs"
}

if (-not (Test-Path -LiteralPath $SourcePath)) {
    throw "Launcher source not found: $SourcePath"
}

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

$cscCandidates = @(
    "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe",
    "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\csc.exe",
    "$env:WINDIR\Microsoft.NET\Framework64\v3.5\csc.exe",
    "$env:WINDIR\Microsoft.NET\Framework\v3.5\csc.exe"
)

$csc = $cscCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (-not $csc) {
    throw "C# compiler not found. Install .NET Framework developer tools or the .NET SDK."
}

$exePath = Join-Path $OutputDir "ReportAutomationLauncher.exe"

& $csc `
    /nologo `
    /target:winexe `
    /platform:anycpu `
    /codepage:65001 `
    "/out:$exePath" `
    /reference:System.dll `
    /reference:System.Core.dll `
    /reference:System.Drawing.dll `
    /reference:System.Windows.Forms.dll `
    /reference:Microsoft.CSharp.dll `
    "$SourcePath"

if ($LASTEXITCODE -ne 0) {
    throw "Launcher build failed with exit code $LASTEXITCODE"
}

Write-Output "Created launcher: $exePath"
