param([string]$ExePath)
$ErrorActionPreference = "Stop"
if (-not $ExePath) { $ExePath = Join-Path $PSScriptRoot '..\report_automation_launcher\bin\ReportAutomationLauncher.exe' }
$source = (Resolve-Path -LiteralPath $ExePath).Path
$directory = Join-Path ([IO.Path]::GetTempPath()) ('ResearchHelper Standalone ' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $directory | Out-Null
$savedEnvironment = @{}
foreach ($name in @('PATH', 'PYTHONHOME', 'PYTHONPATH', 'PYTHONUSERBASE', 'REPORT_AUTOMATION_PYTHON', 'REPORT_AUTOMATION_ENGINE', 'REPORT_AUTOMATION_ADDIN')) {
    $savedEnvironment[$name] = [Environment]::GetEnvironmentVariable($name, 'Process')
    [Environment]::SetEnvironmentVariable($name, (Join-Path $directory 'missing dependency'), 'Process')
}
[Environment]::SetEnvironmentVariable('PATH', "$env:WINDIR\System32;$env:WINDIR", 'Process')
try {
    $copy = Join-Path $directory 'ReportAutomationLauncher.exe'
    Copy-Item -LiteralPath $source -Destination $copy
    $report = Join-Path $directory 'standalone_report.json'
    $process = Start-Process -FilePath $copy -ArgumentList @('--self-check-standalone', '--out', ('"' + $report + '"')) -WorkingDirectory $directory -WindowStyle Hidden -Wait -PassThru
    if ($process.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $report)) { throw "Standalone check failed: exit $($process.ExitCode)" }
    $result = Get-Content -LiteralPath $report -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($result.status -ne 'ready' -or $result.gui -ne 'ready' -or -not $result.bundled -or $result.python -notlike ($result.bundle_root + '\*') -or $result.engine -notlike ($result.bundle_root + '\*') -or $result.addin -notlike ($result.bundle_root + '\*')) { throw 'Standalone initialization failed or dependencies escaped the bundle.' }
    if ($result.runtime.prefix -ne (Join-Path $result.bundle_root 'runtime')) { throw 'Python used an external installation.' }
    # Exercise actual engine commands, not only imports, from the unrelated EXE-only directory.
    $engineDirectory = Split-Path -Parent $result.engine
    $template = Join-Path $directory 'basic report.pptx'
    & $result.python -E -s -X utf8 (Join-Path $engineDirectory 'template_factory.py') --type pptx_report --output $template
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $template)) { throw 'Bundled template generation failed.' }
    $inspection = Join-Path $directory 'template inspection.json'
    & $result.python -E -s -X utf8 (Join-Path $engineDirectory 'template_inspector.py') --template $template --output $inspection
    if ($LASTEXITCODE -ne 0 -or (Get-Content -LiteralPath $inspection -Raw -Encoding UTF8 | ConvertFrom-Json).status -notin @('ready', 'usable_with_warnings')) { throw 'Bundled template inspection failed.' }
    $stamp = Get-Item -LiteralPath (Join-Path $result.bundle_root '.complete')
    $cacheTime = $stamp.LastWriteTimeUtc
    foreach ($flag in @('--self-check-hwp-page-setup', '--self-check-hwp-style-presets', '--self-check-hwp-style-cli')) {
        $check = Start-Process -FilePath $copy -ArgumentList $flag -WorkingDirectory $directory -WindowStyle Hidden -Wait -PassThru
        if ($check.ExitCode -ne 0) { throw "Launcher regression failed: $flag" }
    }
    if ((Get-Item -LiteralPath $stamp.FullName).LastWriteTimeUtc -ne $cacheTime) { throw 'Complete runtime cache was needlessly re-extracted.' }
    Write-Output ($result | ConvertTo-Json -Depth 5)
} finally {
    foreach ($name in $savedEnvironment.Keys) { [Environment]::SetEnvironmentVariable($name, $savedEnvironment[$name], 'Process') }
    $full = [IO.Path]::GetFullPath($directory)
    if (-not $full.StartsWith([IO.Path]::GetFullPath([IO.Path]::GetTempPath()), [StringComparison]::OrdinalIgnoreCase)) { throw 'Unsafe test cleanup path' }
    Remove-Item -LiteralPath $full -Recurse -Force
}
