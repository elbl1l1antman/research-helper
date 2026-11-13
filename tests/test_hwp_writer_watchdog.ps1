param([string]$ExePath)
$ErrorActionPreference = 'Stop'
if (-not $ExePath) { $ExePath = Join-Path $PSScriptRoot '../report_automation_launcher/bin/ReportAutomationLauncher.exe' }
if ($PSVersionTable.PSEdition -eq 'Core') {
    & powershell.exe -NoProfile -File $PSCommandPath -ExePath (Resolve-Path -LiteralPath $ExePath).Path
    if ($LASTEXITCODE -ne 0) { throw "HWP watchdog test failed: $LASTEXITCODE" }
    return
}
$assembly = [Reflection.Assembly]::LoadFrom((Resolve-Path -LiteralPath $ExePath).Path)
$engine = $assembly.GetType('ReportAutomationLauncher.EngineRunner', $true)
$wait = $engine.GetMethod('WaitForHwpWriter', [Reflection.BindingFlags]'NonPublic,Static')
$path = [IO.Path]::GetTempFileName()
$messages = [System.Collections.Generic.List[string]]::new()
$log = [Action[string]]{ param($message) $messages.Add($message) }
try {
    $checkpoints = @'
for ($i = 1; $i -le 8; $i++) {
    $data = @{ status = 'started'; section_count_selected = 3; sections_written = 1;
        progress = @{ table_key = 'T0002'; part = 2; part_count = 19; cell_count = 50; cells_done = $i } }
    [IO.File]::WriteAllText($args[0], (ConvertTo-Json -InputObject $data -Compress))
    Start-Sleep -Milliseconds 450
}
'@
    $childCommand = $checkpoints.Replace('$args[0]', ("'" + $path.Replace("'", "''") + "'"))
    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($childCommand))
    $powershell = (Get-Command powershell.exe).Source
    $active = Start-Process -FilePath $powershell -ArgumentList @('-NoProfile', '-EncodedCommand', $encoded) -WindowStyle Hidden -PassThru
    if (-not $wait.Invoke($null, @($active.PSObject.BaseObject, $path, $log, 2, 100))) { throw 'A progressing child hit the idle deadline.' }
    if ($active.ExitCode -ne 0) { throw 'The progressing child failed.' }

    $idle = Start-Process -FilePath $powershell -ArgumentList @('-NoProfile', '-Command', 'Start-Sleep -Seconds 60') -WindowStyle Hidden -PassThru
    if ($wait.Invoke($null, @($idle.PSObject.BaseObject, $path, $log, 1, 100))) { throw 'A stalled child was reported as complete.' }
    $report = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($report.status -ne 'timed_out' -or $report.document_saved -ne $false -or
        $report.progress.table_key -ne 'T0002' -or $report.progress.cells_done -ne 8 -or $report.errors.Count -ne 1 -or -not $idle.HasExited) {
        throw 'Timeout did not preserve the last checkpoint and report the failure.'
    }
    Write-Output 'HWP progress and idle timeout checks passed.'
} finally {
    foreach ($child in @($active, $idle)) {
        if ($child -and -not $child.HasExited) { $child.Kill(); $child.WaitForExit() }
    }
    Remove-Item -LiteralPath $path -Force
}
