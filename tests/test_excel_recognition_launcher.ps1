param([string]$ExePath)
$ErrorActionPreference = 'Stop'
if (-not $ExePath) { $ExePath = Join-Path $PSScriptRoot '../report_automation_launcher/bin/ReportAutomationLauncher.exe' }
$assembly = [Reflection.Assembly]::LoadFrom((Resolve-Path -LiteralPath $ExePath).Path)
$type = $assembly.GetType('ReportAutomationLauncher.BannerInspector', $true)
$flags = [Reflection.BindingFlags]'NonPublic,Static'
$title = $type.GetMethod('IsTableTitle', $flags)
$table = [string][char]0xD45C
foreach ($text in @("[$table 1] Satisfaction", "<$table 2-3> Awareness", "$table 4. Service", 'Table 5. Awareness', "$([char]0x3008)$table 6$([char]0x3009) Experience", "<$table> Experience")) {
    if (-not $title.Invoke($null, @($text))) { throw "Title variant rejected: $text" }
}
foreach ($text in @(($table + [char]0xBCF8 + [char]0xC218), 'Notes about sample size', 'Source: survey', 'Category')) {
    if ($title.Invoke($null, @($text))) { throw "Non-title accepted: $text" }
}
$check = $type.GetMethod('RunRecognitionSelfCheck', [Reflection.BindingFlags]'Public,Static')
if (-not $check) { throw 'Structural recognition regression check is missing.' }
$check.Invoke($null, @()) | Out-Null
Write-Output 'Launcher structural recognition checks passed.'
