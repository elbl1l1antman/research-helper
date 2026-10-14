param([string]$ExePath)
$ErrorActionPreference = 'Stop'
if (-not $ExePath) { $ExePath = Join-Path $PSScriptRoot '../report_automation_launcher/bin/ReportAutomationLauncher.exe' }
$assembly = [Reflection.Assembly]::LoadFrom((Resolve-Path -LiteralPath $ExePath).Path)
$type = $assembly.GetType('ReportAutomationLauncher.ExcelAddinDiagnostics', $false)
if (-not $type) { throw 'Excel add-in diagnostics are missing.' }
$flags = [Reflection.BindingFlags]'Static,NonPublic'
$open = $type.GetMethod('Open', $flags)
$check = $type.GetMethod('Check', $flags)
if (-not $open -or -not $check) { throw 'Shared open/check entry points are missing.' }

# Public test doubles are required by the dynamic COM call site in the launcher assembly.
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public class ProbeBooks {
    public int Calls;
    public bool Succeed;
    public object Open(string path, int links, bool readOnly) {
        Calls++;
        if (links != 0 || !readOnly) throw new Exception("Add-in must open read-only without updating links.");
        if (Succeed) return new object();
        throw new COMException("fixture: Excel cannot access the file", unchecked((int)0x800A03EC));
    }
}
public class ProbeExcel { public ProbeBooks Workbooks = new ProbeBooks(); }
'@
$dir = Join-Path ([IO.Path]::GetTempPath()) ('RAAddinCheck-' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $dir)
$path = Join-Path $dir 'fixture.xlam'
try {
    # Missing files must fail before touching Excel; no installation is needed for this test.
    $missing = $check.Invoke($null, [object[]]@([string]$path))
    if ($missing['status'] -ne 'blocked' -or $missing['stage'] -ne 'file_check' -or $missing['file_exists']) { throw 'Missing file classification is incorrect.' }
    if ($missing['macro_execution_verified']) { throw 'An open check cannot claim macro execution.' }
    [IO.File]::WriteAllText($path, 'fixture, not an Office document')
    $fake = New-Object ProbeExcel
    $failure = $null
    try { $null = $open.Invoke($null, [object[]]@($fake.PSObject.BaseObject, [string]$path)) } catch {
        $failure = $_.Exception
        while ($failure.InnerException -and $failure.GetType().Name -ne 'ExcelAddinOpenException') { $failure = $failure.InnerException }
    }
    if (-not $failure -or $failure.GetType().Name -ne 'ExcelAddinOpenException') { throw 'COM failure was not translated to an actionable add-in error.' }
    $report = $failure.GetType().GetField('Report').GetValue($failure)
    if ($report['stage'] -ne 'addin_open' -or $report['status'] -ne 'blocked' -or -not $report['file_exists']) { throw 'An existing but inaccessible file must not be reported as missing.' }
    if ($report['hresult'] -ne '0x800A03EC' -or $report['addin_path'] -ne $path) { throw 'COM evidence was lost.' }
    if (-not $report['message'] -or $fake.Workbooks.Calls -ne 1) { throw 'Missing explanation or unexpected retries.' }
    if ($report['macro_execution_verified']) { throw 'Failed open cannot claim macro execution.' }
    $success = New-Object ProbeExcel
    $success.Workbooks.Succeed = $true
    if (-not $open.Invoke($null, [object[]]@($success.PSObject.BaseObject, [string]$path)) -or $success.Workbooks.Calls -ne 1) { throw 'An accessible add-in must be returned unchanged.' }
} finally {
    if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    Remove-Item -LiteralPath $dir -Force
}
Write-Output 'Excel add-in diagnostics regression checks passed (no native Excel required).'
