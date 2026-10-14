param(
    [string]$AddinPath = (Join-Path $PSScriptRoot '..\report_automation_addin\dev\ReportAutomationAddin_dev.xlam'),
    [string]$WorkbookPath = '',
    [int]$ExpectedTables = 127,
    [int]$TimeoutSeconds = 300,
    [switch]$Worker,
    [string]$PidLogFile = ''
)

# Run after the serialized add-in build. No VBProject access or registry changes.
# Fixtures and optional real input are never saved; only our Excel instance quits.
# ASCII source intentionally works in Windows PowerShell 5 without a UTF-8 BOM.
$ErrorActionPreference = 'Stop'
$AddinPath = (Resolve-Path -LiteralPath $AddinPath).Path
if (-not $Worker) {
    # A child process bounds a synchronous Excel.Run hang or hidden compile dialog.
    $runDir = Join-Path ([IO.Path]::GetTempPath()) ('excel-recognition-' + [guid]::NewGuid().ToString('N'))
    [void](New-Item -ItemType Directory -Path $runDir)
    $pidFile = Join-Path $runDir 'owned-excel.pid'
    $stdout = Join-Path $runDir 'stdout.log'
    $stderr = Join-Path $runDir 'stderr.log'
    $child = $null
    try {
        $hostExe = (Get-Process -Id $PID).Path
        $argsList = @('-NoProfile', '-File', ('"' + $PSCommandPath + '"'),
            '-Worker', '-PidLogFile', ('"' + $pidFile + '"'), '-AddinPath', ('"' + $AddinPath + '"'),
            '-ExpectedTables', $ExpectedTables)
        if ($WorkbookPath) { $argsList += @('-WorkbookPath', ('"' + $WorkbookPath + '"')) }
        $child = Start-Process -FilePath $hostExe -ArgumentList $argsList -WindowStyle Hidden -PassThru -RedirectStandardOutput $stdout -RedirectStandardError $stderr
        $null = $child.Handle
        if (-not $child.WaitForExit($TimeoutSeconds * 1000)) {
            # Only the HWND-derived PID created and logged by this worker is eligible.
            if (Test-Path -LiteralPath $pidFile) {
                $identity = (Get-Content -LiteralPath $pidFile -Raw).Trim().Split('|')
                $excelId = [int]$identity[0]
                $ownedExcel = Get-Process -Id $excelId -ErrorAction SilentlyContinue
                if ($null -ne $ownedExcel -and $ownedExcel.ProcessName -eq 'EXCEL' -and $ownedExcel.StartTime -ge $child.StartTime -and $ownedExcel.StartTime.ToUniversalTime().Ticks -eq [long]$identity[1]) {
                    Stop-Process -Id $excelId -Force
                }
                Write-Output "OwnedExcelPID: $excelId (timeout cleanup)"
            }
            Stop-Process -Id $child.Id -Force -ErrorAction SilentlyContinue
            throw "VBA recognition timed out after $TimeoutSeconds seconds; only the test worker and its verified Excel instance were stopped."
        }
        if (Test-Path -LiteralPath $stdout) { Get-Content -LiteralPath $stdout }
        if ($child.ExitCode -ne 0) {
            if (Test-Path -LiteralPath $stderr) { Get-Content -LiteralPath $stderr }
            throw "VBA recognition worker failed with exit code $($child.ExitCode)."
        }
    } finally {
        if ($null -ne $child) { $child.Dispose() }
        # Remove only individually named files in the newly created private directory.
        foreach ($file in @($pidFile, $stdout, $stderr)) {
            if (Test-Path -LiteralPath $file) { Remove-Item -LiteralPath $file -Force }
        }
        Remove-Item -LiteralPath $runDir -Force
    }
    return
}
function U([int[]]$Codes) { -join ($Codes | ForEach-Object { [char]$_ }) }
$tableMark = U @(0xD45C)
$cases = U @(0xC0AC, 0xB840, 0xC218)
$total = U @(0xC804, 0xCCB4)
$sum = U @(0xD569, 0xACC4)
$stub = U @(0xAD6C, 0xBD84)
$item = U @(0xD56D, 0xBAA9)
$ratio = U @(0xBE44, 0xC728)
$source = U @(0xCD9C, 0xCC98)
$unknownWarning = U @(0xD5E4, 0xB354, 0x20, 0xC758, 0xBBF8, 0x20, 0xBBF8, 0xD655, 0xC778)
$titleWarning = U @(0xC81C, 0xBAA9, 0x20, 0xBBF8, 0xD655, 0xC778)
$square = U @(0x25A0)
$groupMark = U @(0x25D0)
$angleOpen = U @(0x3008)
$angleClose = U @(0x3009)
$formSpace = U @(0x3000)
$reportPrefix = (U @(0xBCF4, 0xACE0, 0xC11C)) + '_'
$business = U @(0xC0AC, 0xC5C5, 0xC7A5)
$respondents = U @(0xC751, 0xB2F5, 0xC778, 0xC6D0)
$agree = U @(0xCC2C, 0xC131)
$disagree = U @(0xBC18, 0xB300)

function Assert-Equal($Actual, $Expected, [string]$Message) {
    if ($Actual -ne $Expected) { throw "$Message Expected <$Expected>, got <$Actual>." }
}
function Assert-Contains([string]$Actual, [string]$Expected, [string]$Message) {
    if (-not $Actual.Contains($Expected)) { throw "$Message Missing <$Expected> in <$Actual>." }
}
function Write-Row($Sheet, [int]$Row, [int]$FirstCol, [object[]]$Values) {
    for ($i = 0; $i -lt $Values.Count; $i++) {
        if ($null -ne $Values[$i]) {
            $cell = $Sheet.Cells.Item($Row, $FirstCol + $i)
            try {
                $v = $Values[$i]
                if ($v -is [int]) { $v = [double]$v }
                # Invoke the VARIANT setter directly; PS5 may cache its type after a string cell.
                $cell.GetType().InvokeMember('Value2', [Reflection.BindingFlags]::SetProperty, $null, $cell, @($v))
            } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($cell) }
        }
    }
}
function New-Fixture([string]$Name) {
    $sheet = $fixtures.Worksheets.Add()
    $sheet.Name = $Name
    return $sheet
}
function Read-Tables($Sheet, [string]$Settings = '') {
    $result = $excel.Run($macro, $Sheet.Parent.Name, $Sheet.Name, $Settings)
    if ($null -ne $result) { return ,$result }
}
function Count-Tables($Records) {
    if ($null -eq $Records) { return 0 }
    return $Records.GetLength(0)
}
# First 14 columns retain the VBA record offsets. Diagnostics occupy 14..19.
function Field($Records, [int]$Table, [int]$Index) {
    $Records.GetValue($Records.GetLowerBound(0) + $Table - 1, $Records.GetLowerBound(1) + $Index)
}
function Assert-One($Records, [string]$Name) { Assert-Equal (Count-Tables $Records) 1 "$Name count." }

$excel = $null
$addin = $null
$fixtures = $null
$realInput = $null
$ownedSheets = New-Object System.Collections.ArrayList
try {
    $excel = New-Object -ComObject Excel.Application
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class RecognitionFixtureNative {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr handle, out uint processId);
}
'@
    $ownedExcelId = [uint32]0
    [void][RecognitionFixtureNative]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$ownedExcelId)
    if ($ownedExcelId -eq 0) { throw 'Cannot identify the isolated Excel PID.' }
    Write-Output "OwnedExcelPID: $ownedExcelId"
    if ($PidLogFile) {
        $startTicks = (Get-Process -Id $ownedExcelId).StartTime.ToUniversalTime().Ticks
        [IO.File]::WriteAllText($PidLogFile, "$ownedExcelId|$startTicks", [Text.Encoding]::ASCII)
    }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    # Respect the existing trust/macro policy, including company-managed settings.
    $addin = $excel.Workbooks.Open($AddinPath, 0, $true)
    $macro = "'" + $addin.Name.Replace("'", "''") + "'!ReportAutomation_GetTables"
    $fixtures = $excel.Workbooks.Add()

    $titles = @("<$tableMark> Experience", "[ $tableMark ] Experience", "[ $tableMark 1-2 ] Experience", "$tableMark 1-2. Experience", 'Table 3. Experience', "[$tableMark 1> Experience]",
        "[$tableMark MDBR-6 > Experience]", "[$tableMark MQT2 > Experience]", "$angleOpen${tableMark}1${angleClose} Experience", "[$tableMark${formSpace}1 > Experience]")
    $caseIndex = 0
    foreach ($title in $titles) {
        foreach ($shift in @(1, 2, 3)) {
            $caseIndex++
            $ws = New-Fixture "Title$caseIndex"
            [void]$ownedSheets.Add($ws)
            Write-Row $ws 1 $shift @($title)
            Write-Row $ws 2 2 @($stub, $cases, '%')
            Write-Row $ws 3 2 @('First', 60, 60)
            Write-Row $ws 4 2 @('Second', 40, 40)
            $rec = Read-Tables $ws
            Assert-One $rec $ws.Name
            Assert-Equal (Field $rec 1 2) 'Experience' 'Clean title.'
            if ($title.Contains('MDBR-6')) { Assert-Equal (Field $rec 1 1) 'MDBR-6' 'Alphanumeric ID.' }
            if ($title.Contains('MQT2')) { Assert-Equal (Field $rec 1 1) 'MQT2' 'Compact alphanumeric ID.' }
            Assert-Equal (Field $rec 1 10) 2 'First column, no A margin.'
            Assert-Equal (Field $rec 1 12) 3 'First body row.'
            Assert-Equal (Field $rec 1 14) 'B2:D4' 'Source range.'
            Assert-Equal (Field $rec 1 15) 2 'First body item retained.'
            Assert-Contains (Field $rec 1 16) 'First=60.0%(D3)' 'First body source.'
        }
    }

    $ws = New-Fixture 'LongGaps'
    [void]$ownedSheets.Add($ws)
    Write-Row $ws 1 2 @('Table 1. Long gaps')
    Write-Row $ws 45 2 @($stub, 'N', '%')
    Write-Row $ws 90 2 @('First', 60, 60)
    Write-Row $ws 140 2 @('Second', 40, 40)
    Write-Row $ws 170 2 @('Table 2. Next question')
    Write-Row $ws 200 2 @($stub, 'N', '%')
    Write-Row $ws 240 2 @('Third', 100, 100)
    $rec = Read-Tables $ws
    Assert-Equal (Count-Tables $rec) 2 'Long blank gaps use meaningful rows.'
    Assert-Equal (Field $rec 1 12) 90 'Long-gap first body.'
    Assert-Equal (Field $rec 1 6) 140 'Next title not swallowed across gap.'
    Assert-Equal (Field $rec 1 15) 2 'Separated body rows preserved.'

    foreach ($merged in @($false, $true)) {
        $ws = New-Fixture ('Plain' + $merged)
        [void]$ownedSheets.Add($ws)
        Write-Row $ws 1 2 @('Current question')
        if ($merged) { $ws.Range('B1:E1').Merge() }
        Write-Row $ws 4 2 @($item, $null, $ratio)
        Write-Row $ws 6 2 @('First', $null, 60)
        Write-Row $ws 7 2 @('Second', $null, 40)
        $rec = Read-Tables $ws
        Assert-One $rec $ws.Name
        Assert-Equal (Field $rec 1 2) 'Current question' 'Plain question title.'
        Assert-Equal (Field $rec 1 12) 6 'Spaced first body row.'
        Assert-Equal (Field $rec 1 14) 'B4:D7' 'Internal blank C is preserved; leading blank rows are trimmed.'
        Assert-Equal (Field $rec 1 15) 2 'No-count distribution body.'
    }

    $ws = New-Fixture 'SplitTitle'
    [void]$ownedSheets.Add($ws)
    Write-Row $ws 1 2 @("[ $tableMark 1-2 ]", 'Experience')
    Write-Row $ws 2 2 @($stub, 'N', 'Percentage')
    Write-Row $ws 3 2 @('First', 60, 60)
    Write-Row $ws 4 2 @('Second', 40, 40)
    $rec = Read-Tables $ws
    Assert-One $rec 'Split title'
    Assert-Equal (Field $rec 1 1) '1-2' 'Split ID.'
    Assert-Equal (Field $rec 1 2) 'Experience' 'Split title cleanup.'

    $ws = New-Fixture 'NumericStub'
    [void]$ownedSheets.Add($ws)
    Write-Row $ws 1 1 @('Table 1. Annual response')
    Write-Row $ws 2 1 @((U @(0xC5F0, 0xB3C4)), 'N', '%')
    Write-Row $ws 3 1 @(2023, 100, 60)
    Write-Row $ws 4 1 @(2024, 120, 70)
    $rec = Read-Tables $ws
    Assert-One $rec 'Declared numeric year stub'
    Assert-Equal (Field $rec 1 12) 3 'Numeric-label first body retained.'
    Assert-Equal (Field $rec 1 15) 2 'Numeric labels are categories, not counts.'
    Assert-Contains (Field $rec 1 16) '2023=60.0%(C3)' 'Declared numeric label source.'

    $formats = @(
        @{ Value = 0.633; Format = '0.0%'; Summary = 'First=63.3%(D3)' },
        @{ Value = 63.3; Format = '0.0'; Summary = 'First=63.3%(D3)' },
        @{ Value = 0.633; Format = '0.0"%"'; Summary = 'First=0.6%(D3)' },
        @{ Value = 0.633; Format = '0.0\%'; Summary = 'First=0.6%(D3)' },
        @{ Value = '63.3%'; Format = '@'; Summary = 'First=63.3%(D3)' },
        @{ Value = -0.633; Format = '0.0%;0.0'; Summary = 'First=-0.6%(D3)' }
    )
    foreach ($format in $formats) {
        $ws = New-Fixture ('Percent' + $ownedSheets.Count)
        [void]$ownedSheets.Add($ws)
        Write-Row $ws 1 2 @('Table 1. Formatted percentages')
        Write-Row $ws 2 2 @($stub, 'N', '%')
        $ws.Range('D3').NumberFormat = $format.Format
        Write-Row $ws 3 2 @('First', 60, $format.Value)
        Write-Row $ws 4 2 @('Second', 40, 40)
        $displayBefore = [string]$ws.Range('D3').Text
        $rawBefore = $ws.Range('D3').Value2
        $formatBefore = $ws.Range('D3').NumberFormat
        $rec = Read-Tables $ws
        Assert-One $rec 'Percent formatting'
        $expectedSummary = $format.Summary
        # Korean Excel may translate ASCII backslash to U+20A9 (currency), leaving % active.
        if ($format.Format -eq '0.0\%' -and $formatBefore.Contains([string][char]0x20A9)) {
            $expectedSummary = 'First=63.3%(D3)'
        }
        Assert-Contains (Field $rec 1 16) $expectedSummary ("Percent format <" + $format.Format + ">, actual <" + $formatBefore + ">, raw <" + $rawBefore + ">, display <" + $displayBefore + ">.")
        Assert-Equal $ws.Range('D3').Value2 $rawBefore 'Raw percent value unchanged.'
        Assert-Equal $ws.Range('D3').Text $displayBefore 'Percent display unchanged.'
        Assert-Equal $ws.Range('D3').NumberFormat $formatBefore 'Percent format unchanged.'
    }

    $ws = New-Fixture 'Adjacent'
    [void]$ownedSheets.Add($ws)
    Write-Row $ws 1 2 @('Table 1. First question')
    Write-Row $ws 2 2 @($stub, $cases, '%')
    Write-Row $ws 3 2 @('First', 60, 60)
    Write-Row $ws 4 2 @('Table 2. Second question')
    Write-Row $ws 5 2 @($stub, 'N', 'Percent')
    Write-Row $ws 6 2 @('Second', 40, 40)
    Write-Row $ws 7 2 @("${source}: synthetic", 2025, 99)
    Write-Row $ws 8 2 @($stub, $cases, '%')
    Write-Row $ws 9 2 @('Third', 100, 100)
    $rec = Read-Tables $ws
    Assert-Equal (Count-Tables $rec) 3 'Adjacent independent tables, including titleless.'
    Assert-Equal (Field $rec 1 6) 3 'First table end.'
    Assert-Equal (Field $rec 2 6) 6 'Source note excluded.'
    Assert-Equal (Field $rec 3 14) 'B8:D9' 'Titleless header is not lost.'
    Assert-Contains (Field $rec 3 9) $titleWarning 'Titleless warning.'
    Assert-Equal (Field $rec 3 0) 'T0003' 'Stable vertical key ordering.'

    foreach ($totalLabel in @($total, $sum, 'Total', "$square $total")) {
        $ws = New-Fixture ('Total' + $ownedSheets.Count)
        [void]$ownedSheets.Add($ws)
        Write-Row $ws 1 1 @('Table 1. Experience')
        Write-Row $ws 2 1 @('BASE: all')
        Write-Row $ws 3 1 @($stub, $null, $cases, 'Yes', $null, 'No')
        $ws.Range('D3:E3').Merge()
        $ws.Range('F3:G3').Merge()
        Write-Row $ws 4 3 @($cases, 'N', '%', 'N', '%')
        Write-Row $ws 5 1 @($totalLabel, $null, 100, 65, 65, 35, 35)
        Write-Row $ws 6 2 @('Male', 40, 30, 75, 10, 25)
        $rec = Read-Tables $ws
        Assert-One $rec 'Old wide N/% pairs'
        Assert-Equal (Field $rec 1 12) 5 'Merged header first body.'
        Assert-Equal (Field $rec 1 15) 2 'Only response options, not counts or totals.'
        Assert-Contains (Field $rec 1 16) 'Yes=65.0%(E5)' 'Old wide percentage source.'
        Assert-Equal (Field $rec 1 17) 0 'Response options are not banner groups.'
        Assert-Equal (Field $rec 1 19) 5 'Plain total detection.'
        Assert-Equal $ws.Range('D3:E3').MergeCells $true 'Source merge unchanged.'
        $ws.Range('E5').NumberFormat = '0.0%'
        $ws.Range('E5').Value2 = [double]0.633
        $rec = Read-Tables $ws
        Assert-Contains (Field $rec 1 16) 'Yes=63.3%(E5)' 'Wide real percent format scales once.'
    }

    $ws = New-Fixture 'KisdiSchema'
    [void]$ownedSheets.Add($ws)
    Write-Row $ws 1 1 @("[$tableMark 1> Experience]")
    Write-Row $ws 2 1 @('BASE: all')
    Write-Row $ws 4 3 @("wt$cases", $cases, '%')
    Write-Row $ws 5 1 @("$square$total", $null, 100, 90, 100)
    Write-Row $ws 6 2 @('First', 60, 54, 60)
    Write-Row $ws 7 1 @('Region')
    $ws.Range('A7:B7').Merge()
    Write-Row $ws 8 2 @('Second', 40, 36, 40)
    Write-Row $ws 9 1 @("${source}: fixture", 2025, 99)
    $rec = Read-Tables $ws
    Assert-One $rec 'KISDI columns'
    Assert-Equal (Field $rec 1 12) 5 'Total at first body.'
    Assert-Equal (Field $rec 1 6) 8 'Markerless stub banner does not split body.'
    Assert-Equal (Field $rec 1 15) 2 'Totals are not categories.'
    Assert-Contains (Field $rec 1 16) 'First=60.0%(E6)' 'KISDI percentage source.'

    foreach ($parent in @('Gender', "$groupMark Gender")) {
        $ws = New-Fixture ('Parent' + $ownedSheets.Count)
        [void]$ownedSheets.Add($ws)
        Write-Row $ws 1 1 @('Table 1. Experience')
        Write-Row $ws 2 1 @('BASE: all')
        Write-Row $ws 3 3 @($cases, $parent)
        $ws.Range('D3:G3').Merge()
        Write-Row $ws 4 4 @('Male', $null, 'Female')
        $ws.Range('D4:E4').Merge()
        $ws.Range('F4:G4').Merge()
        Write-Row $ws 5 3 @('N', 'N', '%', 'N', '%')
        Write-Row $ws 6 1 @('Total', $null, 100, 50, 50, 50, 50)
        $rec = Read-Tables $ws
        Assert-One $rec 'Parent banner'
        Assert-Equal (Field $rec 1 17) 1 'Parent banner only.'
        Assert-Equal (Field $rec 1 18) $parent 'Exact parent name preserved.'
    }

    foreach ($withTitle in @($false, $true)) {
        $ws = New-Fixture ('Novel' + $withTitle)
        [void]$ownedSheets.Add($ws)
        $header = 1
        if ($withTitle) { Write-Row $ws 1 2 @('Novel question'); $header = 2 }
        Write-Row $ws $header 2 @($business, $respondents, $agree, $disagree)
        Write-Row $ws ($header + 1) 2 @('Alpha', 10, 6, 4)
        Write-Row $ws ($header + 2) 2 @('Beta', 20, 12, 8)
        $rec = Read-Tables $ws
        Assert-One $rec 'Unknown header retained'
        Assert-Contains (Field $rec 1 9) $unknownWarning 'Unknown header warning.'
        Assert-Equal (Field $rec 1 15) 0 'Unknown columns are not guessed percentages.'
    }

    foreach ($bad in @('INDEX', 'FalseTitle', 'NoBody', 'NoLabel', 'Prose')) {
        $ws = New-Fixture $bad
        [void]$ownedSheets.Add($ws)
        Write-Row $ws 1 3 @('Table 1. Not a table', 3)
        Write-Row $ws 2 3 @('Table 2. Not a table', 7)
        if ($bad -eq 'INDEX') {
            Write-Row $ws 4 2 @($stub, $cases, '%')
            Write-Row $ws 5 2 @('First', 10, 100)
        } elseif ($bad -eq 'NoBody') {
            Write-Row $ws 4 2 @($stub, $cases, '%')
        } elseif ($bad -eq 'NoLabel') {
            Write-Row $ws 4 2 @($stub, $cases, '%')
            Write-Row $ws 5 3 @(100, 50)
        } elseif ($bad -eq 'Prose') {
            Write-Row $ws 4 2 @($stub, $cases, '%')
            Write-Row $ws 5 2 @("${source}: a note", 2025, 50)
        }
        Assert-Equal (Count-Tables (Read-Tables $ws)) 0 "$bad rejection."
    }

    foreach ($name in @('_Report_InsertCells', ($reportPrefix + 'InsertCells'))) {
        $ws = New-Fixture $name
        [void]$ownedSheets.Add($ws)
        Write-Row $ws 1 1 @('table_key', 'title', 'row', 'col', 'raw_value')
        Write-Row $ws 2 1 @('T0001', 'Experience', 1, 1, 60)
        Write-Row $ws 3 1 @('T0001', 'Experience', 2, 1, 40)
        Assert-Equal (Count-Tables (Read-Tables $ws)) 0 'Generated sheet ignored on repeat run.'
    }

    $settings = New-Fixture 'Overrides'
    [void]$ownedSheets.Add($settings)
    Write-Row $settings 3 1 @('table_range.T0001', "'KisdiSchema'!B4:E8")
    $kisdi = $fixtures.Worksheets.Item('KisdiSchema')
    try {
        $rec = Read-Tables $kisdi 'Overrides'
        Assert-Equal (Field $rec 1 14) 'B4:E8' 'Manual range override preserved.'
    } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($kisdi) }

    if ($WorkbookPath) {
        $WorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
        # External input cannot execute macros; the already-loaded add-in remains callable.
        $excel.AutomationSecurity = 3
        $realInput = $excel.Workbooks.Open($WorkbookPath, 0, $true)
        $sheet1 = $realInput.Worksheets.Item('Sheet1')
        try {
            $rec = Read-Tables $sheet1
            Assert-Equal (Count-Tables $rec) $ExpectedTables 'Real Sheet1 count and key ordering.'
            for ($i = 1; $i -le $ExpectedTables; $i++) {
                Assert-Equal (Field $rec $i 0) ('T{0:0000}' -f $i) 'Real table key.'
            }
        } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet1) }
    }
    Write-Output 'PASS: VBA recognition fixtures, geometry, wide pairs, totals, banners, and overrides.'
} finally {
    # The worker closes only its own workbook references and never saves input.
    foreach ($sheet in $ownedSheets) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) }
    foreach ($book in @($realInput, $fixtures, $addin)) {
        if ($null -ne $book) {
            try { $book.Close($false) } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book) }
        }
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
