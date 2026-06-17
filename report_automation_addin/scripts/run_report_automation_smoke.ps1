param(
    [string]$WorkbookPath,
    [string]$AddinPath,
    [int]$ExpectedCreatedSheetCount = -1
)

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent $ScriptDir

if (-not $WorkbookPath) {
    $WorkbookPath = Join-Path $env:TEMP "kisdi_report_automation_test.xlsx"
}

if (-not $AddinPath) {
    $AddinPath = Join-Path $ProjectRoot "dev\ReportAutomationAddin_dev.xlam"
}

if (-not (Test-Path -LiteralPath $WorkbookPath)) {
    throw "Workbook not found: $WorkbookPath"
}

if (-not (Test-Path -LiteralPath $AddinPath)) {
    throw "Add-in not found: $AddinPath"
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1

$addinWorkbook = $null
$dataWorkbook = $null
$verifyWorkbook = $null

try {
    $addinWorkbook = $excel.Workbooks.Open($AddinPath)
    $dataWorkbook = $excel.Workbooks.Open($WorkbookPath)
    $dataWorkbook.Activate() | Out-Null

    Write-Output ("OpenedWorkbook: " + $dataWorkbook.FullName)
    Write-Output ("ReadOnly: " + $dataWorkbook.ReadOnly)
    Write-Output ("FileFormat: " + $dataWorkbook.FileFormat)
    Write-Output ("SheetCountBefore: " + $dataWorkbook.Worksheets.Count)
    if ($dataWorkbook.ReadOnly) {
        throw "Workbook opened as read-only; generated sheets cannot be saved reliably: $WorkbookPath"
    }

    $sheetNamesBefore = @{}
    foreach ($sheet in $dataWorkbook.Worksheets) {
        $sheetNamesBefore[$sheet.Name] = $true
    }
    if ($ExpectedCreatedSheetCount -lt 0) {
        $ExpectedCreatedSheetCount = 9
        if (-not $sheetNamesBefore.ContainsKey("_ReportAutomation_Log")) {
            $ExpectedCreatedSheetCount += 1
        }
    }

    $macroName = "'" + $addinWorkbook.Name + "'!ReportAutomation_GenerateExcelOutputsSilent"
    $excel.Run($macroName) | Out-Null

    Write-Output ("SheetCountAfterRun: " + $dataWorkbook.Worksheets.Count)
    foreach ($sheet in $dataWorkbook.Worksheets) {
        Write-Output (" AfterRunSheet: " + $sheet.Name)
    }

    $dataWorkbook.Save()
    Write-Output ("SavedAfterRun: " + $dataWorkbook.Saved)

    $createdSheets = @()
    foreach ($sheet in $dataWorkbook.Worksheets) {
        if (-not $sheetNamesBefore.ContainsKey($sheet.Name)) {
            $createdSheets += $sheet.Name
        }
    }

    Write-Output ("GeneratedReportSheetCount: " + $createdSheets.Count)
    foreach ($name in $createdSheets) {
        Write-Output (" - " + $name)
    }
    if ($createdSheets.Count -ne $ExpectedCreatedSheetCount) {
        throw "Expected $ExpectedCreatedSheetCount generated sheets, but got $($createdSheets.Count)."
    }

    $dataWorkbook.Close($true)
    $dataWorkbook = $null

    $verifyWorkbook = $excel.Workbooks.Open($WorkbookPath)
    Write-Output ("SheetCountAfterReopen: " + $verifyWorkbook.Worksheets.Count)
    $reopenedSheets = @{}
    foreach ($sheet in $verifyWorkbook.Worksheets) {
        Write-Output (" ReopenedSheet: " + $sheet.Name)
        $reopenedSheets[$sheet.Name] = $true
    }
    $expectedTotalSheetCount = $sheetNamesBefore.Count + $createdSheets.Count
    if ($verifyWorkbook.Worksheets.Count -ne $expectedTotalSheetCount) {
        throw "Expected $expectedTotalSheetCount sheets after reopen, but got $($verifyWorkbook.Worksheets.Count)."
    }
    foreach ($name in $createdSheets) {
        if (-not $reopenedSheets.ContainsKey($name)) {
            throw "Generated sheet was not persisted after reopen: $name"
        }
    }
}
finally {
    if ($null -ne $verifyWorkbook) {
        try { $verifyWorkbook.Close($false) | Out-Null } catch {}
    }
    if ($null -ne $dataWorkbook) {
        try { $dataWorkbook.Close($false) | Out-Null } catch {}
    }
    if ($null -ne $addinWorkbook) {
        try { $addinWorkbook.Close($false) | Out-Null } catch {}
    }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
