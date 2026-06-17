param(
    [string]$OutputDir,
    [string]$ModulePath
)

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent $ScriptDir

if (-not $OutputDir) {
    $OutputDir = Join-Path $ProjectRoot "dev"
}

if (-not $ModulePath) {
    $ModulePath = Join-Path $ProjectRoot "src\ReportAutomationAddin.bas"
}

if (-not (Test-Path -LiteralPath $ModulePath)) {
    throw "VBA module not found: $ModulePath"
}

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

$ModuleText = Get-Content -Raw -Encoding UTF8 -Path $ModulePath
$Version = "unknown"
if ($ModuleText -match 'REPORT_AUTOMATION_VERSION\s+As\s+String\s*=\s*"([^"]+)"') {
    $Version = $Matches[1]
}

$WorkbookPath = Join-Path $OutputDir "ReportAutomationAddin_dev.xlsm"
$AddinPath = Join-Path $OutputDir "ReportAutomationAddin_dev.xlam"
$ImportModulePath = Join-Path $OutputDir "_tmp_ReportAutomationAddin_cp949.bas"
$TempWorkbookPath = Join-Path $OutputDir "_tmp_ReportAutomationAddin_dev.xlsm"
$TempAddinPath = Join-Path $OutputDir "_tmp_ReportAutomationAddin_dev.xlam"

$cp949 = [System.Text.Encoding]::GetEncoding(949)
[System.IO.File]::WriteAllText($ImportModulePath, $ModuleText, $cp949)

function Add-ReportAutomationOptionsForm {
    param($Workbook)

    $form = $Workbook.VBProject.VBComponents.Add(3)
    $form.Name = "ReportAutomationOptionsForm"
    $designer = $form.Designer
    $form.Properties.Item("Caption").Value = "보고서 자동화 옵션"
    $form.Properties.Item("Width").Value = 420
    $form.Properties.Item("Height").Value = 270

    $lblTitle = $designer.Controls.Add("Forms.Label.1", "lblTitle", $true)
    $lblTitle.Caption = "보고서 자동화 실행 옵션"
    $lblTitle.Left = 18
    $lblTitle.Top = 14
    $lblTitle.Width = 360
    $lblTitle.Height = 18

    $lblWorkbook = $designer.Controls.Add("Forms.Label.1", "lblWorkbook", $true)
    $lblWorkbook.Caption = "대상 통합문서"
    $lblWorkbook.Left = 18
    $lblWorkbook.Top = 46
    $lblWorkbook.Width = 90
    $lblWorkbook.Height = 16

    $txtWorkbook = $designer.Controls.Add("Forms.TextBox.1", "txtWorkbook", $true)
    $txtWorkbook.Left = 120
    $txtWorkbook.Top = 43
    $txtWorkbook.Width = 260
    $txtWorkbook.Height = 20
    $txtWorkbook.Locked = $true

    $lblOutput = $designer.Controls.Add("Forms.Label.1", "lblOutput", $true)
    $lblOutput.Caption = "출력 유형"
    $lblOutput.Left = 18
    $lblOutput.Top = 76
    $lblOutput.Width = 90
    $lblOutput.Height = 16

    $txtOutput = $designer.Controls.Add("Forms.TextBox.1", "txtOutput", $true)
    $txtOutput.Left = 120
    $txtOutput.Top = 73
    $txtOutput.Width = 260
    $txtOutput.Height = 20
    $txtOutput.Locked = $true
    $txtOutput.Text = "Excel 산출 시트"

    $lblBanner = $designer.Controls.Add("Forms.Label.1", "lblBanner", $true)
    $lblBanner.Caption = "추출 배너 목록"
    $lblBanner.Left = 18
    $lblBanner.Top = 108
    $lblBanner.Width = 90
    $lblBanner.Height = 16

    $txtBannerSetting = $designer.Controls.Add("Forms.TextBox.1", "txtBannerSetting", $true)
    $txtBannerSetting.Left = 120
    $txtBannerSetting.Top = 105
    $txtBannerSetting.Width = 260
    $txtBannerSetting.Height = 20

    $lblPrefixes = $designer.Controls.Add("Forms.Label.1", "lblPrefixes", $true)
    $lblPrefixes.Caption = "제목 제거 접두어"
    $lblPrefixes.Left = 18
    $lblPrefixes.Top = 140
    $lblPrefixes.Width = 95
    $lblPrefixes.Height = 16

    $txtTitlePrefixes = $designer.Controls.Add("Forms.TextBox.1", "txtTitlePrefixes", $true)
    $txtTitlePrefixes.Left = 120
    $txtTitlePrefixes.Top = 137
    $txtTitlePrefixes.Width = 260
    $txtTitlePrefixes.Height = 20

    $lblHint = $designer.Controls.Add("Forms.Label.1", "lblHint", $true)
    $lblHint.Caption = "쉼표로 구분합니다. 예: 전체,성별,연령대"
    $lblHint.Left = 120
    $lblHint.Top = 162
    $lblHint.Width = 260
    $lblHint.Height = 16
    $lblHint.ForeColor = 8421504

    $cmdRun = $designer.Controls.Add("Forms.CommandButton.1", "cmdRun", $true)
    $cmdRun.Caption = "실행"
    $cmdRun.Left = 210
    $cmdRun.Top = 198
    $cmdRun.Width = 80
    $cmdRun.Height = 28

    $cmdCancel = $designer.Controls.Add("Forms.CommandButton.1", "cmdCancel", $true)
    $cmdCancel.Caption = "취소"
    $cmdCancel.Left = 300
    $cmdCancel.Top = 198
    $cmdCancel.Width = 80
    $cmdCancel.Height = 28

    $code = @'
Option Explicit

Private Sub UserForm_Initialize()
    On Error Resume Next
    If Not ActiveWorkbook Is Nothing Then
        txtWorkbook.Text = ActiveWorkbook.Name
    Else
        txtWorkbook.Text = ""
    End If
    txtOutput.Text = "Excel 산출 시트"
    txtBannerSetting.Text = ReportAutomation_DefaultBannerSetting()
    txtTitlePrefixes.Text = ReportAutomation_DefaultTitlePrefixes()
End Sub

Private Sub cmdRun_Click()
    If Len(Trim$(txtWorkbook.Text)) = 0 Then
        MsgBox "보고서 산출 대상 통합문서를 먼저 활성화하세요.", vbExclamation, "보고서 자동화"
        Exit Sub
    End If
    If ReportAutomation_RunWithOptions(txtBannerSetting.Text, txtTitlePrefixes.Text) Then
        Unload Me
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
'@
    $form.CodeModule.AddFromString($code)
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $null

try {
    $workbook = $excel.Workbooks.Add()
    $sheet = $workbook.Worksheets.Item(1)
    $sheet.Name = "ReportAutomation"
    $sheet.Range("A1").Value = "Report Automation Add-in"
    $sheet.Range("A2").Value = "Version"
    $sheet.Range("B2").Value = $Version
    $sheet.Range("A4").Value = "Build status"
    $sheet.Range("B4").Value = "Before VBA module import"
    $sheet.Range("A6").Value = "Macros"
    $sheet.Range("A7").Value = "ReportAutomation_GenerateExcelOutputs"
    $sheet.Range("A8").Value = "ReportAutomation_About"
    $sheet.Columns("A:B").AutoFit()

    $importedModule = $workbook.VBProject.VBComponents.Import($ImportModulePath)
    $sheet.Range("B4").Value = "VBA module imported"
    $sheet.Range("A5").Value = "VBA module"
    $sheet.Range("B5").Value = $importedModule.Name
    Add-ReportAutomationOptionsForm $workbook

    if (Test-Path -LiteralPath $TempWorkbookPath) { Remove-Item -LiteralPath $TempWorkbookPath -Force }
    if (Test-Path -LiteralPath $TempAddinPath) { Remove-Item -LiteralPath $TempAddinPath -Force }

    $workbook.SaveAs($TempWorkbookPath, 52)
    Copy-Item -LiteralPath $TempWorkbookPath -Destination $WorkbookPath -Force

    $workbook.IsAddin = $true
    $workbook.SaveAs($TempAddinPath, 55)
    Copy-Item -LiteralPath $TempAddinPath -Destination $AddinPath -Force

    $workbook.Close($false)
    $workbook = $null

    Write-Output "Created workbook: $WorkbookPath"
    Write-Output "Created add-in: $AddinPath"
}
finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($false) | Out-Null } catch {}
    }
    foreach ($path in @($ImportModulePath, $TempWorkbookPath, $TempAddinPath)) {
        if (Test-Path -LiteralPath $path) {
            try { Remove-Item -LiteralPath $path -Force } catch {}
        }
    }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
