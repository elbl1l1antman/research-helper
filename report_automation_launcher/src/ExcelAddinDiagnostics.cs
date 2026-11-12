using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Web.Script.Serialization;

namespace ReportAutomationLauncher
{
    internal sealed class ExcelAddinOpenException : InvalidOperationException
    {
        public readonly Dictionary<string, object> Report;

        internal ExcelAddinOpenException(Dictionary<string, object> report, Exception cause)
            : base((string)report["message"], cause) { Report = report; }
    }

    // File.Exists only proves launcher access, not access from the out-of-process Excel server.
    // Keep this shared by diagnostics and actual runs so they report the same failure honestly.
    internal static class ExcelAddinDiagnostics
    {
        internal static void RequireFile(string path)
        {
            if (!File.Exists(path)) throw new ExcelAddinOpenException(NewReport(path), null);
        }

        internal static object Open(object excel, string path)
        {
            var report = NewReport(path);
            if (!(bool)report["file_exists"])
                throw new ExcelAddinOpenException(report, null);
            try
            {
                dynamic xl = excel;
                return xl.Workbooks.Open((string)report["addin_path"], 0, true);
            }
            catch (COMException ex)
            {
                report["stage"] = "addin_open";
                report["hresult"] = "0x" + ex.ErrorCode.ToString("X8");
                report["error"] = ex.Message;
                report["message"] = "Excel에서 보고서 자동화 추가기능을 열지 못했습니다 (" + report["hresult"] + ")." + Environment.NewLine +
                    report["addin_path"] + Environment.NewLine +
                    "런처를 다시 실행해 보세요. 반복되면 산출 결과의 진행 기록과 Excel 연결 오류 기록을 관리자에게 전달하세요. " +
                    "파일 접근 및 차단 상태를 확인할 수 있으며, 런처는 보안 설정을 변경하지 않습니다.";
                throw new ExcelAddinOpenException(report, ex);
            }
        }

        internal static Dictionary<string, object> Check(string path)
        {
            var report = NewReport(path);
            if (!(bool)report["file_exists"]) return report;
            object excel = null;
            object book = null;
            try
            {
                report["stage"] = "excel_create";
                Type type = Type.GetTypeFromProgID("Excel.Application");
                if (type == null) throw new InvalidOperationException("Microsoft Excel 설치를 확인하세요.");
                excel = Activator.CreateInstance(type);
                dynamic xl = excel;
                xl.Visible = false;
                xl.DisplayAlerts = false;
                xl.EnableEvents = false;
                // Diagnostics only need file access, never user add-in code execution.
                xl.AutomationSecurity = 3; // msoAutomationSecurityForceDisable; isolated instance only.
                report["excel_version"] = Convert.ToString(xl.Version);
                // Never lower macro security, register trusted folders, or change the registry.
                book = Open(excel, path);
                report["status"] = "opened";
                report["stage"] = "addin_open";
                report["message"] = "Excel에서 추가기능을 읽기 전용으로 열었습니다. 매크로 실행 가능 여부는 이 검사에서 확인하지 않습니다.";
            }
            catch (ExcelAddinOpenException ex)
            {
                ex.Report["excel_version"] = report["excel_version"];
                report = ex.Report;
            }
            catch (Exception ex)
            {
                report["error"] = ex.Message;
                report["hresult"] = "0x" + ex.HResult.ToString("X8");
                report["message"] = "Excel 연결 검사에 실패했습니다: " + ex.Message;
            }
            finally
            {
                // Only the Excel instance and read-only add-in created by this check are closed.
                AutomationRunner.CloseComWorkbookForLauncher(book, false);
                AutomationRunner.QuitComExcelForLauncher(excel);
                AutomationRunner.ReleaseComForLauncher(book);
                AutomationRunner.ReleaseComForLauncher(excel);
            }
            return report;
        }

        internal static void Save(string output, Dictionary<string, object> report)
        {
            File.WriteAllText(output, new JavaScriptSerializer().Serialize(report), new System.Text.UTF8Encoding(false));
        }

        private static Dictionary<string, object> NewReport(string path)
        {
            string fullPath = string.IsNullOrWhiteSpace(path) ? "" : Path.GetFullPath(path);
            return new Dictionary<string, object> {
                { "status", "blocked" }, { "stage", "file_check" },
                { "addin_path", fullPath }, { "file_exists", File.Exists(fullPath) },
                { "excel_version", "" }, { "hresult", "" }, { "error", "" },
                { "macro_execution_verified", false },
                { "message", "보고서 자동화 추가기능을 찾을 수 없습니다: " + fullPath + Environment.NewLine +
                    "내장 추가기능은 런처를 닫고 다시 실행하세요. 별도 경로를 지정한 경우 해당 경로를 확인하세요." }
            };
        }
    }
}
