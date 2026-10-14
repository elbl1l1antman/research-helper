using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Security.Cryptography;
using System.Threading;
using System.Web.Script.Serialization;

namespace ReportAutomationLauncher
{
    // The distributed EXE owns its runtime. Development builds without the resource keep old resolution.
    internal static class StandaloneBundle
    {
        internal static string Root { get; private set; }
        private const string ResourceName = "ReportAutomation.Payload.zip";

        internal static void Initialize()
        {
            using (Stream payload = Assembly.GetExecutingAssembly().GetManifestResourceStream(ResourceName))
            {
                if (payload == null) return;
                string digest;
                using (var sha = SHA256.Create()) digest = BitConverter.ToString(sha.ComputeHash(payload)).Replace("-", "").ToLowerInvariant();
                string parent = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ResearchHelper", "bundles");
                string target = Path.Combine(parent, digest);
                using (var mutex = new Mutex(false, "Local\\ResearchHelperBundle_" + digest))
                {
                    bool acquired = false;
                    try
                    {
                        try { acquired = mutex.WaitOne(120000); }
                        catch (AbandonedMutexException) { acquired = true; }
                        if (!acquired) throw new TimeoutException("실행 구성요소 준비 시간이 초과되었습니다.");
                        if (!IsComplete(target))
                        {
                            Directory.CreateDirectory(parent);
                            string staging = Path.Combine(parent, "extract-" + Guid.NewGuid().ToString("N"));
                            try
                            {
                                Directory.CreateDirectory(staging);
                                payload.Position = 0;
                                using (var archive = new ZipArchive(payload, ZipArchiveMode.Read, true))
                                {
                                    foreach (ZipArchiveEntry entry in archive.Entries)
                                    {
                                        string destination = ContainedPath(staging, entry.FullName);
                                        if (entry.FullName.EndsWith("/")) Directory.CreateDirectory(destination);
                                        else
                                        {
                                            Directory.CreateDirectory(Path.GetDirectoryName(destination));
                                            entry.ExtractToFile(destination, false);
                                        }
                                    }
                                }
                                ValidateFiles(staging);
                                File.WriteAllText(Path.Combine(staging, ".complete"), digest);
                                if (Directory.Exists(target)) DeleteOwnedDirectory(parent, target);
                                Directory.Move(staging, target);
                            }
                            finally { if (Directory.Exists(staging)) DeleteOwnedDirectory(parent, staging); }
                        }
                        Root = target;
                        // ponytail: retain older content caches; add in-use-aware cleanup if disk usage matters.
                    }
                    finally { if (acquired) mutex.ReleaseMutex(); }
                }
            }
        }

        private static string ContainedPath(string parent, string relative)
        {
            string root = Path.GetFullPath(parent).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            string full = Path.GetFullPath(Path.Combine(root, relative.Replace('/', Path.DirectorySeparatorChar)));
            if (Path.IsPathRooted(relative) || relative.IndexOf(':') >= 0 || !full.StartsWith(root, StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("실행 구성요소에 안전하지 않은 경로가 있습니다.");
            return full;
        }

        private static void DeleteOwnedDirectory(string parent, string target)
        {
            // Never recursively remove an unverified path, even on failure or an interrupted first run.
            string owned = ContainedPath(parent, Path.GetFileName(target));
            if (!string.Equals(owned, Path.GetFullPath(target), StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("캐시 정리 경로가 올바르지 않습니다.");
            Directory.Delete(owned, true);
        }

        private static bool IsComplete(string directory)
        {
            if (!File.Exists(Path.Combine(directory, ".complete"))) return false;
            try { ValidateFiles(directory); return true; }
            catch (IOException) { return false; }
            catch (ArgumentException) { return false; }
            catch (InvalidOperationException) { return false; }
        }

        private static void ValidateFiles(string directory)
        {
            string manifest = Path.Combine(directory, "bundle_manifest.json");
            var data = new JavaScriptSerializer().DeserializeObject(File.ReadAllText(manifest)) as Dictionary<string, object>;
            if (data == null || !data.ContainsKey("files") || !(data["files"] is object[])) throw new InvalidDataException("실행 구성요소 목록이 올바르지 않습니다.");
            foreach (object item in (object[])data["files"])
                if (!(item is string) || !File.Exists(ContainedPath(directory, (string)item))) throw new FileNotFoundException("내장 실행 파일이 누락되었습니다.");
            foreach (string file in new[] { "runtime/python.exe", "report_automation_engine/excel_report_generator.py", "report_automation_addin/dev/ReportAutomationAddin_dev.xlam" })
                if (!File.Exists(ContainedPath(directory, file))) throw new FileNotFoundException("필수 실행 구성요소가 누락되었습니다: " + file);
        }

        internal static string Find(string relative)
        {
            return Root == null ? null : ContainedPath(Root, relative);
        }

        internal static void ConfigurePythonProcess(ProcessStartInfo info)
        {
            if (Root == null) return;
            // Keep the engine script directory for existing local imports, but ignore external Python settings.
            info.Arguments = "-E -s -X utf8 " + info.Arguments;
            info.EnvironmentVariables.Remove("PYTHONHOME");
            info.EnvironmentVariables.Remove("PYTHONPATH");
            info.EnvironmentVariables.Remove("PYTHONUSERBASE");
        }

        internal static void SelfCheck(string output)
        {
            if (Root == null) throw new InvalidOperationException("이 실행 파일에는 독립 실행 구성요소가 없습니다.");
            var start = new ProcessStartInfo(PathResolver.ResolvePythonPath());
            start.Arguments = "-I -X utf8 -c \"import json,sys,openpyxl,pptx,PIL,win32com.client,pythoncom,pywintypes; print(json.dumps({'executable':sys.executable,'prefix':sys.prefix,'imports':'ready'}))\"";
            start.UseShellExecute = false;
            start.CreateNoWindow = true;
            start.RedirectStandardOutput = true;
            start.RedirectStandardError = true;
            using (Process process = Process.Start(start))
            {
                var stdout = process.StandardOutput.ReadToEndAsync();
                var stderr = process.StandardError.ReadToEndAsync();
                if (!process.WaitForExit(60000)) { process.Kill(); throw new TimeoutException("내장 Python 검증 시간이 초과되었습니다."); }
                if (process.ExitCode != 0) throw new InvalidOperationException(stderr.Result);
                var runtime = new JavaScriptSerializer().DeserializeObject(stdout.Result) as Dictionary<string, object>;
                if (runtime == null || !runtime.ContainsKey("prefix") || !runtime.ContainsKey("executable") ||
                    !string.Equals(Path.GetFullPath(Convert.ToString(runtime["prefix"])), Path.Combine(Root, "runtime"), StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(Path.GetFullPath(Convert.ToString(runtime["executable"])), PathResolver.ResolvePythonPath(), StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("내장 Python 이외의 실행 환경이 사용되었습니다.");
                // Construct the actual launcher without showing it or opening user documents.
                using (var form = new MainForm())
                    if (form.Handle == IntPtr.Zero) throw new InvalidOperationException("런처 GUI를 초기화하지 못했습니다.");
                var report = new Dictionary<string, object>
                {
                    { "status", "ready" }, { "bundled", true }, { "gui", "ready" }, { "bundle_root", Root },
                    { "python", PathResolver.ResolvePythonPath() }, { "engine", PathResolver.ResolveExcelEnginePath() },
                    { "addin", PathResolver.ResolveDefaultAddinPath() }, { "runtime", runtime }
                };
                string json = new JavaScriptSerializer().Serialize(report);
                if (string.IsNullOrWhiteSpace(output)) throw new ArgumentException("--out <path> 값을 지정하세요.");
                File.WriteAllText(output, json, System.Text.Encoding.UTF8);
                Console.WriteLine(json);
            }
        }
    }
}
