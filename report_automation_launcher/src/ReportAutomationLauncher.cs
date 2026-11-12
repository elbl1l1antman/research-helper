using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace ReportAutomationLauncher
{
    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            try { StandaloneBundle.Initialize(); }
            catch (Exception ex)
            {
                if (args.Length == 0) MessageBox.Show("실행 구성요소 준비 실패: " + ex.Message, "ReportAutomation", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else Console.Error.WriteLine(ex.Message);
                return 1;
            }
            if (args.Length > 0)
            {
                return RunCommandLine(args);
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
            return 0;
        }

        private static int RunCommandLine(string[] args)
        {
            try
            {
                if (HasFlag(args, "check-excel-addin"))
                {
                    string output = GetArgValue(args, "out");
                    if (string.IsNullOrWhiteSpace(output)) throw new ArgumentException("--out <path> 값을 지정하세요.");
                    string path = GetArgValue(args, "addin");
                    var report = ExcelAddinDiagnostics.Check(string.IsNullOrWhiteSpace(path) ? PathResolver.ResolveDefaultAddinPath() : path);
                    ExcelAddinDiagnostics.Save(output, report);
                    Console.WriteLine(report["message"]);
                    return (string)report["status"] == "opened" ? 0 : 1;
                }
                if (HasFlag(args, "self-check-standalone"))
                {
                    StandaloneBundle.SelfCheck(GetArgValue(args, "out"));
                    return 0;
                }
                if (HasFlag(args, "self-check-excel-recognition"))
                {
                    BannerInspector.RunRecognitionSelfCheck();
                    return 0;
                }
                if (HasFlag(args, "self-check-hwp-page-setup"))
                {
                    EngineRunner.RunHwpPageSetupSelfCheck();
                    Console.WriteLine("HWP page setup self-check passed.");
                    return 0;
                }
                if (HasFlag(args, "self-check-hwp-style-presets"))
                {
                    HwpStylePresetStore.RunSelfCheck();
                    Console.WriteLine("HWP style preset self-check passed.");
                    return 0;
                }
                if (HasFlag(args, "self-check-hwp-style-cli"))
                {
                    EngineRunner.RunHwpStyleCliSelfCheck();
                    Console.WriteLine("HWP style CLI self-check passed.");
                    return 0;
                }

                if (HasFlag(args, "list-banners"))
                {
                    string workbookPath = GetArgValue(args, "workbook");
                    string outputPath = GetArgValue(args, "out");
                    if (string.IsNullOrWhiteSpace(workbookPath))
                    {
                        throw new InvalidOperationException("--workbook <path> 값을 지정하세요.");
                    }
                    if (string.IsNullOrWhiteSpace(outputPath))
                    {
                        throw new InvalidOperationException("--out <path> 값을 지정하세요.");
                    }

                    List<string> banners = BannerInspector.ReadBanners(workbookPath);
                    File.WriteAllLines(outputPath, banners.ToArray(), System.Text.Encoding.UTF8);
                    return 0;
                }

                LauncherOptions options = LauncherOptions.FromArgs(args);
                EngineRunner.RunWorkflow(options, Console.WriteLine);
                return 0;
            }
            catch (Exception ex)
            {
                string outputPath = GetArgValue(args, "out");
                if (!string.IsNullOrWhiteSpace(outputPath))
                {
                    try
                    {
                        File.WriteAllText(outputPath + ".error.txt", ex.ToString(), System.Text.Encoding.UTF8);
                        var addinError = ex as ExcelAddinOpenException;
                        if (addinError != null) ExcelAddinDiagnostics.Save(outputPath + ".excel_addin.json", addinError.Report);
                    }
                    catch
                    {
                    }
                }
                Console.Error.WriteLine(ex.Message);
                return 1;
            }
        }

        private static bool HasFlag(string[] args, string name)
        {
            string flag = "--" + name;
            foreach (string arg in args)
            {
                if (string.Equals(arg, flag, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }

        private static string GetArgValue(string[] args, string name)
        {
            string flag = "--" + name;
            for (int i = 0; i < args.Length - 1; i++)
            {
                if (string.Equals(args[i], flag, StringComparison.OrdinalIgnoreCase))
                {
                    return args[i + 1];
                }
            }
            return "";
        }
    }

    internal sealed class HwpStylePresetStore
    {
        internal const string BuiltInName = "기본 보고서";
        internal const string GoldTableName = "기본 표 · 골드 배너";
        internal const string BlueTableName = "기본 표 · 청회색 머리행";
        internal static readonly string[] StyleNames = { "보고서 본문1", "보고서 본문2", "표보기", "표배너", "표숫자" };
        internal static readonly string[] TableStyleNames = { "표보기", "표배너", "표숫자" };
        internal static readonly Dictionary<string, string> RoleStyleDefaults = new Dictionary<string, string>
        {
            { "표머리행", "표배너" }, { "표가로배너", "표배너" }, { "표세로배너", "표보기" }, { "표BASE", "표배너" }
        };
        internal static readonly string[] MarginSides = { "top", "bottom", "left", "right" };
        internal static readonly Dictionary<string, string> LineGroups = new Dictionary<string, string>
        {
            { "outer", "외곽선" }, { "inner_horizontal", "내부 가로선" }, { "inner_vertical", "내부 세로선" }, { "header_bottom", "머리행 아래선" }
        };
        internal static readonly Dictionary<string, string> LineTypes = new Dictionary<string, string>
        {
            { "none", "없음" }, { "solid", "실선" }, { "dot", "점선" }, { "dash", "파선" }, { "double_slim", "가는 이중선" }
        };
        private readonly string path;
        private readonly JavaScriptSerializer serializer = new JavaScriptSerializer();
        private readonly Action<string> recoveryLog;

        internal HwpStylePresetStore(string storagePath = null, Action<string> recoveryLog = null)
        {
            path = storagePath ?? Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "ResearchHelper",
                "hwp_style_presets.json");
            serializer.MaxJsonLength = int.MaxValue;
            this.recoveryLog = recoveryLog;
        }

        internal string StoragePath { get { return path; } }
        internal string RecoveryWarning { get; private set; }
        internal string RecoveryBackupPath { get; private set; }

        internal List<Dictionary<string, object>> LoadAll()
        {
            var users = LoadUsers();
            var result = new List<Dictionary<string, object>> { DefaultConfig() };
            // Preserve existing user presets even if they happen to use a new built-in name.
            if (!users.Any(item => SameName(item, GoldTableName))) result.Add(DefaultTableConfig(true));
            if (!users.Any(item => SameName(item, BlueTableName))) result.Add(DefaultTableConfig(false));
            result.AddRange(users);
            return result;
        }

        private static bool SameName(Dictionary<string, object> preset, string name)
        {
            return string.Equals(Convert.ToString(preset["preset_name"]), name, StringComparison.OrdinalIgnoreCase);
        }

        internal bool IsBuiltIn(string name)
        {
            return string.Equals(name, BuiltInName, StringComparison.OrdinalIgnoreCase) ||
                ((string.Equals(name, GoldTableName, StringComparison.OrdinalIgnoreCase) ||
                  string.Equals(name, BlueTableName, StringComparison.OrdinalIgnoreCase)) &&
                 !LoadUsers().Any(item => SameName(item, name)));
        }

        internal Dictionary<string, object> Get(string name)
        {
            foreach (Dictionary<string, object> preset in LoadAll())
            {
                if (string.Equals(Convert.ToString(preset["preset_name"]), name, StringComparison.OrdinalIgnoreCase))
                {
                    return Clone(preset);
                }
            }
            return DefaultConfig();
        }

        internal void Save(Dictionary<string, object> config, bool overwrite)
        {
            Dictionary<string, object> normalized = Normalize(config);
            string name = Convert.ToString(normalized["preset_name"]).Trim();
            if (IsBuiltIn(name))
            {
                throw new InvalidOperationException("기본 제공 서식은 수정할 수 없습니다. 다른 이름으로 저장하세요.");
            }

            List<Dictionary<string, object>> users = LoadUsers();
            int existing = users.FindIndex(item => string.Equals(Convert.ToString(item["preset_name"]), name, StringComparison.OrdinalIgnoreCase));
            if (existing >= 0 && !overwrite)
            {
                throw new InvalidOperationException("같은 이름의 서식이 있습니다. 덮어쓰기를 명시적으로 선택하세요.");
            }
            if (existing < 0 && users.Count >= 3)
            {
                throw new InvalidOperationException("사용자 서식은 최대 3개까지 저장할 수 있습니다.");
            }
            if (existing >= 0)
            {
                users[existing] = normalized;
            }
            else
            {
                users.Add(normalized);
            }
            WriteUsers(users);
        }

        internal void Delete(string name)
        {
            if (IsBuiltIn(name))
            {
                throw new InvalidOperationException("기본 제공 서식은 삭제할 수 없습니다.");
            }
            List<Dictionary<string, object>> users = LoadUsers();
            users.RemoveAll(item => string.Equals(Convert.ToString(item["preset_name"]), name, StringComparison.OrdinalIgnoreCase));
            WriteUsers(users);
        }

        internal static Dictionary<string, object> Normalize(Dictionary<string, object> config)
        {
            if (config == null)
            {
                throw new InvalidOperationException("서식 설정이 비어 있습니다.");
            }
            var copy = Clone(config);
            if (Convert.ToString(Value(copy, "schema_version")) != "1.0")
            {
                throw new InvalidOperationException("schema_version은 1.0이어야 합니다.");
            }
            object presetNameValue = Value(copy, "preset_name");
            if (!(presetNameValue is string)) throw new InvalidOperationException("preset_name 값은 문자열이어야 합니다.");
            string presetName = ((string)presetNameValue).Trim();
            if (presetName.Length == 0)
            {
                throw new InvalidOperationException("서식 이름을 입력하세요.");
            }
            copy["preset_name"] = presetName;

            Dictionary<string, object> paragraph = Map(Value(copy, "paragraph_styles"), "paragraph_styles");
            string[] optional = paragraph.Keys.Where(RoleStyleDefaults.ContainsKey).ToArray();
            RequireExactKeys(paragraph, StyleNames.Concat(optional).ToArray(), "paragraph_styles");
            foreach (string name in paragraph.Keys)
            {
                Dictionary<string, object> style = Map(paragraph[name], name);
                RequireText(style, "font_family", name);
                style["font_size_pt"] = Number(style, "font_size_pt", 6, 30, name);
                if (!(Value(style, "bold") is bool)) throw new InvalidOperationException(name + ".bold 값이 올바르지 않습니다.");
                if (!(Value(style, "bullet") is string)) throw new InvalidOperationException(name + ".bullet 값이 올바르지 않습니다.");
                style["line_spacing_percent"] = Number(style, "line_spacing_percent", 80, 300, name);
                style["left_indent_mm"] = Number(style, "left_indent_mm", -30, 100, name);
                style["first_line_indent_mm"] = Number(style, "first_line_indent_mm", -30, 100, name);
                string alignment = Convert.ToString(Value(style, "alignment"));
                if (!(new[] { "left", "center", "right", "justify" }).Contains(alignment))
                    throw new InvalidOperationException(name + ".alignment 값이 올바르지 않습니다.");
            }

            Dictionary<string, object> cells = Map(Value(copy, "table_cell_styles"), "table_cell_styles");
            RequireExactKeys(cells, TableStyleNames.Concat(optional).ToArray(), "table_cell_styles");
            foreach (string name in cells.Keys)
            {
                Dictionary<string, object> style = Map(cells[name], name);
                style["fill_color"] = ColorValue(style, "fill_color", name);
                string vertical = Convert.ToString(Value(style, "vertical_alignment"));
                if (!(new[] { "top", "center", "bottom" }).Contains(vertical))
                    throw new InvalidOperationException(name + ".vertical_alignment 값이 올바르지 않습니다.");
                if (style.ContainsKey("margins_mm"))
                {
                    string field = name + ".margins_mm";
                    var margins = Map(style["margins_mm"], field);
                    RequireExactKeys(margins, MarginSides, field);
                    foreach (string side in MarginSides) margins[side] = Number(margins, side, 0, 20, field);
                }
            }

            Dictionary<string, object> border = Map(Value(copy, "table_border"), "table_border");
            border["inner_width_mm"] = Number(border, "inner_width_mm", 0.1, 5.0, "table_border");
            border["outer_width_mm"] = Number(border, "outer_width_mm", 0.1, 5.0, "table_border");
            border["color"] = ColorValue(border, "color", "table_border");
            if (!border.ContainsKey("show_side_borders")) border["show_side_borders"] = false;
            if (!(border["show_side_borders"] is bool))
                throw new InvalidOperationException("table_border.show_side_borders 값은 부울이어야 합니다.");
            if (border.ContainsKey("edges"))
                foreach (var item in Map(border["edges"], "table_border.edges"))
                {
                    string field = "table_border.edges." + item.Key;
                    if (!LineGroups.ContainsKey(item.Key)) throw new InvalidOperationException(field + " 위치가 올바르지 않습니다.");
                    var line = Map(item.Value, field);
                    if (!LineTypes.ContainsKey(Convert.ToString(Value(line, "type")))) throw new InvalidOperationException(field + " 종류가 올바르지 않습니다.");
                    line["width_mm"] = Number(line, "width_mm", 0.1, 5, field);
                    line["color"] = ColorValue(line, "color", field);
                }
            return copy;
        }

        internal static Dictionary<string, object> DefaultConfig()
        {
            var styles = new Dictionary<string, object>();
            styles["보고서 본문1"] = Style("맑은 고딕", 10, false, "", 160, 0, 0, "left");
            styles["보고서 본문2"] = Style("맑은 고딕", 10, false, "-", 160, 5, -5, "left");
            styles["표보기"] = Style("맑은 고딕", 8.5, false, "", 130, 0, 0, "left");
            styles["표배너"] = Style("맑은 고딕", 8.5, true, "", 130, 0, 0, "center");
            styles["표숫자"] = Style("맑은 고딕", 8.5, false, "", 130, 0, 0, "center");
            return new Dictionary<string, object>
            {
                { "schema_version", "1.0" },
                { "preset_name", BuiltInName },
                { "paragraph_styles", styles },
                { "table_cell_styles", new Dictionary<string, object>
                    {
                        { "표보기", Cell("FFFFFF") }, { "표배너", Cell("E7E7E7") }, { "표숫자", Cell("FFFFFF") }
                    }
                },
                { "table_border", new Dictionary<string, object>
                    {
                        { "inner_width_mm", 0.12 }, { "outer_width_mm", 0.4 }, { "color", "000000" }, { "show_side_borders", false }
                    }
                }
            };
        }

        internal static void WriteConfig(string outputPath, Dictionary<string, object> config)
        {
            var serializer = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            AtomicWrite(outputPath, serializer.Serialize(Normalize(config)));
        }

        private static Dictionary<string, object> DefaultTableConfig(bool gold)
        {
            var config = DefaultConfig();
            config["preset_name"] = gold ? GoldTableName : BlueTableName;
            var styles = Map(config["paragraph_styles"], "paragraph_styles");
            var cells = Map(config["table_cell_styles"], "table_cell_styles");
            string[] roles = { "표보기", "표배너", "표숫자", "표머리행", "표가로배너", "표세로배너", "표BASE" };
            string[] fills = gold
                ? new[] { "FAEDD2", "F1CB7E", "FFFFFF", "F1CB7E", "F5DCA8", "F5DCA8", "F5DCA8" }
                : new[] { "FFFFFF", "F0F4F7", "FFFFFF", "F0F4F7", "F0F4F7", "F7F9FB", "F7F9FB" };
            for (int i = 0; i < roles.Length; i++)
            {
                bool heading = i == 1 || i == 3 || i == 4;
                styles[roles[i]] = Style("맑은 고딕", gold ? 9 : (heading ? 10 : 9.5),
                    heading || i == 6, "", 125, 0, 0, i == 0 || i == 5 ? "left" : "center");
                var cell = Cell(fills[i]);
                cell["margins_mm"] = new Dictionary<string, object>
                {
                    { "top", gold ? 0.5 : 0.3 }, { "bottom", gold ? 0.5 : 0.3 },
                    { "left", gold ? 0.4 : 0.8 }, { "right", gold ? 0.4 : 0.8 }
                };
                cells[roles[i]] = cell;
            }
            var border = Map(config["table_border"], "table_border");
            border["inner_width_mm"] = gold ? 0.12 : 0.1;
            border["outer_width_mm"] = gold ? 0.4 : 0.3;
            border["color"] = gold ? "EEB449" : "666666";
            border["edges"] = new Dictionary<string, object>
            {
                { "outer", new Dictionary<string, object> { { "type", "solid" }, { "width_mm", border["outer_width_mm"] }, { "color", gold ? "ECBB53" : "000000" } } },
                { "header_bottom", new Dictionary<string, object> { { "type", "solid" }, { "width_mm", gold ? 0.3 : 0.1 }, { "color", gold ? "EEB449" : "000000" } } },
                { "inner_vertical", new Dictionary<string, object> { { "type", gold ? "solid" : "dash" }, { "width_mm", border["inner_width_mm"] }, { "color", border["color"] } } }
            };
            return Normalize(config);
        }

        internal static void RunSelfCheck()
        {
            string directory = Path.Combine(Path.GetTempPath(), "ResearchHelperPresetCheck_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try
            {
                string file = Path.Combine(directory, "presets.json");
                var warnings = new List<string>();
                var store = new HwpStylePresetStore(file, warnings.Add);
                Assert(store.LoadAll().Count == 3, "default and two table presets");
                foreach (string name in new[] { GoldTableName, BlueTableName })
                {
                    var builtIn = store.Get(name);
                    var styles = Map(builtIn["table_cell_styles"], "table_cell_styles");
                    Assert((string)Map(styles["표머리행"], "표머리행")["fill_color"] ==
                        (name == GoldTableName ? "F1CB7E" : "F0F4F7"), "table preset header");
                    AssertThrows(delegate { store.Save(builtIn, true); }, "built-in overwrite rejection");
                    AssertThrows(delegate { store.Delete(name); }, "built-in deletion rejection");
                }
                Assert(warnings.Count == 0 && string.IsNullOrWhiteSpace(store.RecoveryWarning), "no warning for missing store");
                object showSideBorders;
                Assert(Map(store.Get(BuiltInName)["table_border"], "table_border").TryGetValue("show_side_borders", out showSideBorders)
                    && showSideBorders is bool && !(bool)showSideBorders, "side borders default off");
                for (int i = 1; i <= 3; i++)
                {
                    Dictionary<string, object> preset = DefaultConfig();
                    preset["preset_name"] = "사용자 " + i;
                    store.Save(preset, false);
                }
                Assert(store.LoadAll().Count == 6, "three user presets in addition to defaults");
                var checkSerializer = new JavaScriptSerializer();
                Dictionary<string, object> legacy = checkSerializer.DeserializeObject(File.ReadAllText(file, System.Text.Encoding.UTF8)) as Dictionary<string, object>;
                foreach (object item in (object[])legacy["presets"])
                    Map(Map(item, "preset")["table_border"], "table_border").Remove("show_side_borders");
                string legacyJson = checkSerializer.Serialize(legacy);
                File.WriteAllText(file, legacyJson, System.Text.Encoding.UTF8);
                Assert(store.LoadAll().Count == 6, "three legacy presets retained");
                for (int i = 1; i <= 3; i++)
                    Assert(!(bool)Map(store.Get("사용자 " + i)["table_border"], "table_border")["show_side_borders"], "legacy side borders default off");
                Assert(warnings.Count == 0 && string.IsNullOrWhiteSpace(store.RecoveryWarning)
                    && File.ReadAllText(file, System.Text.Encoding.UTF8) == legacyJson, "legacy loading without recovery or rewrite");
                Dictionary<string, object> fourth = DefaultConfig();
                fourth["preset_name"] = "사용자 4";
                AssertThrows(delegate { store.Save(fourth, false); }, "fourth-preset rejection");
                Dictionary<string, object> duplicate = DefaultConfig();
                duplicate["preset_name"] = "사용자 1";
                Map(duplicate["table_border"], "table_border")["show_side_borders"] = true;
                Map(Map(duplicate["paragraph_styles"], "paragraph_styles")["보고서 본문1"], "보고서 본문1")["font_size_pt"] = 11.0;
                AssertThrows(delegate { store.Save(duplicate, false); }, "duplicate rejection");
                store.Save(duplicate, true);
                Assert(Convert.ToDouble(Map(Map(store.Get("사용자 1")["paragraph_styles"], "paragraph_styles")["보고서 본문1"], "보고서 본문1")["font_size_pt"]) == 11.0, "duplicate overwrite");
                Assert((bool)Map(store.Get("사용자 1")["table_border"], "table_border")["show_side_borders"], "side borders true roundtrip");
                Map(duplicate["table_border"], "table_border")["show_side_borders"] = false;
                store.Save(duplicate, true);
                Assert(!(bool)Map(store.Get("사용자 1")["table_border"], "table_border")["show_side_borders"], "side borders false roundtrip");
                store.Delete("사용자 3");
                Assert(store.LoadAll().Count == 5, "deletion");

                Dictionary<string, object> partial = checkSerializer.DeserializeObject(File.ReadAllText(file, System.Text.Encoding.UTF8)) as Dictionary<string, object>;
                object[] stored = (object[])partial["presets"];
                Map(Map(Map(stored[1], "preset")["paragraph_styles"], "paragraph_styles")["보고서 본문1"], "보고서 본문1").Remove("font_family");
                File.WriteAllText(file, checkSerializer.Serialize(partial), System.Text.Encoding.UTF8);
                Assert(store.LoadAll().Any(item => Convert.ToString(item["preset_name"]) == "사용자 1"), "valid preset recovery");
                Assert(Directory.GetFiles(directory, "presets.json.corrupt.*.bak").Length == 1, "partial-corrupt backup");
                string partialBackup = store.RecoveryBackupPath;
                Assert(File.Exists(partialBackup) && store.RecoveryWarning.Contains(partialBackup), "partial recovery warning contains backup path");
                Assert(warnings.Count == 1 && warnings[0] == store.RecoveryWarning, "partial recovery logged");
                store.LoadAll();
                Assert(warnings.Count == 1, "recovered store does not repeat warning");

                File.WriteAllText(file, "{broken", System.Text.Encoding.UTF8);
                Assert(store.LoadAll().Count == 3, "corrupt recovery");
                Assert(Directory.GetFiles(directory, "presets.json.corrupt.*.bak").Length == 2, "corrupt backup");
                string corruptBackup = store.RecoveryBackupPath;
                Assert(corruptBackup != partialBackup && File.Exists(corruptBackup) && store.RecoveryWarning.Contains(corruptBackup), "corrupt recovery warning contains backup path");
                Assert(warnings.Count == 2 && warnings[1] == store.RecoveryWarning, "corrupt recovery logged");
                File.WriteAllText(file, "{broken", System.Text.Encoding.UTF8);
                var quietStore = new HwpStylePresetStore(file);
                Assert(quietStore.LoadAll().Count == 3 && !string.IsNullOrWhiteSpace(quietStore.RecoveryWarning), "recovery without interactive logging");
                AssertThrows(delegate { store.Delete(BuiltInName); }, "immutable default");

                foreach (object invalid in new object[] { "false", 0, null })
                {
                    Dictionary<string, object> invalidBoolean = DefaultConfig();
                    Map(invalidBoolean["table_border"], "table_border")["show_side_borders"] = invalid;
                    AssertThrows(delegate { Normalize(invalidBoolean); }, "strict side borders boolean validation");
                }
                Dictionary<string, object> invalidText = DefaultConfig();
                Map(Map(invalidText["paragraph_styles"], "paragraph_styles")["보고서 본문1"], "보고서 본문1")["font_family"] = 123;
                AssertThrows(delegate { Normalize(invalidText); }, "strict string validation");
                Dictionary<string, object> invalidNumber = DefaultConfig();
                Map(Map(invalidNumber["paragraph_styles"], "paragraph_styles")["보고서 본문1"], "보고서 본문1")["font_size_pt"] = double.NaN;
                AssertThrows(delegate { Normalize(invalidNumber); }, "NaN rejection");
                Map(Map(invalidNumber["paragraph_styles"], "paragraph_styles")["보고서 본문1"], "보고서 본문1")["font_size_pt"] = double.PositiveInfinity;
                AssertThrows(delegate { Normalize(invalidNumber); }, "infinite config value rejection");
                AssertThrows(delegate { ParseFiniteDouble("NaN", "grid"); }, "NaN grid value rejection");
                AssertThrows(delegate { ParseFiniteDouble("Infinity", "grid"); }, "infinite grid value rejection");
            }
            finally
            {
                try { Directory.Delete(directory, true); } catch { }
            }
        }

        private List<Dictionary<string, object>> LoadUsers()
        {
            var valid = new List<Dictionary<string, object>>();
            if (!File.Exists(path)) return valid;
            bool damaged = false;
            try
            {
                Dictionary<string, object> root = serializer.DeserializeObject(File.ReadAllText(path, System.Text.Encoding.UTF8)) as Dictionary<string, object>;
                object presetsValue;
                object[] presets = root != null && root.TryGetValue("presets", out presetsValue) ? presetsValue as object[] : null;
                if (presets == null) throw new InvalidDataException("presets 배열이 없습니다.");
                var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (object value in presets)
                {
                    try
                    {
                        Dictionary<string, object> preset = Normalize(Map(value, "preset"));
                        string name = Convert.ToString(preset["preset_name"]);
                        if (string.Equals(name, BuiltInName, StringComparison.OrdinalIgnoreCase) || !names.Add(name) || valid.Count >= 3)
                        {
                            damaged = true;
                            continue;
                        }
                        valid.Add(preset);
                    }
                    catch
                    {
                        damaged = true;
                    }
                }
            }
            catch
            {
                damaged = true;
            }
            if (damaged)
            {
                RecoveryBackupPath = BackupCorrupt();
                WriteUsers(valid);
                RecoveryWarning = "HWP 서식 프리셋 파일의 손상된 항목을 복구했습니다. 백업: " + RecoveryBackupPath;
                if (recoveryLog != null) recoveryLog(RecoveryWarning);
            }
            return valid;
        }

        private void WriteUsers(List<Dictionary<string, object>> users)
        {
            var root = new Dictionary<string, object> { { "schema_version", "1.0" }, { "presets", users.ToArray() } };
            AtomicWrite(path, serializer.Serialize(root));
        }

        private string BackupCorrupt()
        {
            string backup = path + ".corrupt." + DateTime.Now.ToString("yyyyMMddHHmmssfff") + "." + Guid.NewGuid().ToString("N") + ".bak";
            File.Copy(path, backup, false);
            return backup;
        }

        internal static void AtomicWrite(string outputPath, string json)
        {
            string directory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrWhiteSpace(directory)) Directory.CreateDirectory(directory);
            string temp = outputPath + ".tmp." + Guid.NewGuid().ToString("N");
            File.WriteAllText(temp, json, new System.Text.UTF8Encoding(false));
            try
            {
                if (File.Exists(outputPath))
                {
                    string backup = outputPath + ".replace.bak";
                    File.Replace(temp, outputPath, backup, true);
                    if (File.Exists(backup)) File.Delete(backup);
                }
                else
                {
                    File.Move(temp, outputPath);
                }
            }
            finally
            {
                if (File.Exists(temp)) File.Delete(temp);
            }
        }

        private static Dictionary<string, object> Style(string font, double size, bool bold, string bullet, int spacing, double left, double first, string alignment)
        {
            return new Dictionary<string, object> { { "font_family", font }, { "font_size_pt", size }, { "bold", bold }, { "bullet", bullet }, { "line_spacing_percent", spacing }, { "left_indent_mm", left }, { "first_line_indent_mm", first }, { "alignment", alignment } };
        }

        private static Dictionary<string, object> Cell(string color)
        {
            return new Dictionary<string, object> { { "fill_color", color }, { "vertical_alignment", "center" } };
        }

        private static object Value(Dictionary<string, object> map, string key)
        {
            object value;
            if (!map.TryGetValue(key, out value)) throw new InvalidOperationException(key + " 값이 없습니다.");
            return value;
        }

        private static Dictionary<string, object> Map(object value, string field)
        {
            Dictionary<string, object> map = value as Dictionary<string, object>;
            if (map == null) throw new InvalidOperationException(field + " 값이 올바르지 않습니다.");
            return map;
        }

        private static void RequireExactKeys(Dictionary<string, object> map, string[] expected, string field)
        {
            if (map.Count != expected.Length || expected.Any(name => !map.ContainsKey(name)))
                throw new InvalidOperationException(field + "에는 고정된 서식 이름만 사용할 수 있습니다.");
        }

        private static void RequireText(Dictionary<string, object> map, string key, string field)
        {
            object value = Value(map, key);
            if (!(value is string) || string.IsNullOrWhiteSpace((string)value))
                throw new InvalidOperationException(field + "." + key + " 값을 입력하세요.");
        }

        private static double Number(Dictionary<string, object> map, string key, double minimum, double maximum, string field)
        {
            object value = Value(map, key);
            if (value is bool) throw new InvalidOperationException(field + "." + key + " 값이 올바르지 않습니다.");
            double number;
            try { number = Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture); }
            catch { throw new InvalidOperationException(field + "." + key + " 값이 올바르지 않습니다."); }
            if (double.IsNaN(number) || double.IsInfinity(number) || number < minimum || number > maximum)
                throw new InvalidOperationException(field + "." + key + " 범위는 " + minimum + "~" + maximum + "입니다.");
            return number;
        }

        internal static double ParseFiniteDouble(string text, string field)
        {
            double value;
            if ((!double.TryParse(text, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.CurrentCulture, out value) &&
                 !double.TryParse(text, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out value)) ||
                double.IsNaN(value) || double.IsInfinity(value))
                throw new InvalidOperationException(field + " 값은 유한한 숫자여야 합니다.");
            return value;
        }

        private static string ColorValue(Dictionary<string, object> map, string key, string field)
        {
            object value = Value(map, key);
            if (!(value is string)) throw new InvalidOperationException(field + "." + key + " 값은 문자열이어야 합니다.");
            string color = ((string)value).Trim().ToUpperInvariant();
            if (color.Length != 6 || color.Any(ch => !Uri.IsHexDigit(ch)))
                throw new InvalidOperationException(field + "." + key + " 값은 6자리 RGB 색상이어야 합니다.");
            return color;
        }

        private static Dictionary<string, object> Clone(Dictionary<string, object> value)
        {
            var serializer = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            return serializer.DeserializeObject(serializer.Serialize(value)) as Dictionary<string, object>;
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition) throw new InvalidOperationException("Self-check failed: " + message);
        }

        private static void AssertThrows(Action action, string message)
        {
            try { action(); }
            catch { return; }
            throw new InvalidOperationException("Self-check failed: " + message);
        }
    }

    internal static class HwpPageSetup
    {
        internal static readonly string[] NumberKeys =
        {
            "paper_width_mm", "paper_height_mm", "top_mm", "bottom_mm", "left_mm", "right_mm",
            "gutter_mm", "header_mm", "footer_mm"
        };

        internal static Dictionary<string, object> DefaultConfig()
        {
            return new Dictionary<string, object>
            {
                { "paper_width_mm", 210.0 }, { "paper_height_mm", 297.0 }, { "orientation", "portrait" },
                { "top_mm", 20.0 }, { "bottom_mm", 15.0 }, { "left_mm", 24.7 }, { "right_mm", 25.0 },
                { "gutter_mm", 0.0 }, { "header_mm", 10.0 }, { "footer_mm", 15.0 }
            };
        }

        internal static Dictionary<string, object> Normalize(Dictionary<string, object> config)
        {
            if (config == null || config.Count != NumberKeys.Length + 1 || !config.ContainsKey("orientation")
                || NumberKeys.Any(key => !config.ContainsKey(key)))
                throw new InvalidOperationException("편집용지 설정의 필수 키가 없거나 알 수 없는 키가 있습니다.");
            string orientation = config["orientation"] as string;
            if (orientation != "portrait" && orientation != "landscape")
                throw new InvalidOperationException("orientation은 portrait 또는 landscape여야 합니다.");
            var copy = new Dictionary<string, object> { { "orientation", orientation } };
            var units = new Dictionary<string, double>();
            foreach (string key in NumberKeys)
            {
                object value = config[key];
                TypeCode type = value == null ? TypeCode.Empty : Type.GetTypeCode(value.GetType());
                if (type < TypeCode.SByte || type > TypeCode.Decimal)
                    throw new InvalidOperationException(key + " 값은 숫자여야 합니다.");
                double number = Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture);
                bool size = key == "paper_width_mm" || key == "paper_height_mm";
                if (double.IsNaN(number) || double.IsInfinity(number) || (size ? number <= 0 : number < 0))
                    throw new InvalidOperationException(key + " 값은 유한한 " + (size ? "양수" : "0 이상의 수") + "여야 합니다.");
                copy[key] = number;
                // Match the writer's signed-I4 HWPUNIT boundary and per-field rounding.
                double rounded = Math.Round(number * 7200.0 / 25.4);
                if (rounded > int.MaxValue)
                    throw new InvalidOperationException(key + " 값이 HWPUNIT 범위를 초과합니다.");
                units[key] = rounded;
            }
            // Store base sheet dimensions; only the effective landscape dimensions are swapped.
            double width = (double)copy[orientation == "landscape" ? "paper_height_mm" : "paper_width_mm"];
            double height = (double)copy[orientation == "landscape" ? "paper_width_mm" : "paper_height_mm"];
            if (width - (double)copy["left_mm"] - (double)copy["right_mm"] - (double)copy["gutter_mm"] <= 0
                || height - (double)copy["top_mm"] - (double)copy["bottom_mm"] <= 0)
                throw new InvalidOperationException("여백을 제외한 본문 너비와 높이는 0보다 커야 합니다.");
            if (units[orientation == "landscape" ? "paper_height_mm" : "paper_width_mm"] - units["left_mm"] - units["right_mm"] - units["gutter_mm"] <= 0
                || units[orientation == "landscape" ? "paper_width_mm" : "paper_height_mm"] - units["top_mm"] - units["bottom_mm"] <= 0)
                throw new InvalidOperationException("HWPUNIT 변환 후 본문 너비와 높이는 0보다 커야 합니다.");
            // Header/footer live inside top/bottom margins, not additional body deductions.
            if ((double)copy["header_mm"] > (double)copy["top_mm"] || (double)copy["footer_mm"] > (double)copy["bottom_mm"])
                throw new InvalidOperationException("머리말은 위쪽 여백, 꼬리말은 아래쪽 여백 이하여야 합니다.");
            return copy;
        }

        internal static Dictionary<string, object> Load(string path)
        {
            return Normalize(new JavaScriptSerializer().DeserializeObject(
                File.ReadAllText(path, System.Text.Encoding.UTF8)) as Dictionary<string, object>);
        }

        internal static void WriteConfig(string path, Dictionary<string, object> config)
        {
            HwpStylePresetStore.AtomicWrite(path, new JavaScriptSerializer().Serialize(Normalize(config)));
        }
    }

    internal sealed class HwpPageSetupDialog : Form
    {
        private readonly CheckBox overrideCheck = new CheckBox();
        private readonly ComboBox paperCombo = new ComboBox();
        private readonly ComboBox orientationCombo = new ComboBox();
        private readonly TableLayoutPanel editor = new TableLayoutPanel();
        private readonly Dictionary<string, NumericUpDown> numbers = new Dictionary<string, NumericUpDown>();
        // B5 uses ISO dimensions. Presets never change orientation or margins.
        private static readonly decimal[,] PaperSizes = { { 210M, 297M }, { 176M, 250M }, { 297M, 420M }, { 215.9M, 279.4M } };
        private bool loadingPaper;
        internal Dictionary<string, object> SelectedConfig { get; private set; }

        internal HwpPageSetupDialog(Dictionary<string, object> config)
        {
            Text = "HWP 편집용지";
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            LauncherUi.ApplyToForm(this);
            MinimumSize = new Size(640, 440);
            Size = MinimumSize;

            var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(16), ColumnCount = 1, RowCount = 4 };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(root);
            overrideCheck.Text = "템플릿 대신 편집용지 설정 적용";
            overrideCheck.AutoSize = true;
            overrideCheck.CheckedChanged += delegate { editor.Enabled = overrideCheck.Checked; };
            root.Controls.Add(overrideCheck, 0, 0);
            var scope = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            scope.Controls.Add(new Label { Text = "현재 본문 구역", AutoSize = true });
            scope.Controls.Add(new Label { Text = "원본 템플릿 유지", AutoSize = true, Margin = new Padding(16, 0, 0, 0) });
            root.Controls.Add(scope, 0, 1);

            editor.Dock = DockStyle.Fill;
            editor.ColumnCount = 4;
            editor.RowCount = 6;
            for (int column = 0; column < 4; column++) editor.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            for (int row = 0; row < 6; row++) editor.RowStyles.Add(new RowStyle(SizeType.Percent, 100F / 6));
            root.Controls.Add(editor, 0, 2);
            paperCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            paperCombo.Items.AddRange(new object[] { "A4", "B5 (ISO)", "A3", "Letter", "사용자 지정" });
            orientationCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            orientationCombo.Items.AddRange(new object[] { "세로", "가로" });
            AddField("용지", paperCombo, 0, 0);
            AddField("방향", orientationCombo, 2, 0);
            string[] labels = { "기본 너비 (mm)", "기본 높이 (mm)", "위쪽 (mm)", "아래쪽 (mm)", "왼쪽 (mm)", "오른쪽 (mm)", "제본 (mm)", "머리말 (mm)", "꼬리말 (mm)" };
            for (int i = 0; i < HwpPageSetup.NumberKeys.Length; i++)
            {
                string key = HwpPageSetup.NumberKeys[i];
                var number = new NumericUpDown { DecimalPlaces = 2, Increment = 0.1M, Minimum = 0, Maximum = 10000M, AccessibleName = labels[i] };
                numbers.Add(key, number);
                int row = i < 7 ? 1 + i / 2 : 5;
                int column = i < 7 ? (i % 2) * 2 : (i - 7) * 2;
                AddField(labels[i], number, column, row);
            }
            var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.RightToLeft };
            var apply = new Button { Text = "적용", Width = 96 };
            var cancel = new Button { Text = "취소", Width = 96, DialogResult = DialogResult.Cancel };
            apply.Click += delegate
            {
                try
                {
                    SelectedConfig = ReadEditor();
                    DialogResult = DialogResult.OK;
                }
                catch (Exception ex) { MessageBox.Show(this, ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            };
            buttons.Controls.Add(apply);
            buttons.Controls.Add(cancel);
            root.Controls.Add(buttons, 0, 3);
            AcceptButton = apply;
            CancelButton = cancel;
            // ApplyTree sizes nested TextBoxes; NumericUpDown must retain its native edit height.
            LauncherUi.StyleButton(apply, LauncherButtonKind.Secondary);
            LauncherUi.StyleButton(cancel, LauncherButtonKind.Ghost);

            Dictionary<string, object> initial = HwpPageSetup.Normalize(config ?? HwpPageSetup.DefaultConfig());
            loadingPaper = true;
            foreach (string key in HwpPageSetup.NumberKeys) numbers[key].Value = Convert.ToDecimal(initial[key]);
            orientationCombo.SelectedIndex = (string)initial["orientation"] == "landscape" ? 1 : 0;
            paperCombo.SelectedIndex = 4;
            for (int i = 0; i < PaperSizes.GetLength(0); i++)
                if (numbers["paper_width_mm"].Value == PaperSizes[i, 0] && numbers["paper_height_mm"].Value == PaperSizes[i, 1]) paperCombo.SelectedIndex = i;
            loadingPaper = false;
            paperCombo.SelectedIndexChanged += delegate
            {
                int index = paperCombo.SelectedIndex;
                if (loadingPaper || index < 0 || index >= PaperSizes.GetLength(0)) return;
                loadingPaper = true;
                numbers["paper_width_mm"].Value = PaperSizes[index, 0];
                numbers["paper_height_mm"].Value = PaperSizes[index, 1];
                loadingPaper = false;
            };
            numbers["paper_width_mm"].ValueChanged += MarkCustomPaper;
            numbers["paper_height_mm"].ValueChanged += MarkCustomPaper;
            overrideCheck.Checked = config != null;
            editor.Enabled = overrideCheck.Checked;
        }

        private void AddField(string text, Control control, int column, int row)
        {
            editor.Controls.Add(new Label { Text = text, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(3, 6, 3, 6) }, column, row);
            control.Dock = DockStyle.Top;
            control.Margin = new Padding(3, 6, 8, 6);
            editor.Controls.Add(control, column + 1, row);
        }

        private void MarkCustomPaper(object sender, EventArgs e)
        {
            if (!loadingPaper) paperCombo.SelectedIndex = 4;
        }

        internal Dictionary<string, object> ReadEditor()
        {
            if (!overrideCheck.Checked) return null;
            var config = new Dictionary<string, object>();
            foreach (string key in HwpPageSetup.NumberKeys)
            {
                numbers[key].Validate();
                config[key] = Convert.ToDouble(numbers[key].Value);
            }
            config["orientation"] = orientationCombo.SelectedIndex == 1 ? "landscape" : "portrait";
            return HwpPageSetup.Normalize(config);
        }

        internal static void RunSelfCheck()
        {
            using (var dialog = new HwpPageSetupDialog(null))
            {
                if (dialog.ReadEditor() != null || dialog.editor.Enabled) throw new InvalidOperationException("Self-check failed: dialog preserve default");
                dialog.overrideCheck.Checked = true;
                if (!dialog.editor.Enabled || new JavaScriptSerializer().Serialize(dialog.ReadEditor()) != new JavaScriptSerializer().Serialize(HwpPageSetup.Normalize(HwpPageSetup.DefaultConfig())))
                    throw new InvalidOperationException("Self-check failed: dialog override defaults");
                dialog.orientationCombo.SelectedIndex = 1;
                for (int i = 0; i < PaperSizes.GetLength(0); i++)
                {
                    dialog.paperCombo.SelectedIndex = i;
                    if (dialog.numbers["paper_width_mm"].Value != PaperSizes[i, 0] || dialog.numbers["paper_height_mm"].Value != PaperSizes[i, 1]
                        || (string)dialog.ReadEditor()["orientation"] != "landscape")
                        throw new InvalidOperationException("Self-check failed: paper preset/base dimensions");
                }
                dialog.numbers["paper_width_mm"].Value = 220M;
                if (dialog.paperCombo.SelectedIndex != 4) throw new InvalidOperationException("Self-check failed: custom dimensions");
                dialog.overrideCheck.Checked = false;
                if (dialog.ReadEditor() != null || dialog.editor.Enabled) throw new InvalidOperationException("Self-check failed: dialog preserve after override");
                dialog.Size = dialog.MinimumSize;
                dialog.CreateControl();
                CheckLayout(dialog);
                dialog.overrideCheck.Checked = true;
                CheckLayout(dialog);
                Console.WriteLine("Minimum-size page setup layout passed: " + dialog.ClientSize);
            }
        }

        private static void CheckLayout(Control parent)
        {
            parent.PerformLayout();
            foreach (Control child in parent.Controls)
            {
                if (!parent.ClientRectangle.Contains(child.Bounds)) throw new InvalidOperationException("Self-check failed: page setup control clipped: " + child.Text);
                if ((child is Label || child is CheckBox || child is Button) && (child.Width < child.PreferredSize.Width || child.Height < child.PreferredSize.Height))
                    throw new InvalidOperationException("Self-check failed: page setup text clipped: " + child.Text);
                for (int i = 0; i < parent.Controls.Count; i++)
                {
                    Control other = parent.Controls[i];
                    if (other != child && child.Bounds.IntersectsWith(other.Bounds)) throw new InvalidOperationException("Self-check failed: page setup controls overlap");
                }
                CheckLayout(child);
            }
        }
    }

    internal enum LauncherButtonKind
    {
        Primary,
        Secondary,
        Ghost
    }

    internal static class LauncherUi
    {
        public static readonly Color ColorBackground = Color.FromArgb(247, 248, 247);
        public static readonly Color ColorSurface = Color.White;
        public static readonly Color ColorSurfaceAlt = Color.FromArgb(239, 242, 239);
        public static readonly Color ColorSurfaceStrong = Color.FromArgb(224, 235, 229);
        public static readonly Color ColorBorder = Color.FromArgb(204, 212, 207);
        public static readonly Color ColorText = Color.FromArgb(29, 38, 34);
        public static readonly Color ColorMutedText = Color.FromArgb(74, 89, 82);
        public static readonly Color ColorPrimary = Color.FromArgb(25, 83, 65);
        public static readonly Color ColorPrimaryHover = Color.FromArgb(19, 65, 51);
        public static readonly Color ColorSuccess = Color.FromArgb(20, 112, 74);
        public static readonly Color ColorWarning = Color.FromArgb(166, 94, 12);
        public static readonly Color ColorDanger = Color.FromArgb(175, 48, 48);
        public static readonly Color ColorDisabledSurface = Color.FromArgb(235, 238, 236);
        public static readonly Color ColorDisabledText = Color.FromArgb(99, 110, 104);
        public static readonly Color ColorRowAlt = Color.FromArgb(246, 249, 247);

        public const int SpaceXs = 4;
        public const int SpaceSm = 8;
        public const int SpaceMd = 12;
        public const int SpaceLg = 16;
        public const int SpaceXl = 20;
        public const int ControlHeight = 32;

        public static Font BaseFont()
        {
            return new Font("맑은 고딕", 9F, FontStyle.Regular);
        }

        public static Font TitleFont()
        {
            return new Font("맑은 고딕", 16F, FontStyle.Bold);
        }

        public static Font SectionFont()
        {
            return new Font("맑은 고딕", 9.5F, FontStyle.Bold);
        }

        public static Font SmallFont()
        {
            return new Font("맑은 고딕", 8.5F, FontStyle.Regular);
        }

        public static void ApplyToForm(Form form)
        {
            form.BackColor = ColorBackground;
            form.ForeColor = ColorText;
            form.Font = BaseFont();
        }

        public static void ApplyTree(Control root)
        {
            ApplyControl(root);
            foreach (Control child in root.Controls)
            {
                ApplyTree(child);
            }
        }

        public static void StyleButton(Button button, LauncherButtonKind kind)
        {
            button.Height = ControlHeight;
            button.FlatStyle = FlatStyle.Flat;
            button.UseVisualStyleBackColor = false;
            button.Font = kind == LauncherButtonKind.Primary ? SectionFont() : BaseFont();
            button.Padding = new Padding(SpaceSm, 0, SpaceSm, 0);

            Color normalBack;
            Color normalFore;
            Color border;
            Color hoverBack;
            if (kind == LauncherButtonKind.Primary)
            {
                normalBack = ColorPrimary;
                normalFore = Color.White;
                border = ColorPrimary;
                hoverBack = ColorPrimaryHover;
            }
            else if (kind == LauncherButtonKind.Secondary)
            {
                normalBack = ColorSurface;
                normalFore = ColorText;
                border = ColorBorder;
                hoverBack = ColorSurfaceAlt;
            }
            else
            {
                normalBack = ColorSurfaceAlt;
                normalFore = ColorText;
                border = ColorBorder;
                hoverBack = Color.FromArgb(229, 235, 243);
            }

            button.BackColor = normalBack;
            button.ForeColor = normalFore;
            button.FlatAppearance.BorderColor = border;
            button.FlatAppearance.BorderSize = 1;
            button.FlatAppearance.MouseOverBackColor = hoverBack;
            button.FlatAppearance.MouseDownBackColor = kind == LauncherButtonKind.Primary ? ColorPrimaryHover : ColorSurfaceStrong;
            if (string.IsNullOrWhiteSpace(button.AccessibleName))
            {
                button.AccessibleName = button.Text;
            }

            EventHandler enabledHandler = delegate
            {
                if (button.Enabled)
                {
                    button.BackColor = normalBack;
                    button.ForeColor = normalFore;
                    button.FlatAppearance.BorderColor = border;
                }
                else
                {
                    button.BackColor = ColorDisabledSurface;
                    button.ForeColor = ColorDisabledText;
                    button.FlatAppearance.BorderColor = ColorBorder;
                }
            };
            button.EnabledChanged += enabledHandler;
            enabledHandler(button, EventArgs.Empty);
        }

        public static void StyleListItem(ListViewItem item, int index)
        {
            item.UseItemStyleForSubItems = true;
            item.BackColor = index % 2 == 0 ? ColorSurface : ColorRowAlt;
        }

        private static void ApplyControl(Control control)
        {
            if (control is TabPage)
            {
                control.BackColor = ColorBackground;
                control.ForeColor = ColorText;
                return;
            }

            if (control is GroupBox)
            {
                GroupBox groupBox = (GroupBox)control;
                groupBox.BackColor = ColorBackground;
                groupBox.Padding = new Padding(SpaceLg, SpaceXl + SpaceSm, SpaceLg, SpaceLg);
                control.ForeColor = ColorText;
                control.Font = SectionFont();
                groupBox.Paint += PaintGroupBox;
                return;
            }

            if (control is TextBox)
            {
                TextBox textBox = (TextBox)control;
                textBox.BorderStyle = BorderStyle.FixedSingle;
                textBox.BackColor = textBox.ReadOnly ? ColorBackground : ColorSurface;
                textBox.ForeColor = ColorText;
                textBox.Font = BaseFont();
                textBox.MinimumSize = new Size(0, ControlHeight);
                return;
            }

            if (control is ComboBox)
            {
                ComboBox comboBox = (ComboBox)control;
                comboBox.FlatStyle = FlatStyle.Flat;
                comboBox.BackColor = ColorSurface;
                comboBox.ForeColor = ColorText;
                comboBox.Font = BaseFont();
                comboBox.MinimumSize = new Size(0, ControlHeight);
                return;
            }

            if (control is ListView)
            {
                ListView listView = (ListView)control;
                listView.BorderStyle = BorderStyle.FixedSingle;
                listView.BackColor = ColorSurface;
                listView.ForeColor = ColorText;
                listView.GridLines = false;
                listView.HideSelection = false;
                listView.Font = BaseFont();
                return;
            }

            if (control is CheckedListBox)
            {
                CheckedListBox checkedListBox = (CheckedListBox)control;
                checkedListBox.BorderStyle = BorderStyle.FixedSingle;
                checkedListBox.BackColor = ColorSurface;
                checkedListBox.ForeColor = ColorText;
                checkedListBox.Font = BaseFont();
                return;
            }

            if (control is Button)
            {
                Button button = (Button)control;
                StyleButton(button, button.Tag is LauncherButtonKind ? (LauncherButtonKind)button.Tag : LauncherButtonKind.Secondary);
                return;
            }

            if (control is CheckBox)
            {
                CheckBox checkBox = (CheckBox)control;
                checkBox.AutoSize = true;
                checkBox.ForeColor = ColorText;
                checkBox.Font = BaseFont();
                checkBox.Margin = new Padding(0, SpaceSm, SpaceLg, SpaceXs);
                return;
            }

            if (control is FlowLayoutPanel)
            {
                FlowLayoutPanel flow = (FlowLayoutPanel)control;
                if (flow.BackColor == SystemColors.Control || flow.BackColor == Color.Empty)
                {
                    flow.BackColor = ColorBackground;
                }
                flow.WrapContents = true;
                return;
            }

            if (control is Label)
            {
                Label label = (Label)control;
                label.ForeColor = label.ForeColor == Color.Empty || label.ForeColor == Color.DimGray ? ColorMutedText : label.ForeColor;
                if (label.Font == null || label.Font.Style == FontStyle.Regular)
                {
                    label.Font = BaseFont();
                }
            }
        }

        private static void PaintGroupBox(object sender, PaintEventArgs e)
        {
            GroupBox groupBox = (GroupBox)sender;
            e.Graphics.Clear(ColorBackground);

            using (Pen borderPen = new Pen(ColorBorder))
            {
                e.Graphics.DrawLine(borderPen, 0, 12, groupBox.Width, 12);
            }

            Size textSize = TextRenderer.MeasureText(groupBox.Text, groupBox.Font);
            Rectangle textRect = new Rectangle(SpaceMd, 0, textSize.Width + SpaceMd, textSize.Height + 4);
            using (SolidBrush backBrush = new SolidBrush(ColorBackground))
            {
                e.Graphics.FillRectangle(backBrush, textRect);
            }
            TextRenderer.DrawText(
                e.Graphics,
                groupBox.Text,
                groupBox.Font,
                textRect,
                ColorText,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter);
        }
    }

    internal sealed class HwpStylePresetDialog : Form
    {
        private readonly HwpStylePresetStore store;
        private readonly ComboBox presetCombo = new ComboBox();
        private readonly TextBox nameText = new TextBox();
        private readonly DataGridView paragraphGrid = new DataGridView();
        private readonly DataGridView tableGrid = new DataGridView();
        private readonly NumericUpDown innerBorder = new NumericUpDown();
        private readonly NumericUpDown outerBorder = new NumericUpDown();
        private readonly TextBox borderColor = new TextBox();
        private readonly CheckBox showSideBorders = new CheckBox();
        private readonly Button importedLines = new Button { AutoSize = true };
        private readonly HwpStylePreview preview = new HwpStylePreview { Dock = DockStyle.Fill, Margin = Padding.Empty };
        private readonly Label previewNote = new Label { Dock = DockStyle.Fill, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft };
        private readonly System.Windows.Forms.Timer previewTimer = new System.Windows.Forms.Timer { Interval = 100 };
        private readonly JavaScriptSerializer previewSerializer = new JavaScriptSerializer();
        private string savedEditor;
        private Dictionary<string, object> edges = new Dictionary<string, object>();

        internal string SelectedPresetName { get; private set; }

        internal HwpStylePresetDialog(HwpStylePresetStore store, string selectedName)
        {
            this.store = store;
            SelectedPresetName = selectedName;
            Text = "HWP 보고서 서식 설정";
            StartPosition = FormStartPosition.CenterParent;
            MinimumSize = new Size(980, 700);
            Size = new Size(1080, 830);
            BuildUi();
            previewTimer.Tick += delegate { RefreshPreview(); };
            Disposed += delegate { previewTimer.Dispose(); };
            ReloadPresets(selectedName);
            LauncherUi.ApplyToForm(this);
            LauncherUi.ApplyTree(this);
            RefreshPreview();
        }

        private void BuildUi()
        {
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(16), RowCount = 7, ColumnCount = 1 };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 58));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 42));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 226));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(root);

            var top = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            top.Controls.Add(new Label { Text = "저장된 서식", AutoSize = true, Margin = new Padding(0, 8, 8, 0) });
            presetCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            presetCombo.Width = 200;
            presetCombo.SelectedIndexChanged += delegate
            {
                if (presetCombo.SelectedItem == null) return;
                LoadPreset(presetCombo.SelectedItem.ToString());
                savedEditor = previewSerializer.Serialize(ReadEditor(false));
            };
            top.Controls.Add(presetCombo);
            top.Controls.Add(new Label { Text = "서식 이름", AutoSize = true, Margin = new Padding(18, 8, 8, 0) });
            nameText.Width = 220;
            top.Controls.Add(nameText);
            top.Controls.Add(MakeButton("글꼴 선택", ChooseFont));
            top.Controls.Add(MakeButton("HWPX에서 셀 서식 가져오기…", ImportTemplateCell));
            root.Controls.Add(top, 0, 0);

            var help = new Label
            {
                Text = "직접 만들기: 기본 서식을 고르고 글꼴·색·선을 바꾼 뒤, 새 이름으로 [내 서식으로 저장]을 누르세요." + Environment.NewLine +
                       "기존 양식 활용: 아래 표 스타일에서 역할(예: 표머리행)을 고른 뒤, HWPX의 예시 셀과 가져올 속성을 선택하세요." + Environment.NewLine +
                       "가져온 서식은 저장 후 [이 서식 사용]을 눌러 적용합니다. 내 서식은 최대 3개이며, 표의 내용·병합은 복사하지 않습니다.",
                AutoSize = true,
                ForeColor = LauncherUi.ColorMutedText,
                Margin = new Padding(0, 6, 0, 10)
            };
            root.Controls.Add(help, 0, 1);

            ConfigureParagraphGrid();
            root.Controls.Add(paragraphGrid, 0, 2);
            ConfigureTableGrid();
            root.Controls.Add(tableGrid, 0, 3);
            foreach (var editor in new[] { paragraphGrid, tableGrid })
            {
                editor.BackgroundColor = LauncherUi.ColorSurface;
                editor.CellValueChanged += delegate { QueuePreview(); };
                editor.CurrentCellDirtyStateChanged += delegate
                {
                    if (editor.IsCurrentCellDirty && (editor.CurrentCell is DataGridViewCheckBoxCell || editor.CurrentCell is DataGridViewComboBoxCell))
                        editor.CommitEdit(DataGridViewDataErrorContexts.Commit);
                };
            }

            var borderPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            borderPanel.Controls.Add(new Label { Text = "내부선(mm)", AutoSize = true, Margin = new Padding(0, 8, 6, 0) });
            ConfigureNumber(innerBorder, 0.1M, 5M, 0.01M);
            borderPanel.Controls.Add(innerBorder);
            borderPanel.Controls.Add(new Label { Text = "외곽선(mm)", AutoSize = true, Margin = new Padding(16, 8, 6, 0) });
            ConfigureNumber(outerBorder, 0.1M, 5M, 0.01M);
            borderPanel.Controls.Add(outerBorder);
            borderPanel.Controls.Add(new Label { Text = "선 색상(RGB)", AutoSize = true, Margin = new Padding(16, 8, 6, 0) });
            borderColor.Width = 90;
            borderPanel.Controls.Add(borderColor);
            showSideBorders.Text = "좌우 외곽선 표시";
            showSideBorders.AutoSize = true;
            showSideBorders.Margin = new Padding(16, 6, 0, 0);
            borderPanel.Controls.Add(showSideBorders);
            importedLines.Click += ShowImportedLines;
            borderPanel.Controls.Add(importedLines);
            borderPanel.Controls.Add(MakeButton("셀 배경색", ChooseCellFill));
            root.Controls.Add(borderPanel, 0, 4);
            innerBorder.ValueChanged += delegate { QueuePreview(); };
            outerBorder.ValueChanged += delegate { QueuePreview(); };
            borderColor.TextChanged += delegate { QueuePreview(); };
            showSideBorders.CheckedChanged += delegate { QueuePreview(); };
            nameText.TextChanged += delegate { QueuePreview(); };

            var previewArea = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1, Margin = new Padding(3, 8, 3, 0) };
            previewArea.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            previewArea.RowStyles.Add(new RowStyle(SizeType.Absolute, 25));
            previewArea.Controls.Add(preview, 0, 0);
            previewArea.Controls.Add(previewNote, 0, 1);
            root.Controls.Add(previewArea, 0, 5);

            var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.RightToLeft };
            Button close = MakeButton("이 서식 사용", delegate
            {
                ExecuteEditorAction(delegate
                {
                    var edited = ReadEditor();
                    if (previewSerializer.Serialize(edited) != savedEditor)
                        throw new InvalidOperationException("편집한 서식을 적용하려면 [내 서식으로 저장] 또는 [덮어쓰기]를 먼저 누르세요.");
                    SelectedPresetName = Convert.ToString(edited["preset_name"]);
                    DialogResult = DialogResult.OK;
                });
            });
            Button cancel = MakeButton("취소", delegate { DialogResult = DialogResult.Cancel; });
            close.Tag = LauncherButtonKind.Primary;
            Button save = MakeButton("내 서식으로 저장", SaveNew);
            Button overwrite = MakeButton("덮어쓰기", Overwrite);
            Button delete = MakeButton("삭제", DeletePreset);
            Button reset = MakeButton("기본값으로 초기화", ResetEditor);
            buttons.Controls.Add(close);
            buttons.Controls.Add(cancel);
            buttons.Controls.Add(save);
            buttons.Controls.Add(overwrite);
            buttons.Controls.Add(delete);
            buttons.Controls.Add(reset);
            root.Controls.Add(buttons, 0, 6);
        }

        private void ConfigureParagraphGrid()
        {
            paragraphGrid.Dock = DockStyle.Fill;
            paragraphGrid.AllowUserToAddRows = false;
            paragraphGrid.AllowUserToDeleteRows = false;
            paragraphGrid.RowHeadersVisible = false;
            paragraphGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            paragraphGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "StyleName", HeaderText = "스타일", ReadOnly = true, FillWeight = 110 });
            paragraphGrid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "Enabled", HeaderText = "개별", FillWeight = 48,
                ToolTipText = "켜면 이 역할에 별도 서식을 적용합니다. 끄면 공통 표 서식을 사용합니다." });
            paragraphGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Font", HeaderText = "글꼴", FillWeight = 110 });
            paragraphGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Size", HeaderText = "크기(pt)", FillWeight = 65 });
            paragraphGrid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "Bold", HeaderText = "굵게", FillWeight = 48 });
            paragraphGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Bullet", HeaderText = "말머리", FillWeight = 55 });
            paragraphGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Spacing", HeaderText = "줄간격(%)", FillWeight = 75 });
            paragraphGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Left", HeaderText = "왼쪽(mm)", FillWeight = 70 });
            paragraphGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "First", HeaderText = "첫줄(mm)", FillWeight = 70 });
            paragraphGrid.Columns.Add(new DataGridViewComboBoxColumn { Name = "Align", HeaderText = "정렬", DataSource = new[] { "left", "center", "right", "justify" }, FillWeight = 75 });
        }

        private void ConfigureTableGrid()
        {
            tableGrid.Dock = DockStyle.Fill;
            tableGrid.AllowUserToAddRows = false;
            tableGrid.AllowUserToDeleteRows = false;
            tableGrid.RowHeadersVisible = false;
            tableGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            tableGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "StyleName", HeaderText = "표 스타일", ReadOnly = true });
            tableGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Fill", HeaderText = "배경 RGB" });
            tableGrid.Columns.Add(new DataGridViewComboBoxColumn { Name = "Vertical", HeaderText = "세로 정렬", DataSource = new[] { "top", "center", "bottom" } });
            tableGrid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Margins", HeaderText = "여백: 상/하/좌/우(mm)", FillWeight = 150 });
        }

        private static void ConfigureNumber(NumericUpDown control, decimal minimum, decimal maximum, decimal increment)
        {
            control.Minimum = minimum;
            control.Maximum = maximum;
            control.Increment = increment;
            control.DecimalPlaces = 2;
            control.Width = 75;
        }

        private Button MakeButton(string text, EventHandler handler)
        {
            var button = new Button { Text = text, AutoSize = true, Margin = new Padding(8, 4, 0, 0) };
            button.Click += handler;
            return button;
        }

        private void ReloadPresets(string selectedName)
        {
            presetCombo.Items.Clear();
            foreach (Dictionary<string, object> preset in store.LoadAll())
                presetCombo.Items.Add(Convert.ToString(preset["preset_name"]));
            int index = presetCombo.Items.IndexOf(selectedName);
            presetCombo.SelectedIndex = index >= 0 ? index : 0;
        }

        private void LoadPreset(string name)
        {
            Dictionary<string, object> config = store.Get(name);
            nameText.Text = Convert.ToString(config["preset_name"]);
            paragraphGrid.Rows.Clear();
            Dictionary<string, object> styles = (Dictionary<string, object>)config["paragraph_styles"];
            foreach (string styleName in HwpStylePresetStore.StyleNames.Concat(HwpStylePresetStore.RoleStyleDefaults.Keys))
            {
                bool separate = HwpStylePresetStore.RoleStyleDefaults.ContainsKey(styleName);
                bool enabled = styles.ContainsKey(styleName);
                var style = (Dictionary<string, object>)styles[enabled ? styleName : HwpStylePresetStore.RoleStyleDefaults[styleName]];
                int index = paragraphGrid.Rows.Add(styleName, separate ? (object)enabled : null, style["font_family"], style["font_size_pt"], style["bold"], style["bullet"], style["line_spacing_percent"], style["left_indent_mm"], style["first_line_indent_mm"], style["alignment"]);
                paragraphGrid.Rows[index].Cells["Enabled"].ReadOnly = !separate;
            }
            tableGrid.Rows.Clear();
            Dictionary<string, object> cells = (Dictionary<string, object>)config["table_cell_styles"];
            foreach (string styleName in HwpStylePresetStore.TableStyleNames.Concat(HwpStylePresetStore.RoleStyleDefaults.Keys))
            {
                var style = (Dictionary<string, object>)cells[cells.ContainsKey(styleName) ? styleName : HwpStylePresetStore.RoleStyleDefaults[styleName]];
                tableGrid.Rows.Add(styleName, style["fill_color"], style["vertical_alignment"],
                    style.ContainsKey("margins_mm") ? FormatMargins((Dictionary<string, object>)style["margins_mm"]) : "");
            }
            Dictionary<string, object> border = (Dictionary<string, object>)config["table_border"];
            innerBorder.Value = Convert.ToDecimal(border["inner_width_mm"]);
            outerBorder.Value = Convert.ToDecimal(border["outer_width_mm"]);
            borderColor.Text = Convert.ToString(border["color"]);
            showSideBorders.Checked = (bool)border["show_side_borders"];
            edges = border.ContainsKey("edges") ? (Dictionary<string, object>)border["edges"] : new Dictionary<string, object>();
            UpdateImportedLines();
        }

        internal Dictionary<string, object> ReadEditor(bool commitEdits = true)
        {
            if (commitEdits)
            {
                paragraphGrid.EndEdit();
                tableGrid.EndEdit();
            }
            var styles = new Dictionary<string, object>();
            foreach (DataGridViewRow row in paragraphGrid.Rows)
            {
                string name = Convert.ToString(row.Cells["StyleName"].Value);
                if (HwpStylePresetStore.RoleStyleDefaults.ContainsKey(name) && !Convert.ToBoolean(row.Cells["Enabled"].Value)) continue;
                styles[name] = new Dictionary<string, object>
                {
                    { "font_family", CellText(row, "Font") },
                    { "font_size_pt", CellNumber(row, "Size") },
                    { "bold", Convert.ToBoolean(row.Cells["Bold"].Value) },
                    { "bullet", CellText(row, "Bullet", true) },
                    { "line_spacing_percent", CellNumber(row, "Spacing") },
                    { "left_indent_mm", CellNumber(row, "Left") },
                    { "first_line_indent_mm", CellNumber(row, "First") },
                    { "alignment", CellText(row, "Align") }
                };
            }
            var cells = new Dictionary<string, object>();
            foreach (DataGridViewRow row in tableGrid.Rows)
            {
                string name = Convert.ToString(row.Cells["StyleName"].Value);
                if (!styles.ContainsKey(name)) continue;
                var cell = new Dictionary<string, object>
                {
                    { "fill_color", CellText(row, "Fill") },
                    { "vertical_alignment", CellText(row, "Vertical") }
                };
                string margins = CellText(row, "Margins", true);
                if (margins.Length > 0)
                {
                    string[] values = margins.Split('/');
                    if (values.Length != 4) throw new InvalidOperationException(name + " 여백은 상/하/좌/우 네 값을 입력하세요.");
                    cell["margins_mm"] = HwpStylePresetStore.MarginSides.Select((side, i) => new { side, i }).ToDictionary(
                        item => item.side, item => (object)HwpStylePresetStore.ParseFiniteDouble(values[item.i].Trim(), name + " 여백"));
                }
                cells[name] = cell;
            }
            var config = new Dictionary<string, object>
            {
                { "schema_version", "1.0" }, { "preset_name", !commitEdits && string.IsNullOrWhiteSpace(nameText.Text)
                    ? HwpStylePresetStore.BuiltInName : nameText.Text.Trim() },
                { "paragraph_styles", styles }, { "table_cell_styles", cells },
                { "table_border", new Dictionary<string, object>
                    {
                        { "inner_width_mm", Convert.ToDouble(innerBorder.Value) },
                        { "outer_width_mm", Convert.ToDouble(outerBorder.Value) },
                        { "color", borderColor.Text.Trim() },
                        { "show_side_borders", showSideBorders.Checked }
                    }
                }
            };
            if (edges.Count > 0) ((Dictionary<string, object>)config["table_border"])["edges"] = edges;
            return HwpStylePresetStore.Normalize(config);
        }

        private void QueuePreview()
        {
            previewTimer.Stop();
            previewTimer.Start();
        }

        private void RefreshPreview()
        {
            previewTimer.Stop();
            try
            {
                var config = ReadEditor(false);
                preview.UpdateStyle(config);
                bool changed = previewSerializer.Serialize(config) != savedEditor || string.IsNullOrWhiteSpace(nameText.Text);
                previewNote.ForeColor = LauncherUi.ColorMutedText;
                previewNote.Text = (changed ? "편집 중 · 저장 후 적용  |  " : "") + "예시 데이터 · 실제 줄바꿈, 표 폭, 쪽 나눔은 한글에서 달라질 수 있습니다.";
            }
            catch (InvalidOperationException ex)
            {
                previewNote.ForeColor = LauncherUi.ColorWarning;
                previewNote.Text = "마지막 정상 미리보기 · " + ex.Message;
            }
        }

        internal void CheckPreview()
        {
            LoadPreset(HwpStylePresetStore.GoldTableName);
            RefreshPreview();
            if (!PreviewContainsColor("F1CB7E")) throw new InvalidOperationException("Self-check failed: gold preview");
            LoadPreset(HwpStylePresetStore.BlueTableName);
            RefreshPreview();
            if (!PreviewContainsColor("F0F4F7") || PreviewContainsColor("F1CB7E"))
                throw new InvalidOperationException("Self-check failed: preset preview did not change");

            ApplyTemplateCell("표머리행", new Dictionary<string, object> { { "fill_color", "C8E3D5" } });
            if (!previewTimer.Enabled) throw new InvalidOperationException("Self-check failed: imported cell did not schedule preview");
            paragraphGrid.Rows[1].Cells["Bullet"].Value = "※";
            RefreshPreview();
            if (!PreviewContainsColor("C8E3D5") || !preview.AccessibleDescription.Contains("※") ||
                !previewNote.Text.Contains("저장 후 적용"))
                throw new InvalidOperationException("Self-check failed: edited preview");

            borderColor.Text = "invalid";
            RefreshPreview();
            if (!previewNote.Text.Contains("마지막 정상") || !PreviewContainsColor("C8E3D5"))
                throw new InvalidOperationException("Self-check failed: invalid style lost last valid preview");

            LoadPreset(HwpStylePresetStore.BuiltInName);
            RefreshPreview();
        }

        private bool PreviewContainsColor(string hex)
        {
            using (var bitmap = new Bitmap(preview.Width, preview.Height))
            {
                preview.DrawToBitmap(bitmap, preview.ClientRectangle);
                int rgb = ColorTranslator.FromHtml("#" + hex).ToArgb();
                for (int y = 0; y < bitmap.Height; y += 2)
                    for (int x = 0; x < bitmap.Width; x += 2)
                        if (bitmap.GetPixel(x, y).ToArgb() == rgb) return true;
                return false;
            }
        }

        private static string FormatMargins(Dictionary<string, object> values)
        {
            return string.Join(" / ", HwpStylePresetStore.MarginSides.Select(side => Convert.ToString(values[side])));
        }

        private void ChooseFont(object sender, EventArgs e)
        {
            if (paragraphGrid.CurrentRow == null) return;
            ExecuteEditorAction(delegate
            {
                paragraphGrid.EndEdit();
                var row = paragraphGrid.CurrentRow;
                string name = CellText(row, "StyleName");
                float size = (float)Math.Max(6, Math.Min(30, CellNumber(row, "Size")));
                FontStyle weight = Convert.ToBoolean(row.Cells["Bold"].Value) ? FontStyle.Bold : FontStyle.Regular;
                using (var font = new Font(CellText(row, "Font"), size, weight))
                using (var dialog = new FontDialog { Font = font, FontMustExist = true, ShowEffects = false, AllowVerticalFonts = false, MinSize = 6, MaxSize = 30 })
                {
                    if (dialog.ShowDialog(this) != DialogResult.OK) return;
                    row.Cells["Font"].Value = dialog.Font.Name;
                    row.Cells["Size"].Value = dialog.Font.SizeInPoints;
                    row.Cells["Bold"].Value = dialog.Font.Bold;
                    if (HwpStylePresetStore.RoleStyleDefaults.ContainsKey(name)) row.Cells["Enabled"].Value = true;
                }
            });
        }

        private void UpdateImportedLines()
        {
            importedLines.Text = "선 설정 (" + edges.Count + ")";
            QueuePreview();
        }

        private static string LineText(Dictionary<string, object> line)
        {
            return HwpStylePresetStore.LineTypes[Convert.ToString(line["type"])] + " · " + line["width_mm"] + " mm · #" + line["color"];
        }

        private void ShowImportedLines(object sender, EventArgs e)
        {
            var menu = new ContextMenuStrip();
            foreach (var pair in HwpStylePresetStore.LineGroups)
            {
                string key = pair.Key;
                string detail = edges.ContainsKey(key) ? LineText((Dictionary<string, object>)edges[key]) : "기본값";
                menu.Items.Add(pair.Value + ": " + detail, null, delegate { ExecuteEditorAction(() => EditLine(key)); });
            }
            menu.Closed += delegate { menu.Dispose(); };
            menu.Show(importedLines, new Point(0, importedLines.Height));
        }

        private void ChooseCellFill(object sender, EventArgs e)
        {
            tableGrid.EndEdit();
            var row = tableGrid.CurrentRow;
            if (row == null) return;
            string rgb = ChooseRgb(Convert.ToString(row.Cells["Fill"].Value));
            if (rgb != null) ApplyTemplateCell(CellText(row, "StyleName"), new Dictionary<string, object> { { "fill_color", rgb } });
        }

        private string ChooseRgb(string rgb)
        {
            int value;
            Color color = int.TryParse(rgb.Trim().TrimStart('#'), System.Globalization.NumberStyles.HexNumber,
                System.Globalization.CultureInfo.InvariantCulture, out value) && value >= 0 && value <= 0xFFFFFF
                ? Color.FromArgb(255, (value >> 16) & 255, (value >> 8) & 255, value & 255) : Color.White;
            using (var picker = new ColorDialog { Color = color, FullOpen = true })
                return picker.ShowDialog(Form.ActiveForm ?? this) == DialogResult.OK ? (picker.Color.ToArgb() & 0xFFFFFF).ToString("X6") : null;
        }

        private void EditLine(string key)
        {
            var line = edges.ContainsKey(key) ? (Dictionary<string, object>)edges[key] : null;
            var types = HwpStylePresetStore.LineTypes.ToArray();
            using (var dialog = new Form {
                Text = HwpStylePresetStore.LineGroups[key], StartPosition = FormStartPosition.CenterParent,
                ClientSize = new Size(420, 166), FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false, MinimizeBox = false, ShowInTaskbar = false
            })
            {
                var grid = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, Padding = new Padding(12), AutoSize = true };
                grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
                grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
                grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
                dialog.Controls.Add(grid);
                var type = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Fill };
                type.Items.Add("기본값");
                type.Items.AddRange(types.Select(pair => (object)pair.Value).ToArray());
                type.SelectedIndex = line == null ? 0 : Array.FindIndex(types, pair => pair.Key == Convert.ToString(line["type"])) + 1;
                var width = new NumericUpDown { Dock = DockStyle.Fill };
                ConfigureNumber(width, 0.1M, 5M, 0.01M);
                width.Value = line == null ? (key == "outer" ? outerBorder.Value : innerBorder.Value) : Convert.ToDecimal(line["width_mm"]);
                var rgb = new TextBox { Text = line == null ? borderColor.Text : Convert.ToString(line["color"]), Dock = DockStyle.Fill };
                var pick = MakeButton("색상 선택", delegate { rgb.Text = ChooseRgb(rgb.Text) ?? rgb.Text; });
                grid.Controls.Add(new Label { Text = "선 종류", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
                grid.Controls.Add(type, 1, 0);
                grid.Controls.Add(new Label { Text = "굵기(mm)", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 1);
                grid.Controls.Add(width, 1, 1);
                grid.Controls.Add(new Label { Text = "색상(RGB)", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 2);
                grid.Controls.Add(rgb, 1, 2);
                grid.Controls.Add(pick, 2, 2);
                var buttons = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft };
                var cancel = MakeButton("취소", delegate { dialog.DialogResult = DialogResult.Cancel; });
                var apply = MakeButton("적용", delegate
                {
                    if (type.SelectedIndex == 0) edges.Remove(key);
                    else
                    {
                        string color = rgb.Text.Trim().TrimStart('#').ToUpperInvariant();
                        if (!System.Text.RegularExpressions.Regex.IsMatch(color, @"\A[0-9A-F]{6}\z"))
                        {
                            MessageBox.Show(dialog, "색상은 6자리 RGB로 입력하세요.", dialog.Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                        edges[key] = new Dictionary<string, object> { { "type", types[type.SelectedIndex - 1].Key }, { "width_mm", Convert.ToDouble(width.Value) }, { "color", color } };
                    }
                    UpdateImportedLines();
                    dialog.DialogResult = DialogResult.OK;
                });
                Action refresh = () => { width.Enabled = rgb.Enabled = pick.Enabled = type.SelectedIndex > 0 && types[type.SelectedIndex - 1].Key != "none"; };
                type.SelectedIndexChanged += delegate { refresh(); };
                refresh();
                buttons.Controls.Add(apply);
                buttons.Controls.Add(cancel);
                grid.Controls.Add(buttons, 0, 3);
                grid.SetColumnSpan(buttons, 3);
                dialog.AcceptButton = apply;
                dialog.CancelButton = cancel;
                LauncherUi.ApplyToForm(dialog);
                LauncherUi.ApplyTree(dialog);
                dialog.ShowDialog(this);
            }
        }

        private async void ImportTemplateCell(object sender, EventArgs e)
        {
            if (tableGrid.CurrentRow == null) return;
            string target = CellText(tableGrid.CurrentRow, "StyleName");
            string preset = Convert.ToString(presetCombo.SelectedItem);
            using (var file = new OpenFileDialog { Filter = "한글 문서 (*.hwpx)|*.hwpx", Title = target + "에 가져올 HWPX 선택" })
            {
                if (file.ShowDialog(this) != DialogResult.OK) return;
                string output = null;
                var button = (Button)sender;
                button.Enabled = false;
                try
                {
                    output = Path.GetTempFileName();
                    string input = file.FileName;
                    await System.Threading.Tasks.Task.Run(() => MainForm.RunTemplateTool("hwp_template_probe.py",
                        "--cell-style-output " + MainForm.QuoteArg(output) + " " + MainForm.QuoteArg(input), 30000));
                    if (IsDisposed || Convert.ToString(presetCombo.SelectedItem) != preset) return;
                    var reader = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
                    var catalog = (Dictionary<string, object>)reader.DeserializeObject(File.ReadAllText(output));
                    var values = SelectTemplateCell(target, (object[])catalog["tables"]);
                    if (values != null) ApplyTemplateCell(target, values);
                }
                catch (Exception ex)
                {
                    if (!IsDisposed) MessageBox.Show(this, ex.Message, "셀 서식 가져오기", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                finally
                {
                    if (output != null)
                    {
                        try { File.Delete(output); } catch (IOException) { } catch (UnauthorizedAccessException) { }
                    }
                    if (!button.IsDisposed) button.Enabled = true;
                }
            }
        }

        private Dictionary<string, object> SelectTemplateCell(string target, object[] source)
        {
            var tables = source.Cast<Dictionary<string, object>>().Where(t => ((object[])t["cells"]).Length > 0).ToArray();
            if (tables.Length == 0) throw new InvalidOperationException("가져올 수 있는 셀 서식이 없습니다.");
            using (var dialog = new Form { Text = target + "에 가져올 셀과 속성", StartPosition = FormStartPosition.CenterParent, Size = new Size(680, 490), MinimumSize = new Size(480, 390) })
            {
                var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 5, Padding = new Padding(12) };
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                layout.RowStyles.Add(new RowStyle(SizeType.Percent, 55));
                layout.RowStyles.Add(new RowStyle(SizeType.Percent, 45));
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                dialog.Controls.Add(layout);
                var table = new ComboBox { Dock = DockStyle.Top, DropDownStyle = ComboBoxStyle.DropDownList };
                var cells = new ListBox { Dock = DockStyle.Fill, HorizontalScrollbar = true };
                var fields = new CheckedListBox { Dock = DockStyle.Fill, CheckOnClick = true };
                var lineTarget = new ComboBox { Width = 180, DropDownStyle = ComboBoxStyle.DropDownList };
                var groupKeys = HwpStylePresetStore.LineGroups.Keys.ToArray();
                lineTarget.Items.AddRange(groupKeys.Select(key => (object)HwpStylePresetStore.LineGroups[key]).ToArray());
                var linePanel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
                linePanel.Controls.Add(new Label { Text = "선 적용 위치", AutoSize = true, Margin = new Padding(0, 6, 8, 0) });
                linePanel.Controls.Add(lineTarget);
                var accept = new Button { Text = "선택 적용", AutoSize = true, Anchor = AnchorStyles.Right };
                layout.Controls.Add(table, 0, 0);
                layout.Controls.Add(cells, 0, 1);
                layout.Controls.Add(fields, 0, 2);
                layout.Controls.Add(linePanel, 0, 3);
                layout.Controls.Add(accept, 0, 4);
                var labels = new Dictionary<string, string> { { "font_family", "글꼴" }, { "font_size_pt", "크기(pt)" }, { "bold", "굵게" },
                    { "alignment", "가로 정렬" }, { "vertical_alignment", "세로 정렬" }, { "fill_color", "배경 RGB" } };
                labels.Add("margins_mm", "여백(상/하/좌/우 mm)");
                labels.Add("line_top", "위쪽 선");
                labels.Add("line_bottom", "아래쪽 선");
                labels.Add("line_left", "왼쪽 선");
                labels.Add("line_right", "오른쪽 선");
                var keys = new List<string>();
                Dictionary<string, object>[] anchors = null;
                Dictionary<string, object> style = null;
                cells.SelectedIndexChanged += delegate
                {
                    fields.Items.Clear();
                    keys.Clear();
                    if (cells.SelectedIndex < 0) return;
                    style = (Dictionary<string, object>)anchors[cells.SelectedIndex]["style"];
                    foreach (var item in labels.Where(item => style.ContainsKey(item.Key)))
                    {
                        keys.Add(item.Key);
                        object value = style[item.Key];
                        string detail = item.Key.StartsWith("line_", StringComparison.Ordinal) ? LineText((Dictionary<string, object>)value)
                            : item.Key == "margins_mm" ? FormatMargins((Dictionary<string, object>)value) : Convert.ToString(value);
                        fields.Items.Add(item.Value + ": " + detail);
                    }
                };
                table.SelectedIndexChanged += delegate
                {
                    cells.Items.Clear();
                    anchors = ((object[])tables[table.SelectedIndex]["cells"]).Cast<Dictionary<string, object>>().ToArray();
                    cells.Items.AddRange(anchors.Select(c => (object)string.Format("{0}행 {1}열 ({2}×{3})  {4}", c["row"], c["col"], c["rowspan"], c["colspan"], c["text"])).ToArray());
                    cells.SelectedIndex = 0;
                };
                Dictionary<string, object> result = null;
                accept.Click += delegate
                {
                    if (fields.CheckedIndices.Count == 0) { MessageBox.Show(dialog, "가져올 속성을 하나 이상 선택하세요."); return; }
                    var picked = fields.CheckedIndices.Cast<int>().Select(i => keys[i]).ToArray();
                    var lines = picked.Where(key => key.StartsWith("line_", StringComparison.Ordinal)).ToArray();
                    if (lines.Length > 1 || (lines.Length == 1 && lineTarget.SelectedIndex < 0))
                    {
                        MessageBox.Show(dialog, "기준 선 하나와 적용 위치를 선택하세요.");
                        return;
                    }
                    result = picked.Except(lines).ToDictionary(key => key, key => style[key]);
                    if (lines.Length == 1) result["edge"] = new Dictionary<string, object> { { groupKeys[lineTarget.SelectedIndex], style[lines[0]] } };
                    dialog.DialogResult = DialogResult.OK;
                };
                table.Items.AddRange(tables.Select(t => (object)string.Format("표 {0} ({1}행 × {2}열)", t["table_index"], t["rows"], t["cols"])).ToArray());
                table.SelectedIndex = 0;
                LauncherUi.ApplyToForm(dialog);
                LauncherUi.ApplyTree(dialog);
                return dialog.ShowDialog(this) == DialogResult.OK ? result : null;
            }
        }

        internal void ApplyTemplateCell(string target, Dictionary<string, object> values)
        {
            var paragraph = paragraphGrid.Rows.Cast<DataGridViewRow>().Single(row => CellText(row, "StyleName") == target);
            var cell = tableGrid.Rows.Cast<DataGridViewRow>().Single(row => CellText(row, "StyleName") == target);
            foreach (var item in new[] { "font_family:Font", "font_size_pt:Size", "bold:Bold", "alignment:Align" })
            {
                string[] mapping = item.Split(':');
                if (values.ContainsKey(mapping[0])) paragraph.Cells[mapping[1]].Value = values[mapping[0]];
            }
            if (values.ContainsKey("fill_color")) cell.Cells["Fill"].Value = values["fill_color"];
            if (values.ContainsKey("vertical_alignment")) cell.Cells["Vertical"].Value = values["vertical_alignment"];
            if (values.ContainsKey("margins_mm")) cell.Cells["Margins"].Value = FormatMargins((Dictionary<string, object>)values["margins_mm"]);
            if (values.ContainsKey("edge"))
                foreach (var line in (Dictionary<string, object>)values["edge"]) edges[line.Key] = line.Value;
            UpdateImportedLines();
            if (values.Keys.Any(key => key != "edge") && HwpStylePresetStore.RoleStyleDefaults.ContainsKey(target)) paragraph.Cells["Enabled"].Value = true;
        }

        private static string CellText(DataGridViewRow row, string name, bool allowEmpty = false)
        {
            string value = Convert.ToString(row.Cells[name].Value).Trim();
            if (!allowEmpty && value.Length == 0) throw new InvalidOperationException(name + " 값을 입력하세요.");
            return value;
        }

        private static double CellNumber(DataGridViewRow row, string name)
        {
            return HwpStylePresetStore.ParseFiniteDouble(Convert.ToString(row.Cells[name].Value), name);
        }

        private void SaveNew(object sender, EventArgs e)
        {
            ExecuteEditorAction(delegate
            {
                Dictionary<string, object> config = ReadEditor();
                string name = Convert.ToString(config["preset_name"]);
                bool exists = store.LoadAll().Any(item => string.Equals(Convert.ToString(item["preset_name"]), name, StringComparison.OrdinalIgnoreCase));
                if (exists)
                {
                    if (store.IsBuiltIn(name))
                        throw new InvalidOperationException("기본 제공 서식은 수정할 수 없습니다. 다른 이름을 입력하세요.");
                    if (MessageBox.Show(this, "같은 이름의 서식을 덮어쓸까요?", "서식 덮어쓰기", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
                }
                store.Save(config, exists);
                SelectedPresetName = name;
                ReloadPresets(name);
            });
        }

        private void Overwrite(object sender, EventArgs e)
        {
            ExecuteEditorAction(delegate
            {
                string selected = presetCombo.SelectedItem == null ? "" : presetCombo.SelectedItem.ToString();
                if (store.IsBuiltIn(selected))
                    throw new InvalidOperationException("기본 제공 서식은 수정할 수 없습니다. 다른 이름으로 저장하세요.");
                nameText.Text = selected;
                if (MessageBox.Show(this, "'" + selected + "' 서식을 덮어쓸까요?", "서식 덮어쓰기", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
                store.Save(ReadEditor(), true);
                SelectedPresetName = selected;
                ReloadPresets(selected);
            });
        }

        private void DeletePreset(object sender, EventArgs e)
        {
            ExecuteEditorAction(delegate
            {
                string selected = presetCombo.SelectedItem == null ? "" : presetCombo.SelectedItem.ToString();
                if (store.IsBuiltIn(selected))
                    throw new InvalidOperationException("기본 제공 서식은 삭제할 수 없습니다.");
                if (MessageBox.Show(this, "'" + selected + "' 서식을 삭제할까요?", "서식 삭제", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
                store.Delete(selected);
                SelectedPresetName = HwpStylePresetStore.BuiltInName;
                ReloadPresets(SelectedPresetName);
            });
        }

        private void ResetEditor(object sender, EventArgs e)
        {
            string name = nameText.Text;
            LoadPreset(HwpStylePresetStore.BuiltInName);
            if (!string.Equals(name, HwpStylePresetStore.BuiltInName, StringComparison.OrdinalIgnoreCase)) nameText.Text = name;
        }

        private void ExecuteEditorAction(Action action)
        {
            try { action(); }
            catch (Exception ex) { MessageBox.Show(this, ex.Message, "서식 설정", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }
    }

    internal sealed class MainForm : Form
    {
        private readonly HwpStylePresetStore hwpStylePresetStore;
        private readonly TextBox workbookPathText = new TextBox();
        private readonly ComboBox outputTypeCombo = new ComboBox();
        private readonly TextBox hwpTemplateText = new TextBox();
        private readonly TextBox pptTemplateText = new TextBox();
        private readonly TextBox hwpTableStyleProfileText = new TextBox();
        private readonly Button inspectTemplateButton = new Button();
        private readonly Button createHwpTemplateButton = new Button();
        private readonly Button useBasicHwpTemplateButton = new Button();
        private readonly Button createPptTemplateButton = new Button();
        private readonly Button createChartTemplateButton = new Button();
        private readonly Button autoFixTemplateButton = new Button();
        private readonly Button openTemplateGuideButton = new Button();
        private readonly Label templateStatusLabel = new Label();
        private readonly CheckBox hwpVisibleCheck = new CheckBox();
        private readonly CheckBox hwpKeepOpenOnErrorCheck = new CheckBox();
        private readonly ComboBox hwpMaxSectionsCombo = new ComboBox();
        private readonly ComboBox hwpDispatchModeCombo = new ComboBox();
        private readonly ComboBox hwpStylePresetCombo = new ComboBox();
        private readonly Button hwpStyleSettingsButton = new Button();
        private readonly Button hwpPageSetupButton = new Button();
        private Dictionary<string, object> hwpPageSetup;
        private readonly Button hwpEnvironmentCheckButton = new Button();
        private readonly TextBox bannerText = new TextBox();
        private readonly CheckedListBox bannerList = new CheckedListBox();
        private readonly TabControl workflowTabs = new TabControl();
        private readonly ListView tablePreviewList = new ListView();
        private readonly Button selectTableRangeButton = new Button();
        private readonly Button clearTableRangeButton = new Button();
        private readonly Label dataStatusLabel = new Label();
        private readonly Label workspaceStatus = new Label();
        private readonly ToolTip actionTips = new ToolTip();
        private readonly Button reloadBannerButton = new Button();
        private readonly Button recommendedBannerButton = new Button();
        private readonly Button selectAllBannerButton = new Button();
        private readonly Button clearBannerButton = new Button();
        private readonly Button moveBannerUpButton = new Button();
        private readonly Button moveBannerDownButton = new Button();
        private readonly Button deleteBannerButton = new Button();
        private readonly Label bannerStatusLabel = new Label();
        private readonly TextBox titlePrefixesText = new TextBox();
        private readonly ComboBox reportProfileCombo = new ComboBox();
        private readonly ComboBox styleProfileCombo = new ComboBox();
        private readonly TextBox narrativeExampleText = new TextBox();
        private readonly ComboBox tableExampleCombo = new ComboBox();
        private readonly DataGridView tableExampleGrid = new DataGridView();
        private readonly TextBox tableExampleOutput = new TextBox();
        private readonly CheckBox analysisCheck = new CheckBox();
        private readonly CheckBox chartCheck = new CheckBox();
        private readonly CheckBox tableCheck = new CheckBox();
        private readonly CheckBox qaCheck = new CheckBox();
        private readonly CheckBox draftTextCheck = new CheckBox();
        private readonly ComboBox decimalPlacesCombo = new ComboBox();
        private readonly ComboBox chartOutputCombo = new ComboBox();
        private readonly ComboBox tableInsertModeCombo = new ComboBox();
        private readonly CheckBox llmEnabledCheck = new CheckBox();
        private readonly ComboBox llmProviderCombo = new ComboBox();
        private readonly TextBox llmModelText = new TextBox();
        private readonly TextBox llmApiKeyText = new TextBox();
        private readonly CheckBox copyWorkbookCheck = new CheckBox();
        private readonly CheckBox keepExcelOpenCheck = new CheckBox();
        private readonly Button runButton = new Button();
        private bool reportRunning;
        private readonly Button openWorkbookButton = new Button();
        private readonly Button openDraftButton = new Button();
        private readonly Button openHwpOutputButton = new Button();
        private readonly Button openHwpReportButton = new Button();
        private readonly Button openPreflightButton = new Button();
        private readonly Button copyDraftButton = new Button();
        private readonly TextBox logText = new TextBox();
        private readonly TextBox resultSummaryText = new TextBox();
        private readonly ListView readinessList = new ListView();
        private readonly TextBox draftPreviewText = new TextBox();
        private readonly TabControl draftReviewTabs = new TabControl();
        private readonly ListView sentenceReviewList = new ListView();
        private readonly TextBox sentenceEditText = new TextBox();
        private readonly Button applySentenceEditButton = new Button();
        private readonly Button copySelectedSentenceButton = new Button();
        private readonly Button exportReviewedDraftButton = new Button();
        private readonly ComboBox qaFilterCombo = new ComboBox();
        private readonly ListView qaIssueList = new ListView();
        private readonly Label draftPreviewStatusLabel = new Label();
        private readonly TextBox dashboardWorkbookText = new TextBox();
        private readonly Button dashboardBrowseButton = new Button();
        private readonly Button dashboardInspectButton = new Button();
        private readonly ComboBox dashboardSheetCombo = new ComboBox();
        private readonly ComboBox dashboardEntityColumnCombo = new ComboBox();
        private readonly ComboBox dashboardOutputModeCombo = new ComboBox();
        private readonly ComboBox dashboardPageSizeCombo = new ComboBox();
        private readonly ComboBox dashboardDesignCombo = new ComboBox();
        private readonly ComboBox dashboardFontCombo = new ComboBox();
        private readonly TextBox dashboardTemplateText = new TextBox();
        private readonly Button dashboardTemplateBrowseButton = new Button();
        private readonly CheckedListBox dashboardEntityList = new CheckedListBox();
        private readonly CheckedListBox dashboardColumnList = new CheckedListBox();
        private readonly ListView dashboardColumnPreviewList = new ListView();
        private readonly TextBox dashboardNarrativeTemplateText = new TextBox();
        private readonly TextBox dashboardStatusText = new TextBox();
        private readonly Button dashboardSelectAllEntitiesButton = new Button();
        private readonly Button dashboardClearEntitiesButton = new Button();
        private readonly Button dashboardSelectAllColumnsButton = new Button();
        private readonly Button dashboardClearColumnsButton = new Button();
        private readonly Button dashboardGenerateButton = new Button();
        private readonly Button dashboardOpenOutputButton = new Button();
        private readonly ComboBox[] dashboardKpiColumnCombos = new ComboBox[6];
        private readonly TextBox[] dashboardKpiLabelTexts = new TextBox[6];
        private readonly TextBox[] dashboardKpiUnitTexts = new TextBox[6];
        private readonly ComboBox[] dashboardChartTypeCombos = new ComboBox[4];
        private readonly TextBox[] dashboardChartTitleTexts = new TextBox[4];
        private readonly TextBox[] dashboardChartColumnsTexts = new TextBox[4];
        private readonly TextBox[] dashboardChartLabelsTexts = new TextBox[4];
        private readonly List<DraftSentenceItem> draftSentenceItems = new List<DraftSentenceItem>();
        private readonly List<DraftQaIssue> draftQaIssues = new List<DraftQaIssue>();
        private readonly HashSet<string> recommendedBanners = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private string lastTemplateStatus = "미검사";
        private string currentDraftPath = "";
        private DashboardWorkbookInfo currentDashboardInfo;
        private string lastDashboardOutputPath = "";
        private readonly List<TablePreview> currentTablePreviews = new List<TablePreview>();

        public MainForm()
        {
            hwpStylePresetStore = new HwpStylePresetStore(recoveryLog: Log);
            Text = "Research Helper | 보고서 자동화";
            MinimumSize = new Size(980, 720);
            Size = new Size(1120, 820);
            StartPosition = FormStartPosition.CenterScreen;
            LauncherUi.ApplyToForm(this);

            var root = new TableLayoutPanel();
            root.Dock = DockStyle.Fill;
            root.Padding = new Padding(LauncherUi.SpaceXl);
            root.ColumnCount = 1;
            root.RowCount = 3;
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(root);

            root.Controls.Add(BuildHeaderPanel(), 0, 0);

            workflowTabs.Dock = DockStyle.Fill;
            workflowTabs.DrawMode = TabDrawMode.OwnerDrawFixed;
            workflowTabs.SizeMode = TabSizeMode.Fixed;
            workflowTabs.ItemSize = new Size(146, 37);
            workflowTabs.DrawItem += WorkflowTabs_DrawItem;
            workflowTabs.TabPages.Add(CreateStepPage("1  원본 파일", BuildFileGroup()));
            workflowTabs.TabPages.Add(CreateStepPage("2  표 · 배너", BuildDataReviewGroup()));
            workflowTabs.TabPages.Add(CreateStepPage("3  보고서 설정", BuildRulesPage()));
            workflowTabs.TabPages.Add(CreateStepPage("4  산출 결과", BuildRunGroup()));
            workflowTabs.TabPages.Add(CreateStepPage("기관 대시보드 PPT", BuildDashboardPage()));
            workflowTabs.SelectedIndexChanged += delegate { UpdateWorkflowStatus(); };
            root.Controls.Add(workflowTabs, 0, 1);

            var footer = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, AutoSize = true, ColumnCount = 2,
                Margin = new Padding(0, LauncherUi.SpaceSm, 0, 0)
            };
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            workspaceStatus.Dock = DockStyle.Fill;
            workspaceStatus.AutoEllipsis = true;
            workspaceStatus.TextAlign = ContentAlignment.MiddleLeft;
            workspaceStatus.ForeColor = LauncherUi.ColorMutedText;
            workspaceStatus.AccessibleName = "현재 보고서 작업 상태";
            runButton.AutoSize = true;
            runButton.MinimumSize = new Size(176, 34);
            runButton.Tag = LauncherButtonKind.Primary;
            runButton.Click += RunButton_Click;
            footer.Controls.Add(workspaceStatus, 0, 0);
            footer.Controls.Add(runButton, 1, 0);
            root.Controls.Add(footer, 0, 2);
            FormClosing += delegate(object sender, FormClosingEventArgs args)
            {
                if (reportRunning && args.CloseReason == CloseReason.UserClosing) args.Cancel = true;
            };
            Disposed += delegate { actionTips.Dispose(); };

            workbookPathText.TextChanged += delegate { UpdateWorkflowStatus(); };
            outputTypeCombo.SelectedIndexChanged += delegate { UpdateWorkflowStatus(); };
            hwpTemplateText.TextChanged += delegate { UpdateWorkflowStatus(); };
            pptTemplateText.TextChanged += delegate { UpdateWorkflowStatus(); };

            bannerText.Text = "전체";
            titlePrefixesText.Text = "";
            outputTypeCombo.SelectedIndex = 0;
            reportProfileCombo.SelectedIndex = 0;
            styleProfileCombo.SelectedIndex = 0;
            decimalPlacesCombo.SelectedIndex = 1;
            chartOutputCombo.SelectedIndex = 0;
            tableInsertModeCombo.SelectedIndex = 0;
            llmProviderCombo.SelectedIndex = 0;
            llmModelText.Text = "gpt-4.1-mini";
            hwpMaxSectionsCombo.SelectedIndex = 0;
            hwpDispatchModeCombo.SelectedIndex = 0;
            RefreshHwpStylePresetCombo(HwpStylePresetStore.BuiltInName);
            dashboardOutputModeCombo.SelectedIndex = 0;
            dashboardPageSizeCombo.SelectedIndex = 0;
            dashboardDesignCombo.SelectedIndex = 0;
            dashboardFontCombo.SelectedIndex = 0;
            analysisCheck.Checked = true;
            chartCheck.Checked = true;
            tableCheck.Checked = true;
            qaCheck.Checked = true;
            draftTextCheck.Checked = true;
            copyWorkbookCheck.Checked = true;
            keepExcelOpenCheck.Checked = true;
            LauncherUi.ApplyTree(this);
            UpdateWorkflowStatus();
            UpdateReadinessChecklist();
        }

        internal void CheckWorkflowLayout()
        {
            Size = MinimumSize;
            if (runButton.Enabled || runButton.Text != "Excel 산출 시트 만들기")
                throw new InvalidOperationException("원본 파일이 없는 Excel 산출 상태가 올바르지 않습니다.");

            // Existing paths stand in for inputs; no Excel or Hancom session is opened.
            string bundledAddin = PathResolver.ResolveDefaultAddinPath();
            workbookPathText.Text = bundledAddin;
            if (!runButton.Enabled || ReadOptions().AddinPath != bundledAddin)
                throw new InvalidOperationException("원본 파일만 선택한 Excel 산출에 내장 추가기능이 자동 지정되지 않았습니다.");
            if (workflowTabs.TabPages[0].Controls[0].Controls.Count != 1 ||
                readinessList.Items.Cast<ListViewItem>().Any(item => item.Text == "추가기능"))
                throw new InvalidOperationException("Excel 추가기능 수동 설정이 일반 작업 화면에 남아 있습니다.");
            outputTypeCombo.SelectedIndex = 1;
            if (runButton.Enabled || runButton.Text != "한글 보고서 만들기")
                throw new InvalidOperationException("양식 없는 한글 보고서의 준비 상태가 올바르지 않습니다.");
            hwpTemplateText.Text = bundledAddin;
            if (!runButton.Enabled)
                throw new InvalidOperationException("입력이 준비된 한글 보고서의 실행 버튼이 비활성 상태입니다.");

            var examples = new HashSet<string>();
            for (int index = 0; index < reportProfileCombo.Items.Count; index++)
            {
                reportProfileCombo.SelectedIndex = index;
                for (int style = 0; style < styleProfileCombo.Items.Count; style++)
                {
                    styleProfileCombo.SelectedIndex = style;
                    if (narrativeExampleText.Lines.Length < 2 || !examples.Add(narrativeExampleText.Text))
                        throw new InvalidOperationException("보고서 유형·문체별 분석문 예시가 올바르지 않습니다.");
                }
            }
            var expected = new[] { "52.0%로 가장 높게", "복수응답 기준", "52.0%로 가장 높게", "72.6점", "78.4점" };
            for (int index = 0; index < expected.Length; index++)
            {
                tableExampleCombo.SelectedIndex = index;
                if (tableExampleGrid.Rows.Count != 3 || tableExampleOutput.Lines.Length != 2 ||
                    !tableExampleOutput.Text.Contains(expected[index]))
                    throw new InvalidOperationException("집계표 유형별 실제 산출 예시가 올바르지 않습니다.");
            }
            tableExampleCombo.SelectedIndex = 2;
            if (tableExampleGrid.Rows[0].Cells[3].Style.BackColor != LauncherUi.ColorSurfaceStrong ||
                tableExampleGrid.Rows[1].Cells[3].Style.BackColor == LauncherUi.ColorSurfaceStrong ||
                tableExampleGrid.Rows[0].Cells[2].Style.BackColor == LauncherUi.ColorSurfaceStrong)
                throw new InvalidOperationException("교차표의 전체 비율 셀만 표시되어야 합니다.");

            for (int index = 0; index < workflowTabs.TabCount; index++)
            {
                workflowTabs.SelectedIndex = index;
                PerformLayout();
                if (index == 1 && selectTableRangeButton.Parent == bannerList.Parent)
                    throw new InvalidOperationException("표 범위와 배너 도구가 같은 작업 영역에 있습니다.");
                if (index == 4 && runButton.Visible)
                    throw new InvalidOperationException("대시보드에 한글 보고서 버튼이 표시됩니다.");
            }
            foreach (var button in new[] { runButton, selectTableRangeButton, dashboardBrowseButton, dashboardInspectButton, dashboardTemplateBrowseButton })
            {
                if (button.Width < TextRenderer.MeasureText(button.Text, button.Font).Width)
                    throw new InvalidOperationException("최소 창 크기에서 버튼명이 잘립니다: " + button.Text);
            }
        }

        private static TabPage CreateStepPage(string title, Control content)
        {
            var page = new TabPage(title);
            page.Padding = new Padding(LauncherUi.SpaceMd);
            page.BackColor = LauncherUi.ColorBackground;
            page.ForeColor = LauncherUi.ColorText;
            content.Dock = DockStyle.Fill;
            page.Controls.Add(content);
            return page;
        }

        private Control BuildHeaderPanel()
        {
            var panel = new TableLayoutPanel();
            panel.Dock = DockStyle.Top;
            panel.AutoSize = true;
            panel.Margin = new Padding(0, 0, 0, LauncherUi.SpaceSm);
            panel.ColumnCount = 2;
            panel.RowCount = 1;
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            var title = new Label();
            title.Text = "보고서 자동화";
            title.Font = LauncherUi.TitleFont();
            title.AutoSize = true;
            title.ForeColor = LauncherUi.ColorText;
            title.Margin = new Padding(0, 0, 0, LauncherUi.SpaceXs);
            panel.Controls.Add(title, 0, 0);
            panel.Controls.Add(new Label
            {
                Text = "Research Helper", AutoSize = true, Anchor = AnchorStyles.Right,
                ForeColor = LauncherUi.ColorMutedText, Margin = new Padding(8, 0, 0, 0)
            }, 1, 0);
            return panel;
        }

        private void WorkflowTabs_DrawItem(object sender, DrawItemEventArgs e)
        {
            TabPage page = workflowTabs.TabPages[e.Index];
            bool selected = e.Index == workflowTabs.SelectedIndex;
            Rectangle bounds = e.Bounds;
            Color backColor = selected ? LauncherUi.ColorSurface : LauncherUi.ColorSurfaceAlt;
            Color foreColor = selected ? LauncherUi.ColorPrimary : LauncherUi.ColorMutedText;

            using (SolidBrush backBrush = new SolidBrush(backColor))
            {
                e.Graphics.FillRectangle(backBrush, bounds);
                if (selected)
                {
                    using (Pen accent = new Pen(LauncherUi.ColorPrimary, 3))
                    {
                        e.Graphics.DrawLine(accent, bounds.Left + 12, bounds.Bottom - 3, bounds.Right - 12, bounds.Bottom - 3);
                    }
                }
                TextRenderer.DrawText(e.Graphics, page.Text, workflowTabs.Font, bounds, foreColor,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
                if ((e.State & DrawItemState.Focus) != 0) e.DrawFocusRectangle();
            }
        }

        private void UpdateWorkflowStatus()
        {
            bool hasWorkbook = File.Exists(workbookPathText.Text.Trim());
            int tableCount = tablePreviewList.Items.Count;
            int checkedBanners = CountCheckedBanners();

            bool dashboard = workflowTabs.SelectedIndex == 4;
            bool hwp = outputTypeCombo.Text.IndexOf("HWP", StringComparison.OrdinalIgnoreCase) >= 0;
            bool needsTemplate = hwp && !File.Exists(hwpTemplateText.Text.Trim());
            runButton.Visible = !dashboard;
            runButton.Text = reportRunning ? "보고서 산출 중…" : hwp ? "한글 보고서 만들기" : "Excel 산출 시트 만들기";
            runButton.AccessibleName = runButton.Text;
            runButton.Enabled = !reportRunning && hasWorkbook && !needsTemplate;
            workspaceStatus.Text = dashboard ? "기관 대시보드 · PowerPoint" :
                reportRunning ? "산출 중 · 진행 기록은 산출 결과에서 확인" :
                !hasWorkbook ? "원본 Excel 미선택" :
                needsTemplate ? "한글 양식 미선택 · 보고서 설정" :
                Path.GetFileName(workbookPathText.Text.Trim()) + "  ·  표 " + tableCount + "개  ·  " +
                    (checkedBanners == 0 ? "전체 기준" : "배너 " + checkedBanners + "개");
            UpdateReadinessChecklist();
        }

        private void UpdateReadinessChecklist()
        {
            if (readinessList.Columns.Count == 0)
            {
                return;
            }

            bool workbookReady = File.Exists(workbookPathText.Text.Trim());
            int tableCount = tablePreviewList.Items.Count;
            int checkedBanners = CountCheckedBanners();
            bool bannerPreviewReady = bannerList.Items.Count > 0;
            bool outputReady = IsOutputReadyForSelectedOutput();

            readinessList.BeginUpdate();
            try
            {
                readinessList.Items.Clear();
                AddReadinessItem("집계표", workbookReady, workbookReady ? Path.GetFileName(workbookPathText.Text.Trim()) : "집계표 엑셀 파일을 선택하세요.");
                AddReadinessItem("표 탐지", tableCount > 0, tableCount > 0 ? tableCount + "개 표를 발견했습니다." : "파일 등록 후 표 목록을 확인하세요.");
                AddReadinessItem("분석 배너", bannerPreviewReady, BuildBannerReadinessText(checkedBanners, bannerPreviewReady));
                AddReadinessItem("산출 방식", outputReady, BuildOutputReadinessText(outputReady));
                AddReadinessItem("템플릿", IsTemplateReadyForSelectedOutput(), BuildTemplateReadinessText());
            }
            finally
            {
                readinessList.EndUpdate();
            }
        }

        private void AddReadinessItem(string name, bool ready, string detail)
        {
            var item = new ListViewItem(name);
            item.SubItems.Add(ready ? "완료" : "확인 필요");
            item.SubItems.Add(detail);
            LauncherUi.StyleListItem(item, readinessList.Items.Count);
            item.UseItemStyleForSubItems = false;
            item.SubItems[1].ForeColor = ready ? LauncherUi.ColorSuccess : LauncherUi.ColorWarning;
            readinessList.Items.Add(item);
        }

        private string BuildBannerReadinessText(int checkedBanners, bool bannerPreviewReady)
        {
            if (!bannerPreviewReady)
            {
                return "집계표를 읽은 뒤 분석에 사용할 가로배너를 선택하세요.";
            }
            if (checkedBanners == 0)
            {
                return "선택 배너가 없어 전체 기준으로 실행합니다.";
            }
            return checkedBanners + "개 배너를 분석 대상으로 사용합니다.";
        }

        private int CountCheckedBanners()
        {
            int count = 0;
            for (int i = 0; i < bannerList.Items.Count; i++)
            {
                if (bannerList.GetItemChecked(i))
                {
                    count++;
                }
            }
            return count;
        }

        private bool IsOutputReadyForSelectedOutput()
        {
            string output = outputTypeCombo.Text;
            if (output.StartsWith("Excel", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            if (output.IndexOf("HWP", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }
            return false;
        }

        private string BuildOutputReadinessText(bool ready)
        {
            string output = outputTypeCombo.Text;
            if (output.StartsWith("Excel", StringComparison.OrdinalIgnoreCase))
            {
                return "Excel 산출 시트 기준으로 실행합니다.";
            }
            if (output.IndexOf("HWP", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "Excel 산출 후 아래한글 COM으로 HWPX 초본을 생성합니다.";
            }
            return ready ? output + " 기준으로 실행합니다." : "현재 런처 직접 생성은 Excel/HWPX만 지원합니다.";
        }

        private bool IsTemplateReadyForSelectedOutput()
        {
            string output = outputTypeCombo.Text;
            if (output.StartsWith("Excel", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            if (output.IndexOf("HWP", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return File.Exists(hwpTemplateText.Text.Trim()) && IsTemplateStatusUsable();
            }
            if (output.IndexOf("PowerPoint", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return File.Exists(pptTemplateText.Text.Trim()) && IsTemplateStatusUsable();
            }
            return false;
        }

        private bool IsTemplateStatusUsable()
        {
            return lastTemplateStatus.IndexOf("ready", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   lastTemplateStatus.IndexOf("usable_with_warnings", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private string BuildTemplateReadinessText()
        {
            string output = outputTypeCombo.Text;
            if (output.StartsWith("Excel", StringComparison.OrdinalIgnoreCase))
            {
                return "Excel 산출은 외부 문서 템플릿이 필요하지 않습니다.";
            }
            if (output.IndexOf("HWP", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return File.Exists(hwpTemplateText.Text.Trim()) ? "HWPX/HWP 템플릿 상태: " + lastTemplateStatus : "HWPX/HWP 템플릿을 선택하거나 기본 템플릿을 생성하세요.";
            }
            if (output.IndexOf("PowerPoint", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return File.Exists(pptTemplateText.Text.Trim()) ? "PPTX 템플릿 상태: " + lastTemplateStatus : "PPTX 템플릿을 선택하거나 기본 템플릿을 생성하세요.";
            }
            return "템플릿 검사가 필요합니다.";
        }

        private Control BuildFileGroup()
        {
            var page = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, AutoScroll = true };
            var source = new GroupBox { Text = "원본 집계표", Dock = DockStyle.Top, AutoSize = true };
            var input = CreateGrid(2);
            source.Controls.Add(input);
            AddPathRow(input, 0, "Excel 집계표", workbookPathText, "Excel 파일 선택…", BrowseWorkbook);
            var checks = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            copyWorkbookCheck.Text = "작업 복사본에 저장";
            keepExcelOpenCheck.Text = "산출 후 Excel 열어두기";
            checks.Controls.Add(copyWorkbookCheck);
            checks.Controls.Add(keepExcelOpenCheck);
            AddLabel(input, 1, "산출 파일");
            input.Controls.Add(checks, 1, 1);
            input.SetColumnSpan(checks, 2);
            page.Controls.Add(source, 0, 0);

            return page;
        }

        private Control BuildOutputGroup()
        {
            var group = new GroupBox();
            group.Text = "출력 파일";
            group.Dock = DockStyle.Top;
            group.AutoSize = true;
            group.Padding = new Padding(10);

            var grid = CreateGrid(9);
            group.Controls.Add(grid);

            AddLabel(grid, 0, "출력 형식");
            outputTypeCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            outputTypeCombo.Items.Add("Excel 산출 시트");
            outputTypeCombo.Items.Add("HWPX 보고서");
            outputTypeCombo.Items.Add("PowerPoint 보고서");
            outputTypeCombo.Dock = DockStyle.Fill;
            grid.Controls.Add(outputTypeCombo, 1, 0);
            grid.SetColumnSpan(outputTypeCombo, 2);

            AddPathRow(grid, 1, "한글 보고서 양식", hwpTemplateText, "한글 양식 선택…", BrowseHwpTemplate);
            AddPathRow(grid, 2, "PowerPoint 양식", pptTemplateText, "PPTX 양식 선택…", BrowsePptTemplate);
            AddPathRow(grid, 3, "표 구조 프로필 (고급)", hwpTableStyleProfileText, "JSON 프로필 선택…", BrowseHwpTableStyleProfile);

            var templateButtons = new FlowLayoutPanel();
            templateButtons.Dock = DockStyle.Fill;
            templateButtons.AutoSize = true;
            inspectTemplateButton.Text = "양식 검사";
            inspectTemplateButton.Click += InspectTemplateButton_Click;
            useBasicHwpTemplateButton.Text = "기본 한글 양식 사용";
            useBasicHwpTemplateButton.Click += UseBasicHwpTemplateButton_Click;
            createHwpTemplateButton.Text = "편집용 양식 사본 저장…";
            createHwpTemplateButton.Click += CreateHwpTemplateButton_Click;
            createPptTemplateButton.Text = "PPTX 기본 양식 저장…";
            createPptTemplateButton.Click += CreatePptTemplateButton_Click;
            createChartTemplateButton.Text = "차트 검토 양식 저장…";
            createChartTemplateButton.Click += CreateChartTemplateButton_Click;
            autoFixTemplateButton.Text = "보정 사본 저장…";
            autoFixTemplateButton.Click += AutoFixTemplateButton_Click;
            openTemplateGuideButton.Text = "양식 가이드 열기";
            openTemplateGuideButton.Click += OpenTemplateGuideButton_Click;
            templateButtons.Controls.Add(useBasicHwpTemplateButton);
            templateButtons.Controls.Add(createHwpTemplateButton);
            templateButtons.Controls.Add(inspectTemplateButton);
            templateButtons.Controls.Add(autoFixTemplateButton);
            templateButtons.Controls.Add(createPptTemplateButton);
            templateButtons.Controls.Add(createChartTemplateButton);
            templateButtons.Controls.Add(openTemplateGuideButton);
            foreach (Button button in templateButtons.Controls) button.AutoSize = true;
            AddLabel(grid, 4, "양식 도구");
            grid.Controls.Add(templateButtons, 1, 4);
            grid.SetColumnSpan(templateButtons, 2);

            templateStatusLabel.Text = "템플릿을 선택하거나 기본 템플릿을 생성하세요.";
            templateStatusLabel.AutoSize = true;
            templateStatusLabel.ForeColor = LauncherUi.ColorMutedText;
            templateStatusLabel.Margin = new Padding(0, 4, 0, 0);
            grid.Controls.Add(templateStatusLabel, 1, 5);
            grid.SetColumnSpan(templateStatusLabel, 2);

            var hwpOptions = new FlowLayoutPanel();
            hwpOptions.Dock = DockStyle.Fill;
            hwpOptions.AutoSize = true;
            hwpVisibleCheck.Text = "아래한글 창 표시";
            hwpKeepOpenOnErrorCheck.Text = "실패 시 문서 유지";
            var hwpLimitLabel = new Label();
            hwpLimitLabel.Text = "출력 문항 수";
            hwpLimitLabel.AutoSize = true;
            hwpLimitLabel.Margin = new Padding(12, 7, 6, 0);
            hwpMaxSectionsCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            hwpMaxSectionsCombo.Width = 120;
            hwpMaxSectionsCombo.Items.Add("첫 문항만");
            hwpMaxSectionsCombo.Items.Add("첫 3개 문항");
            hwpMaxSectionsCombo.Items.Add("모든 문항");
            var hwpDispatchLabel = new Label();
            hwpDispatchLabel.Text = "한글 연결 방식";
            hwpDispatchLabel.AutoSize = true;
            hwpDispatchLabel.Margin = new Padding(12, 7, 6, 0);
            hwpDispatchModeCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            hwpDispatchModeCombo.Width = 135;
            hwpDispatchModeCombo.Items.Add("ensure_dispatch");
            hwpDispatchModeCombo.Items.Add("dispatch");
            hwpDispatchModeCombo.Items.Add("dispatch_ex");
            hwpEnvironmentCheckButton.Text = "한글 연결 확인";
            hwpEnvironmentCheckButton.Click += HwpEnvironmentCheckButton_Click;
            var hwpStyleLabel = new Label();
            hwpStyleLabel.Text = "본문·표 서식";
            hwpStyleLabel.AutoSize = true;
            hwpStyleLabel.Margin = new Padding(12, 7, 6, 0);
            hwpStylePresetCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            hwpStylePresetCombo.Width = 182;
            hwpStyleSettingsButton.Text = "내 서식 추가·편집…";
            hwpStyleSettingsButton.Click += HwpStyleSettingsButton_Click;
            hwpPageSetupButton.Text = "용지·여백 설정…";
            hwpPageSetupButton.Click += delegate
            {
                using (var dialog = new HwpPageSetupDialog(hwpPageSetup))
                {
                    if (dialog.ShowDialog(this) == DialogResult.OK) hwpPageSetup = dialog.SelectedConfig;
                }
            };
            hwpStyleSettingsButton.AutoSize = true;
            hwpPageSetupButton.AutoSize = true;
            hwpEnvironmentCheckButton.AutoSize = true;
            hwpOptions.Controls.Add(hwpStyleLabel);
            hwpOptions.Controls.Add(hwpStylePresetCombo);
            hwpOptions.Controls.Add(hwpStyleSettingsButton);
            hwpOptions.Controls.Add(hwpPageSetupButton);
            hwpOptions.Controls.Add(hwpLimitLabel);
            hwpOptions.Controls.Add(hwpMaxSectionsCombo);
            AddLabel(grid, 6, "한글 서식");
            grid.Controls.Add(hwpOptions, 1, 6);
            grid.SetColumnSpan(hwpOptions, 2);

            var components = new FlowLayoutPanel();
            components.Dock = DockStyle.Fill;
            components.AutoSize = true;
            analysisCheck.Text = "본문 분석문";
            chartCheck.Text = "차트 데이터";
            tableCheck.Text = "보고서 삽입표";
            qaCheck.Text = "검토·출처·수정 기록";
            draftTextCheck.Text = "문장 초안 텍스트 파일";
            components.Controls.Add(analysisCheck);
            components.Controls.Add(chartCheck);
            components.Controls.Add(tableCheck);
            components.Controls.Add(qaCheck);
            components.Controls.Add(draftTextCheck);
            AddLabel(grid, 7, "포함 항목");
            grid.Controls.Add(components, 1, 7);
            grid.SetColumnSpan(components, 2);

            var connection = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            connection.Controls.Add(hwpDispatchLabel);
            connection.Controls.Add(hwpDispatchModeCombo);
            connection.Controls.Add(hwpEnvironmentCheckButton);
            connection.Controls.Add(hwpVisibleCheck);
            connection.Controls.Add(hwpKeepOpenOnErrorCheck);
            AddLabel(grid, 8, "한글 실행 옵션");
            grid.Controls.Add(connection, 1, 8);
            grid.SetColumnSpan(connection, 2);

            outputTypeCombo.SelectedIndexChanged += delegate
            {
                bool hwp = outputTypeCombo.SelectedIndex == 1;
                bool ppt = outputTypeCombo.SelectedIndex == 2;
                grid.SuspendLayout();
                foreach (Control control in grid.Controls)
                {
                    int row = grid.GetRow(control);
                    control.Visible = row == 0 || row == 7 ||
                        (hwp && (row == 1 || row == 3 || row == 6 || row == 8)) ||
                        (ppt && row == 2) || ((hwp || ppt) && (row == 4 || row == 5));
                }
                useBasicHwpTemplateButton.Visible = createHwpTemplateButton.Visible = hwp;
                createPptTemplateButton.Visible = ppt;
                createChartTemplateButton.Visible = ppt;
                grid.ResumeLayout(true);
            };
            return group;
        }

        private Control BuildDataReviewGroup()
        {
            var panel = new TableLayoutPanel();
            panel.Dock = DockStyle.Fill;
            panel.ColumnCount = 2;
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 71));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 29));
            panel.RowCount = 3;
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            dataStatusLabel.Text = "원본 Excel 미선택";
            dataStatusLabel.AutoSize = true;
            dataStatusLabel.ForeColor = LauncherUi.ColorMutedText;
            dataStatusLabel.Margin = new Padding(0, 2, 0, 6);
            panel.Controls.Add(dataStatusLabel, 0, 0);

            tablePreviewList.Dock = DockStyle.Fill;
            tablePreviewList.View = View.Details;
            tablePreviewList.FullRowSelect = true;
            tablePreviewList.GridLines = true;
            tablePreviewList.Columns.Add("순번", 48);
            tablePreviewList.Columns.Add("표 번호", 65);
            tablePreviewList.Columns.Add("표 제목", 248);
            tablePreviewList.Columns.Add("시트", 90);
            tablePreviewList.Columns.Add("자동 범위", 92);
            tablePreviewList.Columns.Add("사용 범위", 110);
            tablePreviewList.Columns.Add("크기", 70);
            tablePreviewList.Columns.Add("보정", 60);
            tablePreviewList.Columns.Add("삽입", 60);
            tablePreviewList.Columns.Add("인식 확인", 180);
            tablePreviewList.Columns[5].DisplayIndex = 3;
            tablePreviewList.AccessibleName = "감지한 통계표 목록";
            panel.Controls.Add(tablePreviewList, 0, 2);

            var buttons = new FlowLayoutPanel();
            buttons.Dock = DockStyle.Fill;
            buttons.AutoSize = true;
            selectTableRangeButton.Text = "Excel에서 표 범위 지정…";
            selectTableRangeButton.AutoSize = true;
            selectTableRangeButton.Click += delegate { SelectTableRangeInExcel(); };
            clearTableRangeButton.Text = "감지 범위 복원";
            clearTableRangeButton.AutoSize = true;
            clearTableRangeButton.Click += delegate { ClearSelectedTableRange(); };
            reloadBannerButton.Text = "표·배너 다시 읽기";
            reloadBannerButton.AutoSize = true;
            reloadBannerButton.Click += delegate { LoadBannerPreviewAsync(); };
            buttons.Controls.Add(reloadBannerButton);
            buttons.Controls.Add(selectTableRangeButton);
            buttons.Controls.Add(clearTableRangeButton);
            panel.Controls.Add(buttons, 0, 1);

            var banners = BuildBannerPanel();
            banners.Margin = new Padding(12, 0, 0, 0);
            panel.Controls.Add(banners, 1, 0);
            panel.SetRowSpan(banners, 3);
            return panel;
        }

        private Control BuildRulesPage()
        {
            var tabs = new TabControl { Dock = DockStyle.Fill };
            foreach (var page in new[]
            {
                CreateStepPage("파일 · 서식", BuildOutputGroup()),
                CreateStepPage("분석문", BuildOptionsGroup()),
                CreateStepPage("수치 · 고급", BuildAdvancedOptionsGroup())
            })
            {
                page.AutoScroll = true;
                page.Controls[0].Dock = DockStyle.Top;
                tabs.TabPages.Add(page);
            }
            tabs.TabPages.Insert(2, CreateStepPage("집계표별 산출 예시", BuildTableExamplePage()));
            return tabs;
        }

        private Control BuildAdvancedOptionsGroup()
        {
            var group = new GroupBox();
            group.Text = "수치·산출 옵션";
            group.Dock = DockStyle.Top;
            group.AutoSize = true;
            group.Padding = new Padding(10);

            var grid = CreateGrid(6);
            group.Controls.Add(grid);

            AddLabel(grid, 0, "수치 표기");
            decimalPlacesCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            decimalPlacesCombo.Items.Add("소수점 없음");
            decimalPlacesCombo.Items.Add("소수점 1자리");
            decimalPlacesCombo.Items.Add("소수점 2자리");
            decimalPlacesCombo.Dock = DockStyle.Fill;
            grid.Controls.Add(decimalPlacesCombo, 1, 0);
            grid.SetColumnSpan(decimalPlacesCombo, 2);

            AddLabel(grid, 1, "차트 출력");
            chartOutputCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            chartOutputCombo.Items.Add("Excel 차트 데이터");
            chartOutputCombo.Items.Add("HWP 메타파일 붙여넣기용");
            chartOutputCombo.Items.Add("PPTX 차트 객체");
            chartOutputCombo.Dock = DockStyle.Fill;
            grid.Controls.Add(chartOutputCombo, 1, 1);
            grid.SetColumnSpan(chartOutputCombo, 2);

            AddLabel(grid, 2, "삽입표 방식");
            tableInsertModeCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            tableInsertModeCombo.Items.Add("Excel 삽입표 시트");
            tableInsertModeCombo.Items.Add("HWP 표 객체");
            tableInsertModeCombo.Items.Add("선택 붙여넣기");
            tableInsertModeCombo.Dock = DockStyle.Fill;
            grid.Controls.Add(tableInsertModeCombo, 1, 2);
            grid.SetColumnSpan(tableInsertModeCombo, 2);

            llmEnabledCheck.Text = "AI로 문장 다듬기";
            llmEnabledCheck.AutoSize = true;
            AddLabel(grid, 3, "문장 보조");
            grid.Controls.Add(llmEnabledCheck, 1, 3);
            grid.SetColumnSpan(llmEnabledCheck, 2);

            var llmPanel = new FlowLayoutPanel();
            llmPanel.Dock = DockStyle.Fill;
            llmPanel.AutoSize = true;
            llmProviderCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            llmProviderCombo.Width = 130;
            llmProviderCombo.Items.Add("OpenAI");
            llmProviderCombo.Items.Add("Claude");
            llmProviderCombo.Items.Add("사용자 지정");
            llmModelText.Width = 220;
            llmPanel.Controls.Add(llmProviderCombo);
            llmPanel.Controls.Add(llmModelText);
            AddLabel(grid, 4, "제공자/모델");
            grid.Controls.Add(llmPanel, 1, 4);
            grid.SetColumnSpan(llmPanel, 2);

            llmApiKeyText.Dock = DockStyle.Fill;
            llmApiKeyText.PasswordChar = '*';
            AddLabel(grid, 5, "API 키");
            grid.Controls.Add(llmApiKeyText, 1, 5);
            grid.SetColumnSpan(llmApiKeyText, 2);

            return group;
        }

        private Control BuildRunGroup()
        {
            var panel = new TableLayoutPanel();
            panel.Dock = DockStyle.Fill;
            panel.ColumnCount = 1;
            panel.RowCount = 3;
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 172));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            var overview = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 2, Margin = new Padding(0) };
            overview.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 57));
            overview.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 43));
            overview.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            overview.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            overview.Controls.Add(new Label { Text = "산출 전 확인", Font = LauncherUi.SectionFont(), AutoSize = true }, 0, 0);
            overview.Controls.Add(new Label { Text = "최근 산출 요약", Font = LauncherUi.SectionFont(), AutoSize = true }, 1, 0);
            panel.Controls.Add(overview, 0, 0);

            readinessList.Dock = DockStyle.Fill;
            readinessList.View = View.Details;
            readinessList.FullRowSelect = true;
            readinessList.GridLines = true;
            readinessList.HeaderStyle = ColumnHeaderStyle.Nonclickable;
            readinessList.Columns.Add("항목", 82);
            readinessList.Columns.Add("상태", 72);
            readinessList.Columns.Add("확인 내용", 286);
            readinessList.SizeChanged += delegate
            {
                readinessList.Columns[2].Width = Math.Max(145, readinessList.ClientSize.Width - 158 - SystemInformation.VerticalScrollBarWidth);
            };
            overview.Controls.Add(readinessList, 0, 1);

            logText.Multiline = true;
            logText.ReadOnly = true;
            logText.ScrollBars = ScrollBars.Vertical;
            logText.Dock = DockStyle.Fill;
            logText.Margin = new Padding(0, 0, 0, 8);

            resultSummaryText.Multiline = true;
            resultSummaryText.ReadOnly = true;
            resultSummaryText.ScrollBars = ScrollBars.Vertical;
            resultSummaryText.Dock = DockStyle.Fill;
            resultSummaryText.Text = "아직 실행 결과가 없습니다.";
            overview.Controls.Add(resultSummaryText, 1, 1);

            var resultButtons = new FlowLayoutPanel();
            resultButtons.Dock = DockStyle.Fill;
            resultButtons.AutoSize = true;
            resultButtons.Margin = new Padding(0, LauncherUi.SpaceXs, 0, LauncherUi.SpaceXs);
            openWorkbookButton.Text = "산출 Excel 열기";
            openWorkbookButton.Enabled = false;
            openWorkbookButton.Click += OpenWorkbookButton_Click;
            openDraftButton.Text = "문장 초안 파일 열기";
            openDraftButton.Enabled = false;
            openDraftButton.Click += OpenDraftButton_Click;
            openHwpOutputButton.Text = "한글 보고서 열기";
            openHwpOutputButton.Enabled = false;
            openHwpOutputButton.Click += OpenHwpOutputButton_Click;
            openHwpReportButton.Text = "한글 생성 기록 열기";
            openHwpReportButton.Enabled = false;
            openHwpReportButton.Click += OpenHwpReportButton_Click;
            openPreflightButton.Text = "QA 검사 기록 열기";
            openPreflightButton.Enabled = false;
            openPreflightButton.Click += delegate { OpenPathFromButton(openPreflightButton); };
            copyDraftButton.Text = "전체 초안 복사";
            copyDraftButton.AutoSize = true;
            copyDraftButton.Enabled = false;
            copyDraftButton.Click += CopyDraftButton_Click;
            draftPreviewStatusLabel.AutoSize = true;
            draftPreviewStatusLabel.ForeColor = LauncherUi.ColorMutedText;
            draftPreviewStatusLabel.Margin = new Padding(10, 7, 0, 0);
            draftPreviewStatusLabel.Text = "문장 초안 없음";
            foreach (var button in new[] { openHwpOutputButton, openWorkbookButton, openDraftButton, openPreflightButton, openHwpReportButton })
            {
                button.AutoSize = true;
                resultButtons.Controls.Add(button);
            }
            panel.Controls.Add(resultButtons, 0, 1);

            draftPreviewText.Multiline = true;
            draftPreviewText.ReadOnly = true;
            draftPreviewText.ScrollBars = ScrollBars.Vertical;
            draftPreviewText.Dock = DockStyle.Fill;
            draftPreviewText.Text = "아직 문장 초안이 없습니다.";

            var draft = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2 };
            draft.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            draft.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var draftTools = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, Margin = new Padding(0) };
            draftTools.Controls.Add(copyDraftButton);
            draftTools.Controls.Add(draftPreviewStatusLabel);
            draft.Controls.Add(draftTools, 0, 0);
            draft.Controls.Add(draftPreviewText, 0, 1);
            draftReviewTabs.Dock = DockStyle.Fill;
            draftReviewTabs.TabPages.Add(CreateStepPage("전체 초안", draft));
            draftReviewTabs.TabPages.Add(CreateStepPage("문장 수정", BuildSentenceReviewPanel()));
            draftReviewTabs.TabPages.Add(CreateStepPage("문장 검토 경고", BuildQaReviewPanel()));
            draftReviewTabs.TabPages.Add(CreateStepPage("진행 기록", logText));
            panel.Controls.Add(draftReviewTabs, 0, 2);
            return panel;
        }

        private Control BuildDashboardPage()
        {
            var root = new TableLayoutPanel();
            root.Dock = DockStyle.Fill;
            root.ColumnCount = 1;
            root.RowCount = 2;
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 76));

            var inputs = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2 };
            inputs.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            inputs.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            inputs.Controls.Add(BuildDashboardFileGroup(), 0, 0);
            inputs.Controls.Add(BuildDashboardSelectionGroup(), 0, 1);
            var output = CreateStepPage("PPT 출력", BuildDashboardCommandGroup());
            output.Controls[0].Dock = DockStyle.Top;
            var tabs = new TabControl { Dock = DockStyle.Fill };
            tabs.TabPages.Add(CreateStepPage("원자료 · 기관", inputs));
            tabs.TabPages.Add(CreateStepPage("지표 배치", BuildDashboardMappingGroup()));
            tabs.TabPages.Add(output);
            root.Controls.Add(tabs, 0, 0);

            dashboardStatusText.Multiline = true;
            dashboardStatusText.ReadOnly = true;
            dashboardStatusText.ScrollBars = ScrollBars.Vertical;
            dashboardStatusText.Dock = DockStyle.Fill;
            dashboardStatusText.Text = "대시보드 원자료 미선택";
            root.Controls.Add(dashboardStatusText, 0, 1);
            return root;
        }

        private Control BuildDashboardFileGroup()
        {
            var group = new GroupBox();
            group.Text = "대시보드 원자료";
            group.Dock = DockStyle.Top;
            group.AutoSize = true;
            group.Padding = new Padding(10);

            var grid = CreateGrid(3);
            group.Controls.Add(grid);
            AddLabel(grid, 0, "Excel 원자료");
            dashboardWorkbookText.Dock = DockStyle.Fill;
            grid.Controls.Add(dashboardWorkbookText, 1, 0);
            dashboardBrowseButton.Text = "Excel 파일 선택…";
            dashboardBrowseButton.Dock = DockStyle.Fill;
            dashboardBrowseButton.Click += DashboardBrowseButton_Click;
            grid.Controls.Add(dashboardBrowseButton, 2, 0);

            AddLabel(grid, 1, "원자료 시트");
            dashboardSheetCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            dashboardSheetCombo.Dock = DockStyle.Fill;
            dashboardSheetCombo.SelectedIndexChanged += DashboardSheetCombo_SelectedIndexChanged;
            grid.Controls.Add(dashboardSheetCombo, 1, 1);
            dashboardInspectButton.Text = "원자료 다시 읽기";
            dashboardInspectButton.Dock = DockStyle.Fill;
            dashboardInspectButton.Click += DashboardInspectButton_Click;
            grid.Controls.Add(dashboardInspectButton, 2, 1);

            AddLabel(grid, 2, "작업용 PPT 템플릿");
            dashboardTemplateText.Dock = DockStyle.Fill;
            grid.Controls.Add(dashboardTemplateText, 1, 2);
            dashboardTemplateBrowseButton.Text = "PPT 양식 선택…";
            dashboardTemplateBrowseButton.Dock = DockStyle.Fill;
            dashboardTemplateBrowseButton.Click += DashboardTemplateBrowseButton_Click;
            grid.Controls.Add(dashboardTemplateBrowseButton, 2, 2);

            return group;
        }

        private Control BuildDashboardSelectionGroup()
        {
            var split = new TableLayoutPanel();
            split.Dock = DockStyle.Fill;
            split.ColumnCount = 3;
            split.RowCount = 1;
            split.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));
            split.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
            split.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));

            var previewGroup = new GroupBox();
            previewGroup.Text = "열 미리보기";
            previewGroup.Dock = DockStyle.Fill;
            dashboardColumnPreviewList.Dock = DockStyle.Fill;
            dashboardColumnPreviewList.View = View.Details;
            dashboardColumnPreviewList.FullRowSelect = true;
            dashboardColumnPreviewList.GridLines = true;
            dashboardColumnPreviewList.Columns.Add("열", 150);
            dashboardColumnPreviewList.Columns.Add("유형", 70);
            dashboardColumnPreviewList.Columns.Add("예시값", 140);
            dashboardColumnPreviewList.Columns.Add("결측", 60);
            previewGroup.Controls.Add(dashboardColumnPreviewList);
            split.Controls.Add(previewGroup, 0, 0);

            var entityGroup = new GroupBox();
            entityGroup.Text = "기관/기업 선택";
            entityGroup.Dock = DockStyle.Fill;
            var entityPanel = new TableLayoutPanel();
            entityPanel.Dock = DockStyle.Fill;
            entityPanel.RowCount = 3;
            entityPanel.ColumnCount = 1;
            entityPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            entityPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            entityPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var entityTop = new FlowLayoutPanel();
            entityTop.Dock = DockStyle.Fill;
            entityTop.AutoSize = true;
            var entityLabel = new Label();
            entityLabel.Text = "기관명 열";
            entityLabel.AutoSize = true;
            entityLabel.Margin = new Padding(0, 7, 6, 0);
            dashboardEntityColumnCombo.Width = 160;
            dashboardEntityColumnCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            dashboardEntityColumnCombo.SelectedIndexChanged += delegate { PopulateDashboardEntities(); };
            entityTop.Controls.Add(entityLabel);
            entityTop.Controls.Add(dashboardEntityColumnCombo);
            entityPanel.Controls.Add(entityTop, 0, 0);
            dashboardEntityList.Dock = DockStyle.Fill;
            dashboardEntityList.CheckOnClick = true;
            entityPanel.Controls.Add(dashboardEntityList, 0, 1);
            var entityButtons = new FlowLayoutPanel();
            entityButtons.Dock = DockStyle.Fill;
            entityButtons.AutoSize = true;
            dashboardSelectAllEntitiesButton.Text = "모든 기관 체크";
            dashboardSelectAllEntitiesButton.AutoSize = true;
            dashboardSelectAllEntitiesButton.Click += delegate { SetChecked(dashboardEntityList, true); };
            dashboardClearEntitiesButton.Text = "체크 해제";
            dashboardClearEntitiesButton.AutoSize = true;
            dashboardClearEntitiesButton.Click += delegate { SetChecked(dashboardEntityList, false); };
            entityButtons.Controls.Add(dashboardSelectAllEntitiesButton);
            entityButtons.Controls.Add(dashboardClearEntitiesButton);
            entityPanel.Controls.Add(entityButtons, 0, 2);
            entityGroup.Controls.Add(entityPanel);
            split.Controls.Add(entityGroup, 1, 0);

            var columnGroup = new GroupBox();
            columnGroup.Text = "사용할 열 선택";
            columnGroup.Dock = DockStyle.Fill;
            var columnPanel = new TableLayoutPanel();
            columnPanel.Dock = DockStyle.Fill;
            columnPanel.RowCount = 2;
            columnPanel.ColumnCount = 1;
            columnPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            columnPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            dashboardColumnList.Dock = DockStyle.Fill;
            dashboardColumnList.CheckOnClick = true;
            dashboardColumnList.ItemCheck += delegate { BeginInvoke(new Action(RefreshDashboardMappingCandidates)); };
            columnPanel.Controls.Add(dashboardColumnList, 0, 0);
            var columnButtons = new FlowLayoutPanel();
            columnButtons.Dock = DockStyle.Fill;
            columnButtons.AutoSize = true;
            dashboardSelectAllColumnsButton.Text = "모든 열 체크";
            dashboardSelectAllColumnsButton.AutoSize = true;
            dashboardSelectAllColumnsButton.Click += delegate { SetChecked(dashboardColumnList, true); RefreshDashboardMappingCandidates(); };
            dashboardClearColumnsButton.Text = "체크 해제";
            dashboardClearColumnsButton.AutoSize = true;
            dashboardClearColumnsButton.Click += delegate { SetChecked(dashboardColumnList, false); RefreshDashboardMappingCandidates(); };
            columnButtons.Controls.Add(dashboardSelectAllColumnsButton);
            columnButtons.Controls.Add(dashboardClearColumnsButton);
            columnPanel.Controls.Add(columnButtons, 0, 1);
            columnGroup.Controls.Add(columnPanel);
            split.Controls.Add(columnGroup, 2, 0);
            return split;
        }

        private Control BuildDashboardMappingGroup()
        {
            var tabs = new TabControl();
            tabs.Dock = DockStyle.Fill;
            tabs.TabPages.Add(CreateStepPage("KPI", BuildDashboardKpiPanel()));
            tabs.TabPages.Add(CreateStepPage("차트", BuildDashboardChartPanel()));
            tabs.TabPages.Add(CreateStepPage("분석문", BuildDashboardNarrativePanel()));
            return tabs;
        }

        private Control BuildDashboardKpiPanel()
        {
            var grid = CreateGrid(6);
            for (int i = 0; i < 6; i++)
            {
                AddLabel(grid, i, "KPI " + (i + 1));
                var row = new FlowLayoutPanel();
                row.Dock = DockStyle.Fill;
                row.AutoSize = true;
                dashboardKpiLabelTexts[i] = new TextBox();
                dashboardKpiLabelTexts[i].Width = 120;
                dashboardKpiLabelTexts[i].Text = DefaultKpiLabel(i);
                dashboardKpiColumnCombos[i] = new ComboBox();
                dashboardKpiColumnCombos[i].Width = 180;
                dashboardKpiColumnCombos[i].DropDownStyle = ComboBoxStyle.DropDownList;
                dashboardKpiUnitTexts[i] = new TextBox();
                dashboardKpiUnitTexts[i].Width = 50;
                dashboardKpiUnitTexts[i].Text = DefaultKpiUnit(i);
                row.Controls.Add(dashboardKpiLabelTexts[i]);
                row.Controls.Add(dashboardKpiColumnCombos[i]);
                row.Controls.Add(dashboardKpiUnitTexts[i]);
                grid.Controls.Add(row, 1, i);
                grid.SetColumnSpan(row, 2);
            }
            return grid;
        }

        private Control BuildDashboardChartPanel()
        {
            var grid = CreateGrid(4);
            for (int i = 0; i < 4; i++)
            {
                AddLabel(grid, i, "차트 " + (i + 1));
                var row = new FlowLayoutPanel();
                row.Dock = DockStyle.Fill;
                row.AutoSize = true;
                dashboardChartTitleTexts[i] = new TextBox();
                dashboardChartTitleTexts[i].Width = 140;
                dashboardChartTitleTexts[i].Text = DefaultChartTitle(i);
                dashboardChartTypeCombos[i] = new ComboBox();
                dashboardChartTypeCombos[i].Width = 90;
                dashboardChartTypeCombos[i].DropDownStyle = ComboBoxStyle.DropDownList;
                dashboardChartTypeCombos[i].Items.Add("auto");
                dashboardChartTypeCombos[i].Items.Add("column");
                dashboardChartTypeCombos[i].Items.Add("pie");
                dashboardChartTypeCombos[i].Items.Add("line");
                dashboardChartTypeCombos[i].Items.Add("progress");
                dashboardChartTypeCombos[i].SelectedIndex = i == 3 ? 3 : 0;
                dashboardChartColumnsTexts[i] = new TextBox();
                dashboardChartColumnsTexts[i].Width = 260;
                dashboardChartColumnsTexts[i].Text = DefaultChartColumns(i);
                dashboardChartLabelsTexts[i] = new TextBox();
                dashboardChartLabelsTexts[i].Width = 180;
                dashboardChartLabelsTexts[i].Text = DefaultChartLabels(i);
                row.Controls.Add(dashboardChartTitleTexts[i]);
                row.Controls.Add(dashboardChartTypeCombos[i]);
                row.Controls.Add(dashboardChartColumnsTexts[i]);
                row.Controls.Add(dashboardChartLabelsTexts[i]);
                grid.Controls.Add(row, 1, i);
                grid.SetColumnSpan(row, 2);
            }
            return grid;
        }

        private Control BuildDashboardNarrativePanel()
        {
            dashboardNarrativeTemplateText.Multiline = true;
            dashboardNarrativeTemplateText.ScrollBars = ScrollBars.Vertical;
            dashboardNarrativeTemplateText.Dock = DockStyle.Fill;
            dashboardNarrativeTemplateText.Text = "{{기관명}}은 매출액 {{매출액}}, 만족도 {{만족도}}%를 기록했다.";
            return dashboardNarrativeTemplateText;
        }

        private Control BuildDashboardCommandGroup()
        {
            var panel = new FlowLayoutPanel();
            panel.Dock = DockStyle.Fill;
            panel.AutoSize = true;
            panel.BackColor = LauncherUi.ColorSurfaceAlt;
            panel.Padding = new Padding(LauncherUi.SpaceSm);
            var outputLabel = new Label();
            outputLabel.Text = "기관 범위";
            outputLabel.AutoSize = true;
            outputLabel.Margin = new Padding(0, 7, 6, 0);
            dashboardOutputModeCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            dashboardOutputModeCombo.Width = 150;
            dashboardOutputModeCombo.Items.Add("단일 기관");
            dashboardOutputModeCombo.Items.Add("여러 기관 일괄");
            var sizeLabel = new Label();
            sizeLabel.Text = "용지";
            sizeLabel.AutoSize = true;
            sizeLabel.Margin = new Padding(12, 7, 6, 0);
            dashboardPageSizeCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            dashboardPageSizeCombo.Width = 90;
            dashboardPageSizeCombo.Items.Add("A4");
            dashboardPageSizeCombo.Items.Add("B5");
            var designLabel = new Label();
            designLabel.Text = "디자인";
            designLabel.AutoSize = true;
            designLabel.Margin = new Padding(12, 7, 6, 0);
            dashboardDesignCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            dashboardDesignCombo.Width = 130;
            dashboardDesignCombo.Items.Add("모던 블루");
            dashboardDesignCombo.Items.Add("모던 민트");
            dashboardDesignCombo.Items.Add("그래파이트");
            var fontLabel = new Label();
            fontLabel.Text = "글꼴";
            fontLabel.AutoSize = true;
            fontLabel.Margin = new Padding(12, 7, 6, 0);
            dashboardFontCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            dashboardFontCombo.Width = 130;
            dashboardFontCombo.Items.Add("맑은 고딕");
            dashboardFontCombo.Items.Add("나눔고딕");
            dashboardFontCombo.Items.Add("Noto Sans CJK KR");
            dashboardFontCombo.Items.Add("Arial");
            dashboardGenerateButton.Text = "대시보드 PPT 만들기…";
            dashboardGenerateButton.AutoSize = true;
            dashboardGenerateButton.Tag = LauncherButtonKind.Primary;
            dashboardGenerateButton.Click += DashboardGenerateButton_Click;
            dashboardOpenOutputButton.Text = "산출 PPT 열기";
            dashboardOpenOutputButton.AutoSize = true;
            dashboardOpenOutputButton.Enabled = false;
            dashboardOpenOutputButton.Click += DashboardOpenOutputButton_Click;
            panel.Controls.Add(outputLabel);
            panel.Controls.Add(dashboardOutputModeCombo);
            panel.Controls.Add(sizeLabel);
            panel.Controls.Add(dashboardPageSizeCombo);
            panel.Controls.Add(designLabel);
            panel.Controls.Add(dashboardDesignCombo);
            panel.Controls.Add(fontLabel);
            panel.Controls.Add(dashboardFontCombo);
            panel.Controls.Add(dashboardGenerateButton);
            panel.Controls.Add(dashboardOpenOutputButton);
            return panel;
        }

        private Control BuildSentenceReviewPanel()
        {
            var panel = new TableLayoutPanel();
            panel.Dock = DockStyle.Fill;
            panel.ColumnCount = 1;
            panel.RowCount = 3;
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 55));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 45));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            sentenceReviewList.Dock = DockStyle.Fill;
            sentenceReviewList.View = View.Details;
            sentenceReviewList.FullRowSelect = true;
            sentenceReviewList.GridLines = true;
            sentenceReviewList.MultiSelect = false;
            sentenceReviewList.Columns.Add("No", 46);
            sentenceReviewList.Columns.Add("상태", 70);
            sentenceReviewList.Columns.Add("제목", 260);
            sentenceReviewList.Columns.Add("문장", 520);
            sentenceReviewList.Columns.Add("출처", 110);
            sentenceReviewList.SelectedIndexChanged += SentenceReviewList_SelectedIndexChanged;
            panel.Controls.Add(sentenceReviewList, 0, 0);

            sentenceEditText.Multiline = true;
            sentenceEditText.ScrollBars = ScrollBars.Vertical;
            sentenceEditText.Dock = DockStyle.Fill;
            sentenceEditText.Text = "문장 목록에서 항목을 선택하면 여기에서 수정할 수 있습니다.";
            panel.Controls.Add(sentenceEditText, 0, 1);

            var buttons = new FlowLayoutPanel();
            buttons.Dock = DockStyle.Fill;
            buttons.AutoSize = true;
            applySentenceEditButton.Text = "수정 문장 반영";
            applySentenceEditButton.AutoSize = true;
            applySentenceEditButton.Enabled = false;
            applySentenceEditButton.Click += ApplySentenceEditButton_Click;
            copySelectedSentenceButton.Text = "선택 문장 복사";
            copySelectedSentenceButton.AutoSize = true;
            copySelectedSentenceButton.Enabled = false;
            copySelectedSentenceButton.Click += CopySelectedSentenceButton_Click;
            exportReviewedDraftButton.Text = "검토한 초안 저장…";
            exportReviewedDraftButton.AutoSize = true;
            exportReviewedDraftButton.Enabled = false;
            exportReviewedDraftButton.Click += ExportReviewedDraftButton_Click;
            buttons.Controls.Add(applySentenceEditButton);
            buttons.Controls.Add(copySelectedSentenceButton);
            buttons.Controls.Add(exportReviewedDraftButton);
            panel.Controls.Add(buttons, 0, 2);
            return panel;
        }

        private Control BuildQaReviewPanel()
        {
            var panel = new TableLayoutPanel();
            panel.Dock = DockStyle.Fill;
            panel.ColumnCount = 1;
            panel.RowCount = 2;
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            var top = new FlowLayoutPanel();
            top.Dock = DockStyle.Fill;
            top.AutoSize = true;
            var label = new Label();
            label.Text = "필터";
            label.AutoSize = true;
            label.Margin = new Padding(0, 7, 8, 0);
            qaFilterCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            qaFilterCombo.Width = 180;
            qaFilterCombo.Items.Add("전체");
            qaFilterCombo.Items.Add("확인 필요");
            qaFilterCombo.Items.Add("출처 없음");
            qaFilterCombo.Items.Add("문장 짧음");
            qaFilterCombo.Items.Add("수치 없음");
            qaFilterCombo.Items.Add("종결 표현 확인");
            qaFilterCombo.SelectedIndex = 0;
            qaFilterCombo.SelectedIndexChanged += delegate { PopulateQaIssues(); };
            top.Controls.Add(label);
            top.Controls.Add(qaFilterCombo);
            panel.Controls.Add(top, 0, 0);

            qaIssueList.Dock = DockStyle.Fill;
            qaIssueList.View = View.Details;
            qaIssueList.FullRowSelect = true;
            qaIssueList.GridLines = true;
            qaIssueList.MultiSelect = false;
            qaIssueList.Columns.Add("No", 46);
            qaIssueList.Columns.Add("유형", 110);
            qaIssueList.Columns.Add("제목", 260);
            qaIssueList.Columns.Add("내용", 560);
            qaIssueList.SelectedIndexChanged += QaIssueList_SelectedIndexChanged;
            panel.Controls.Add(qaIssueList, 0, 1);
            return panel;
        }

        private Control BuildOptionsGroup()
        {
            var group = new GroupBox();
            group.Text = "분석문 작성 기준";
            group.Dock = DockStyle.Top;
            group.AutoSize = true;
            group.Padding = new Padding(10);

            var grid = CreateGrid(5);
            group.Controls.Add(grid);

            AddLabel(grid, 0, "보고서 유형 (참고)");
            reportProfileCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            reportProfileCombo.Items.Add("인식도/만족도 조사형");
            reportProfileCombo.Items.Add("산업 실태조사형");
            reportProfileCombo.Items.Add("정책 수요조사형");
            reportProfileCombo.Items.Add("일반 빈도/교차표형");
            reportProfileCombo.Dock = DockStyle.Fill;
            grid.Controls.Add(reportProfileCombo, 1, 0);
            grid.SetColumnSpan(reportProfileCombo, 2);

            AddLabel(grid, 1, "문체 (참고)");
            styleProfileCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            styleProfileCombo.Items.Add("공식 보고서체");
            styleProfileCombo.Items.Add("간결 요약체");
            styleProfileCombo.Items.Add("상세 해석체");
            styleProfileCombo.Items.Add("검토 메모체");
            styleProfileCombo.Dock = DockStyle.Fill;
            grid.Controls.Add(styleProfileCombo, 1, 1);
            grid.SetColumnSpan(styleProfileCombo, 2);

            AddLabel(grid, 2, "제목에서 뺄 접두어");
            titlePrefixesText.Dock = DockStyle.Fill;
            grid.Controls.Add(titlePrefixesText, 1, 2);
            grid.SetColumnSpan(titlePrefixesText, 2);

            AddLabel(grid, 3, "분석문 작성 예시");
            var caption = new Label
            {
                Text = "작성 참고 · 디지털 상담 지원 서비스(가상, 500개 기업)",
                AutoSize = true, ForeColor = LauncherUi.ColorMutedText,
                Margin = new Padding(3, 8, 0, 5)
            };
            grid.Controls.Add(caption, 1, 3);
            grid.SetColumnSpan(caption, 2);
            narrativeExampleText.Multiline = true;
            narrativeExampleText.ReadOnly = true;
            narrativeExampleText.ScrollBars = ScrollBars.Vertical;
            narrativeExampleText.Dock = DockStyle.Fill;
            narrativeExampleText.AccessibleName = "선택한 보고서 유형의 분석문 예시";
            grid.RowStyles[4].SizeType = SizeType.Absolute;
            grid.RowStyles[4].Height = 120;
            grid.Controls.Add(narrativeExampleText, 1, 4);
            grid.SetColumnSpan(narrativeExampleText, 2);
            var focus = new[]
            {
                "이용자 평가: 만족·불만족 비율과 환산 점수",
                "산업 현황: 전체 이용률과 기업 규모별 차이",
                "정책 수요: 현재 미이용 기업의 향후 이용 의향",
                "빈도·교차표: 사례수와 비율, 집단별 분모"
            };
            var examples = new[,]
            {
                {
                    "이용 기업 310개를 대상으로 조사한 결과, 만족 응답은 68.4%로 불만족 응답(12.9%)보다 높게 나타남. 100점 환산 평균은 72.6점임.",
                    "이용 기업 만족 68.4%, 불만족 12.9%. 100점 환산 평균 72.6점.",
                    "이용 기업 310개 중 68.4%가 만족, 12.9%가 불만족으로 응답함. 평균은 72.6점(100점 환산)이며, 미이용 기업은 이 평가에 포함되지 않음.",
                    "확인: 만족 68.4%, 불만족 12.9%, 100점 평균 72.6점. 본문에 ‘이용 기업 기준(n=310)’을 표시할 것."
                },
                {
                    "전체 응답 기업의 서비스 이용률은 62.0%로 나타남. 중대형 기업은 72.0%로 소규모 기업(52.0%)보다 20.0%p 높았음.",
                    "이용률 전체 62.0%. 중대형(72.0%)이 소규모(52.0%)보다 20.0%p 높음.",
                    "전체 이용률은 62.0%이며, 규모별로는 중대형 72.0%, 소규모 52.0%로 20.0%p 차이가 남. 이 표만으로는 기업 규모가 이용률 차이의 원인인지 판단할 수 없음.",
                    "확인: 전체 62.0%, 중대형 72.0%, 소규모 52.0%. 격차는 20.0%p로 표기하고 원인으로 단정하지 말 것."
                },
                {
                    "현재 미이용 기업 190개 중 54.2%가 향후 이용 의향이 있다고 응답함. 기존 이용 실적과 별도로 신규 지원 수요를 검토할 때 참고할 수 있음.",
                    "미이용 기업의 54.2%가 향후 이용 의향을 밝힘.",
                    "미이용 기업 190개 중 103개(54.2%)가 향후 이용 의향을 보임. 이는 미이용 집단의 잠재 수요이며, 전체 기업의 이용률(62.0%)이나 향후 실제 이용률과는 구분해야 함.",
                    "검토: 이용 의향 54.2%의 분모는 미이용 기업 190개. 지원 규모를 산정하기 전에 미이용 사유도 확인할 것."
                },
                {
                    "응답 기업 500개 중 이용 310개(62.0%), 미이용 190개(38.0%)로 집계됨. 규모별 이용률은 소규모 52.0%, 중대형 72.0%임.",
                    "이용 310개(62.0%), 미이용 190개(38.0%). 규모별 이용률은 소규모 52.0%, 중대형 72.0%.",
                    "전체 500개 중 310개(62.0%)가 이용함. 규모별 분모는 각각 250개로, 이용 기업은 소규모 130개(52.0%), 중대형 180개(72.0%)임.",
                    "대조: 이용 310개 + 미이용 190개 = 500개. 교차표의 소규모 52.0%, 중대형 72.0%는 각 집단 250개 기준."
                }
            };
            EventHandler updateExample = delegate
            {
                int type = reportProfileCombo.SelectedIndex;
                int style = styleProfileCombo.SelectedIndex;
                narrativeExampleText.Text = type < 0 || style < 0 ? "" :
                    focus[type] + Environment.NewLine + Environment.NewLine + examples[type, style];
            };
            reportProfileCombo.SelectedIndexChanged += updateExample;
            styleProfileCombo.SelectedIndexChanged += updateExample;
            return group;
        }

        private Control BuildTableExamplePage()
        {
            var panel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 2 };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 59));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 41));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var toolbar = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
            toolbar.Controls.Add(new Label { Text = "집계표 유형", AutoSize = true, Margin = new Padding(0, 8, 8, 0) });
            tableExampleCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            tableExampleCombo.Width = 214;
            tableExampleCombo.AccessibleName = "산출 예시에 사용할 집계표 유형";
            tableExampleCombo.Items.AddRange(new object[] { "단일응답 비율표", "복수응답 비율표", "배너 교차표", "만족도 척도표", "100점 점수표" });
            toolbar.Controls.Add(tableExampleCombo);
            panel.Controls.Add(toolbar, 0, 0);
            panel.SetColumnSpan(toolbar, 2);

            var input = new GroupBox { Text = "가상 집계표 · 문장에 사용한 값 강조", Dock = DockStyle.Fill };
            tableExampleGrid.Dock = DockStyle.Fill;
            tableExampleGrid.AccessibleName = "가상 입력 집계표";
            tableExampleGrid.ReadOnly = true;
            tableExampleGrid.AllowUserToAddRows = false;
            tableExampleGrid.AllowUserToDeleteRows = false;
            tableExampleGrid.AllowUserToResizeRows = false;
            tableExampleGrid.RowHeadersVisible = false;
            tableExampleGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            tableExampleGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            tableExampleGrid.EnableHeadersVisualStyles = false;
            tableExampleGrid.ColumnHeadersDefaultCellStyle.BackColor = LauncherUi.ColorSurfaceAlt;
            tableExampleGrid.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            tableExampleGrid.BackgroundColor = LauncherUi.ColorSurface;
            tableExampleGrid.BorderStyle = BorderStyle.FixedSingle;
            tableExampleGrid.GridColor = LauncherUi.ColorBorder;
            tableExampleGrid.DefaultCellStyle.SelectionBackColor = LauncherUi.ColorSurfaceStrong;
            tableExampleGrid.DefaultCellStyle.SelectionForeColor = LauncherUi.ColorText;
            tableExampleGrid.AlternatingRowsDefaultCellStyle.BackColor = LauncherUi.ColorRowAlt;
            input.Controls.Add(tableExampleGrid);
            panel.Controls.Add(input, 0, 1);

            var output = new GroupBox { Text = "Excel 매크로 산출 · 분석문_기본", Dock = DockStyle.Fill };
            var result = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3 };
            result.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            result.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            result.RowStyles.Add(new RowStyle(SizeType.Absolute, 70));
            var title = new Label { AutoSize = true, Font = LauncherUi.SectionFont() };
            tableExampleOutput.Multiline = true;
            tableExampleOutput.ReadOnly = true;
            tableExampleOutput.ScrollBars = ScrollBars.Vertical;
            tableExampleOutput.Dock = DockStyle.Fill;
            tableExampleOutput.AccessibleName = "집계표에서 자동 생성한 기본 분석문";
            var source = new Label { Dock = DockStyle.Fill, ForeColor = LauncherUi.ColorMutedText, Margin = new Padding(3, 8, 3, 0) };
            result.Controls.Add(title, 0, 0);
            result.Controls.Add(tableExampleOutput, 0, 1);
            result.Controls.Add(source, 0, 2);
            output.Controls.Add(result);
            panel.Controls.Add(output, 1, 1);

            var titles = new[] { "주 이용 채널", "필요한 지원(복수응답)", "규모별 주 이용 채널", "서비스 이용 만족도", "세부 항목 만족도(100점)" };
            var tables = new[]
            {
                new[] { new[] { "구분", "사례수", "%" }, new[] { "웹사이트", "33", "33.0" },
                    new[] { "오프라인 창구", "15", "15.0" }, new[] { "모바일 앱", "52", "52.0" } },
                new[] { new[] { "구분", "사례수", "%" }, new[] { "컨설팅", "27", "27.0" },
                    new[] { "인력 양성", "54", "54.0" }, new[] { "사업비 지원", "40", "40.0" } },
                new[] { new[] { "구분", "사례수", "모바일 앱\nN", "모바일 앱\n%", "웹사이트\nN", "웹사이트\n%", "오프라인 창구\nN", "오프라인 창구\n%" },
                    new[] { "전체", "100", "52", "52.0", "33", "33.0", "15", "15.0" },
                    new[] { "소규모", "40", "16", "40.0", "16", "40.0", "8", "20.0" },
                    new[] { "중대형", "60", "36", "60.0", "17", "28.3", "7", "11.7" } },
                new[] { new[] { "구분", "사례수", "만족(합)\nN", "만족(합)\n%", "불만족(합)\nN", "불만족(합)\n%", "100점\n평균" },
                    new[] { "전체", "100", "68", "68.0", "13", "13.0", "72.6" },
                    new[] { "소규모", "40", "24", "60.0", "8", "20.0", "65.0" },
                    new[] { "중대형", "60", "44", "73.3", "5", "8.3", "77.7" } },
                new[] { new[] { "항목", "사례수", "100점 점수" }, new[] { "처리 속도", "100", "69.2" },
                    new[] { "상담 품질", "100", "78.4" }, new[] { "이용 편의", "100", "74.1" } }
            };
            // These are the VBA outputs for the sample tables, not a second implementation of the writer.
            var sentences = new[,]
            {
                { "주 이용 채널은 '모바일 앱'이 52.0%로 가장 높게 나타남",
                    "다음으로 '웹사이트'(33.0%), '오프라인 창구'(15.0%) 순으로 나타남" },
                { "필요한 지원(복수응답)은 복수응답 기준으로 '인력 양성'이 54.0%로 가장 높게 나타남",
                    "다음으로 '사업비 지원'(40.0%), '컨설팅'(27.0%) 순으로 나타남" },
                { "규모별 주 이용 채널은 '모바일 앱'이 52.0%로 가장 높게 나타남",
                    "다음으로 '웹사이트'(33.0%), '오프라인 창구'(15.0%) 순으로 나타남" },
                { "서비스 이용 만족도에 대해 조사한 결과, '만족(합)' 응답은 68.0%, 100점 평균은 72.6점으로 나타남",
                    "반면, '불만족(합)'은 13.0%로 나타남" },
                { "‘세부 항목 만족도’를 100점 환산 기준으로 분석한 결과, '상담 품질'이 78.4점으로 가장 높게 나타남",
                    "그다음으로는 '이용 편의'(74.1점), '처리 속도'(69.2점) 등의 순으로 높게 나타남" }
            };
            var sources = new[]
            {
                "문장에 사용: 응답 항목별 %를 높은 순서로 정렬. 사례수는 순위에 사용하지 않음.",
                "문장에 사용: 응답 항목별 %와 제목의 ‘복수응답’ 표시. 비율 합계가 100%를 넘을 수 있음.",
                "문장에 사용: ‘전체’ 행의 %만 비교. 사례수와 소규모·중대형의 차이는 기본 문장에 넣지 않음.",
                "문장에 사용: ‘전체’ 행의 만족(합) %, 100점 평균, 불만족(합) %. 점수를 비율 순위에 넣지 않음.",
                "문장에 사용: 항목별 100점 점수. 비율(%) 대신 점수를 순위로 비교."
            };
            tableExampleCombo.SelectedIndexChanged += delegate
            {
                int index = tableExampleCombo.SelectedIndex;
                if (index < 0) return;
                title.Text = titles[index];
                source.Text = sources[index];
                tableExampleOutput.Text = sentences[index, 0] + Environment.NewLine + sentences[index, 1];
                tableExampleGrid.Columns.Clear();
                for (int column = 0; column < tables[index][0].Length; column++)
                {
                    tableExampleGrid.Columns.Add(new DataGridViewTextBoxColumn
                    {
                        HeaderText = tables[index][0][column], MinimumWidth = 54, SortMode = DataGridViewColumnSortMode.NotSortable,
                        DefaultCellStyle = new DataGridViewCellStyle
                        {
                            Alignment = column == 0 ? DataGridViewContentAlignment.MiddleLeft : DataGridViewContentAlignment.MiddleRight
                        }
                    });
                }
                for (int row = 1; row < tables[index].Length; row++)
                    tableExampleGrid.Rows.Add(tables[index][row]);
                if (index == 2 || index == 3)
                {
                    foreach (int column in new[] { 3, 5, index == 2 ? 7 : 6 })
                        tableExampleGrid.Rows[0].Cells[column].Style.BackColor = LauncherUi.ColorSurfaceStrong;
                }
                else
                {
                    foreach (DataGridViewRow row in tableExampleGrid.Rows)
                        row.Cells[2].Style.BackColor = LauncherUi.ColorSurfaceStrong;
                }
                tableExampleGrid.ClearSelection();
            };
            tableExampleCombo.SelectedIndex = 0;
            return panel;
        }

        private Control BuildBannerPanel()
        {
            var panel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 7 };
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            panel.Controls.Add(new Label { Text = "분석 배너", Font = LauncherUi.SectionFont(), AutoSize = true }, 0, 0);

            bannerList.Dock = DockStyle.Fill;
            bannerList.CheckOnClick = true;
            bannerList.IntegralHeight = false;
            bannerList.AccessibleName = "분석에 포함할 가로 배너";
            bannerList.ItemCheck += BannerList_ItemCheck;
            panel.Controls.Add(bannerList, 0, 2);

            var selection = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, Margin = new Padding(0) };
            recommendedBannerButton.Text = "추천 체크";
            recommendedBannerButton.Click += delegate { SelectRecommendedBanners(); };
            selectAllBannerButton.Text = "모두 체크";
            selectAllBannerButton.Click += delegate { SetAllBannersChecked(true); };
            clearBannerButton.Text = "체크 해제";
            clearBannerButton.Click += delegate { SetAllBannersChecked(false); };
            foreach (var button in new[] { recommendedBannerButton, selectAllBannerButton, clearBannerButton })
            {
                button.AutoSize = true;
                selection.Controls.Add(button);
            }
            panel.Controls.Add(selection, 0, 1);

            var reorder = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, Margin = new Padding(0) };
            moveBannerUpButton.Text = "↑";
            moveBannerUpButton.Width = 34;
            moveBannerUpButton.AccessibleName = "선택한 배너 위로 이동";
            actionTips.SetToolTip(moveBannerUpButton, moveBannerUpButton.AccessibleName);
            moveBannerUpButton.Click += delegate { MoveSelectedBanner(-1); };
            moveBannerDownButton.Text = "↓";
            moveBannerDownButton.Width = 34;
            moveBannerDownButton.AccessibleName = "선택한 배너 아래로 이동";
            actionTips.SetToolTip(moveBannerDownButton, moveBannerDownButton.AccessibleName);
            moveBannerDownButton.Click += delegate { MoveSelectedBanner(1); };
            deleteBannerButton.Text = "배너 목록에서 제거";
            deleteBannerButton.AutoSize = true;
            deleteBannerButton.Click += delegate { DeleteSelectedBanners(); };
            reorder.Controls.Add(moveBannerUpButton);
            reorder.Controls.Add(moveBannerDownButton);
            reorder.Controls.Add(deleteBannerButton);
            panel.Controls.Add(reorder, 0, 3);
            bannerStatusLabel.Dock = DockStyle.Fill;
            bannerStatusLabel.Height = 22;
            bannerStatusLabel.AutoEllipsis = true;
            bannerStatusLabel.ForeColor = LauncherUi.ColorMutedText;
            panel.Controls.Add(bannerStatusLabel, 0, 4);
            panel.Controls.Add(new Label { Text = "분석에 적용할 배너", AutoSize = true, ForeColor = LauncherUi.ColorMutedText }, 0, 5);
            bannerText.Dock = DockStyle.Fill;
            bannerText.AccessibleName = "분석에 사용할 배너 목록";
            panel.Controls.Add(bannerText, 0, 6);
            return panel;
        }

        private static TableLayoutPanel CreateGrid(int rows)
        {
            var grid = new TableLayoutPanel();
            grid.Dock = DockStyle.Fill;
            grid.AutoSize = true;
            grid.Margin = new Padding(0);
            grid.Padding = new Padding(LauncherUi.SpaceXs, LauncherUi.SpaceXs, LauncherUi.SpaceXs, 0);
            grid.ColumnCount = 3;
            grid.RowCount = rows;
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 148));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 138));
            for (int i = 0; i < rows; i++)
            {
                grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            }
            return grid;
        }

        private static void AddLabel(TableLayoutPanel grid, int row, string text)
        {
            var label = new Label();
            label.Text = text;
            label.TextAlign = ContentAlignment.MiddleLeft;
            label.Dock = DockStyle.Fill;
            label.Font = LauncherUi.SectionFont();
            label.ForeColor = LauncherUi.ColorMutedText;
            label.Margin = new Padding(0, LauncherUi.SpaceSm, LauncherUi.SpaceMd, LauncherUi.SpaceSm);
            grid.Controls.Add(label, 0, row);
        }

        private static void AddPathRow(TableLayoutPanel grid, int row, string label, TextBox textBox, string buttonText, EventHandler handler)
        {
            AddLabel(grid, row, label);
            textBox.AccessibleName = label;
            textBox.Dock = DockStyle.Fill;
            textBox.Margin = new Padding(0, LauncherUi.SpaceXs, LauncherUi.SpaceSm, LauncherUi.SpaceXs);
            grid.Controls.Add(textBox, 1, row);

            var button = new Button();
            button.Text = buttonText;
            button.Dock = DockStyle.Fill;
            button.Margin = new Padding(0, LauncherUi.SpaceXs, 0, LauncherUi.SpaceXs);
            button.Click += handler;
            grid.Controls.Add(button, 2, row);
        }

        private void BrowseWorkbook(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = "집계표 엑셀 파일 선택";
                dialog.Filter = "Excel files (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls|All files (*.*)|*.*";
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    workbookPathText.Text = dialog.FileName;
                    LoadWorkbookPreviewAsync();
                    workflowTabs.SelectedIndex = 1;
                }
            }
        }

        private void LoadBannerPreviewAsync()
        {
            LoadWorkbookPreviewAsync();
        }

        private void LoadWorkbookPreviewAsync()
        {
            string path = workbookPathText.Text.Trim();
            if (string.IsNullOrWhiteSpace(path))
            {
                MessageBox.Show(this, "집계표 엑셀 파일을 먼저 선택하세요.", "배너 확인", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (!File.Exists(path))
            {
                MessageBox.Show(this, "선택한 집계표 파일을 찾을 수 없습니다.", "배너 확인", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            SetBannerPreviewEnabled(false);
            bannerStatusLabel.Text = "읽는 중...";
            dataStatusLabel.Text = "집계표 구조를 읽는 중...";
            tablePreviewList.Items.Clear();
            Log("표 목록과 배너 목록을 읽습니다.");

            var thread = new Thread(delegate()
            {
                try
                {
                    WorkbookPreview preview = BannerInspector.ReadPreview(path);
                    BeginInvoke(new Action(delegate()
                    {
                        bannerList.Items.Clear();
                        recommendedBanners.Clear();
                        foreach (string banner in preview.PrimaryBanners)
                        {
                            recommendedBanners.Add(banner);
                        }
                        foreach (string banner in preview.Banners)
                        {
                            bool isRecommended = recommendedBanners.Contains(banner) || preview.PrimaryBanners.Count == 0;
                            bannerList.Items.Add(banner, isRecommended);
                        }
                        UpdateBannerTextFromCheckedList();
                        PopulateTablePreview(preview.Tables);
                        UpdateBannerStatusText();
                        UpdateWorkflowStatus();
                        dataStatusLabel.Text = preview.Tables.Count + "개 표, " + preview.Banners.Count + "개 배너를 발견했습니다.";
                        Log("표 " + preview.Tables.Count + "개, 배너 " + preview.Banners.Count + "개, 추천 배너 " + preview.PrimaryBanners.Count + "개를 발견했습니다.");
                    }));
                }
                catch (Exception ex)
                {
                    BeginInvoke(new Action(delegate()
                    {
                        bannerStatusLabel.Text = "확인 실패";
                        dataStatusLabel.Text = "집계표 구조 확인에 실패했습니다.";
                        UpdateWorkflowStatus();
                        Log("배너 확인 실패: " + ex.Message);
                        MessageBox.Show(this, ex.Message, "배너 확인 실패", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }));
                }
                finally
                {
                    BeginInvoke(new Action(delegate()
                    {
                        SetBannerPreviewEnabled(true);
                    }));
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }

        private void PopulateTablePreview(List<TablePreview> tables)
        {
            var snapshot = new List<TablePreview>(tables);
            tablePreviewList.BeginUpdate();
            try
            {
                tablePreviewList.Items.Clear();
                currentTablePreviews.Clear();
                int index = 1;
                foreach (TablePreview table in snapshot)
                {
                    currentTablePreviews.Add(table);
                    var item = new ListViewItem(index.ToString());
                    item.SubItems.Add(table.TableNo);
                    item.SubItems.Add(table.Title);
                    item.SubItems.Add(table.SheetName);
                    item.SubItems.Add(table.AutoRange);
                    item.SubItems.Add(table.FinalRange);
                    item.SubItems.Add(table.RowCount + "x" + table.ColCount);
                    item.SubItems.Add(table.IsManual ? "수동" : "자동");
                    item.SubItems.Add("포함");
                    item.SubItems.Add(table.RecognitionWarning ?? "");
                    item.Tag = table;
                    LauncherUi.StyleListItem(item, index);
                    tablePreviewList.Items.Add(item);
                    index++;
                }
            }
            finally
            {
                tablePreviewList.EndUpdate();
            }
        }

        private void SelectTableRangeInExcel()
        {
            if (tablePreviewList.SelectedItems.Count != 1)
            {
                MessageBox.Show(this, "범위를 보정할 표 하나를 선택하세요.", "표 범위", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string path = workbookPathText.Text.Trim();
            if (!File.Exists(path))
            {
                MessageBox.Show(this, "집계표 엑셀 파일을 찾을 수 없습니다.", "표 범위", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            TablePreview table = tablePreviewList.SelectedItems[0].Tag as TablePreview;
            object excel = null;
            object workbook = null;
            object selected = null;
            try
            {
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    throw new InvalidOperationException("Excel.Application COM 개체를 찾을 수 없습니다.");
                }

                excel = Activator.CreateInstance(excelType);
                dynamic xl = excel;
                xl.Visible = true;
                workbook = xl.Workbooks.Open(path, 0, true);
                dynamic wb = workbook;
                wb.Worksheets[table.SheetName].Activate();
                wb.Worksheets[table.SheetName].Range[table.FinalRange].Select();
                object missing = Type.Missing;
                selected = xl.InputBox("삽입할 표 범위를 드래그한 뒤 확인을 누르세요.", "표 범위 선택", missing, missing, missing, missing, missing, 8);
                if (selected is bool && !(bool)selected)
                {
                    return;
                }

                dynamic range = selected;
                if (Convert.ToInt32(range.Areas.Count) != 1)
                {
                    throw new InvalidOperationException("서로 떨어진 여러 범위는 사용할 수 없습니다. 연속된 표 범위 하나를 선택하세요.");
                }
                table.SheetName = Convert.ToString(range.Worksheet.Name);
                table.FinalRange = Convert.ToString(range.Address(false, false));
                table.RowCount = Convert.ToInt32(range.Rows.Count);
                table.ColCount = Convert.ToInt32(range.Columns.Count);
                table.IsManual = true;
                PopulateTablePreview(currentTablePreviews);
                SelectTablePreview(table.TableKey);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "표 범위 선택 실패", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                AutomationRunner.CloseComWorkbookForLauncher(workbook, false);
                AutomationRunner.QuitComExcelForLauncher(excel);
                AutomationRunner.ReleaseComForLauncher(selected);
                AutomationRunner.ReleaseComForLauncher(workbook);
                AutomationRunner.ReleaseComForLauncher(excel);
            }
        }

        private void ClearSelectedTableRange()
        {
            if (tablePreviewList.SelectedItems.Count != 1)
            {
                return;
            }
            TablePreview table = tablePreviewList.SelectedItems[0].Tag as TablePreview;
            table.SheetName = table.AutoSheetName;
            table.FinalRange = table.AutoRange;
            table.RowCount = table.AutoRowCount;
            table.ColCount = table.AutoColCount;
            table.IsManual = false;
            PopulateTablePreview(currentTablePreviews);
            SelectTablePreview(table.TableKey);
        }

        private void SelectTablePreview(string tableKey)
        {
            foreach (ListViewItem item in tablePreviewList.Items)
            {
                TablePreview table = item.Tag as TablePreview;
                if (table != null && table.TableKey == tableKey)
                {
                    item.Selected = true;
                    item.EnsureVisible();
                    return;
                }
            }
        }

        private void BannerList_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            BeginInvoke(new Action(UpdateBannerTextFromCheckedList));
        }

        private void UpdateBannerTextFromCheckedList()
        {
            if (bannerList.Items.Count == 0)
            {
                bannerStatusLabel.Text = "";
                return;
            }

            var selected = new List<string>();
            for (int i = 0; i < bannerList.Items.Count; i++)
            {
                if (bannerList.GetItemChecked(i))
                {
                    selected.Add(bannerList.Items[i].ToString());
                }
            }
            bannerText.Text = selected.Count == 0 ? "전체" : string.Join(",", selected.ToArray());
            UpdateBannerStatusText();
            UpdateWorkflowStatus();
        }

        private void SetAllBannersChecked(bool isChecked)
        {
            for (int i = 0; i < bannerList.Items.Count; i++)
            {
                bannerList.SetItemChecked(i, isChecked);
            }
            UpdateBannerTextFromCheckedList();
        }

        private void SelectRecommendedBanners()
        {
            bool hasRecommended = recommendedBanners.Count > 0;
            for (int i = 0; i < bannerList.Items.Count; i++)
            {
                string banner = bannerList.Items[i].ToString();
                bannerList.SetItemChecked(i, !hasRecommended || recommendedBanners.Contains(banner));
            }
            UpdateBannerTextFromCheckedList();
        }

        private void MoveSelectedBanner(int direction)
        {
            int oldIndex = bannerList.SelectedIndex;
            if (oldIndex < 0)
            {
                return;
            }

            int newIndex = oldIndex + direction;
            if (newIndex < 0 || newIndex >= bannerList.Items.Count)
            {
                return;
            }

            object item = bannerList.Items[oldIndex];
            bool isChecked = bannerList.GetItemChecked(oldIndex);
            bannerList.Items.RemoveAt(oldIndex);
            bannerList.Items.Insert(newIndex, item);
            bannerList.SetItemChecked(newIndex, isChecked);
            bannerList.SelectedIndex = newIndex;
            UpdateBannerTextFromCheckedList();
        }

        private void DeleteSelectedBanners()
        {
            int index = bannerList.SelectedIndex;
            if (index < 0)
            {
                return;
            }

            bannerList.Items.RemoveAt(index);
            if (bannerList.Items.Count > 0)
            {
                bannerList.SelectedIndex = Math.Min(index, bannerList.Items.Count - 1);
            }
            UpdateBannerTextFromCheckedList();
        }

        private void UpdateBannerStatusText()
        {
            if (bannerList.Items.Count == 0)
            {
                bannerStatusLabel.Text = "";
                return;
            }

            int checkedCount = 0;
            for (int i = 0; i < bannerList.Items.Count; i++)
            {
                if (bannerList.GetItemChecked(i))
                {
                    checkedCount++;
                }
            }

            int recommendedCount = 0;
            foreach (object item in bannerList.Items)
            {
                if (recommendedBanners.Contains(item.ToString()))
                {
                    recommendedCount++;
                }
            }

            bannerStatusLabel.Text = bannerList.Items.Count + "개 발견 / " + recommendedCount + "개 추천 / " + checkedCount + "개 선택";
        }

        private void SetBannerPreviewEnabled(bool enabled)
        {
            reloadBannerButton.Enabled = enabled;
            recommendedBannerButton.Enabled = enabled;
            selectAllBannerButton.Enabled = enabled;
            clearBannerButton.Enabled = enabled;
            moveBannerUpButton.Enabled = enabled;
            moveBannerDownButton.Enabled = enabled;
            deleteBannerButton.Enabled = enabled;
            bannerList.Enabled = enabled;
        }

        private void BrowseHwpTemplate(object sender, EventArgs e)
        {
            BrowseTemplate(hwpTemplateText, "HWP/HWPX 템플릿 선택", "HWP/HWPX files (*.hwp;*.hwpx)|*.hwp;*.hwpx|All files (*.*)|*.*", "HWPX/HWP 템플릿을 선택했습니다. 검사를 실행하세요.");
        }

        private void BrowsePptTemplate(object sender, EventArgs e)
        {
            BrowseTemplate(pptTemplateText, "PowerPoint 템플릿 선택", "PowerPoint files (*.pptx;*.ppt)|*.pptx;*.ppt|All files (*.*)|*.*", "PPTX 템플릿을 선택했습니다. 검사를 실행하세요.");
        }

        private void BrowseHwpTableStyleProfile(object sender, EventArgs e)
        {
            BrowseTemplate(hwpTableStyleProfileText, "HWP 표 스타일 profile 선택", "JSON files (*.json)|*.json|All files (*.*)|*.*", "HWP 표 스타일 profile을 선택했습니다.");
        }

        private void BrowseTemplate(TextBox targetTextBox, string title, string filter, string statusMessage)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = title;
                dialog.Filter = filter;
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    targetTextBox.Text = dialog.FileName;
                    ResetTemplateStatus(statusMessage);
                }
            }
        }

        private void ResetTemplateStatus(string message)
        {
            lastTemplateStatus = "미검사";
            templateStatusLabel.Text = message;
            templateStatusLabel.ForeColor = LauncherUi.ColorMutedText;
            UpdateWorkflowStatus();
        }

        private void InspectTemplateButton_Click(object sender, EventArgs e)
        {
            string templatePath = SelectedTemplatePath();
            if (string.IsNullOrWhiteSpace(templatePath))
            {
                MessageBox.Show(this, "검사할 HWPX/HWP 또는 PPTX 템플릿을 먼저 선택하세요.", "템플릿 검사", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (!File.Exists(templatePath))
            {
                MessageBox.Show(this, "선택한 템플릿 파일을 찾을 수 없습니다.", "템플릿 검사", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                InspectTemplate(templatePath, SelectedTemplateType(templatePath), true);
            }
            catch (Exception ex)
            {
                lastTemplateStatus = "검사 실패";
                templateStatusLabel.Text = "템플릿 검사 실패: " + ex.Message;
                templateStatusLabel.ForeColor = LauncherUi.ColorDanger;
                UpdateWorkflowStatus();
                MessageBox.Show(this, ex.Message, "템플릿 검사 실패", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void CreateHwpTemplateButton_Click(object sender, EventArgs e)
        {
            CreateTemplate("hwpx_report", "report_template_basic.hwpx", "HWPX files (*.hwpx)|*.hwpx", hwpTemplateText);
        }

        private void UseBasicHwpTemplateButton_Click(object sender, EventArgs e)
        {
            string engineDirectory = Path.GetDirectoryName(PathResolver.ResolveEngineToolPath("template_factory.py"));
            string path = Path.GetFullPath(Path.Combine(engineDirectory, "..", "templates", "report_template_basic.hwpx"));
            try
            {
                InspectTemplate(path, "hwpx_report", false);
                if (!IsTemplateStatusUsable()) throw new InvalidOperationException("기본 한글 양식을 확인할 수 없습니다.");
                hwpTemplateText.Text = path;
                templateStatusLabel.Text = "기본 한글 양식을 사용합니다. 본문·표 서식은 아래에서 별도로 선택할 수 있습니다.";
            }
            catch (Exception ex)
            {
                ResetTemplateStatus("기본 한글 양식 확인 실패");
                MessageBox.Show(this, ex.Message, "기본 한글 양식", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void CreatePptTemplateButton_Click(object sender, EventArgs e)
        {
            CreateTemplate("pptx_report", "report_template_basic.pptx", "PowerPoint files (*.pptx)|*.pptx", pptTemplateText);
        }

        private void CreateChartTemplateButton_Click(object sender, EventArgs e)
        {
            CreateTemplate("chart_review", "chart_review_template_basic.pptx", "PowerPoint files (*.pptx)|*.pptx", pptTemplateText);
        }

        private void CreateTemplate(string templateType, string defaultName, string filter, TextBox targetTextBox)
        {
            using (var dialog = new SaveFileDialog())
            {
                dialog.Title = "기본 템플릿 저장";
                dialog.FileName = defaultName;
                dialog.Filter = filter + "|All files (*.*)|*.*";
                dialog.OverwritePrompt = true;
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                try
                {
                    string outputPath = dialog.FileName;
                    string arguments = "--type " + QuoteArg(templateType) + " --output " + QuoteArg(outputPath);
                    RunTemplateTool("template_factory.py", arguments, 60000);
                    targetTextBox.Text = outputPath;
                    Log("기본 템플릿 생성: " + outputPath);
                    InspectTemplate(outputPath, templateType, false);
                    MessageBox.Show(this, "기본 템플릿을 생성했습니다." + Environment.NewLine + outputPath, "템플릿 생성", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, ex.Message, "템플릿 생성 실패", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void AutoFixTemplateButton_Click(object sender, EventArgs e)
        {
            string templatePath = SelectedTemplatePath();
            if (string.IsNullOrWhiteSpace(templatePath) || !File.Exists(templatePath))
            {
                MessageBox.Show(this, "자동 보정할 템플릿 파일을 먼저 선택하세요.", "템플릿 자동 보정", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string templateType = SelectedTemplateType(templatePath);
            string extension = Path.GetExtension(templatePath);
            string defaultPath = Path.Combine(
                Path.GetDirectoryName(templatePath),
                Path.GetFileNameWithoutExtension(templatePath) + "_template_ready" + extension);

            using (var dialog = new SaveFileDialog())
            {
                dialog.Title = "자동 보정 사본 저장";
                dialog.FileName = Path.GetFileName(defaultPath);
                dialog.InitialDirectory = Path.GetDirectoryName(defaultPath);
                dialog.Filter = extension.Equals(".pptx", StringComparison.OrdinalIgnoreCase) ? "PowerPoint files (*.pptx)|*.pptx|All files (*.*)|*.*" : "HWPX files (*.hwpx)|*.hwpx|All files (*.*)|*.*";
                dialog.OverwritePrompt = true;
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                try
                {
                    string outputPath = dialog.FileName;
                    string arguments = "--template " + QuoteArg(templatePath) +
                                      " --type " + QuoteArg(templateType) +
                                      " --output " + QuoteArg(outputPath);
                    RunTemplateTool("template_autofix.py", arguments, 60000);
                    if (extension.Equals(".pptx", StringComparison.OrdinalIgnoreCase))
                    {
                        pptTemplateText.Text = outputPath;
                    }
                    else
                    {
                        hwpTemplateText.Text = outputPath;
                    }
                    Log("템플릿 자동 보정 사본 생성: " + outputPath);
                    InspectTemplate(outputPath, templateType, false);
                    MessageBox.Show(this, "원본을 보존하고 자동 보정 사본을 생성했습니다." + Environment.NewLine + outputPath, "템플릿 자동 보정", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, ex.Message, "템플릿 자동 보정 실패", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void OpenTemplateGuideButton_Click(object sender, EventArgs e)
        {
            string guidePath = Path.Combine(Path.GetTempPath(), "report_automation_template_guide.txt");
            File.WriteAllText(guidePath, BuildTemplateGuideText(), System.Text.Encoding.UTF8);

            var info = new ProcessStartInfo(guidePath);
            info.UseShellExecute = true;
            Process.Start(info);
        }

        private string SelectedTemplatePath()
        {
            string output = outputTypeCombo.Text;
            if (output.IndexOf("HWP", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return hwpTemplateText.Text.Trim();
            }
            if (output.IndexOf("PowerPoint", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return pptTemplateText.Text.Trim();
            }
            if (File.Exists(hwpTemplateText.Text.Trim()))
            {
                return hwpTemplateText.Text.Trim();
            }
            return pptTemplateText.Text.Trim();
        }

        private string SelectedTemplateType(string templatePath)
        {
            string extension = Path.GetExtension(templatePath);
            if (extension.Equals(".hwp", StringComparison.OrdinalIgnoreCase) || extension.Equals(".hwpx", StringComparison.OrdinalIgnoreCase))
            {
                return "hwpx_report";
            }
            if (Path.GetFileNameWithoutExtension(templatePath).IndexOf("chart", StringComparison.OrdinalIgnoreCase) >= 0 ||
                Path.GetFileNameWithoutExtension(templatePath).IndexOf("차트", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "chart_review";
            }
            return "pptx_report";
        }

        private void InspectTemplate(string templatePath, string templateType, bool showMessage)
        {
            string reportPath = Path.Combine(Path.GetTempPath(), "report_automation_template_report.json");
            string arguments = "--template " + QuoteArg(templatePath) +
                              " --type " + QuoteArg(templateType) +
                              " --output " + QuoteArg(reportPath);
            RunTemplateTool("template_inspector.py", arguments, 60000);

            string json = File.Exists(reportPath) ? File.ReadAllText(reportPath, System.Text.Encoding.UTF8) : "";
            string status = ReadJsonStringValue(json, "status");
            string detectedType = ReadJsonStringValue(json, "template_type");
            string found = ReadJsonArrayValue(json, "found_placeholders");
            string missingRequired = ReadJsonArrayValue(json, "missing_required");
            string missingRecommended = ReadJsonArrayValue(json, "missing_recommended");

            if (string.IsNullOrWhiteSpace(status))
            {
                status = "검사 결과 없음";
            }
            if (string.IsNullOrWhiteSpace(detectedType))
            {
                detectedType = templateType;
            }

            lastTemplateStatus = status;
            templateStatusLabel.Text = "상태: " + status + " / 유형: " + detectedType +
                                       " / 발견: " + EmptyToDash(found) +
                                       " / 누락 필수: " + EmptyToDash(missingRequired);
            templateStatusLabel.ForeColor = IsTemplateStatusUsable() ? LauncherUi.ColorSuccess : LauncherUi.ColorWarning;
            UpdateWorkflowStatus();

            Log("템플릿 검사: " + status + " (" + detectedType + ")");
            if (showMessage)
            {
                string message = "상태: " + status + Environment.NewLine +
                                 "유형: " + detectedType + Environment.NewLine +
                                 "발견 placeholder: " + EmptyToDash(found) + Environment.NewLine +
                                 "누락 필수 필드: " + EmptyToDash(missingRequired) + Environment.NewLine +
                                 "누락 권장 필드: " + EmptyToDash(missingRecommended);
                MessageBox.Show(this, message, "템플릿 검사 결과", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        internal static void RunTemplateTool(string scriptName, string arguments, int timeoutMs)
        {
            string pythonPath = PathResolver.ResolvePythonPath();
            string scriptPath = PathResolver.ResolveEngineToolPath(scriptName);
            if (string.IsNullOrWhiteSpace(pythonPath) || !File.Exists(pythonPath))
            {
                throw new FileNotFoundException("Python 실행 파일을 찾지 못했습니다. REPORT_AUTOMATION_PYTHON 환경변수를 확인하세요.");
            }
            if (string.IsNullOrWhiteSpace(scriptPath) || !File.Exists(scriptPath))
            {
                throw new FileNotFoundException("템플릿 도구를 찾지 못했습니다.", scriptName);
            }

            var startInfo = new ProcessStartInfo();
            startInfo.FileName = pythonPath;
            startInfo.Arguments = QuoteArg(scriptPath) + " " + arguments;
            startInfo.UseShellExecute = false;
            startInfo.CreateNoWindow = true;
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = true;
            startInfo.StandardOutputEncoding = System.Text.Encoding.UTF8;
            startInfo.StandardErrorEncoding = System.Text.Encoding.UTF8;

            StandaloneBundle.ConfigurePythonProcess(startInfo);
            using (Process process = Process.Start(startInfo))
            {
                var stdoutTask = process.StandardOutput.ReadToEndAsync();
                var stderrTask = process.StandardError.ReadToEndAsync();
                if (!process.WaitForExit(timeoutMs))
                {
                    try { process.Kill(); } catch { }
                    throw new TimeoutException("템플릿 도구 실행 시간이 초과되었습니다.");
                }
                string stdout = stdoutTask.Result;
                string stderr = stderrTask.Result;
                if (process.ExitCode != 0)
                {
                    throw new InvalidOperationException(string.IsNullOrWhiteSpace(stderr) ? stdout.Trim() : stderr.Trim());
                }
            }
        }

        internal static string QuoteArg(string value)
        {
            return "\"" + (value ?? "").Replace("\"", "\\\"") + "\"";
        }

        private static string EmptyToDash(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? "-" : value;
        }

        private static string ReadJsonStringValue(string json, string key)
        {
            int start = JsonValueStart(json, key, '"');
            if (start < 0)
            {
                return "";
            }
            int end = json.IndexOf('"', start + 1);
            if (end < 0)
            {
                return "";
            }
            return json.Substring(start + 1, end - start - 1);
        }

        private static string ReadJsonArrayValue(string json, string key)
        {
            int start = JsonValueStart(json, key, '[');
            int end = json.IndexOf(']', start + 1);
            if (start < 0 || end < 0)
            {
                return "";
            }
            string raw = json.Substring(start + 1, end - start - 1).Trim();
            if (string.IsNullOrWhiteSpace(raw))
            {
                return "";
            }
            return raw.Replace("\"", "").Replace(",", ", ");
        }

        private static string ReadJsonPrimitiveValue(string json, string key)
        {
            string marker = "\"" + key + "\"";
            int keyIndex = json.IndexOf(marker, StringComparison.Ordinal);
            if (keyIndex < 0)
            {
                return "";
            }
            int colonIndex = json.IndexOf(':', keyIndex + marker.Length);
            if (colonIndex < 0)
            {
                return "";
            }
            int start = colonIndex + 1;
            while (start < json.Length && char.IsWhiteSpace(json[start]))
            {
                start++;
            }
            int end = start;
            while (end < json.Length && ",}\r\n ".IndexOf(json[end]) < 0)
            {
                end++;
            }
            return json.Substring(start, end - start).Trim().Trim('"');
        }

        private static int JsonValueStart(string json, string key, char valueStart)
        {
            string marker = "\"" + key + "\"";
            int keyIndex = json.IndexOf(marker, StringComparison.Ordinal);
            if (keyIndex < 0)
            {
                return -1;
            }
            int colonIndex = json.IndexOf(':', keyIndex + marker.Length);
            return colonIndex < 0 ? -1 : json.IndexOf(valueStart, colonIndex + 1);
        }

        private static string BuildTemplateGuideText()
        {
            return "보고서 자동화 템플릿 가이드" + Environment.NewLine +
                   Environment.NewLine +
                   "1. 최소 HWPX 템플릿" + Environment.NewLine +
                   "- [기본 한글 양식 사용]은 내장 양식을 바로 선택합니다. 편집하려면 [편집용 양식 사본 저장]을 누르세요." + Environment.NewLine +
                   "- 한글 문서에서 본문이 들어갈 위치에 {{BODY}} 한 줄을 유지합니다." + Environment.NewLine +
                   "- 표지, 글꼴, 여백, 제목 스타일, 로고, 쪽번호는 자유롭게 수정할 수 있습니다." + Environment.NewLine +
                   "- 본문·표 서식은 문서 양식과 별개입니다. 골드 배너 또는 청회색 머리행을 선택할 수 있습니다." + Environment.NewLine +
                   "- 예: [내 서식 추가·편집] → 아래 표 스타일에서 '표머리행' 선택 → [HWPX에서 셀 서식 가져오기] → 머리행 셀과 배경색 선택 → [선택 적용]." + Environment.NewLine +
                   "- 새 이름을 입력하고 [내 서식으로 저장] → [이 서식 사용] 순서로 적용합니다. 최대 3개를 저장하며, .hwp 파일은 .hwpx 사본으로 저장 후 가져옵니다." + Environment.NewLine +
                   Environment.NewLine +
                   "2. 최소 PPTX 보고서 템플릿" + Environment.NewLine +
                   "- 반복 슬라이드에 {{SECTION_TITLE}}, {{NARRATIVE}}, {{TABLE}}, {{CHART}} 텍스트 상자를 둡니다." + Environment.NewLine +
                   "- shape 이름이 RA_로 시작하는 요소는 삭제하거나 이름을 바꾸지 않습니다." + Environment.NewLine +
                   Environment.NewLine +
                   "3. 최소 차트 검토 PPTX 템플릿" + Environment.NewLine +
                   "- 반복 슬라이드에 {{CHART_TITLE}}, {{CHART}}, {{CHART_NOTE}}를 둡니다." + Environment.NewLine +
                   Environment.NewLine +
                   "4. 수정 가능 영역" + Environment.NewLine +
                   "- 글꼴, 색상, 배경, 마스터 슬라이드, 표지, 페이지 번호, placeholder의 위치와 크기" + Environment.NewLine +
                   Environment.NewLine +
                   "5. 수정 금지 영역" + Environment.NewLine +
                   "- 중괄호 placeholder 텍스트, RA_로 시작하는 shape/bookmark 이름, 반복 슬라이드, {{BODY}}" + Environment.NewLine;
        }

        private void RunButton_Click(object sender, EventArgs e)
        {
            LauncherOptions options;
            try
            {
                options = ReadOptions();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "입력 확인", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            reportRunning = true;
            workflowTabs.SelectedIndex = 3;
            draftReviewTabs.SelectedIndex = 3;
            UpdateWorkflowStatus();
            resultSummaryText.Text = options.OutputType.IndexOf("HWP", StringComparison.OrdinalIgnoreCase) >= 0
                ? "원본 확인 → Excel 산출 시트 자동 생성 → QA → 한글 보고서"
                : "원본 확인 → Excel 산출 시트 자동 생성 → QA";
            draftPreviewText.Text = "문장 초안을 기다리는 중입니다.";
            draftPreviewStatusLabel.Text = "실행 중...";
            openWorkbookButton.Enabled = false;
            openDraftButton.Enabled = false;
            openHwpOutputButton.Enabled = false;
            openHwpReportButton.Enabled = false;
            openPreflightButton.Enabled = false;
            copyDraftButton.Enabled = false;
            Log("실행을 시작합니다.");

            var thread = new Thread(delegate()
            {
                try
                {
                    EngineRunner.RunWorkflow(options, Log);
                    BeginInvoke(new Action(delegate()
                    {
                        resultSummaryText.Text = BuildResultSummary(options);
                        PopulateResultFiles(options);
                        MessageBox.Show(this, "보고서 자동화 산출이 완료되었습니다.", "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }));
                }
                catch (Exception ex)
                {
                    Log("오류: " + ex.Message);
                    var addinError = ex as ExcelAddinOpenException;
                    string diagnosticPath = null;
                    if (addinError != null)
                    {
                        try
                        {
                            string path = Path.Combine(Path.GetDirectoryName(options.WorkbookPath), "excel_addin_error_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".json");
                            ExcelAddinDiagnostics.Save(path, addinError.Report);
                            diagnosticPath = path;
                            Log("Excel 연결 오류 기록: " + diagnosticPath);
                        }
                        catch (Exception saveError) { Log("오류 기록 저장 실패: " + saveError.Message); }
                    }
                    BeginInvoke(new Action(delegate()
                    {
                        PopulateResultFiles(options);
                        string errorMessage = ex.Message + (diagnosticPath == null ? "" :
                            Environment.NewLine + "Excel 연결 오류 기록: " + diagnosticPath);
                        resultSummaryText.Text = errorMessage +
                            (string.IsNullOrWhiteSpace(options.LastGeneratedWorkbookPath) ? "" :
                                Environment.NewLine + "저장된 산출 Excel: " + options.LastGeneratedWorkbookPath) +
                            (string.IsNullOrWhiteSpace(options.LastPreflightReportPath) ? "" :
                                Environment.NewLine + "QA 검사 기록: " + options.LastPreflightReportPath);
                        MessageBox.Show(this, errorMessage, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }));
                }
                finally
                {
                    BeginInvoke(new Action(delegate()
                    {
                        reportRunning = false;
                        UpdateWorkflowStatus();
                    }));
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }

        private static string BuildResultSummary(LauncherOptions options)
        {
            var lines = new List<string>();
            lines.Add("완료되었습니다.");
            lines.Add("");
            lines.Add("출력 형식: " + options.OutputType);
            lines.Add("보고서 유형: " + options.ReportProfile);
            lines.Add("문체: " + options.StyleProfile);
            lines.Add("수치 표기: 소수점 " + options.DecimalPlaces + "자리");
            lines.Add("차트 출력: " + options.ChartOutputMode);
            lines.Add("삽입표 방식: " + options.TableInsertMode);
            lines.Add("LLM 문장 고도화: " + (options.UseLlm ? options.LlmProvider + " / " + options.LlmModel : "사용 안 함"));
            lines.Add("작업 방식: " + (options.CopyWorkbook ? "원본 옆 복사본 생성" : "원본 파일 직접 산출"));
            lines.Add("선택 배너: " + options.BannerSetting);
            lines.Add("구성요소: " + BuildComponentSummary(options));
            if (!string.IsNullOrWhiteSpace(options.LastGeneratedWorkbookPath))
            {
                lines.Add("산출 엑셀: " + options.LastGeneratedWorkbookPath);
            }
            if (!string.IsNullOrWhiteSpace(options.LastDraftTextPath))
            {
                lines.Add("문장 초안: " + options.LastDraftTextPath);
            }
            else if (options.GenerateDraftText)
            {
                lines.Add("문장 초안: 생성하지 못했습니다. 로그를 확인하세요.");
            }
            if (!string.IsNullOrWhiteSpace(options.LastReportPackagePath))
            {
                lines.Add("Report package: " + options.LastReportPackagePath);
            }
            if (!string.IsNullOrWhiteSpace(options.LastPreflightReportPath))
            {
                lines.Add("Preflight: " + options.LastPreflightReportPath);
                lines.Add(BuildPreflightSummary(options.LastPreflightReportPath));
            }
            if (!string.IsNullOrWhiteSpace(options.LastHwpOutputPath))
            {
                lines.Add("HWPX 초본: " + options.LastHwpOutputPath);
            }
            else if (options.OutputType.IndexOf("HWP", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                lines.Add("HWPX 초본: 생성하지 못했습니다. 로그와 writer report를 확인하세요.");
            }
            if (!string.IsNullOrWhiteSpace(options.LastHwpWriterReportPath))
            {
                lines.Add("HWPX writer report: " + options.LastHwpWriterReportPath);
                lines.Add(BuildHwpWriterReportSummary(options.LastHwpWriterReportPath));
            }
            lines.Add("");
            lines.Add("Excel에서 보고서_분석문, 보고서_차트데이터, 보고서_삽입표, 보고서_QA 시트를 확인하세요.");
            return string.Join(Environment.NewLine, lines.ToArray());
        }

        private static string BuildPreflightSummary(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                return "문서 생성 준비 상태: 확인 실패";
            }
            string json = File.ReadAllText(path, System.Text.Encoding.UTF8);
            return "문서 생성 준비 상태: " + EmptyToDash(ReadJsonStringValue(json, "status")) +
                   " / 문장 " + EmptyToDash(ReadJsonPrimitiveValue(json, "section_count")) +
                   " / 차트 " + EmptyToDash(ReadJsonPrimitiveValue(json, "chart_candidate_count")) +
                   " / 삽입표 " + EmptyToDash(ReadJsonPrimitiveValue(json, "table_count")) +
                   " / 경고 " + EmptyToDash(ReadJsonPrimitiveValue(json, "qa_warning_count")) +
                   " / 오류 " + EmptyToDash(ReadJsonPrimitiveValue(json, "qa_error_count"));
        }

        private static string BuildHwpWriterReportSummary(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                return "HWP COM: report 없음";
            }
            try
            {
                var serializer = new JavaScriptSerializer();
                serializer.MaxJsonLength = int.MaxValue;
                var root = serializer.DeserializeObject(File.ReadAllText(path, System.Text.Encoding.UTF8)) as Dictionary<string, object>;
                if (root == null)
                {
                    return "HWP COM: report 파싱 실패";
                }

                string status = JsonString(root, "status");
                string stage = JsonString(root, "stage");
                string action = JsonString(root, "action");
                string dispatchMode = "";
                string progId = "";
                string lastStep = "";

                Dictionary<string, object> com = JsonObject(root, "com");
                if (com != null)
                {
                    dispatchMode = JsonString(com, "dispatch_mode");
                    progId = JsonString(com, "current_prog_id");
                    lastStep = LastComStepSummary(com);
                }

                return "HWP COM: status=" + EmptyToDash(status) +
                       " / dispatch=" + EmptyToDash(dispatchMode) +
                       " / progId=" + EmptyToDash(progId) +
                       " / stage=" + EmptyToDash(stage) +
                       " / action=" + EmptyToDash(action) +
                       " / lastStep=" + EmptyToDash(lastStep);
            }
            catch (Exception ex)
            {
                return "HWP COM: report 파싱 실패 - " + ex.Message;
            }
        }

        private static string LastComStepSummary(Dictionary<string, object> com)
        {
            object stepsValue;
            if (!com.TryGetValue("steps", out stepsValue))
            {
                return "";
            }
            object[] steps = stepsValue as object[];
            if (steps == null || steps.Length == 0)
            {
                return "";
            }
            Dictionary<string, object> last = steps[steps.Length - 1] as Dictionary<string, object>;
            if (last == null)
            {
                return "";
            }
            string name = JsonString(last, "name");
            string status = JsonString(last, "status");
            string mode = JsonString(last, "dispatch_mode");
            return string.IsNullOrWhiteSpace(mode) ? name + ":" + status : name + ":" + status + ":" + mode;
        }

        private static Dictionary<string, object> JsonObject(Dictionary<string, object> values, string key)
        {
            object value;
            return values.TryGetValue(key, out value) ? value as Dictionary<string, object> : null;
        }

        private static string JsonString(Dictionary<string, object> values, string key)
        {
            object value;
            return values.TryGetValue(key, out value) && value != null ? Convert.ToString(value) : "";
        }

        private static string BuildComponentSummary(LauncherOptions options)
        {
            var components = new List<string>();
            if (options.IncludeAnalysis) components.Add("분석문");
            if (options.IncludeCharts) components.Add("차트 데이터");
            if (options.IncludeTables) components.Add("삽입용 집계표");
            if (options.IncludeQa) components.Add("QA/출처/수정이력");
            if (options.GenerateDraftText) components.Add("문장 초안 TXT");
            return components.Count == 0 ? "선택 없음" : string.Join(", ", components.ToArray());
        }

        private void PopulateResultFiles(LauncherOptions options)
        {
            openWorkbookButton.Tag = options.LastGeneratedWorkbookPath;
            openWorkbookButton.Enabled = !string.IsNullOrWhiteSpace(options.LastGeneratedWorkbookPath) && File.Exists(options.LastGeneratedWorkbookPath);

            openDraftButton.Tag = options.LastDraftTextPath;
            openDraftButton.Enabled = !string.IsNullOrWhiteSpace(options.LastDraftTextPath) && File.Exists(options.LastDraftTextPath);

            openHwpOutputButton.Tag = options.LastHwpOutputPath;
            openHwpOutputButton.Enabled = !string.IsNullOrWhiteSpace(options.LastHwpOutputPath) && File.Exists(options.LastHwpOutputPath);

            openHwpReportButton.Tag = options.LastHwpWriterReportPath;
            openHwpReportButton.Enabled = !string.IsNullOrWhiteSpace(options.LastHwpWriterReportPath) && File.Exists(options.LastHwpWriterReportPath);

            openPreflightButton.Tag = options.LastPreflightReportPath;
            openPreflightButton.Enabled = File.Exists(options.LastPreflightReportPath);

            LoadDraftPreview(options.LastDraftTextPath);
        }

        private void LoadDraftPreview(string path)
        {
            copyDraftButton.Enabled = false;
            if (string.IsNullOrWhiteSpace(path))
            {
                ClearDraftReviewState();
                draftPreviewText.Text = "문장 초안이 생성되지 않았습니다.";
                draftPreviewStatusLabel.Text = "초안 없음";
                return;
            }
            if (!File.Exists(path))
            {
                ClearDraftReviewState();
                draftPreviewText.Text = "문장 초안 파일을 찾지 못했습니다." + Environment.NewLine + path;
                draftPreviewStatusLabel.Text = "초안 파일 없음";
                return;
            }

            string text = File.ReadAllText(path, System.Text.Encoding.UTF8);
            int originalLength = text.Length;
            currentDraftPath = path;
            LoadSentenceReview(text);
            const int previewLimit = 60000;
            if (text.Length > previewLimit)
            {
                text = text.Substring(0, previewLimit) + Environment.NewLine + Environment.NewLine + "... 미리보기는 여기까지 표시합니다. 전체 내용은 TXT 파일을 열어 확인하세요.";
            }

            draftPreviewText.Text = text;
            copyDraftButton.Enabled = draftPreviewText.TextLength > 0;
            draftPreviewStatusLabel.Text = originalLength.ToString("N0") + "자 초안, " + draftSentenceItems.Count + "개 문장, " + draftQaIssues.Count + "개 QA 경고";
        }

        private void ClearDraftReviewState()
        {
            currentDraftPath = "";
            draftSentenceItems.Clear();
            draftQaIssues.Clear();
            sentenceReviewList.Items.Clear();
            qaIssueList.Items.Clear();
            sentenceEditText.Text = "문장 목록에서 항목을 선택하면 여기에서 수정할 수 있습니다.";
            applySentenceEditButton.Enabled = false;
            copySelectedSentenceButton.Enabled = false;
            exportReviewedDraftButton.Enabled = false;
        }

        private void LoadSentenceReview(string text)
        {
            draftSentenceItems.Clear();
            draftQaIssues.Clear();

            string currentTitle = "";
            int index = 1;
            string[] lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();
                if (line.Length == 0)
                {
                    continue;
                }
                if (line.StartsWith("▶", StringComparison.Ordinal))
                {
                    currentTitle = line.TrimStart('▶').Trim();
                    continue;
                }
                if (line.StartsWith("[source:", StringComparison.OrdinalIgnoreCase))
                {
                    if (draftSentenceItems.Count > 0)
                    {
                        draftSentenceItems[draftSentenceItems.Count - 1].Source = line.Trim('[', ']');
                    }
                    continue;
                }
                if (currentTitle.Length == 0 && line.EndsWith("자동 생성 보고서 본문", StringComparison.Ordinal))
                {
                    continue;
                }

                var item = new DraftSentenceItem();
                item.Index = index++;
                item.Title = currentTitle.Length == 0 ? "(제목 없음)" : currentTitle;
                item.Text = line;
                item.IsEdited = false;
                draftSentenceItems.Add(item);
            }

            BuildDraftQaIssues();
            PopulateSentenceReviewList();
            PopulateQaIssues();
            exportReviewedDraftButton.Enabled = draftSentenceItems.Count > 0;
        }

        private void PopulateSentenceReviewList()
        {
            sentenceReviewList.BeginUpdate();
            try
            {
                sentenceReviewList.Items.Clear();
                foreach (DraftSentenceItem sentence in draftSentenceItems)
                {
                    var item = new ListViewItem(sentence.Index.ToString());
                    item.SubItems.Add(sentence.IsEdited ? "수정됨" : "초안");
                    item.SubItems.Add(sentence.Title);
                    item.SubItems.Add(sentence.Text);
                    item.SubItems.Add(sentence.Source);
                    item.Tag = sentence;
                    LauncherUi.StyleListItem(item, sentence.Index);
                    sentenceReviewList.Items.Add(item);
                }
            }
            finally
            {
                sentenceReviewList.EndUpdate();
            }

            applySentenceEditButton.Enabled = false;
            copySelectedSentenceButton.Enabled = false;
            if (sentenceReviewList.Items.Count > 0)
            {
                sentenceReviewList.Items[0].Selected = true;
            }
        }

        private void BuildDraftQaIssues()
        {
            draftQaIssues.Clear();
            foreach (DraftSentenceItem sentence in draftSentenceItems)
            {
                string text = sentence.Text == null ? "" : sentence.Text.Trim();
                if (sentence.Title == "(제목 없음)")
                {
                    AddQaIssue(sentence, "확인 필요", "문장에 연결된 표 제목을 확인하지 못했습니다.");
                }
                if (string.IsNullOrWhiteSpace(sentence.Source))
                {
                    AddQaIssue(sentence, "출처 없음", "문장 출처(source)가 없습니다.");
                }
                if (text.Length < 25)
                {
                    AddQaIssue(sentence, "문장 짧음", "보고서 본문으로 쓰기에는 문장이 짧습니다.");
                }
                if (!text.Contains("%") && !ContainsDigit(text))
                {
                    AddQaIssue(sentence, "수치 없음", "수치나 비율이 포함되지 않아 집계표 기반 문장인지 확인이 필요합니다.");
                }
                if (!EndsWithReportStyle(text))
                {
                    AddQaIssue(sentence, "종결 표현 확인", "보고서체 종결 표현이 자연스러운지 확인하세요.");
                }
            }
        }

        private void AddQaIssue(DraftSentenceItem sentence, string type, string message)
        {
            var issue = new DraftQaIssue();
            issue.Sentence = sentence;
            issue.Type = type;
            issue.Message = message;
            draftQaIssues.Add(issue);
        }

        private static bool ContainsDigit(string text)
        {
            foreach (char ch in text)
            {
                if (char.IsDigit(ch))
                {
                    return true;
                }
            }
            return false;
        }

        private static bool EndsWithReportStyle(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            string value = text.Trim();
            return value.EndsWith("나타남", StringComparison.Ordinal) ||
                value.EndsWith("나타났다", StringComparison.Ordinal) ||
                value.EndsWith("높게 나타남", StringComparison.Ordinal) ||
                value.EndsWith("낮게 나타남", StringComparison.Ordinal) ||
                value.EndsWith("수준임", StringComparison.Ordinal) ||
                value.EndsWith("확인됨", StringComparison.Ordinal) ||
                value.EndsWith(".", StringComparison.Ordinal);
        }

        private void PopulateQaIssues()
        {
            string filter = qaFilterCombo.SelectedItem == null ? "전체" : qaFilterCombo.SelectedItem.ToString();
            qaIssueList.BeginUpdate();
            try
            {
                qaIssueList.Items.Clear();
                int index = 1;
                foreach (DraftQaIssue issue in draftQaIssues)
                {
                    if (filter != "전체" && issue.Type != filter)
                    {
                        continue;
                    }

                    var item = new ListViewItem(index.ToString());
                    item.SubItems.Add(issue.Type);
                    item.SubItems.Add(issue.Sentence.Title);
                    item.SubItems.Add(issue.Message + " / " + issue.Sentence.Text);
                    item.Tag = issue;
                    LauncherUi.StyleListItem(item, index);
                    qaIssueList.Items.Add(item);
                    index++;
                }
            }
            finally
            {
                qaIssueList.EndUpdate();
            }
        }

        private void OpenWorkbookButton_Click(object sender, EventArgs e)
        {
            OpenPathFromButton(openWorkbookButton);
        }

        private void OpenDraftButton_Click(object sender, EventArgs e)
        {
            OpenPathFromButton(openDraftButton);
        }

        private void OpenHwpOutputButton_Click(object sender, EventArgs e)
        {
            OpenPathFromButton(openHwpOutputButton);
        }

        private void OpenHwpReportButton_Click(object sender, EventArgs e)
        {
            OpenPathFromButton(openHwpReportButton);
        }

        private void CopyDraftButton_Click(object sender, EventArgs e)
        {
            if (draftPreviewText.TextLength > 0)
            {
                Clipboard.SetText(draftPreviewText.Text);
                draftPreviewStatusLabel.Text = "미리보기 내용을 클립보드에 복사했습니다.";
            }
        }

        private void OpenPathFromButton(Button button)
        {
            string path = button.Tag == null ? "" : button.Tag.ToString();
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                MessageBox.Show(this, "열 파일을 찾을 수 없습니다.", "파일 열기", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                var info = new ProcessStartInfo(path);
                info.UseShellExecute = true;
                Process.Start(info);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "파일 열기 실패", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void DashboardBrowseButton_Click(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = "대시보드 원자료 Excel 선택";
                dialog.Filter = "Excel files (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|All files (*.*)|*.*";
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    dashboardWorkbookText.Text = dialog.FileName;
                    LoadDashboardInspectionAsync();
                }
            }
        }

        private void DashboardInspectButton_Click(object sender, EventArgs e)
        {
            LoadDashboardInspectionAsync();
        }

        private void DashboardTemplateBrowseButton_Click(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = "작업용 PPT 템플릿 선택";
                dialog.Filter = "PowerPoint files (*.pptx)|*.pptx|All files (*.*)|*.*";
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    dashboardTemplateText.Text = dialog.FileName;
                }
            }
        }

        private void LoadDashboardInspectionAsync()
        {
            string path = dashboardWorkbookText.Text.Trim();
            if (string.IsNullOrWhiteSpace(path) && File.Exists(workbookPathText.Text.Trim()))
            {
                path = workbookPathText.Text.Trim();
                dashboardWorkbookText.Text = path;
            }
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                MessageBox.Show(this, "대시보드 원자료 Excel 파일을 선택하세요.", "대시보드", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            dashboardInspectButton.Enabled = false;
            dashboardGenerateButton.Enabled = false;
            dashboardStatusText.Text = "Excel 원자료를 읽는 중입니다...";

            var thread = new Thread(delegate()
            {
                try
                {
                    string outputPath = Path.Combine(Path.GetTempPath(), "dashboard_excel_inspect.json");
                    RunPythonTool("dashboard_package.py", "--excel " + QuoteArg(path) + " --inspect-output " + QuoteArg(outputPath), 120000);
                    DashboardWorkbookInfo info = DashboardWorkbookInfo.Load(outputPath);
                    BeginInvoke(new Action(delegate()
                    {
                        currentDashboardInfo = info;
                        PopulateDashboardWorkbook(info);
                        dashboardStatusText.Text = "sheet " + info.Sheets.Count + "개를 읽었습니다. 사용할 sheet, 기관, 열을 선택하세요.";
                    }));
                }
                catch (Exception ex)
                {
                    BeginInvoke(new Action(delegate()
                    {
                        dashboardStatusText.Text = "대시보드 원자료 확인 실패: " + ex.Message;
                        MessageBox.Show(this, ex.Message, "대시보드 원자료 확인 실패", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }));
                }
                finally
                {
                    BeginInvoke(new Action(delegate()
                    {
                        dashboardInspectButton.Enabled = true;
                        dashboardGenerateButton.Enabled = true;
                    }));
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }

        private void PopulateDashboardWorkbook(DashboardWorkbookInfo info)
        {
            dashboardSheetCombo.Items.Clear();
            foreach (DashboardSheetInfo sheet in info.Sheets)
            {
                dashboardSheetCombo.Items.Add(sheet.Name);
            }
            if (dashboardSheetCombo.Items.Count > 0)
            {
                dashboardSheetCombo.SelectedIndex = 0;
            }
        }

        private void DashboardSheetCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            PopulateDashboardSheet();
        }

        private void PopulateDashboardSheet()
        {
            DashboardSheetInfo sheet = SelectedDashboardSheet();
            dashboardColumnPreviewList.Items.Clear();
            dashboardEntityColumnCombo.Items.Clear();
            dashboardColumnList.Items.Clear();
            dashboardEntityList.Items.Clear();
            if (sheet == null)
            {
                return;
            }

            foreach (DashboardColumnInfo column in sheet.Columns)
            {
                var item = new ListViewItem(column.Name);
                item.SubItems.Add(column.InferredType);
                item.SubItems.Add(column.Sample);
                item.SubItems.Add(column.MissingCount.ToString());
                LauncherUi.StyleListItem(item, dashboardColumnPreviewList.Items.Count + 1);
                dashboardColumnPreviewList.Items.Add(item);
                dashboardEntityColumnCombo.Items.Add(column.Name);
                dashboardColumnList.Items.Add(column.Name, true);
            }

            int entityIndex = FindColumnIndex(sheet, "기관명", "기업명", "회사명", "업체명", "기관");
            if (entityIndex >= 0)
            {
                dashboardEntityColumnCombo.SelectedIndex = entityIndex;
            }
            else if (dashboardEntityColumnCombo.Items.Count > 0)
            {
                dashboardEntityColumnCombo.SelectedIndex = 0;
            }
            RefreshDashboardMappingCandidates();
        }

        private DashboardSheetInfo SelectedDashboardSheet()
        {
            if (currentDashboardInfo == null || dashboardSheetCombo.SelectedItem == null)
            {
                return null;
            }
            string name = dashboardSheetCombo.SelectedItem.ToString();
            foreach (DashboardSheetInfo sheet in currentDashboardInfo.Sheets)
            {
                if (string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    return sheet;
                }
            }
            return null;
        }

        private void PopulateDashboardEntities()
        {
            dashboardEntityList.Items.Clear();
            DashboardSheetInfo sheet = SelectedDashboardSheet();
            if (sheet == null || dashboardEntityColumnCombo.SelectedItem == null)
            {
                return;
            }
            string column = dashboardEntityColumnCombo.SelectedItem.ToString();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Dictionary<string, object> row in sheet.Preview)
            {
                object value;
                if (row.TryGetValue(column, out value))
                {
                    string text = Convert.ToString(value ?? "").Trim();
                    if (text.Length > 0 && seen.Add(text))
                    {
                        dashboardEntityList.Items.Add(text, true);
                    }
                }
            }
        }

        private void RefreshDashboardMappingCandidates()
        {
            var selected = SelectedDashboardColumns();
            foreach (ComboBox combo in dashboardKpiColumnCombos)
            {
                if (combo == null)
                {
                    continue;
                }
                string previous = combo.SelectedItem == null ? "" : combo.SelectedItem.ToString();
                combo.Items.Clear();
                combo.Items.Add("");
                foreach (string column in selected)
                {
                    combo.Items.Add(column);
                }
                int index = previous.Length == 0 ? -1 : combo.Items.IndexOf(previous);
                combo.SelectedIndex = index >= 0 ? index : 0;
            }
            ApplyDefaultDashboardMapping();
        }

        private void ApplyDefaultDashboardMapping()
        {
            string[] guesses = { "매출액", "예상매출액", "종사자", "만족도", "평가점수", "성과점수" };
            for (int i = 0; i < dashboardKpiColumnCombos.Length; i++)
            {
                ComboBox combo = dashboardKpiColumnCombos[i];
                if (combo == null || combo.SelectedIndex > 0)
                {
                    continue;
                }
                int index = GuessComboIndex(combo, guesses[i]);
                if (index >= 0)
                {
                    combo.SelectedIndex = index;
                }
            }
        }

        private List<string> SelectedDashboardColumns()
        {
            var selected = new List<string>();
            for (int i = 0; i < dashboardColumnList.Items.Count; i++)
            {
                if (dashboardColumnList.GetItemChecked(i))
                {
                    selected.Add(dashboardColumnList.Items[i].ToString());
                }
            }
            return selected;
        }

        private List<string> SelectedDashboardEntities()
        {
            var selected = new List<string>();
            for (int i = 0; i < dashboardEntityList.Items.Count; i++)
            {
                if (dashboardEntityList.GetItemChecked(i))
                {
                    selected.Add(dashboardEntityList.Items[i].ToString());
                }
            }
            if (dashboardOutputModeCombo.SelectedIndex == 0 && selected.Count > 1)
            {
                return new List<string> { selected[0] };
            }
            return selected;
        }

        private void DashboardGenerateButton_Click(object sender, EventArgs e)
        {
            try
            {
                GenerateDashboardPpt();
            }
            catch (Exception ex)
            {
                dashboardStatusText.Text = "대시보드 PPT 생성 실패: " + ex.Message;
                MessageBox.Show(this, ex.Message, "대시보드 PPT 생성 실패", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void GenerateDashboardPpt()
        {
            string excelPath = dashboardWorkbookText.Text.Trim();
            if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath))
            {
                throw new InvalidOperationException("대시보드 원자료 Excel 파일을 선택하세요.");
            }
            if (dashboardSheetCombo.SelectedItem == null || dashboardEntityColumnCombo.SelectedItem == null)
            {
                throw new InvalidOperationException("sheet와 기관명 열을 선택하세요.");
            }

            string directory = Path.GetDirectoryName(excelPath);
            string stem = Path.GetFileNameWithoutExtension(excelPath);
            string selectionPath = Path.Combine(directory, stem + "_dashboard_data_selection.json");
            string mappingPath = Path.Combine(directory, stem + "_dashboard_mapping.json");
            string packagePath = Path.Combine(directory, stem + "_dashboard_package.json");
            string preflightPath = Path.Combine(directory, stem + "_dashboard_preflight_report.json");
            string outputPath = Path.Combine(directory, stem + "_dashboard_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".pptx");

            WriteJson(selectionPath, BuildDashboardSelection(excelPath));
            WriteJson(mappingPath, BuildDashboardMapping());
            string pageSize = dashboardPageSizeCombo.SelectedItem == null ? "A4" : dashboardPageSizeCombo.SelectedItem.ToString();

            dashboardStatusText.Text = "dashboard package를 생성합니다...";
            RunPythonTool("dashboard_package.py",
                "--excel " + QuoteArg(excelPath) +
                " --selection " + QuoteArg(selectionPath) +
                " --mapping " + QuoteArg(mappingPath) +
                " --page-size " + QuoteArg(pageSize) +
                " --package-output " + QuoteArg(packagePath) +
                " --preflight-output " + QuoteArg(preflightPath),
                120000);

            dashboardStatusText.Text = "대시보드 PPTX를 생성합니다...";
            string templateArg = "";
            if (!string.IsNullOrWhiteSpace(dashboardTemplateText.Text.Trim()) && File.Exists(dashboardTemplateText.Text.Trim()))
            {
                templateArg = " --template " + QuoteArg(dashboardTemplateText.Text.Trim());
            }
            RunPythonTool("dashboard_writer.py",
                "--package " + QuoteArg(packagePath) +
                " --preflight " + QuoteArg(preflightPath) +
                templateArg +
                " --output " + QuoteArg(outputPath),
                120000);

            lastDashboardOutputPath = outputPath;
            dashboardOpenOutputButton.Enabled = File.Exists(outputPath);
            dashboardStatusText.Text = "대시보드 PPT 생성 완료" + Environment.NewLine +
                                       "선택: " + selectionPath + Environment.NewLine +
                                       "매핑: " + mappingPath + Environment.NewLine +
                                       "Package: " + packagePath + Environment.NewLine +
                                       "Preflight: " + preflightPath + Environment.NewLine +
                                       "PPTX: " + outputPath;
        }

        private Dictionary<string, object> BuildDashboardSelection(string excelPath)
        {
            return new Dictionary<string, object>
            {
                {"workbook_path", excelPath},
                {"selected_sheet", dashboardSheetCombo.SelectedItem.ToString()},
                {"header_row", 1},
                {"entity_name_column", dashboardEntityColumnCombo.SelectedItem.ToString()},
                {"selected_entity_names", SelectedDashboardEntities()},
                {"selected_columns", SelectedDashboardColumns()},
                {"inferred_column_types", new Dictionary<string, object>()},
                {"user_column_types", new Dictionary<string, object>()}
            };
        }

        private Dictionary<string, object> BuildDashboardMapping()
        {
            var profileFields = new List<Dictionary<string, object>>();
            AddProfileIfSelected(profileFields, "업종");
            AddProfileIfSelected(profileFields, "지역");
            AddProfileIfSelected(profileFields, "담당기관");

            var kpis = new List<Dictionary<string, object>>();
            for (int i = 0; i < 6; i++)
            {
                string column = dashboardKpiColumnCombos[i].SelectedItem == null ? "" : dashboardKpiColumnCombos[i].SelectedItem.ToString();
                if (string.IsNullOrWhiteSpace(column))
                {
                    continue;
                }
                kpis.Add(new Dictionary<string, object>
                {
                    {"label", dashboardKpiLabelTexts[i].Text.Trim()},
                    {"value_column", column},
                    {"unit", dashboardKpiUnitTexts[i].Text.Trim()},
                    {"decimals", SelectedDecimalPlaces()}
                });
            }

            var charts = new List<Dictionary<string, object>>();
            for (int i = 0; i < 4; i++)
            {
                List<string> columns = SplitCsv(dashboardChartColumnsTexts[i].Text);
                if (columns.Count == 0)
                {
                    continue;
                }
                charts.Add(new Dictionary<string, object>
                {
                    {"title", dashboardChartTitleTexts[i].Text.Trim()},
                    {"chart_type", dashboardChartTypeCombos[i].SelectedItem == null ? "auto" : dashboardChartTypeCombos[i].SelectedItem.ToString()},
                    {"value_columns", columns},
                    {"category_labels", SplitCsv(dashboardChartLabelsTexts[i].Text)}
                });
            }

            return new Dictionary<string, object>
            {
                {"output_mode", dashboardOutputModeCombo.SelectedIndex == 0 ? "single" : "batch"},
                {"style_preset", SelectedDashboardStylePreset()},
                {"font_family", SelectedDashboardFontFamily()},
                {"profile_fields", profileFields},
                {"kpi_slots", kpis},
                {"chart_slots", charts},
                {"narrative_template", dashboardNarrativeTemplateText.Text.Trim()}
            };
        }

        private string SelectedDashboardStylePreset()
        {
            string value = dashboardDesignCombo.SelectedItem == null ? "" : dashboardDesignCombo.SelectedItem.ToString();
            if (value.IndexOf("민트", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "modern_mint";
            }
            if (value.IndexOf("그래파이트", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "graphite";
            }
            return "modern_blue";
        }

        private string SelectedDashboardFontFamily()
        {
            string value = dashboardFontCombo.SelectedItem == null ? "" : dashboardFontCombo.SelectedItem.ToString();
            if (string.IsNullOrWhiteSpace(value) || value == "맑은 고딕")
            {
                return "Malgun Gothic";
            }
            return value;
        }

        private void AddProfileIfSelected(List<Dictionary<string, object>> profileFields, string column)
        {
            if (SelectedDashboardColumns().Contains(column))
            {
                profileFields.Add(new Dictionary<string, object> { { "label", column }, { "column", column } });
            }
        }

        private void DashboardOpenOutputButton_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(lastDashboardOutputPath) && File.Exists(lastDashboardOutputPath))
            {
                var info = new ProcessStartInfo(lastDashboardOutputPath);
                info.UseShellExecute = true;
                Process.Start(info);
            }
        }

        private void RunPythonTool(string scriptName, string arguments, int timeoutMs)
        {
            string pythonPath = PathResolver.ResolvePythonPath();
            string scriptPath = PathResolver.ResolveEngineToolPath(scriptName);
            if (string.IsNullOrWhiteSpace(pythonPath) || !File.Exists(pythonPath))
            {
                throw new FileNotFoundException("Python 실행 파일을 찾지 못했습니다.");
            }
            if (string.IsNullOrWhiteSpace(scriptPath) || !File.Exists(scriptPath))
            {
                throw new FileNotFoundException("Python 도구를 찾지 못했습니다.", scriptName);
            }

            var startInfo = new ProcessStartInfo();
            startInfo.FileName = pythonPath;
            startInfo.Arguments = QuoteArg(scriptPath) + " " + arguments;
            startInfo.UseShellExecute = false;
            startInfo.CreateNoWindow = true;
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = true;
            startInfo.StandardOutputEncoding = System.Text.Encoding.UTF8;
            startInfo.StandardErrorEncoding = System.Text.Encoding.UTF8;
            StandaloneBundle.ConfigurePythonProcess(startInfo);
            using (Process process = Process.Start(startInfo))
            {
                if (!process.WaitForExit(timeoutMs))
                {
                    try { process.Kill(); } catch { }
                    throw new TimeoutException("Python 도구 실행 시간이 초과되었습니다.");
                }
                string stdout = process.StandardOutput.ReadToEnd();
                string stderr = process.StandardError.ReadToEnd();
                if (process.ExitCode != 0)
                {
                    throw new InvalidOperationException(string.IsNullOrWhiteSpace(stderr) ? stdout.Trim() : stderr.Trim());
                }
            }
        }

        private static void WriteJson(string path, object value)
        {
            var serializer = new JavaScriptSerializer();
            serializer.MaxJsonLength = int.MaxValue;
            File.WriteAllText(path, serializer.Serialize(value), System.Text.Encoding.UTF8);
        }

        private static List<string> SplitCsv(string text)
        {
            var values = new List<string>();
            foreach (string part in (text ?? "").Split(','))
            {
                string value = part.Trim();
                if (value.Length > 0)
                {
                    values.Add(value);
                }
            }
            return values;
        }

        private static void SetChecked(CheckedListBox list, bool check)
        {
            for (int i = 0; i < list.Items.Count; i++)
            {
                list.SetItemChecked(i, check);
            }
        }

        private static int FindColumnIndex(DashboardSheetInfo sheet, params string[] candidates)
        {
            for (int i = 0; i < sheet.Columns.Count; i++)
            {
                foreach (string candidate in candidates)
                {
                    if (sheet.Columns[i].Name.IndexOf(candidate, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return i;
                    }
                }
            }
            return -1;
        }

        private static int GuessComboIndex(ComboBox combo, string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return combo.Items.Count > 0 ? 0 : -1;
            }
            for (int i = 0; i < combo.Items.Count; i++)
            {
                if (combo.Items[i].ToString().IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return i;
                }
            }
            return combo.Items.Count > 0 ? 0 : -1;
        }

        private static string DefaultKpiLabel(int index)
        {
            string[] values = { "매출액", "예상 매출액", "종사자 수", "만족도", "평가점수", "성과점수" };
            return values[index];
        }

        private static string DefaultKpiUnit(int index)
        {
            string[] values = { "원", "원", "명", "%", "점", "점" };
            return values[index];
        }

        private static string DefaultChartTitle(int index)
        {
            string[] values = { "매출/예상매출", "만족/불만족", "평가/성과", "연도별 추이" };
            return values[index];
        }

        private static string DefaultChartColumns(int index)
        {
            string[] values = { "매출액,예상매출액", "만족도,불만족도", "평가점수,성과점수", "매출2022,매출2023,매출2024" };
            return values[index];
        }

        private static string DefaultChartLabels(int index)
        {
            string[] values = { "매출,예상", "만족,불만족", "평가,성과", "2022,2023,2024" };
            return values[index];
        }

        private void SentenceReviewList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (sentenceReviewList.SelectedItems.Count == 0)
            {
                applySentenceEditButton.Enabled = false;
                copySelectedSentenceButton.Enabled = false;
                return;
            }

            DraftSentenceItem sentence = sentenceReviewList.SelectedItems[0].Tag as DraftSentenceItem;
            if (sentence == null)
            {
                return;
            }

            sentenceEditText.Text = sentence.Text;
            applySentenceEditButton.Enabled = true;
            copySelectedSentenceButton.Enabled = true;
        }

        private void ApplySentenceEditButton_Click(object sender, EventArgs e)
        {
            if (sentenceReviewList.SelectedItems.Count == 0)
            {
                return;
            }

            DraftSentenceItem sentence = sentenceReviewList.SelectedItems[0].Tag as DraftSentenceItem;
            if (sentence == null)
            {
                return;
            }

            sentence.Text = sentenceEditText.Text.Trim();
            sentence.IsEdited = true;
            sentenceReviewList.SelectedItems[0].SubItems[1].Text = "수정됨";
            sentenceReviewList.SelectedItems[0].SubItems[3].Text = sentence.Text;
            BuildDraftQaIssues();
            PopulateQaIssues();
            draftPreviewText.Text = BuildReviewedDraftText();
            draftPreviewStatusLabel.Text = draftSentenceItems.Count + "개 문장, " + draftQaIssues.Count + "개 QA 경고";
        }

        private void CopySelectedSentenceButton_Click(object sender, EventArgs e)
        {
            if (sentenceReviewList.SelectedItems.Count == 0)
            {
                return;
            }

            DraftSentenceItem sentence = sentenceReviewList.SelectedItems[0].Tag as DraftSentenceItem;
            if (sentence != null && !string.IsNullOrWhiteSpace(sentence.Text))
            {
                Clipboard.SetText(sentence.Text);
                draftPreviewStatusLabel.Text = "선택 문장을 클립보드에 복사했습니다.";
            }
        }

        private void ExportReviewedDraftButton_Click(object sender, EventArgs e)
        {
            if (draftSentenceItems.Count == 0)
            {
                MessageBox.Show(this, "저장할 문장 초안이 없습니다.", "검토본 저장", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string basePath = currentDraftPath;
            if (string.IsNullOrWhiteSpace(basePath))
            {
                basePath = Path.Combine(Path.GetTempPath(), "report_automation_draft.txt");
            }

            string directory = Path.GetDirectoryName(basePath);
            string name = Path.GetFileNameWithoutExtension(basePath);
            string outputPath = Path.Combine(directory, name + "_reviewed.txt");
            File.WriteAllText(outputPath, BuildReviewedDraftText(), System.Text.Encoding.UTF8);
            draftPreviewStatusLabel.Text = "검토본 저장: " + outputPath;

            var info = new ProcessStartInfo(outputPath);
            info.UseShellExecute = true;
            Process.Start(info);
        }

        private void QaIssueList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (qaIssueList.SelectedItems.Count == 0)
            {
                return;
            }

            DraftQaIssue issue = qaIssueList.SelectedItems[0].Tag as DraftQaIssue;
            if (issue == null || issue.Sentence == null)
            {
                return;
            }

            SelectSentence(issue.Sentence);
            draftReviewTabs.SelectedIndex = 1;
        }

        private void SelectSentence(DraftSentenceItem sentence)
        {
            foreach (ListViewItem item in sentenceReviewList.Items)
            {
                if (object.ReferenceEquals(item.Tag, sentence))
                {
                    item.Selected = true;
                    item.Focused = true;
                    item.EnsureVisible();
                    break;
                }
            }
        }

        private string BuildReviewedDraftText()
        {
            var lines = new List<string>();
            string lastTitle = "";
            foreach (DraftSentenceItem sentence in draftSentenceItems)
            {
                if (sentence.Title != lastTitle)
                {
                    if (lines.Count > 0)
                    {
                        lines.Add("");
                    }
                    lines.Add("▶ " + sentence.Title);
                    lastTitle = sentence.Title;
                }

                lines.Add(sentence.Text);
                if (!string.IsNullOrWhiteSpace(sentence.Source))
                {
                    lines.Add("[" + sentence.Source.Trim('[', ']') + "]");
                }
            }

            return string.Join(Environment.NewLine, lines.ToArray()) + Environment.NewLine;
        }

        private void RefreshHwpStylePresetCombo(string selectedName)
        {
            hwpStylePresetCombo.Items.Clear();
            foreach (Dictionary<string, object> preset in hwpStylePresetStore.LoadAll())
            {
                hwpStylePresetCombo.Items.Add(Convert.ToString(preset["preset_name"]));
            }
            int index = hwpStylePresetCombo.Items.IndexOf(selectedName);
            hwpStylePresetCombo.SelectedIndex = index >= 0 ? index : 0;
        }

        private void HwpStyleSettingsButton_Click(object sender, EventArgs e)
        {
            string selected = hwpStylePresetCombo.SelectedItem == null
                ? HwpStylePresetStore.BuiltInName
                : hwpStylePresetCombo.SelectedItem.ToString();
            using (var dialog = new HwpStylePresetDialog(hwpStylePresetStore, selected))
            {
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    RefreshHwpStylePresetCombo(dialog.SelectedPresetName);
                }
                else
                {
                    RefreshHwpStylePresetCombo(selected);
                }
            }
        }

        private LauncherOptions ReadOptions()
        {
            var options = new LauncherOptions();
            options.WorkbookPath = workbookPathText.Text.Trim();
            options.OutputType = outputTypeCombo.SelectedItem == null ? "" : outputTypeCombo.SelectedItem.ToString();
            options.ReportProfile = reportProfileCombo.SelectedItem == null ? "" : reportProfileCombo.SelectedItem.ToString();
            options.StyleProfile = styleProfileCombo.SelectedItem == null ? "" : styleProfileCombo.SelectedItem.ToString();
            options.HwpTemplatePath = hwpTemplateText.Text.Trim();
            options.PptTemplatePath = pptTemplateText.Text.Trim();
            options.HwpTableStyleProfilePath = hwpTableStyleProfileText.Text.Trim();
            options.BannerSetting = bannerText.Text.Trim();
            options.TitlePrefixes = titlePrefixesText.Text.Trim();
            options.DecimalPlaces = SelectedDecimalPlaces();
            options.ChartOutputMode = chartOutputCombo.SelectedItem == null ? "" : chartOutputCombo.SelectedItem.ToString();
            options.TableInsertMode = tableInsertModeCombo.SelectedItem == null ? "" : tableInsertModeCombo.SelectedItem.ToString();
            options.UseLlm = llmEnabledCheck.Checked;
            options.LlmProvider = llmProviderCombo.SelectedItem == null ? "" : llmProviderCombo.SelectedItem.ToString();
            options.LlmModel = llmModelText.Text.Trim();
            options.LlmApiKey = llmApiKeyText.Text.Trim();
            options.IncludeAnalysis = analysisCheck.Checked;
            options.IncludeCharts = chartCheck.Checked;
            options.IncludeTables = tableCheck.Checked;
            options.IncludeQa = qaCheck.Checked;
            options.GenerateDraftText = draftTextCheck.Checked;
            options.CopyWorkbook = copyWorkbookCheck.Checked;
            options.KeepExcelOpen = keepExcelOpenCheck.Checked;
            options.HwpVisible = hwpVisibleCheck.Checked;
            options.HwpKeepOpenOnError = hwpKeepOpenOnErrorCheck.Checked;
            options.HwpMaxSections = SelectedHwpMaxSections();
            options.HwpDispatchMode = SelectedHwpDispatchMode();
            options.HwpStylePresetName = hwpStylePresetCombo.SelectedItem == null
                ? HwpStylePresetStore.BuiltInName
                : hwpStylePresetCombo.SelectedItem.ToString();
            options.HwpStyleConfig = hwpStylePresetStore.Get(options.HwpStylePresetName);
            options.HwpPageSetup = hwpPageSetup == null ? null : HwpPageSetup.Normalize(hwpPageSetup);
            options.TableRangeOverrides = string.Join(Environment.NewLine, currentTablePreviews
                .Where(table => table.IsManual)
                .Select(table => table.TableKey + "=" + table.SheetName + "!" + table.FinalRange)
                .ToArray());
            options.Validate();
            return options;
        }

        private void HwpEnvironmentCheckButton_Click(object sender, EventArgs e)
        {
            hwpEnvironmentCheckButton.Enabled = false;
            string reportPath = Path.Combine(ResolveHwpDiagnosticOutputDirectory(), "hwp_environment_report_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".json");
            var options = new LauncherOptions();
            options.HwpVisible = hwpVisibleCheck.Checked;
            options.HwpDispatchMode = SelectedHwpDispatchMode();
            resultSummaryText.Text = "HWP COM environment diagnostics running..." + Environment.NewLine + reportPath;
            Log("HWP COM environment diagnostics started.");

            var thread = new Thread(delegate()
            {
                try
                {
                    string summary = EngineRunner.TryRunHwpEnvironmentDiagnostics(options, reportPath, Log);
                    BeginInvoke(new Action(delegate()
                    {
                        resultSummaryText.Text = summary;
                        openHwpReportButton.Tag = reportPath;
                        openHwpReportButton.Enabled = File.Exists(reportPath);
                    }));
                }
                catch (Exception ex)
                {
                    Log("HWP COM diagnostics failed: " + ex.Message);
                    BeginInvoke(new Action(delegate()
                    {
                        resultSummaryText.Text = "HWP COM diagnostics failed." + Environment.NewLine + ex.Message;
                    }));
                }
                finally
                {
                    BeginInvoke(new Action(delegate()
                    {
                        hwpEnvironmentCheckButton.Enabled = true;
                    }));
                }
            });
            thread.IsBackground = true;
            thread.Start();
        }

        private string ResolveHwpDiagnosticOutputDirectory()
        {
            string workbookPath = workbookPathText.Text.Trim();
            if (!string.IsNullOrWhiteSpace(workbookPath))
            {
                try
                {
                    string directory = Path.GetDirectoryName(Path.GetFullPath(workbookPath));
                    if (!string.IsNullOrWhiteSpace(directory) && Directory.Exists(directory))
                    {
                        return directory;
                    }
                }
                catch
                {
                }
            }
            return Directory.Exists(Environment.CurrentDirectory) ? Environment.CurrentDirectory : AppDomain.CurrentDomain.BaseDirectory;
        }

        private string SelectedHwpDispatchMode()
        {
            return hwpDispatchModeCombo.SelectedItem == null ? "ensure_dispatch" : hwpDispatchModeCombo.SelectedItem.ToString();
        }

        private int SelectedHwpMaxSections()
        {
            string value = hwpMaxSectionsCombo.SelectedItem == null ? "" : hwpMaxSectionsCombo.SelectedItem.ToString();
            if (value.StartsWith("3", StringComparison.OrdinalIgnoreCase))
            {
                return 3;
            }
            if (value.IndexOf("전체", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return 0;
            }
            return 1;
        }

        private int SelectedDecimalPlaces()
        {
            if (decimalPlacesCombo.SelectedIndex < 0)
            {
                return 1;
            }
            return decimalPlacesCombo.SelectedIndex;
        }

        private void Log(string message)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action<string>(Log), message);
                return;
            }
            logText.AppendText(DateTime.Now.ToString("HH:mm:ss") + "  " + message + Environment.NewLine);
        }
    }

    internal sealed class LauncherOptions
    {
        public string WorkbookPath;
        public string AddinPath;
        public string OutputType = "Excel 산출 시트";
        public string ReportProfile = "인식도/만족도 조사형";
        public string StyleProfile = "공식 보고서체";
        public string HwpTemplatePath;
        public string PptTemplatePath;
        public string HwpTableStyleProfilePath;
        public string HwpStylePresetName = HwpStylePresetStore.BuiltInName;
        public string HwpStyleConfigPath;
        public Dictionary<string, object> HwpStyleConfig;
        public Dictionary<string, object> HwpPageSetup;
        public string HwpPageSetupPath;
        public string BannerSetting = "전체";
        public string TitlePrefixes = "";
        public int DecimalPlaces = 1;
        public string ChartOutputMode = "Excel 차트 데이터";
        public string TableInsertMode = "Excel 삽입표 시트";
        public bool UseLlm;
        public string LlmProvider = "OpenAI";
        public string LlmModel = "gpt-4.1-mini";
        public string LlmApiKey;
        public bool IncludeAnalysis = true;
        public bool IncludeCharts = true;
        public bool IncludeTables = true;
        public bool IncludeQa = true;
        public bool GenerateDraftText = true;
        public bool CopyWorkbook = true;
        public bool KeepExcelOpen;
        public bool HwpVisible;
        public bool HwpKeepOpenOnError;
        public int HwpMaxSections = 1;
        public string HwpDispatchMode = "ensure_dispatch";
        public string TableRangeOverrides = "";
        public string LastGeneratedWorkbookPath;
        public string LastDraftTextPath;
        public string LastReportPackagePath;
        public string LastPreflightReportPath;
        public string LastHwpOutputPath;
        public string LastHwpWriterReportPath;

        public void Validate()
        {
            if (string.IsNullOrWhiteSpace(WorkbookPath))
            {
                throw new InvalidOperationException("집계표 엑셀 파일을 선택하세요.");
            }
            WorkbookPath = Path.GetFullPath(WorkbookPath);
            if (!File.Exists(WorkbookPath))
            {
                throw new FileNotFoundException("집계표 엑셀 파일을 찾을 수 없습니다.", WorkbookPath);
            }
            AddinPath = Path.GetFullPath(string.IsNullOrWhiteSpace(AddinPath) ? PathResolver.ResolveDefaultAddinPath() : AddinPath);
            if (!string.IsNullOrWhiteSpace(HwpTemplatePath))
            {
                HwpTemplatePath = Path.GetFullPath(HwpTemplatePath);
            }
            if (!string.IsNullOrWhiteSpace(PptTemplatePath))
            {
                PptTemplatePath = Path.GetFullPath(PptTemplatePath);
            }
            if (!string.IsNullOrWhiteSpace(HwpTableStyleProfilePath))
            {
                HwpTableStyleProfilePath = Path.GetFullPath(HwpTableStyleProfilePath);
                if (!File.Exists(HwpTableStyleProfilePath))
                {
                    throw new FileNotFoundException("HWP 표 스타일 profile 파일을 찾을 수 없습니다.", HwpTableStyleProfilePath);
                }
            }
            if (!string.IsNullOrWhiteSpace(HwpStyleConfigPath))
            {
                HwpStyleConfigPath = Path.GetFullPath(HwpStyleConfigPath);
                if (!File.Exists(HwpStyleConfigPath))
                    throw new FileNotFoundException("HWP 서식 설정 파일을 찾을 수 없습니다.", HwpStyleConfigPath);
                var serializer = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
                HwpStyleConfig = HwpStylePresetStore.Normalize(
                    serializer.DeserializeObject(File.ReadAllText(HwpStyleConfigPath, System.Text.Encoding.UTF8)) as Dictionary<string, object>);
            }
            if (HwpStyleConfig != null)
            {
                HwpStyleConfig = HwpStylePresetStore.Normalize(HwpStyleConfig);
                HwpStylePresetName = Convert.ToString(HwpStyleConfig["preset_name"]);
            }
            if (HwpPageSetup != null)
                HwpPageSetup = ReportAutomationLauncher.HwpPageSetup.Normalize(HwpPageSetup);
            else
                HwpPageSetupPath = null;
            if (OutputType.IndexOf("HWP", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                if (string.IsNullOrWhiteSpace(HwpTemplatePath))
                {
                    throw new InvalidOperationException("HWPX 보고서 생성을 위해 HWPX 템플릿을 선택하세요.");
                }
                if (!File.Exists(HwpTemplatePath))
                {
                    throw new FileNotFoundException("HWPX 템플릿 파일을 찾을 수 없습니다.", HwpTemplatePath);
                }
            }
            else if (!OutputType.StartsWith("Excel", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("현재 런처 직접 생성은 Excel 산출 시트와 HWPX 보고서만 지원합니다.");
            }
            if (string.IsNullOrWhiteSpace(BannerSetting))
            {
                BannerSetting = "전체";
            }
            if (string.IsNullOrWhiteSpace(ReportProfile))
            {
                ReportProfile = "인식도/만족도 조사형";
            }
            if (string.IsNullOrWhiteSpace(StyleProfile))
            {
                StyleProfile = "공식 보고서체";
            }
            if (string.IsNullOrWhiteSpace(TitlePrefixes))
            {
                TitlePrefixes = "";
            }
            if (DecimalPlaces < 0 || DecimalPlaces > 2)
            {
                DecimalPlaces = 1;
            }
            if (string.IsNullOrWhiteSpace(ChartOutputMode))
            {
                ChartOutputMode = "Excel 차트 데이터";
            }
            if (string.IsNullOrWhiteSpace(TableInsertMode))
            {
                TableInsertMode = "Excel 삽입표 시트";
            }
            if (string.IsNullOrWhiteSpace(LlmProvider))
            {
                LlmProvider = "OpenAI";
            }
            if (string.IsNullOrWhiteSpace(LlmModel))
            {
                LlmModel = "gpt-4.1-mini";
            }
            HwpDispatchMode = NormalizeHwpDispatchMode(HwpDispatchMode);
        }

        public static LauncherOptions FromArgs(string[] args)
        {
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var flags = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < args.Length; i++)
            {
                string arg = args[i];
                if (!arg.StartsWith("--", StringComparison.Ordinal))
                {
                    continue;
                }
                string key = arg.Substring(2);
                if (i + 1 < args.Length && !args[i + 1].StartsWith("--", StringComparison.Ordinal))
                {
                    values[key] = args[++i];
                }
                else
                {
                    flags.Add(key);
                }
            }

            var options = new LauncherOptions();
            options.WorkbookPath = Get(values, "workbook", "");
            options.AddinPath = Get(values, "addin", "");
            options.OutputType = Get(values, "output", "Excel 산출 시트");
            options.ReportProfile = Get(values, "report-profile", "인식도/만족도 조사형");
            options.StyleProfile = Get(values, "style-profile", "공식 보고서체");
            options.HwpTemplatePath = Get(values, "hwp-template", "");
            options.PptTemplatePath = Get(values, "ppt-template", "");
            options.HwpTableStyleProfilePath = Get(values, "hwp-table-style-profile", "");
            options.HwpStyleConfigPath = Get(values, "hwp-style-config", "");
            options.HwpStylePresetName = Get(values, "hwp-style-preset", HwpStylePresetStore.BuiltInName);
            if (flags.Contains("hwp-page-setup"))
                throw new InvalidOperationException("--hwp-page-setup <path> 값을 지정하세요.");
            options.HwpPageSetupPath = Get(values, "hwp-page-setup", "");
            if (values.ContainsKey("hwp-page-setup"))
            {
                if (string.IsNullOrWhiteSpace(options.HwpPageSetupPath))
                    throw new InvalidOperationException("--hwp-page-setup <path> 값을 지정하세요.");
                options.HwpPageSetupPath = Path.GetFullPath(options.HwpPageSetupPath);
                options.HwpPageSetup = ReportAutomationLauncher.HwpPageSetup.Load(options.HwpPageSetupPath);
            }
            options.BannerSetting = Get(values, "banner", "전체");
            options.TitlePrefixes = Get(values, "prefixes", "");
            options.DecimalPlaces = GetInt(values, "decimal-places", 1);
            options.ChartOutputMode = Get(values, "chart-output", "Excel 차트 데이터");
            options.TableInsertMode = Get(values, "table-insert", "Excel 삽입표 시트");
            options.UseLlm = flags.Contains("use-llm");
            options.LlmProvider = Get(values, "llm-provider", "OpenAI");
            options.LlmModel = Get(values, "llm-model", "gpt-4.1-mini");
            options.GenerateDraftText = !flags.Contains("no-draft");
            options.CopyWorkbook = !flags.Contains("no-copy");
            options.KeepExcelOpen = flags.Contains("keep-open");
            options.HwpVisible = flags.Contains("hwp-visible");
            options.HwpKeepOpenOnError = flags.Contains("hwp-keep-open-on-error");
            options.HwpMaxSections = GetInt(values, "hwp-max-sections", 1);
            options.HwpDispatchMode = Get(values, "hwp-dispatch-mode", "ensure_dispatch");
            if (options.OutputType.IndexOf("HWP", StringComparison.OrdinalIgnoreCase) >= 0 && string.IsNullOrWhiteSpace(options.HwpStyleConfigPath))
            {
                options.HwpStyleConfig = new HwpStylePresetStore(recoveryLog: Console.Error.WriteLine).Get(options.HwpStylePresetName);
            }
            options.Validate();
            return options;
        }

        private static string NormalizeHwpDispatchMode(string value)
        {
            string mode = string.IsNullOrWhiteSpace(value) ? "ensure_dispatch" : value.Trim().ToLowerInvariant().Replace("-", "_");
            if (mode == "ensure_dispatch" || mode == "dispatch" || mode == "dispatch_ex")
            {
                return mode;
            }
            return "ensure_dispatch";
        }

        private static string Get(Dictionary<string, string> values, string key, string fallback)
        {
            string value;
            return values.TryGetValue(key, out value) ? value : fallback;
        }

        private static int GetInt(Dictionary<string, string> values, string key, int fallback)
        {
            string value;
            int parsed;
            if (values.TryGetValue(key, out value) && int.TryParse(value, out parsed))
            {
                return parsed;
            }
            return fallback;
        }
    }

    internal sealed class DraftSentenceItem
    {
        public int Index;
        public string Title = "";
        public string Text = "";
        public string Source = "";
        public bool IsEdited;
    }

    internal sealed class DraftQaIssue
    {
        public DraftSentenceItem Sentence;
        public string Type = "";
        public string Message = "";
    }

    internal sealed class DashboardWorkbookInfo
    {
        public string Workbook = "";
        public readonly List<DashboardSheetInfo> Sheets = new List<DashboardSheetInfo>();

        public static DashboardWorkbookInfo Load(string path)
        {
            var serializer = new JavaScriptSerializer();
            serializer.MaxJsonLength = int.MaxValue;
            var root = serializer.DeserializeObject(File.ReadAllText(path, System.Text.Encoding.UTF8)) as Dictionary<string, object>;
            var info = new DashboardWorkbookInfo();
            if (root == null)
            {
                return info;
            }
            info.Workbook = GetString(root, "workbook");
            object sheetsValue;
            if (root.TryGetValue("sheets", out sheetsValue))
            {
                object[] sheets = sheetsValue as object[];
                if (sheets != null)
                {
                    foreach (object sheetObj in sheets)
                    {
                        Dictionary<string, object> rawSheet = sheetObj as Dictionary<string, object>;
                        if (rawSheet == null)
                        {
                            continue;
                        }
                        var sheet = new DashboardSheetInfo();
                        sheet.Name = GetString(rawSheet, "name");
                        LoadColumns(sheet, rawSheet);
                        LoadPreview(sheet, rawSheet);
                        info.Sheets.Add(sheet);
                    }
                }
            }
            return info;
        }

        private static void LoadColumns(DashboardSheetInfo sheet, Dictionary<string, object> rawSheet)
        {
            object value;
            if (!rawSheet.TryGetValue("columns", out value))
            {
                return;
            }
            object[] columns = value as object[];
            if (columns == null)
            {
                return;
            }
            foreach (object columnObj in columns)
            {
                Dictionary<string, object> rawColumn = columnObj as Dictionary<string, object>;
                if (rawColumn == null)
                {
                    continue;
                }
                var column = new DashboardColumnInfo();
                column.Name = GetString(rawColumn, "name");
                column.InferredType = GetString(rawColumn, "inferred_type");
                column.Sample = GetString(rawColumn, "sample");
                column.MissingCount = GetInt(rawColumn, "missing_count");
                sheet.Columns.Add(column);
            }
        }

        private static void LoadPreview(DashboardSheetInfo sheet, Dictionary<string, object> rawSheet)
        {
            object value;
            if (!rawSheet.TryGetValue("preview", out value))
            {
                return;
            }
            object[] rows = value as object[];
            if (rows == null)
            {
                return;
            }
            foreach (object rowObj in rows)
            {
                Dictionary<string, object> rawRow = rowObj as Dictionary<string, object>;
                if (rawRow != null)
                {
                    sheet.Preview.Add(rawRow);
                }
            }
        }

        private static string GetString(Dictionary<string, object> values, string key)
        {
            object value;
            return values.TryGetValue(key, out value) && value != null ? Convert.ToString(value) : "";
        }

        private static int GetInt(Dictionary<string, object> values, string key)
        {
            object value;
            if (!values.TryGetValue(key, out value) || value == null)
            {
                return 0;
            }
            try
            {
                return Convert.ToInt32(value);
            }
            catch
            {
                return 0;
            }
        }
    }

    internal sealed class DashboardSheetInfo
    {
        public string Name = "";
        public readonly List<DashboardColumnInfo> Columns = new List<DashboardColumnInfo>();
        public readonly List<Dictionary<string, object>> Preview = new List<Dictionary<string, object>>();
    }

    internal sealed class DashboardColumnInfo
    {
        public string Name = "";
        public string InferredType = "";
        public string Sample = "";
        public int MissingCount;
    }

    internal static class AutomationRunner
    {
        public static string Run(LauncherOptions options, Action<string> log)
        {
            options.Validate();
            ExcelAddinDiagnostics.RequireFile(options.AddinPath);
            string workbookToOpen = options.WorkbookPath;
            bool legacyFormat = Path.GetExtension(workbookToOpen).Equals(".xls", StringComparison.OrdinalIgnoreCase);
            bool originalEnableEvents = true;
            log("추가기능: " + options.AddinPath);

            object excel = null;
            object workbook = null;
            object addin = null;

            try
            {
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    throw new InvalidOperationException("Excel.Application COM 개체를 찾을 수 없습니다. Microsoft Excel 설치 상태를 확인하세요.");
                }

                excel = Activator.CreateInstance(excelType);
                dynamic xl = excel;
                xl.DisplayAlerts = false;
                xl.Visible = false;

                log("Excel을 시작했습니다.");
                // Validate actual Excel access before copying/opening the user's input workbook.
                addin = ExcelAddinDiagnostics.Open(excel, options.AddinPath);
                log("Excel 추가기능을 열었습니다.");
                if (legacyFormat)
                {
                    // The cell export can exceed the 65,536 rows of an Excel 97–2003 sheet.
                    int previousSecurity = xl.AutomationSecurity;
                    originalEnableEvents = xl.EnableEvents;
                    xl.EnableEvents = false;
                    try
                    {
                        xl.AutomationSecurity = 3; // Do not run macros in the source workbook or its copy.
                        workbook = xl.Workbooks.Open(options.WorkbookPath, 0, true);
                        dynamic legacy = workbook;
                        bool hasMacros = legacy.HasVBProject;
                        workbookToOpen = WorkingCopyPath(options.WorkbookPath, hasMacros ? ".xlsm" : ".xlsx");
                        if (File.Exists(workbookToOpen)) throw new IOException("작업 사본이 이미 있습니다: " + workbookToOpen);
                        legacy.SaveAs(workbookToOpen, hasMacros ? 52 : 51);
                        // SaveAs alone leaves the workbook in 65,536-row compatibility mode.
                        legacy.Close(false);
                        ReleaseComForLauncher(workbook);
                        workbook = null;
                        workbook = xl.Workbooks.Open(workbookToOpen, 0, false);
                        dynamic modern = workbook;
                        if ((int)modern.Worksheets.Item(1).Rows.Count <= 65536)
                            throw new InvalidOperationException("XLS 작업 사본이 Excel 호환 모드에서 벗어나지 못했습니다.");
                    }
                    finally
                    {
                        xl.AutomationSecurity = previousSecurity;
                    }
                    log("구형 XLS를 현대 Excel 작업 사본으로 변환했습니다. 원본은 변경하지 않았습니다.");
                }
                else
                {
                    workbookToOpen = options.CopyWorkbook ? CreateWorkingCopy(options.WorkbookPath) : options.WorkbookPath;
                    workbook = xl.Workbooks.Open(workbookToOpen);
                }
                log("대상 파일: " + workbookToOpen);
                dynamic wb = workbook;
                if ((bool)wb.ReadOnly) throw new IOException("산출할 Excel 파일이 읽기 전용으로 열렸습니다: " + workbookToOpen);
                wb.Activate();

                string macroName = "'" + Path.GetFileName(options.AddinPath).Replace("'", "''") + "'!ReportAutomation_RunWithOptionsAndRangesSilent";
                log("보고서 자동화 매크로를 실행합니다.");
                object result;
                try { result = xl.Run(macroName, options.BannerSetting, options.TitlePrefixes, options.TableRangeOverrides); }
                catch (COMException ex)
                {
                    throw new InvalidOperationException("Excel 산출 매크로 실행 실패 (0x" + ex.ErrorCode.ToString("X8") + "): " + ex.Message, ex);
                }
                if (!(result is bool) || !(bool)result)
                    throw new InvalidOperationException("Excel 산출 시트 생성 실패: " + (result as string ?? "매크로 완료 응답을 받지 못했습니다."));

                wb.Save();
                WriteLauncherConfig(workbookToOpen, options);
                log("저장 완료");

                if (options.KeepExcelOpen)
                {
                    try
                    {
                        dynamic addinBook = addin;
                        addinBook.Close(false);
                        addin = null;
                    }
                    catch
                    {
                    }
                    xl.DisplayAlerts = true;
                    xl.Visible = true;
                    wb.Activate();
                    if (legacyFormat) xl.EnableEvents = originalEnableEvents;
                    log("Excel 창을 열어둡니다.");
                    workbook = null;
                    excel = null;
                }
                else
                {
                    dynamic addinBook = addin;
                    addinBook.Close(false);
                    addin = null;
                    wb.Close(true);
                    workbook = null;
                    xl.Quit();
                    excel = null;
                    log("Excel을 종료했습니다.");
                }
            }
            finally
            {
                CloseComWorkbookForLauncher(addin, false);
                // Success saves explicitly above. Never save partial changes while unwinding a failure.
                CloseComWorkbookForLauncher(workbook, false);
                QuitComExcelForLauncher(excel);
                ReleaseComForLauncher(addin);
                ReleaseComForLauncher(workbook);
                ReleaseComForLauncher(excel);
            }

            return workbookToOpen;
        }

        private static string CreateWorkingCopy(string sourcePath)
        {
            string copyPath = WorkingCopyPath(sourcePath, Path.GetExtension(sourcePath));
            File.Copy(sourcePath, copyPath, true);
            return copyPath;
        }

        private static string WorkingCopyPath(string sourcePath, string extension)
        {
            return Path.Combine(Path.GetDirectoryName(sourcePath),
                Path.GetFileNameWithoutExtension(sourcePath) + "_report_alpha_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + extension);
        }

        internal static void WriteLauncherConfig(string workbookPath, LauncherOptions options)
        {
            string path = Path.Combine(Path.GetDirectoryName(workbookPath), Path.GetFileNameWithoutExtension(workbookPath) + "_launcher_config.txt");
            using (var writer = new StreamWriter(path, false, System.Text.Encoding.UTF8))
            {
                writer.WriteLine("Workbook=" + workbookPath);
                writer.WriteLine("GeneratedWorkbook=" + workbookPath);
                writer.WriteLine("Addin=" + options.AddinPath);
                writer.WriteLine("OutputType=" + options.OutputType);
                writer.WriteLine("ReportProfile=" + options.ReportProfile);
                writer.WriteLine("StyleProfile=" + options.StyleProfile);
                writer.WriteLine("HwpTemplate=" + options.HwpTemplatePath);
                writer.WriteLine("PptTemplate=" + options.PptTemplatePath);
                writer.WriteLine("HwpTableStyleProfile=" + options.HwpTableStyleProfilePath);
                writer.WriteLine("HwpStylePreset=" + options.HwpStylePresetName);
                writer.WriteLine("HwpStyleConfig=" + options.HwpStyleConfigPath);
                writer.WriteLine("HwpPageSetup=" + (options.HwpPageSetup == null ? "" : options.HwpPageSetupPath));
                writer.WriteLine("BannerSetting=" + options.BannerSetting);
                writer.WriteLine("TitlePrefixes=" + options.TitlePrefixes);
                writer.WriteLine("DecimalPlaces=" + options.DecimalPlaces);
                writer.WriteLine("ChartOutputMode=" + options.ChartOutputMode);
                writer.WriteLine("TableInsertMode=" + options.TableInsertMode);
                writer.WriteLine("HwpMaxSections=" + options.HwpMaxSections);
                writer.WriteLine("HwpDispatchMode=" + options.HwpDispatchMode);
                writer.WriteLine("TableRangeOverrides=" + (options.TableRangeOverrides ?? "").Replace("\r", "").Replace("\n", ";"));
                writer.WriteLine("UseLlm=" + options.UseLlm);
                writer.WriteLine("LlmProvider=" + options.LlmProvider);
                writer.WriteLine("LlmModel=" + options.LlmModel);
                writer.WriteLine("LlmApiKeyConfigured=" + !string.IsNullOrWhiteSpace(options.LlmApiKey));
                writer.WriteLine("IncludeAnalysis=" + options.IncludeAnalysis);
                writer.WriteLine("IncludeCharts=" + options.IncludeCharts);
                writer.WriteLine("IncludeTables=" + options.IncludeTables);
                writer.WriteLine("IncludeQa=" + options.IncludeQa);
                writer.WriteLine("GenerateDraftText=" + options.GenerateDraftText);
                writer.WriteLine("DraftText=" + options.LastDraftTextPath);
                writer.WriteLine("ReportPackage=" + options.LastReportPackagePath);
                writer.WriteLine("PreflightReport=" + options.LastPreflightReportPath);
                writer.WriteLine("CreatedAt=" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            }
        }

        internal static void CloseComWorkbookForLauncher(object workbook, bool save)
        {
            if (workbook == null)
            {
                return;
            }
            try
            {
                dynamic wb = workbook;
                wb.Close(save);
            }
            catch
            {
            }
        }

        internal static void QuitComExcelForLauncher(object excel)
        {
            if (excel == null)
            {
                return;
            }
            try
            {
                dynamic xl = excel;
                xl.Quit();
            }
            catch
            {
            }
        }

        internal static void ReleaseComForLauncher(object value)
        {
            if (value == null)
            {
                return;
            }
            try
            {
                if (Marshal.IsComObject(value))
                {
                    Marshal.FinalReleaseComObject(value);
                }
            }
            catch
            {
            }
        }
    }

    internal static class EngineRunner
    {
        internal static void RunWorkflow(LauncherOptions options, Action<string> log)
        {
            bool hwp = options.OutputType.IndexOf("HWP", StringComparison.OrdinalIgnoreCase) >= 0;
            if (hwp)
            {
                log("한글 양식의 본문 위치를 확인합니다.");
                RequireHwpTemplate(options.HwpTemplatePath);
            }

            log("1단계: 원본 집계표에서 Excel 산출 시트를 자동 생성합니다.");
            string workbookPath = AutomationRunner.Run(options, log);
            options.LastGeneratedWorkbookPath = workbookPath;
            log("2단계: 생성된 산출 시트에서 분석문·표를 읽어 QA를 검사합니다.");
            TryGenerateReportPackage(workbookPath, options, log);
            if (hwp)
            {
                RequireHwpPreflight(options);
                log("3단계: QA를 통과한 산출 데이터로 한글 보고서를 만듭니다.");
                TryGenerateHwpDocument(options, log);
                if (string.IsNullOrWhiteSpace(options.LastHwpOutputPath))
                    throw new InvalidOperationException("한글 보고서를 저장하지 못했습니다. 진행 기록과 한글 생성 기록을 확인하세요.");
            }
            if (options.GenerateDraftText)
                options.LastDraftTextPath = TryGenerateDraft(workbookPath, log);
            AutomationRunner.WriteLauncherConfig(workbookPath, options);
        }

        private static void RequireHwpTemplate(string path)
        {
            if (!Path.GetExtension(path).Equals(".hwpx", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("한글 양식을 .hwpx 사본으로 저장하고 본문 삽입 위치에 {{BODY}}를 지정하세요.");
            string reportPath = Path.GetTempFileName();
            try
            {
                MainForm.RunTemplateTool("template_inspector.py",
                    "--template " + Quote(path) + " --type hwpx_report --output " + Quote(reportPath), 30000);
                string status = GetString(ReadWorkflowReport(reportPath), "status");
                if (status != "ready" && status != "usable_with_warnings")
                    throw new InvalidOperationException("한글 양식에 본문 삽입 위치 {{BODY}}가 없습니다. 양식 검사 또는 보정 사본 저장을 사용하세요.");
            }
            finally
            {
                try { File.Delete(reportPath); } catch (IOException) { } catch (UnauthorizedAccessException) { }
            }
        }

        private static void RequireHwpPreflight(LauncherOptions options)
        {
            if (!File.Exists(options.LastReportPackagePath) || !File.Exists(options.LastPreflightReportPath))
                throw new InvalidOperationException("Excel 산출은 완료됐지만 QA 검사 결과를 만들지 못했습니다. 진행 기록을 확인하세요.");
            var report = ReadWorkflowReport(options.LastPreflightReportPath);
            string status = GetString(report, "status");
            if (status == "ready" || status == "ready_with_warnings") return;
            object value;
            var errors = report.TryGetValue("errors", out value) ? value as object[] : null;
            var messages = (errors ?? new object[0]).OfType<Dictionary<string, object>>()
                .Select(error => (GetString(error, "table_key") + " " + GetString(error, "message")).Trim())
                .Where(message => message.Length > 0).Distinct().Take(3).ToArray();
            throw new InvalidOperationException("QA에서 한글 보고서 생성을 중단했습니다." +
                (messages.Length == 0 ? " QA 검사 기록을 확인하세요." : Environment.NewLine + string.Join(Environment.NewLine, messages)));
        }

        private static Dictionary<string, object> ReadWorkflowReport(string path)
        {
            var serializer = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            var report = serializer.DeserializeObject(File.ReadAllText(path, System.Text.Encoding.UTF8)) as Dictionary<string, object>;
            if (report == null) throw new InvalidOperationException("검사 결과가 비어 있습니다: " + path);
            return report;
        }

        internal static void RunHwpPageSetupSelfCheck()
        {
            Dictionary<string, object> defaults = HwpPageSetup.Normalize(HwpPageSetup.DefaultConfig());
            double[] expectedDefaults = { 210, 297, 20, 15, 24.7, 25, 0, 10, 15 };
            CheckPageSetup(defaults.Count == 10 && (string)defaults["orientation"] == "portrait", "exact default schema");
            for (int i = 0; i < HwpPageSetup.NumberKeys.Length; i++)
                CheckPageSetup((double)defaults[HwpPageSetup.NumberKeys[i]] == expectedDefaults[i], "default " + HwpPageSetup.NumberKeys[i]);
            foreach (string key in HwpPageSetup.NumberKeys)
            {
                foreach (object bad in new object[] { null, true, "1", double.NaN, double.PositiveInfinity, double.NegativeInfinity, -0.1 })
                {
                    var invalid = new Dictionary<string, object>(defaults);
                    invalid[key] = bad;
                    RejectPageSetup(delegate { HwpPageSetup.Normalize(invalid); }, "invalid " + key);
                }
            }
            foreach (string key in new[] { "paper_width_mm", "paper_height_mm" })
            {
                var invalid = new Dictionary<string, object>(defaults);
                invalid[key] = 0.0;
                RejectPageSetup(delegate { HwpPageSetup.Normalize(invalid); }, "zero " + key);
            }
            var changed = new Dictionary<string, object>(defaults);
            changed.Remove("top_mm");
            RejectPageSetup(delegate { HwpPageSetup.Normalize(changed); }, "missing key");
            changed = new Dictionary<string, object>(defaults);
            changed["scope"] = "body";
            RejectPageSetup(delegate { HwpPageSetup.Normalize(changed); }, "extra key");
            foreach (object bad in new object[] { null, true, "Landscape", "" })
            {
                changed = new Dictionary<string, object>(defaults);
                changed["orientation"] = bad;
                RejectPageSetup(delegate { HwpPageSetup.Normalize(changed); }, "invalid orientation");
            }
            changed = new Dictionary<string, object>(defaults);
            changed["header_mm"] = 20.0;
            CheckPageSetup(HwpPageSetup.Normalize(changed) != null, "header/footer equality accepted");
            changed["header_mm"] = 20.01;
            RejectPageSetup(delegate { HwpPageSetup.Normalize(changed); }, "header exceeds top");
            changed["header_mm"] = 20.0;
            changed["footer_mm"] = 15.01;
            RejectPageSetup(delegate { HwpPageSetup.Normalize(changed); }, "footer exceeds bottom");
            changed = new Dictionary<string, object>(defaults);
            changed["paper_height_mm"] = 36.0;
            CheckPageSetup(HwpPageSetup.Normalize(changed) != null, "header/footer not subtracted twice");
            changed["paper_height_mm"] = 35.0;
            RejectPageSetup(delegate { HwpPageSetup.Normalize(changed); }, "zero body height");
            changed["paper_height_mm"] = 34.0;
            RejectPageSetup(delegate { HwpPageSetup.Normalize(changed); }, "negative body height");
            changed = new Dictionary<string, object>(defaults);
            changed["left_mm"] = 100.0;
            changed["right_mm"] = 100.0;
            changed["gutter_mm"] = 10.0;
            RejectPageSetup(delegate { HwpPageSetup.Normalize(changed); }, "gutter consumes body width");
            changed["gutter_mm"] = 11.0;
            RejectPageSetup(delegate { HwpPageSetup.Normalize(changed); }, "negative body width");
            changed["left_mm"] = 200.0;
            changed["right_mm"] = 50.0;
            changed["gutter_mm"] = 0.0;
            changed["orientation"] = "landscape";
            Dictionary<string, object> landscape = HwpPageSetup.Normalize(changed);
            CheckPageSetup((double)landscape["paper_width_mm"] == 210 && (double)landscape["paper_height_mm"] == 297, "landscape retains base dimensions");
            changed["orientation"] = "portrait";
            RejectPageSetup(delegate { HwpPageSetup.Normalize(changed); }, "orientation swaps effective width");
            changed = new Dictionary<string, object>(defaults);
            changed["top_mm"] = 110.0;
            changed["bottom_mm"] = 110.0;
            CheckPageSetup(HwpPageSetup.Normalize(changed) != null, "portrait effective height");
            changed["orientation"] = "landscape";
            RejectPageSetup(delegate { HwpPageSetup.Normalize(changed); }, "orientation swaps effective height");
            changed = new Dictionary<string, object>(defaults);
            foreach (string key in HwpPageSetup.NumberKeys.Skip(2)) changed[key] = 0.0;
            CheckPageSetup(HwpPageSetup.Normalize(changed) != null, "all margins may be zero");
            changed = new Dictionary<string, object>(defaults);
            changed["paper_width_mm"] = 49.700001;
            RejectPageSetup(delegate { HwpPageSetup.Normalize(changed); }, "rounded body width is zero");
            changed = new Dictionary<string, object>(defaults);
            changed["paper_height_mm"] = 35.000001;
            RejectPageSetup(delegate { HwpPageSetup.Normalize(changed); }, "rounded body height is zero");
            foreach (string key in HwpPageSetup.NumberKeys)
            {
                changed = new Dictionary<string, object>(defaults);
                changed[key] = (int.MaxValue + 1.0) * 25.4 / 7200.0;
                RejectPageSetup(delegate { HwpPageSetup.Normalize(changed); }, "signed-I4 overflow " + key);
            }
            changed = new Dictionary<string, object>(defaults);
            changed["paper_width_mm"] = int.MaxValue * 25.4 / 7200.0;
            CheckPageSetup(HwpPageSetup.Normalize(changed) != null, "signed-I4 equality accepted");

            string directory = Path.Combine(Path.GetTempPath(), "ResearchHelper PageSetup Check " + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try
            {
                var options = new LauncherOptions { LastReportPackagePath = Path.Combine(directory, "report_package.json") };
                string emitted = Path.Combine(directory, "hwp_page_setup.json");
                PrepareHwpPageSetup(options);
                CheckPageSetup(options.HwpPageSetup == null && options.HwpPageSetupPath == null && !File.Exists(emitted), "fresh preserve mode does not emit JSON");
                options.HwpPageSetup = defaults;
                PrepareHwpPageSetup(options);
                CheckPageSetup(options.HwpPageSetupPath == emitted && File.Exists(emitted), "config next to package");
                var serializer = new JavaScriptSerializer();
                CheckPageSetup(serializer.Serialize(HwpPageSetup.Load(emitted)) == serializer.Serialize(defaults), "default save/load roundtrip");
                options.HwpPageSetup = landscape;
                PrepareHwpPageSetup(options);
                CheckPageSetup(serializer.Serialize(HwpPageSetup.Load(emitted)) == serializer.Serialize(landscape), "landscape overwrite roundtrip");
                CheckPageSetup(BuildHwpWriterArguments("writer.py", options, "output.hwpx", "report.json", "plan.json").Contains(" --page-setup " + Quote(emitted)), "quoted writer argument");

                string input = Path.Combine(directory, "input setup.json");
                HwpPageSetup.WriteConfig(input, defaults);
                string workbook = Path.Combine(directory, "workbook.xlsx");
                string addin = Path.Combine(directory, "addin.xlam");
                File.WriteAllText(workbook, "non-COM validation fixture");
                File.WriteAllText(addin, "non-COM validation fixture");
                var args = new[] { "--workbook", workbook, "--addin", addin, "--hwp-page-setup", input };
                LauncherOptions loaded = LauncherOptions.FromArgs(args);
                CheckPageSetup(loaded.GenerateDraftText && loaded.HwpPageSetupPath == input && serializer.Serialize(loaded.HwpPageSetup) == serializer.Serialize(defaults), "CLI loads config for draft-enabled run");
                loaded.LastReportPackagePath = options.LastReportPackagePath;
                PrepareHwpPageSetup(loaded);
                CheckPageSetup(BuildHwpWriterArguments("writer.py", loaded, "output.hwpx", "report.json", "plan.json").Contains(" --page-setup " + Quote(emitted)), "CLI config copied and forwarded");
                loaded = LauncherOptions.FromArgs(args.Concat(new[] { "--no-draft" }).ToArray());
                CheckPageSetup(!loaded.GenerateDraftText && loaded.HwpPageSetup != null, "new run options retain page setup without TXT draft");
                RejectPageSetup(delegate { LauncherOptions.FromArgs(new[] { "--hwp-page-setup" }); }, "missing CLI path");
                RejectPageSetup(delegate { LauncherOptions.FromArgs(new[] { "--hwp-page-setup", "" }); }, "empty CLI path");
                HwpStylePresetStore.AtomicWrite(input, "{\"orientation\":\"portrait\"}");
                RejectPageSetup(delegate { LauncherOptions.FromArgs(args); }, "CLI rejects malformed schema");

                string saved = File.ReadAllText(emitted);
                options.HwpPageSetup = null;
                CheckPageSetup(!BuildHwpWriterArguments("writer.py", options, "output.hwpx", "report.json", "plan.json").Contains(" --page-setup "), "stale path never forwarded before preparation");
                PrepareHwpPageSetup(options);
                CheckPageSetup(options.HwpPageSetupPath == null && File.ReadAllText(emitted) == saved, "preserve clears path without rewriting old JSON");
                LauncherOptions preserved = LauncherOptions.FromArgs(new[] { "--workbook", workbook, "--addin", addin });
                preserved.HwpPageSetupPath = emitted;
                preserved.Validate();
                CheckPageSetup(preserved.HwpPageSetup == null && preserved.HwpPageSetupPath == null, "validation cannot revive stale config");
                CheckPageSetup(!BuildHwpWriterArguments("writer.py", preserved, "output.hwpx", "report.json", "plan.json").Contains(" --page-setup "), "CLI preserve mode omits argument");
                preserved.HwpPageSetupPath = emitted;
                AutomationRunner.WriteLauncherConfig(workbook, preserved);
                CheckPageSetup(File.ReadAllLines(Path.Combine(directory, "workbook_launcher_config.txt")).Contains("HwpPageSetup="), "preserve config log omits stale path");
                HwpPageSetupDialog.RunSelfCheck();
                using (var dialog = new HwpPageSetupDialog(landscape))
                    CheckPageSetup(serializer.Serialize(dialog.ReadEditor()) == serializer.Serialize(landscape), "configured dialog reopen");
            }
            finally
            {
                Directory.Delete(directory, true);
            }
        }

        private static void CheckPageSetup(bool condition, string message)
        {
            if (!condition) throw new InvalidOperationException("Self-check failed: " + message);
        }

        private static void RejectPageSetup(Action action, string message)
        {
            try { action(); }
            catch (InvalidOperationException) { return; }
            throw new InvalidOperationException("Self-check failed: accepted " + message);
        }

        internal static void RunHwpStyleCliSelfCheck()
        {
            string directory = Path.Combine(Path.GetTempPath(), "ResearchHelperStyleCliCheck_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try
            {
                var options = new LauncherOptions();
                options.LastReportPackagePath = Path.Combine(directory, "report_package.json");
                options.LastPreflightReportPath = Path.Combine(directory, "preflight_report.json");
                options.HwpStyleConfig = HwpStylePresetStore.DefaultConfig();
                File.WriteAllText(options.LastReportPackagePath, "{}", System.Text.Encoding.UTF8);
                File.WriteAllText(options.LastPreflightReportPath, "{\"status\":\"ready_with_warnings\"}");
                RequireHwpPreflight(options);
                File.WriteAllText(options.LastPreflightReportPath,
                    "{\"status\":\"blocked\",\"errors\":[{\"table_key\":\"T0001\",\"message\":\"missing source range\"}]}");
                bool qaBlocked = false;
                try { RequireHwpPreflight(options); }
                catch (InvalidOperationException ex) { qaBlocked = ex.Message.Contains("T0001 missing source range"); }
                if (!qaBlocked) throw new InvalidOperationException("Self-check failed: blocked QA did not surface the table error");
                PrepareHwpStyleConfig(options);
                string expected = Path.Combine(directory, "hwp_style_config.json");
                if (!string.Equals(options.HwpStyleConfigPath, expected, StringComparison.OrdinalIgnoreCase) || !File.Exists(expected))
                    throw new InvalidOperationException("Self-check failed: normalized style config emission");
                string arguments = BuildHwpWriterArguments("writer.py", options, "output.hwpx", "writer_report.json", "render_plan.json");
                if (arguments.IndexOf(" --style-config " + Quote(expected), StringComparison.Ordinal) < 0)
                    throw new InvalidOperationException("Self-check failed: --style-config forwarding");
                var store = new HwpStylePresetStore(Path.Combine(directory, "presets.json"));
                using (var dialog = new HwpStylePresetDialog(store, HwpStylePresetStore.BuiltInName))
                {
                    dialog.ApplyTemplateCell("표머리행", new Dictionary<string, object> {
                        { "fill_color", "DDEBF7" }, { "font_size_pt", 11.0 },
                        { "margins_mm", new Dictionary<string, object> { { "top", 0.5 }, { "bottom", 0.5 }, { "left", 1.0 }, { "right", 2.0 } } },
                        { "edge", new Dictionary<string, object> { { "header_bottom", new Dictionary<string, object> {
                            { "type", "dot" }, { "width_mm", 0.4 }, { "color", "1F4E79" } } } } }
                    });
                    var edited = dialog.ReadEditor();
                    edited["preset_name"] = "역할 검증";
                    store.Save(edited, false);
                    var saved = store.Get("역할 검증");
                    var cells = (Dictionary<string, object>)saved["table_cell_styles"];
                    var styles = (Dictionary<string, object>)saved["paragraph_styles"];
                    var header = (Dictionary<string, object>)cells["표머리행"];
                    var savedBorder = (Dictionary<string, object>)saved["table_border"];
                    if (cells.ContainsKey("표가로배너") || cells.ContainsKey("표세로배너") ||
                        Convert.ToString(header["fill_color"]) != "DDEBF7" ||
                        Convert.ToDouble(((Dictionary<string, object>)header["margins_mm"])["right"]) != 2.0 ||
                        Convert.ToString(((Dictionary<string, object>)((Dictionary<string, object>)savedBorder["edges"])["header_bottom"])["type"]) != "dot" ||
                        Convert.ToString(((Dictionary<string, object>)cells["표배너"])["fill_color"]) != "E7E7E7" ||
                        Convert.ToDouble(((Dictionary<string, object>)styles["표머리행"])["font_size_pt"]) != 11.0)
                        throw new InvalidOperationException("Self-check failed: selected role properties did not round-trip independently");
                    dialog.Size = dialog.MinimumSize;
                    dialog.PerformLayout();
                    var root = (TableLayoutPanel)dialog.Controls[0];
                    root.PerformLayout();
                    dialog.CheckPreview();
                    var borderPanel = (FlowLayoutPanel)root.GetControlFromPosition(0, 4);
                    borderPanel.PerformLayout();
                    foreach (Control control in borderPanel.Controls)
                        if (!borderPanel.ClientRectangle.Contains(control.Bounds))
                            throw new InvalidOperationException("Self-check failed: border control clipped at minimum dialog size");
                    CheckBox checkbox = borderPanel.Controls.OfType<CheckBox>().Single();
                    if (checkbox.Width < checkbox.PreferredSize.Width || checkbox.Height < checkbox.PreferredSize.Height
                        || root.GetRowHeights()[4] < borderPanel.Height + borderPanel.Margin.Vertical)
                        throw new InvalidOperationException("Self-check failed: side borders checkbox or row clipped");
                    Console.WriteLine("Minimum-size border layout passed: " + borderPanel.ClientSize + ", checkbox " + checkbox.Bounds);
                }
            }
            finally
            {
                try { Directory.Delete(directory, true); } catch { }
            }
        }

        public static string TryGenerateDraft(string workbookPath, Action<string> log)
        {
            try
            {
                string pythonPath = PathResolver.ResolvePythonPath();
                string enginePath = PathResolver.ResolveExcelEnginePath();
                string configPath = PathResolver.ResolveDefaultStyleConfigPath();

                if (string.IsNullOrWhiteSpace(pythonPath) || !File.Exists(pythonPath))
                {
                    log("Python 실행 파일을 찾지 못해 문장 초안 생성을 건너뜁니다.");
                    return "";
                }
                if (string.IsNullOrWhiteSpace(enginePath) || !File.Exists(enginePath))
                {
                    log("Python 문장 엔진을 찾지 못해 문장 초안 생성을 건너뜁니다.");
                    return "";
                }
                if (string.IsNullOrWhiteSpace(configPath) || !File.Exists(configPath))
                {
                    log("문장 스타일 설정 파일을 찾지 못해 문장 초안 생성을 건너뜁니다.");
                    return "";
                }

                string outputPath = Path.Combine(
                    Path.GetDirectoryName(workbookPath),
                    Path.GetFileNameWithoutExtension(workbookPath) + "_draft.txt");

                var startInfo = new ProcessStartInfo();
                startInfo.FileName = pythonPath;
                startInfo.Arguments = Quote(enginePath) +
                                      " --excel " + Quote(workbookPath) +
                                      " --config " + Quote(configPath) +
                                      " --output " + Quote(outputPath) +
                                      " --max-tables 30";
                startInfo.UseShellExecute = false;
                startInfo.CreateNoWindow = true;
                startInfo.RedirectStandardOutput = true;
                startInfo.RedirectStandardError = true;
                startInfo.StandardOutputEncoding = System.Text.Encoding.UTF8;
                startInfo.StandardErrorEncoding = System.Text.Encoding.UTF8;

                log("Python 문장 초안을 생성합니다.");
                StandaloneBundle.ConfigurePythonProcess(startInfo);
                using (Process process = Process.Start(startInfo))
                {
                    var stdoutTask = process.StandardOutput.ReadToEndAsync();
                    var stderrTask = process.StandardError.ReadToEndAsync();
                    if (!process.WaitForExit(120000))
                    {
                        try { process.Kill(); } catch { }
                        log("문장 초안 생성 시간이 길어져 이번 실행에서는 건너뜁니다.");
                        return "";
                    }

                    string stdout = stdoutTask.Result;
                    string stderr = stderrTask.Result;

                    if (process.ExitCode != 0)
                    {
                        log("문장 초안 생성 실패: " + (string.IsNullOrWhiteSpace(stderr) ? stdout : stderr).Trim());
                        return "";
                    }
                }

                if (File.Exists(outputPath))
                {
                    log("문장 초안 저장: " + outputPath);
                    return outputPath;
                }
                log("문장 초안 생성이 완료되었지만 출력 파일을 찾지 못했습니다.");
            }
            catch (Exception ex)
            {
                log("문장 초안 생성 실패: " + ex.Message);
            }
            return "";
        }

        public static void TryGenerateReportPackage(string workbookPath, LauncherOptions options, Action<string> log)
        {
            try
            {
                string pythonPath = PathResolver.ResolvePythonPath();
                string toolPath = PathResolver.ResolveEngineToolPath("report_package.py");
                if (string.IsNullOrWhiteSpace(pythonPath) || !File.Exists(pythonPath) || string.IsNullOrWhiteSpace(toolPath) || !File.Exists(toolPath))
                {
                    log("report package 도구를 찾지 못해 preflight 생성을 건너뜁니다.");
                    return;
                }

                string directory = Path.GetDirectoryName(workbookPath);
                string stem = Path.GetFileNameWithoutExtension(workbookPath);
                string packagePath = Path.Combine(directory, stem + "_report_package.json");
                string preflightPath = Path.Combine(directory, stem + "_preflight_report.json");

                var startInfo = new ProcessStartInfo();
                startInfo.FileName = pythonPath;
                startInfo.Arguments = Quote(toolPath) +
                                      " --excel " + Quote(workbookPath) +
                                      " --package-output " + Quote(packagePath) +
                                      " --preflight-output " + Quote(preflightPath) +
                                      " --hwp-template " + Quote(options.OutputType.IndexOf("HWP", StringComparison.OrdinalIgnoreCase) >= 0 ? options.HwpTemplatePath : "") +
                                      " --output-type " + Quote(options.OutputType) +
                                      " --report-profile " + Quote(options.ReportProfile) +
                                      " --style-profile " + Quote(options.StyleProfile) +
                                      " --banner " + Quote(options.BannerSetting) +
                                      " --decimal-places " + Quote(options.DecimalPlaces.ToString());
                startInfo.UseShellExecute = false;
                startInfo.CreateNoWindow = true;
                startInfo.RedirectStandardOutput = true;
                startInfo.RedirectStandardError = true;
                startInfo.StandardOutputEncoding = System.Text.Encoding.UTF8;
                startInfo.StandardErrorEncoding = System.Text.Encoding.UTF8;

                log("report package와 preflight를 생성합니다.");
                StandaloneBundle.ConfigurePythonProcess(startInfo);
                using (Process process = Process.Start(startInfo))
                {
                    var stdoutTask = process.StandardOutput.ReadToEndAsync();
                    var stderrTask = process.StandardError.ReadToEndAsync();
                    if (!process.WaitForExit(120000))
                    {
                        try { process.Kill(); } catch { }
                        log("preflight 생성 시간이 길어져 이번 실행에서는 건너뜁니다.");
                        return;
                    }
                    string stdout = stdoutTask.Result;
                    string stderr = stderrTask.Result;
                    if (process.ExitCode != 0)
                    {
                        log("preflight 생성 실패: " + (string.IsNullOrWhiteSpace(stderr) ? stdout : stderr).Trim());
                        return;
                    }
                }

                options.LastReportPackagePath = File.Exists(packagePath) ? packagePath : "";
                options.LastPreflightReportPath = File.Exists(preflightPath) ? preflightPath : "";
                if (!string.IsNullOrWhiteSpace(options.LastPreflightReportPath))
                {
                    log("preflight 저장: " + options.LastPreflightReportPath);
                }
            }
            catch (Exception ex)
            {
                log("preflight 생성 실패: " + ex.Message);
            }
        }

        public static void TryGenerateHwpDocument(LauncherOptions options, Action<string> log)
        {
            try
            {
                string pythonPath = PathResolver.ResolvePythonPath();
                string toolPath = PathResolver.ResolveEngineToolPath("hwp_com_writer.py");
                if (string.IsNullOrWhiteSpace(pythonPath) || !File.Exists(pythonPath) || string.IsNullOrWhiteSpace(toolPath) || !File.Exists(toolPath))
                {
                    log("HWPX writer 도구를 찾지 못해 HWPX 생성을 건너뜁니다.");
                    return;
                }
                if (string.IsNullOrWhiteSpace(options.LastReportPackagePath) || !File.Exists(options.LastReportPackagePath))
                {
                    log("report_package.json이 없어 HWPX 생성을 건너뜁니다.");
                    return;
                }
                if (string.IsNullOrWhiteSpace(options.LastPreflightReportPath) || !File.Exists(options.LastPreflightReportPath))
                {
                    log("preflight_report.json이 없어 HWPX 생성을 건너뜁니다.");
                    return;
                }
                if (string.IsNullOrWhiteSpace(options.HwpTemplatePath) || !File.Exists(options.HwpTemplatePath))
                {
                    log("HWPX 템플릿이 없어 HWPX 생성을 건너뜁니다.");
                    return;
                }

                PrepareHwpStyleConfig(options);
                log("HWP 서식 설정 저장: " + options.HwpStyleConfigPath);
                PrepareHwpPageSetup(options);
                if (options.HwpPageSetup != null) log("HWP 편집용지 설정 저장: " + options.HwpPageSetupPath);

                string directory = Path.GetDirectoryName(options.LastGeneratedWorkbookPath);
                string stem = Path.GetFileNameWithoutExtension(options.LastGeneratedWorkbookPath);
                string outputPath = Path.Combine(directory, stem + "_hwp_report_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".hwpx");
                string reportPath = Path.Combine(directory, stem + "_hwp_writer_report.json");
                string renderPlanPath = Path.Combine(directory, stem + "_hwp_render_plan.json");

                if (!RunHwpEnvironmentCheck(pythonPath, toolPath, reportPath, options, log))
                {
                    options.LastHwpWriterReportPath = File.Exists(reportPath) ? reportPath : "";
                    return;
                }

                var startInfo = new ProcessStartInfo();
                startInfo.FileName = pythonPath;
                startInfo.Arguments = BuildHwpWriterArguments(toolPath, options, outputPath, reportPath, renderPlanPath);
                startInfo.UseShellExecute = false;
                startInfo.CreateNoWindow = true;
                startInfo.RedirectStandardOutput = true;
                startInfo.RedirectStandardError = true;
                startInfo.StandardOutputEncoding = System.Text.Encoding.UTF8;
                startInfo.StandardErrorEncoding = System.Text.Encoding.UTF8;

                log("아래한글 COM으로 HWPX 초본을 생성합니다.");
                StandaloneBundle.ConfigurePythonProcess(startInfo);
                using (Process process = Process.Start(startInfo))
                {
                    var stdoutTask = process.StandardOutput.ReadToEndAsync();
                    var stderrTask = process.StandardError.ReadToEndAsync();
                    if (!process.WaitForExit(180000))
                    {
                        try { process.Kill(); } catch { }
                        log("HWPX 생성 시간이 길어져 이번 실행에서는 중단했습니다.");
                        options.LastHwpWriterReportPath = File.Exists(reportPath) ? reportPath : "";
                        return;
                    }
                    string stdout = stdoutTask.Result;
                    string stderr = stderrTask.Result;
                    if (process.ExitCode != 0)
                    {
                        log("HWPX 생성 실패: " + (string.IsNullOrWhiteSpace(stderr) ? stdout : stderr).Trim());
                        options.LastHwpWriterReportPath = File.Exists(reportPath) ? reportPath : "";
                        return;
                    }
                }

                options.LastHwpOutputPath = File.Exists(outputPath) ? outputPath : "";
                options.LastHwpWriterReportPath = File.Exists(reportPath) ? reportPath : "";
                if (!string.IsNullOrWhiteSpace(options.LastHwpOutputPath))
                {
                    log("HWPX 초본 저장: " + options.LastHwpOutputPath);
                }
                if (!string.IsNullOrWhiteSpace(options.LastHwpWriterReportPath))
                {
                    log("HWPX writer report 저장: " + options.LastHwpWriterReportPath);
                }
                if (File.Exists(renderPlanPath))
                {
                    log("HWPX render plan 저장: " + renderPlanPath);
                }
            }
            catch (Exception ex)
            {
                log("HWPX 생성 실패: " + ex.Message);
            }
        }

        private static void PrepareHwpStyleConfig(LauncherOptions options)
        {
            Dictionary<string, object> config = options.HwpStyleConfig ?? HwpStylePresetStore.DefaultConfig();
            options.HwpStyleConfigPath = Path.Combine(Path.GetDirectoryName(options.LastReportPackagePath), "hwp_style_config.json");
            HwpStylePresetStore.WriteConfig(options.HwpStyleConfigPath, config);
            options.HwpStylePresetName = Convert.ToString(config["preset_name"]);
        }

        private static string BuildHwpWriterArguments(string toolPath, LauncherOptions options, string outputPath, string reportPath, string renderPlanPath)
        {
            return Quote(toolPath) +
                   " --package " + Quote(options.LastReportPackagePath) +
                   " --preflight " + Quote(options.LastPreflightReportPath) +
                   " --template " + Quote(options.HwpTemplatePath) +
                   " --output " + Quote(outputPath) +
                   " --visible " + Quote(options.HwpVisible ? "true" : "false") +
                   " --report-output " + Quote(reportPath) +
                   " --render-plan-output " + Quote(renderPlanPath) +
                   " --max-sections " + Quote(options.HwpMaxSections.ToString()) +
                   " --dispatch-mode " + Quote(options.HwpDispatchMode) +
                   " --style-config " + Quote(options.HwpStyleConfigPath) +
                   (options.HwpPageSetup == null ? "" : OptionalArgument(" --page-setup ", options.HwpPageSetupPath)) +
                   OptionalArgument(" --table-style-profile ", options.HwpTableStyleProfilePath) +
                   (options.HwpKeepOpenOnError ? " --keep-open-on-error" : "");
        }

        private static void PrepareHwpPageSetup(LauncherOptions options)
        {
            // Null is authoritative: never reuse a previous run's file in template-preserve mode.
            options.HwpPageSetupPath = null;
            if (options.HwpPageSetup == null) return;
            string path = Path.Combine(Path.GetDirectoryName(options.LastReportPackagePath), "hwp_page_setup.json");
            HwpPageSetup.WriteConfig(path, options.HwpPageSetup);
            options.HwpPageSetupPath = path;
        }

        public static string TryRunHwpEnvironmentDiagnostics(LauncherOptions options, string outputReportPath, Action<string> log)
        {
            string pythonPath = PathResolver.ResolvePythonPath();
            string toolPath = PathResolver.ResolveEngineToolPath("hwp_com_writer.py");
            if (string.IsNullOrWhiteSpace(pythonPath) || !File.Exists(pythonPath) || string.IsNullOrWhiteSpace(toolPath) || !File.Exists(toolPath))
            {
                throw new FileNotFoundException("HWPX writer tool or Python runtime was not found.");
            }

            string outputDirectory = Path.GetDirectoryName(outputReportPath);
            if (!string.IsNullOrWhiteSpace(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }
            else
            {
                outputDirectory = Environment.CurrentDirectory;
            }

            var results = new List<Dictionary<string, object>>();
            foreach (string mode in OrderedDispatchModes(options.HwpDispatchMode))
            {
                log("HWP COM diagnostics mode: " + mode);
                string modeReportPath = Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(outputReportPath) + "_" + mode + ".json");
                results.Add(RunHwpEnvironmentMode(pythonPath, toolPath, modeReportPath, options.HwpVisible, mode));
            }

            string status = "blocked";
            foreach (Dictionary<string, object> result in results)
            {
                if (GetString(result, "status") == "ready")
                {
                    status = "ready";
                    break;
                }
            }

            var report = new Dictionary<string, object>();
            report["schema_version"] = "1.0";
            report["status"] = status;
            report["created_at"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            report["preferred_dispatch_mode"] = options.HwpDispatchMode;
            report["timeout_ms"] = 15000;
            report["results"] = results;

            var serializer = new JavaScriptSerializer();
            serializer.MaxJsonLength = int.MaxValue;
            File.WriteAllText(outputReportPath, serializer.Serialize(report), System.Text.Encoding.UTF8);
            log("HWP COM environment diagnostics saved: " + outputReportPath);
            return BuildHwpEnvironmentDiagnosticsSummary(outputReportPath, results);
        }

        private static List<string> OrderedDispatchModes(string preferred)
        {
            var modes = new List<string>();
            string normalized = string.IsNullOrWhiteSpace(preferred) ? "ensure_dispatch" : preferred.Trim().ToLowerInvariant().Replace("-", "_");
            AddModeOnce(modes, normalized);
            AddModeOnce(modes, "ensure_dispatch");
            AddModeOnce(modes, "dispatch");
            AddModeOnce(modes, "dispatch_ex");
            return modes;
        }

        private static void AddModeOnce(List<string> modes, string mode)
        {
            if ((mode == "ensure_dispatch" || mode == "dispatch" || mode == "dispatch_ex") && !modes.Contains(mode))
            {
                modes.Add(mode);
            }
        }

        private static Dictionary<string, object> RunHwpEnvironmentMode(string pythonPath, string toolPath, string reportPath, bool visible, string mode)
        {
            var result = new Dictionary<string, object>();
            result["dispatch_mode"] = mode;
            result["report_path"] = reportPath;
            result["status"] = "started";

            var startInfo = new ProcessStartInfo();
            startInfo.FileName = pythonPath;
            startInfo.Arguments = Quote(toolPath) +
                                  " --check-environment" +
                                  " --visible " + Quote(visible ? "true" : "false") +
                                  " --dispatch-mode " + Quote(mode) +
                                  " --report-output " + Quote(reportPath);
            startInfo.UseShellExecute = false;
            startInfo.CreateNoWindow = true;
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = true;
            startInfo.StandardOutputEncoding = System.Text.Encoding.UTF8;
            startInfo.StandardErrorEncoding = System.Text.Encoding.UTF8;
            HashSet<int> hwpBefore = SnapshotProcessIds("Hwp");

            StandaloneBundle.ConfigurePythonProcess(startInfo);
            using (Process process = Process.Start(startInfo))
            {
                if (!process.WaitForExit(15000))
                {
                    try { process.Kill(); } catch { }
                    result["killed_hwp_processes"] = KillNewProcesses("Hwp", hwpBefore);
                    result["status"] = "timeout";
                    result["timed_out"] = true;
                    MergeHwpEnvironmentModeSummary(result, reportPath);
                    return result;
                }
                string stdout = process.StandardOutput.ReadToEnd();
                string stderr = process.StandardError.ReadToEnd();
                result["exit_code"] = process.ExitCode;
                result["stdout"] = stdout.Trim();
                result["stderr"] = stderr.Trim();
                MergeHwpEnvironmentModeSummary(result, reportPath);
                if (GetString(result, "status") == "started")
                {
                    result["status"] = process.ExitCode == 0 ? "ready" : "failed";
                }
                return result;
            }
        }

        private static HashSet<int> SnapshotProcessIds(string processName)
        {
            var ids = new HashSet<int>();
            try
            {
                foreach (Process process in Process.GetProcessesByName(processName))
                {
                    ids.Add(process.Id);
                }
            }
            catch
            {
            }
            return ids;
        }

        private static int KillNewProcesses(string processName, HashSet<int> before)
        {
            int killed = 0;
            try
            {
                foreach (Process process in Process.GetProcessesByName(processName))
                {
                    if (!before.Contains(process.Id))
                    {
                        try
                        {
                            process.Kill();
                            killed++;
                        }
                        catch
                        {
                        }
                    }
                }
            }
            catch
            {
            }
            return killed;
        }

        private static void MergeHwpEnvironmentModeSummary(Dictionary<string, object> result, string reportPath)
        {
            if (!File.Exists(reportPath))
            {
                return;
            }
            try
            {
                var serializer = new JavaScriptSerializer();
                serializer.MaxJsonLength = int.MaxValue;
                var report = serializer.DeserializeObject(File.ReadAllText(reportPath, System.Text.Encoding.UTF8)) as Dictionary<string, object>;
                if (report == null)
                {
                    return;
                }
                result["writer_status"] = GetString(report, "status");
                result["stage"] = GetString(report, "stage");
                result["action"] = GetString(report, "action");
                Dictionary<string, object> com = GetObject(report, "com");
                if (com != null)
                {
                    result["current_prog_id"] = GetString(com, "current_prog_id");
                    result["prog_id"] = GetString(com, "prog_id");
                    result["last_step"] = LastEnvironmentStepSummary(com);
                    if (GetString(result, "status") == "started" && GetString(report, "status") == "ready")
                    {
                        result["status"] = "ready";
                    }
                }
            }
            catch (Exception ex)
            {
                result["parse_error"] = ex.Message;
            }
        }

        private static string LastEnvironmentStepSummary(Dictionary<string, object> com)
        {
            object stepsValue;
            if (!com.TryGetValue("steps", out stepsValue))
            {
                return "";
            }
            object[] steps = stepsValue as object[];
            if (steps == null || steps.Length == 0)
            {
                return "";
            }
            Dictionary<string, object> last = steps[steps.Length - 1] as Dictionary<string, object>;
            if (last == null)
            {
                return "";
            }
            return GetString(last, "name") + ":" + GetString(last, "status");
        }

        private static Dictionary<string, object> GetObject(Dictionary<string, object> values, string key)
        {
            object value;
            return values.TryGetValue(key, out value) ? value as Dictionary<string, object> : null;
        }

        private static string GetString(Dictionary<string, object> values, string key)
        {
            object value;
            return values.TryGetValue(key, out value) && value != null ? Convert.ToString(value) : "";
        }

        private static string BuildHwpEnvironmentDiagnosticsSummary(string outputReportPath, List<Dictionary<string, object>> results)
        {
            var lines = new List<string>();
            lines.Add("HWP COM environment diagnostics");
            lines.Add("Report: " + outputReportPath);
            lines.Add("");
            foreach (Dictionary<string, object> result in results)
            {
                lines.Add(GetString(result, "dispatch_mode") +
                          ": status=" + GetString(result, "status") +
                          " / stage=" + GetString(result, "stage") +
                          " / action=" + GetString(result, "action") +
                          " / progId=" + GetString(result, "current_prog_id") +
                          " / lastStep=" + GetString(result, "last_step"));
            }
            return string.Join(Environment.NewLine, lines.ToArray());
        }

        private static bool RunHwpEnvironmentCheck(string pythonPath, string toolPath, string reportPath, LauncherOptions options, Action<string> log)
        {
            var startInfo = new ProcessStartInfo();
            startInfo.FileName = pythonPath;
            startInfo.Arguments = Quote(toolPath) +
                                  " --check-environment" +
                                  " --visible " + Quote(options.HwpVisible ? "true" : "false") +
                                  " --dispatch-mode " + Quote(options.HwpDispatchMode) +
                                  " --report-output " + Quote(reportPath);
            startInfo.UseShellExecute = false;
            startInfo.CreateNoWindow = true;
            startInfo.RedirectStandardOutput = true;
            startInfo.RedirectStandardError = true;
            startInfo.StandardOutputEncoding = System.Text.Encoding.UTF8;
            startInfo.StandardErrorEncoding = System.Text.Encoding.UTF8;

            log("아래한글 COM 환경을 확인합니다.");
            StandaloneBundle.ConfigurePythonProcess(startInfo);
            using (Process process = Process.Start(startInfo))
            {
                var stdoutTask = process.StandardOutput.ReadToEndAsync();
                var stderrTask = process.StandardError.ReadToEndAsync();
                if (!process.WaitForExit(60000))
                {
                    try { process.Kill(); } catch { }
                    log("아래한글 COM 환경 확인 시간이 초과되었습니다.");
                    return false;
                }
                string stdout = stdoutTask.Result;
                string stderr = stderrTask.Result;
                if (process.ExitCode != 0)
                {
                    log("아래한글 COM 환경 확인 실패: " + (string.IsNullOrWhiteSpace(stderr) ? stdout : stderr).Trim());
                    return false;
                }
            }
            log("아래한글 COM 환경 확인 완료");
            return true;
        }

        private static string Quote(string value)
        {
            return "\"" + (value ?? "").Replace("\"", "\\\"") + "\"";
        }

        private static string OptionalArgument(string name, string value)
        {
            return string.IsNullOrWhiteSpace(value) ? "" : name + Quote(value);
        }
    }

    internal sealed class WorkbookPreview
    {
        public readonly List<string> Banners = new List<string>();
        public readonly List<string> PrimaryBanners = new List<string>();
        public readonly List<TablePreview> Tables = new List<TablePreview>();
    }

    internal sealed class TablePreview
    {
        public string TableKey;
        public string SheetName;
        public string AutoSheetName;
        public string TableNo;
        public string Title;
        public int Row;
        public int RowCount;
        public int ColCount;
        public int AutoRowCount;
        public int AutoColCount;
        public string RecognitionWarning;
        public string AutoRange;
        public string FinalRange;
        public bool IsManual;
    }

    internal static class BannerInspector
    {
        private const int MaxPreviewRows = 5000;
        private const int MaxPreviewColumns = 300;

        public static List<string> ReadBanners(string workbookPath)
        {
            WorkbookPreview preview = ReadPreview(workbookPath);
            if (preview.Tables.Count == 0)
            {
                throw new InvalidOperationException("집계표에서 데이터가 있는 표를 찾지 못했습니다. 제목·헤더·표 범위를 확인하세요.");
            }
            return preview.Banners;
        }

        public static WorkbookPreview ReadPreview(string workbookPath)
        {
            workbookPath = Path.GetFullPath(workbookPath);
            object excel = null;
            object workbook = null;

            try
            {
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    throw new InvalidOperationException("Excel.Application COM 개체를 찾을 수 없습니다.");
                }

                excel = Activator.CreateInstance(excelType);
                dynamic xl = excel;
                xl.DisplayAlerts = false;
                xl.Visible = false;

                workbook = xl.Workbooks.Open(workbookPath, 0, true);
                dynamic wb = workbook;
                var preview = new WorkbookPreview();
                var found = preview.Banners;
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var primary = preview.PrimaryBanners;
                var primarySeen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                found.Add("전체");
                seen.Add("전체");
                primary.Add("전체");
                primarySeen.Add("전체");

                dynamic dataSheet = ResolvePreviewWorksheet(wb);
                if (dataSheet != null)
                {
                    ScanWorksheet(dataSheet, found, seen, primary, primarySeen, preview.Tables);
                }

                return preview;
            }
            finally
            {
                AutomationRunner.CloseComWorkbookForLauncher(workbook, false);
                AutomationRunner.QuitComExcelForLauncher(excel);
                AutomationRunner.ReleaseComForLauncher(workbook);
                AutomationRunner.ReleaseComForLauncher(excel);
            }
        }

        private static void ScanWorksheet(dynamic ws, List<string> found, HashSet<string> seen, List<string> primary, HashSet<string> primarySeen, List<TablePreview> tables)
        {
            SheetSnapshot sheet;
            if (!TryReadSheetSnapshot(ws, out sheet))
            {
                return;
            }

            foreach (TableBlock block in DetectBlocks(sheet))
            {
                AddTablePreview(tables, Convert.ToString(ws.Name), block);
                // Header extent comes from the first numeric body row, not a required total-row symbol.
                int beforeHeader = block.TitleRow > 0 ? block.TitleRow : block.HeaderRow - 1;
                foreach (string banner in FindBannerGroups(sheet, beforeHeader, block.BodyRow, block.LastCol, block.FirstCol))
                {
                    string choice = "가로: " + NormalizeBannerName(banner);
                    if (seen.Add(choice.Replace(" ", ""))) found.Add(choice);
                    if (IsPrimaryBannerCandidate(banner) && primarySeen.Add(choice.Replace(" ", ""))) primary.Add(choice);
                }
                foreach (string banner in FindVerticalBannerNames(sheet, beforeHeader, block.BodyRow, block.EndRow, block.FirstCol, block.LastCol))
                {
                    string choice = "세로: " + banner;
                    if (seen.Add(choice.Replace(" ", ""))) found.Add(choice);
                }
            }
        }

        private static bool IsTableTitle(string text)
        {
            return TableTitlePattern.IsMatch(NormalizeCell(text));
        }

        private static readonly System.Text.RegularExpressions.Regex TableTitlePattern = new System.Text.RegularExpressions.Regex(
            @"^\s*[\[<〈《(]?\s*(?:표|table)\s*(?:[<〈]?\s*(?<no>(?=[A-Z0-9_.-]*[0-9])[A-Z0-9]+(?:[_.-][A-Z0-9]+)*)\s*[>\]〉》):：.]?|[>\]〉》)])\s*",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        private sealed class TableBlock
        {
            public int TitleRow, HeaderRow, BodyRow, EndRow, FirstCol, LastCol, SourceRow;
            public bool InferredHeader;
            public string Title;
        }

        private static string NormalizeCell(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "";
            return System.Text.RegularExpressions.Regex.Replace((text ?? "").Normalize(System.Text.NormalizationForm.FormKC), @"\s+", " ").Trim();
        }

        private static string HeaderToken(string text)
        {
            string token = NormalizeCell(text).Replace(" ", "").ToLowerInvariant();
            if ((token.StartsWith("(") && token.EndsWith(")")) || (token.StartsWith("[") && token.EndsWith("]"))) token = token.Substring(1, token.Length - 2);
            return token;
        }

        private static bool IsCaseHeader(string text)
        {
            string token = System.Text.RegularExpressions.Regex.Replace(HeaderToken(text), @"\([^)]*\)$", "");
            return new[] { "사례수", "표본수", "응답수", "빈도", "빈도수", "케이스수", "n", "casecount", "wt사례수" }.Contains(token);
        }

        private static bool IsPercentHeader(string text)
        {
            return new[] { "%", "비율", "비율(%)", "백분율", "구성비", "비중", "percent", "pct" }.Contains(HeaderToken(text));
        }

        private static bool IsStubHeader(string text)
        {
            return new[] { "구분", "항목", "응답항목", "응답", "보기", "category", "item" }.Contains(HeaderToken(text));
        }

        private static bool IsMetadata(string text)
        {
            string token = HeaderToken(text).TrimStart('[');
            return token.StartsWith("base") || token.StartsWith("단위") || token.StartsWith("출처") ||
                token.StartsWith("주:") || token.StartsWith("주)") || token.StartsWith("source:") || token.StartsWith("note:");
        }

        public static void RunRecognitionSelfCheck()
        {
            Func<object[][], SheetSnapshot> snapshot = rows =>
            {
                int columns = rows.Max(row => row.Length);
                Array values = Array.CreateInstance(typeof(object), new[] { rows.Length, columns }, new[] { 1, 1 });
                for (int r = 0; r < rows.Length; r++)
                    for (int c = 0; c < rows[r].Length; c++) values.SetValue(rows[r][c], r + 1, c + 1);
                return new SheetSnapshot(1, 1, rows.Length, columns, values);
            };
            Action<bool, string> check = (passed, label) => { if (!passed) throw new InvalidOperationException("Excel recognition check failed: " + label); };
            var shifted = snapshot(new[] {
                new object[] { "", "<표 2-3> 서비스 만족도" },
                new object[] { "", "구분", "응답 수(명)", "만족", "불만족" },
                new object[] { "", "전체", 20, 70, 30 },
                new object[] { "", "남성", 10, 60, 40 }
            });
            var blocks = DetectBlocks(shifted);
            check(blocks.Count == 1 && blocks[0].TitleRow == 1 && blocks[0].FirstCol == 2 && blocks[0].LastCol == 5, "shifted title/header/body");
            check(FindTotalRow(shifted, 1, 4) == 3 && ParseTableNo(blocks[0].Title) == "2-3", "plain total and table number");
            var adjacent = snapshot(new[] {
                new object[] { "", "Q1. 서비스 평가" }, new object[] { "" }, new object[] { "" },
                new object[] { "항목", "비율(%)" }, new object[] { "첫 항목", 80 }, new object[] { "둘째 항목", 20 },
                new object[] { "" }, new object[] { "Table 2. Usage" }, new object[] { "항목", "%" },
                new object[] { "전체", 100 }, new object[] { "이용", 100 }
            });
            blocks = DetectBlocks(adjacent);
            check(blocks.Count == 2 && blocks[0].TitleRow == 1 && blocks[0].BodyRow == 5 && blocks[0].EndRow == 6 && blocks[1].TitleRow == 8, "plain title, gaps and independent tables");
            var hierarchical = snapshot(new[] {
                new object[] { "[ 표 3 ] 분포" }, new object[] { "", "성별", "", "지역", "" },
                new object[] { "구분", "남성", "여성", "수도권", "비수도권" },
                new object[] { "", "%", "%", "%", "%" },
                new object[] { "전체", 10, 60, 20, 40 }, new object[] { "이용", 5, 50, 15, 30 }
            });
            blocks = DetectBlocks(hierarchical);
            check(blocks.Count == 1 && blocks[0].TitleRow == 1, "hierarchical header is not a title");
            var banners = FindBannerGroups(hierarchical, 1, 5, 5);
            check(banners.Count == 2 && banners.Contains("성별") && banners.Contains("지역"), "markerless parent banners without child options");
            var flat = snapshot(new[] { new object[] { "표 4 분포" }, new object[] { "사례수", "만족", "불만족" }, new object[] { "", "%", "%" }, new object[] { "전체", 60, 40 } });
            check(FindBannerGroups(flat, 1, 4, 3).Count == 0, "response options are not parent banners");
            var rowGroups = snapshot(new[] {
                new object[] { "<일반현황> 이용 경험" }, new object[] { "", "", "빈도수", "경험" },
                new object[] { "", "", "", "있음", "없음" }, new object[] { "전체", "", 20, 60, 40 },
                new object[] { "종목", "A", 10, 60, 40 }, new object[] { "", "B", 10, 60, 40 },
                new object[] { "▣ 성 별 ▣", "남성", 20, 60, 40 }, new object[] { "", "고립 항목", 5, 60, 40 },
            });
            blocks = DetectBlocks(rowGroups);
            check(blocks.Count == 1 && FindVerticalBannerNames(rowGroups, 1, 4, 8, 1, 5).SequenceEqual(new[] { "종목", "성 별" }),
                "marked and unmarked row groups with 빈도수 and no percent header");
            var aggregates = snapshot(new[] {
                new object[] { "[표 8] 척도" }, new object[] { "구분", "◐TOP2 (4+5)", "◐긍정 (합)", "◐증상 없음" },
                new object[] { "N", "%", "%", "%" }, new object[] { "전체", 70, 60, 30 }
            });
            check(FindBannerGroups(aggregates, 1, 4, 4).Count == 0, "marked response options and scale summaries are not selectable groups");
            var titleless = snapshot(new[] { new object[] { "항목", "비율" }, new object[] { "이용", 30 }, new object[] { "미이용", 70 } });
            blocks = DetectBlocks(titleless);
            var previews = new List<TablePreview>();
            AddTablePreview(previews, "자료", blocks[0]);
            check(blocks.Count == 1 && previews[0].AutoRange == "A1:B3" && !string.IsNullOrEmpty(previews[0].RecognitionWarning), "missing title stays reviewable");
            var novel = snapshot(new[] { new object[] { "표 5 실태" }, new object[] { "사업장", "응답인원", "찬성", "반대" }, new object[] { "전체", 10, 60, 40 }, new object[] { "A기관", 5, 70, 30 } });
            blocks = DetectBlocks(novel);
            check(blocks.Count == 1 && blocks[0].InferredHeader && blocks[0].TitleRow == 1, "unfamiliar headers require structural evidence");
            check(ParseTableTitle("<표> 이용 경험") == "이용 경험", "unnumbered table marker");
            var years = snapshot(new[] { new object[] { "표 6 연도 비교" }, new object[] { "연도", "N", "%" }, new object[] { 2023, 100, 60 }, new object[] { 2024, 120, 70 } });
            blocks = DetectBlocks(years);
            check(blocks.Count == 1 && blocks[0].BodyRow == 3, "numeric year labels before case-count column");
            var index = snapshot(new[] { new object[] { "INDEX" }, new object[] { "[표 1] 서비스", "GO" }, new object[] { "사례수에 관한 안내" } });
            check(DetectBlocks(index).Count == 0, "contents and narrative mentions are not tables");
        }

        private static bool IsNumericCell(string text)
        {
            double number;
            return double.TryParse(NormalizeCell(text).Trim('(', ')', '%').Replace(",", ""),
                System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out number);
        }

        private static string RowText(SheetSnapshot sheet, int row)
        {
            var parts = new List<string>();
            for (int col = sheet.FirstCol; col <= sheet.LastCol; col++)
            {
                string text = NormalizeCell(sheet.Text(row, col));
                if (text.Length > 0 && !parts.Contains(text)) parts.Add(text);
            }
            return string.Join(" ", parts);
        }

        private static bool IsDataRow(SheetSnapshot sheet, int row)
        {
            bool label = false;
            for (int col = sheet.FirstCol; col <= sheet.LastCol; col++)
            {
                string text = NormalizeCell(sheet.Text(row, col));
                if (text.Length == 0) continue;
                if (!label)
                {
                    if (IsNumericCell(text) || IsMetadata(text) || IsCaseHeader(text) || IsPercentHeader(text) || IsStubHeader(text) || text == "표") return false;
                    label = true;
                }
                else if (IsNumericCell(text)) return true;
            }
            return false;
        }

        private static int FirstBodyRow(SheetSnapshot sheet, int header)
        {
            int meaningful = 0;
            int countCol = 0;
            for (int col = sheet.FirstCol; col <= sheet.LastCol; col++)
                if (IsCaseHeader(sheet.Text(header, col))) { countCol = col; break; }
            for (int row = header + 1; row <= sheet.LastRow; row++)
            {
                if (IsDataRow(sheet, row)) return row;
                // Numeric years/codes are labels only when present left of an explicit count column.
                if (countCol > sheet.FirstCol)
                {
                    bool label = false;
                    for (int col = sheet.FirstCol; col < countCol; col++)
                        if (sheet.Text(row, col).Length > 0) { label = IsNumericCell(sheet.Text(row, col)); break; }
                    if (label && IsNumericCell(sheet.Text(row, countCol)))
                        for (int col = countCol + 1; col <= sheet.LastCol; col++)
                            if (IsNumericCell(sheet.Text(row, col))) return row;
                }
                string text = RowText(sheet, row);
                if (IsTableTitle(text) || text.StartsWith("출처") || ++meaningful > 32 && text.Length > 0) break;
                if (text.Length == 0) meaningful--;
            }
            return 0;
        }

        private static List<TableBlock> DetectBlocks(SheetSnapshot sheet)
        {
            var blocks = new List<TableBlock>();
            // Pair a header with labeled numeric data. Title spelling alone must not turn an INDEX into tables.
            for (int row = sheet.FirstRow; row <= sheet.LastRow; row++)
            {
                if (IsDataRow(sheet, row)) continue;
                int nonempty = 0;
                bool header = false;
                for (int col = sheet.FirstCol; col <= sheet.LastCol; col++)
                {
                    string text = sheet.Text(row, col);
                    if (text.Length > 0) nonempty++;
                    header |= IsCaseHeader(text) || IsPercentHeader(text) || IsStubHeader(text);
                }
                if (nonempty < 2 || IsTableTitle(RowText(sheet, row))) continue;
                int body = FirstBodyRow(sheet, row);
                if (body == 0 || blocks.Any(b => b.BodyRow == body)) continue; // Multiple header levels belong to one table.
                if (!header)
                {
                    // An unfamiliar header needs two real data rows and a boundary, not just text over numbers.
                    bool immediate = true;
                    for (int h = row + 1; h < body; h++)
                        if (RowText(sheet, h).Length > 0 && !IsMetadata(RowText(sheet, h))) immediate = false;
                    int second = body + 1;
                    while (second <= sheet.LastRow && RowText(sheet, second).Length == 0) second++;
                    int previous = row - 1;
                    while (previous >= sheet.FirstRow && RowText(sheet, previous).Length == 0) previous--;
                    bool boundary = previous < sheet.FirstRow || row - previous > 2 || IsTableTitle(RowText(sheet, previous)) ||
                        (blocks.Count == 0 && !IsDataRow(sheet, previous) && RowText(sheet, previous).Length > 0 && !IsMetadata(RowText(sheet, previous)));
                    if (!immediate || !boundary || second > sheet.LastRow || !IsDataRow(sheet, second)) continue;
                }
                blocks.Add(new TableBlock { HeaderRow = row, BodyRow = body, InferredHeader = !header });
            }
            for (int i = 0; i < blocks.Count; i++)
            {
                TableBlock block = blocks[i];
                block.FirstCol = sheet.LastCol;
                block.LastCol = sheet.FirstCol;
                for (int col = sheet.FirstCol; col <= sheet.LastCol; col++)
                    if (sheet.Text(block.BodyRow, col).Length > 0) { block.FirstCol = Math.Min(block.FirstCol, col); block.LastCol = Math.Max(block.LastCol, col); }
                int lower = i == 0 ? sheet.FirstRow : blocks[i - 1].BodyRow + 1;
                int plainRow = 0;
                string plainTitle = "";
                for (int row = block.HeaderRow - 1; row >= lower; row--)
                {
                    string text = RowText(sheet, row);
                    if (text.Length == 0 || IsMetadata(text) || IsDataRow(sheet, row)) continue;
                    if (IsTableTitle(text)) { block.TitleRow = row; block.Title = text; break; }
                    int count = 0, first = 0;
                    for (int col = sheet.FirstCol; col <= sheet.LastCol; col++)
                        if (sheet.Text(row, col).Length > 0) { if (first == 0) first = col; count++; }
                    if (plainRow == 0 && count <= 2 && first <= block.FirstCol + 1 && !IsStubHeader(text) && !IsCaseHeader(text) && !IsPercentHeader(text))
                    { plainRow = row; plainTitle = text; }
                }
                if (block.TitleRow == 0) { block.TitleRow = plainRow; block.Title = plainTitle; }
                block.SourceRow = block.TitleRow > 0 ? block.TitleRow + 1 : block.HeaderRow;
                while (block.SourceRow < block.HeaderRow && RowText(sheet, block.SourceRow).Length == 0) block.SourceRow++;
            }
            for (int i = 0; i < blocks.Count; i++)
            {
                TableBlock block = blocks[i];
                block.EndRow = i + 1 < blocks.Count ? (blocks[i + 1].TitleRow > 0 ? blocks[i + 1].TitleRow : blocks[i + 1].HeaderRow) - 1 : sheet.LastRow;
                while (block.EndRow > block.BodyRow && RowText(sheet, block.EndRow).Length == 0) block.EndRow--;
                for (int row = block.HeaderRow; row <= block.EndRow; row++)
                    if (!IsMetadata(RowText(sheet, row)))
                        for (int col = sheet.FirstCol; col <= sheet.LastCol; col++)
                            if (sheet.Text(row, col).Length > 0) { block.FirstCol = Math.Min(block.FirstCol, col); block.LastCol = Math.Max(block.LastCol, col); }
            }
            return blocks;
        }

        private static void AddTablePreview(List<TablePreview> tables, string sheetName, TableBlock block)
        {
            if (tables.Count >= 500)
            {
                return;
            }

            var preview = new TablePreview();
            preview.TableKey = "T" + (tables.Count + 1).ToString("0000");
            preview.SheetName = sheetName;
            preview.AutoSheetName = sheetName;
            preview.Row = block.TitleRow > 0 ? block.TitleRow : block.HeaderRow;
            int firstRow = block.SourceRow;
            preview.RowCount = block.EndRow - firstRow + 1;
            preview.ColCount = block.LastCol - block.FirstCol + 1;
            preview.AutoRowCount = preview.RowCount;
            preview.AutoColCount = preview.ColCount;
            preview.AutoRange = ExcelColumnName(block.FirstCol) + firstRow + ":" + ExcelColumnName(block.LastCol) + block.EndRow;
            preview.FinalRange = preview.AutoRange;
            preview.TableNo = ParseTableNo(block.Title);
            preview.Title = string.IsNullOrWhiteSpace(block.Title) ? "제목 미확인" : ParseTableTitle(block.Title);
            preview.RecognitionWarning = block.TitleRow == 0 ? "제목 없음: 범위 확인 필요" : !IsTableTitle(block.Title) ? "구조로 추정: 제목·범위 확인" : "";
            if (block.InferredHeader) preview.RecognitionWarning += (preview.RecognitionWarning.Length > 0 ? "; " : "") + "비표준 헤더: 열 의미 확인";
            tables.Add(preview);
        }

        private static dynamic ResolvePreviewWorksheet(dynamic workbook)
        {
            foreach (dynamic ws in workbook.Worksheets)
            {
                if (Convert.ToString(ws.Name).IndexOf("배너조정", StringComparison.OrdinalIgnoreCase) >= 0 && HasTableTitles(ws))
                {
                    return ws;
                }
            }
            try
            {
                dynamic active = workbook.ActiveSheet;
                if (active != null && HasTableTitles(active))
                {
                    return active;
                }
            }
            catch
            {
            }
            try
            {
                dynamic sheet1 = workbook.Worksheets["Sheet1"];
                if (HasTableTitles(sheet1))
                {
                    return sheet1;
                }
            }
            catch
            {
            }
            foreach (dynamic ws in workbook.Worksheets)
            {
                if (!IsGeneratedSheet(Convert.ToString(ws.Name)) && HasTableTitles(ws))
                {
                    return ws;
                }
            }
            return null;
        }

        private static bool HasTableTitles(dynamic worksheet)
        {
            try
            {
                if (IsGeneratedSheet(Convert.ToString(worksheet.Name)) || string.Equals(Convert.ToString(worksheet.Name), "INDEX", StringComparison.OrdinalIgnoreCase) || Convert.ToString(worksheet.Name) == "목차") return false;
                SheetSnapshot sheet;
                return TryReadSheetSnapshot(worksheet, out sheet) && DetectBlocks(sheet).Count > 0;
            }
            catch
            {
            }
            return false;
        }

        private static string ExcelColumnName(int column)
        {
            string name = "";
            while (column > 0)
            {
                column--;
                name = (char)('A' + column % 26) + name;
                column /= 26;
            }
            return name;
        }

        private static string ParseTableNo(string rawTitle)
        {
            return TableTitlePattern.Match(NormalizeCell(rawTitle)).Groups["no"].Value;
        }

        private static string ParseTableTitle(string rawTitle)
        {
            string text = TableTitlePattern.Replace(NormalizeCell(rawTitle), "", 1).TrimEnd(']').Trim();
            if (text.StartsWith("[", StringComparison.Ordinal))
            {
                int sectionClose = text.IndexOf("]", StringComparison.Ordinal);
                if (sectionClose >= 0 && sectionClose + 1 < text.Length)
                {
                    text = text.Substring(sectionClose + 1).Trim();
                }
            }
            int divider = text.IndexOf("─", StringComparison.Ordinal);
            if (divider > 0)
            {
                text = text.Substring(0, divider).Trim();
            }
            int question = text.IndexOf("[ 문", StringComparison.Ordinal);
            if (question > 0)
            {
                text = text.Substring(0, question).Trim();
            }
            return text.Length == 0 ? rawTitle : text;
        }

        private static bool TryReadSheetSnapshot(dynamic ws, out SheetSnapshot sheet)
        {
            sheet = null;
            try
            {
                dynamic used = ws.UsedRange;
                int firstRow = Convert.ToInt32(used.Row);
                int firstCol = Convert.ToInt32(used.Column);
                int lastRow = firstRow + Convert.ToInt32(used.Rows.Count) - 1;
                int lastCol = firstCol + Convert.ToInt32(used.Columns.Count) - 1;
                if (lastRow > MaxPreviewRows)
                {
                    lastRow = MaxPreviewRows;
                }
                if (lastCol > MaxPreviewColumns)
                {
                    lastCol = MaxPreviewColumns;
                }
                if (lastRow < firstRow || lastCol < 1)
                {
                    return false;
                }

                int rowCount = lastRow - firstRow + 1;
                int colCount = lastCol - firstCol + 1;
                dynamic limitedRange = ws.Range[ws.Cells[firstRow, firstCol], ws.Cells[lastRow, lastCol]];
                object values = limitedRange.Value2;
                sheet = new SheetSnapshot(firstRow, firstCol, rowCount, colCount, values);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static int FindTotalRow(SheetSnapshot sheet, int startRow, int endRow)
        {
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = sheet.FirstCol; col <= sheet.LastCol; col++)
                {
                    string text = HeaderToken(sheet.Text(row, col)).TrimStart('■', '▣', '●');
                    if ((text == "전체" || text == "합계" || text == "총계" || text == "total") && IsDataRow(sheet, row)) return row;
                }
            }
            return 0;
        }

        private static List<string> FindVerticalBannerNames(SheetSnapshot sheet, int startRow, int bodyRow, int endRow, int firstCol, int lastCol)
        {
            var groups = new List<string>();
            int countCol = 0;
            for (int col = firstCol; col <= lastCol && countCol == 0; col++)
                for (int row = startRow + 1; row < bodyRow; row++)
                    if (IsCaseHeader(sheet.Text(row, col))) { countCol = col; break; }
            if (countCol < firstCol + 2 || countCol == lastCol) return groups;
            // shortcut: adjacent group/member columns only, extend with verified geometry for other row layouts.
            int anchor = countCol - 2;
            Func<int, bool> isMember = row => NormalizeCell(sheet.Text(row, anchor + 1)).Length > 0 &&
                IsNumericCell(sheet.Text(row, countCol)) &&
                Enumerable.Range(countCol + 1, lastCol - countCol).Any(col => IsNumericCell(sheet.Text(row, col)));
            for (int row = bodyRow; row <= endRow; row++)
            {
                string raw = NormalizeCell(sheet.Text(row, anchor));
                string name = NormalizeBannerName(raw);
                if (new[] { "전체", "합계", "총계", "total" }.Contains(HeaderToken(name)) || !IsGroupNameCandidate(name) || !isMember(row)) continue;
                if (raw.StartsWith("▣", StringComparison.Ordinal) ||
                    (row < endRow && NormalizeCell(sheet.Text(row + 1, anchor)).Length == 0 && isMember(row + 1)))
                    AddBanner(groups, name);
            }
            return groups;
        }

        private static List<string> FindBannerGroups(SheetSnapshot sheet, int startRow, int totalRow, int lastCol, int firstCol = 1)
        {
            var groups = new List<string>();
            if (totalRow == 0)
            {
                return groups;
            }

            int headerStart = startRow + 1;
            int headerEnd = totalRow - 1;
            if (headerStart > headerEnd)
            {
                return groups;
            }

            // A parent banner needs a child-label level. Flat response options above % are not banners.
            int bannerRow = 0;
            for (int row = headerStart; row <= headerEnd && bannerRow == 0; row++)
                for (int col = firstCol; col <= lastCol; col++)
                {
                    string candidate = sheet.Text(row, col);
                    if (!IsGroupNameCandidate(candidate) || IsScaleSummaryName(candidate)) continue;
                    if (HasBannerChildren(sheet, row, col, headerEnd, lastCol, IsMarkedBanner(candidate) ? 1 : 2)) bannerRow = row;
                    if (bannerRow > 0) break;
                }
            if (bannerRow == 0) return groups;
            var colGroupName = new string[lastCol + 1];
            string lastSeenGroupName = "";
            for (int col = firstCol; col <= lastCol; col++)
            {
                {
                    string text = sheet.Text(bannerRow, col);
                    if (IsGroupNameCandidate(text) && !IsScaleSummaryName(text) &&
                        HasBannerChildren(sheet, bannerRow, col, headerEnd, lastCol, IsMarkedBanner(text) ? 1 : 2))
                    {
                        colGroupName[col] = text;
                        lastSeenGroupName = text;
                    }
                    else if (!string.IsNullOrWhiteSpace(text)) lastSeenGroupName = "";
                }
                if (string.IsNullOrEmpty(colGroupName[col]) && IsPercentMeasureColumn(sheet, startRow, totalRow, col))
                {
                    colGroupName[col] = lastSeenGroupName;
                }
            }

            string currentName = "";
            int percentCount = 0;
            for (int col = firstCol; col <= lastCol; col++)
            {
                string name = colGroupName[col] ?? "";
                if (!string.Equals(name, currentName, StringComparison.Ordinal))
                {
                    if (!string.IsNullOrWhiteSpace(currentName) && percentCount > 0)
                    {
                        AddBanner(groups, currentName);
                    }
                    currentName = name;
                    percentCount = 0;
                }
                if (IsPercentMeasureColumn(sheet, startRow, totalRow, col))
                {
                    percentCount++;
                }
            }

            if (!string.IsNullOrWhiteSpace(currentName) && percentCount > 0)
            {
                AddBanner(groups, currentName);
            }

            return groups;
        }

        private static bool HasBannerChildren(SheetSnapshot sheet, int row, int col, int headerEnd, int lastCol, int minimumMeasures)
        {
            // Blank cells following a parent anchor are its span, ending at the next heading.
            int end = col;
            while (end < lastCol && string.IsNullOrWhiteSpace(sheet.Text(row, end + 1))) end++;
            if (Enumerable.Range(col, end - col + 1).Count(c => IsPercentMeasureColumn(sheet, row, headerEnd + 1, c)) < minimumMeasures) return false;
            for (int child = row + 1; child <= headerEnd; child++)
                for (int c = col; c <= end; c++)
                    if (IsGroupNameCandidate(sheet.Text(child, c)) && NormalizeCell(sheet.Text(child, c)) != NormalizeCell(sheet.Text(row, col))) return true;
            return false;
        }

        private static bool IsMarkedBanner(string value)
        {
            return value.StartsWith("◐", StringComparison.Ordinal) || value.StartsWith("▣", StringComparison.Ordinal);
        }

        private static bool IsScaleSummaryName(string value)
        {
            string key = NormalizeBannerName(value).Split('(')[0].Trim().Replace(" ", "");
            return System.Text.RegularExpressions.Regex.IsMatch(key, @"^(TOP|BOT|BOTTOM|MID|LOW)\d*$", System.Text.RegularExpressions.RegexOptions.IgnoreCase) ||
                new[] { "긍정", "부정", "보통", "인지", "비인지", "만족", "불만족", "동의", "비동의", "찬성", "반대", "도움이됨", "해로움", "필요함", "필요하지않음" }.Contains(key);
        }

        private static bool IsPercentMeasureColumn(SheetSnapshot sheet, int startRow, int totalRow, int col)
        {
            for (int row = startRow + 1; row <= totalRow - 1; row++)
            {
                if (IsPercentHeader(sheet.Text(row, col)))
                {
                    return true;
                }
            }
            return false;
        }

        private static bool IsGroupNameCandidate(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }
            string normalized = NormalizeCell(text);
            return normalized != "계" && normalized != "합계" && HeaderToken(normalized) != "n/%" && !IsPercentHeader(normalized) && !IsCaseHeader(normalized) && !IsStubHeader(normalized) && !IsMetadata(normalized) && !IsNumericCell(normalized) && !IsTableTitle(normalized);
        }

        private static void AddBanner(List<string> groups, string name)
        {
            name = (name ?? "").Trim();
            if (name.Length == 0)
            {
                return;
            }
            foreach (string existing in groups)
            {
                if (string.Equals(existing, name, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }
            }
            groups.Add(name);
        }

        private static bool IsPrimaryBannerCandidate(string banner)
        {
            string name = NormalizeBannerName(banner);
            if (name.Length == 0)
            {
                return false;
            }
            if (string.Equals(name, "전체", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            string compact = name.Replace(" ", "");
            string[] excludedTokens = new string[]
            {
                "BOT", "BOTTOM", "MID", "TOP", "LOW",
                "평균", "점수", "환산", "100점", "5점",
                "사례수", "BASE", "N=", "종합",
                "긍정", "부정", "보통", "인지", "비인지",
                "만족", "불만족", "동의", "비동의",
                "필요함", "불필요"
            };
            if (ContainsAny(compact, excludedTokens))
            {
                return false;
            }

            string[] preferredTokens = new string[]
            {
                "성별", "연령", "연령대", "지역", "권역", "수도권",
                "사업체", "기업", "기관", "업종", "산업", "직종",
                "종사자", "규모", "대분류", "중분류", "소분류",
                "구분", "유형", "참여", "경험", "이용"
            };
            return ContainsAny(compact, preferredTokens);
        }

        private static string NormalizeBannerName(string banner)
        {
            return NormalizeCell(banner).Trim('◐', '▣', ' ').Trim();
        }

        private static bool ContainsAny(string text, string[] tokens)
        {
            foreach (string token in tokens)
            {
                if (text.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }
            return false;
        }

        private static bool IsGeneratedSheet(string name)
        {
            return name.StartsWith("보고서_", StringComparison.OrdinalIgnoreCase) ||
                   name.StartsWith("_Report", StringComparison.OrdinalIgnoreCase);
        }

        private sealed class SheetSnapshot
        {
            private readonly int firstRow;
            private readonly int firstCol;
            private readonly int rowCount;
            private readonly int colCount;
            private readonly object values;

            public SheetSnapshot(int firstRow, int firstCol, int rowCount, int colCount, object values)
            {
                this.firstRow = firstRow;
                this.firstCol = firstCol;
                this.rowCount = rowCount;
                this.colCount = colCount;
                this.values = values;
            }

            public int FirstRow { get { return firstRow; } }
            public int FirstCol { get { return firstCol; } }
            public int LastRow { get { return firstRow + rowCount - 1; } }
            public int LastCol { get { return firstCol + colCount - 1; } }

            public string Text(int row, int col)
            {
                int r = row - firstRow + 1;
                int c = col - firstCol + 1;
                if (r < 1 || r > rowCount || c < 1 || c > colCount)
                {
                    return "";
                }

                object value = null;
                Array array = values as Array;
                if (array != null && array.Rank == 2)
                {
                    value = array.GetValue(r, c);
                }
                else if (rowCount == 1 && colCount == 1)
                {
                    value = values;
                }
                return Convert.ToString(value ?? "").Trim();
            }
        }
    }

    internal static class PathResolver
    {
        public static string ResolveDefaultAddinPath()
        {
            if (StandaloneBundle.Root != null) return StandaloneBundle.Find("report_automation_addin/dev/ReportAutomationAddin_dev.xlam");
            string env = Environment.GetEnvironmentVariable("REPORT_AUTOMATION_ADDIN");
            if (!string.IsNullOrWhiteSpace(env) && File.Exists(env))
            {
                return env;
            }

            var candidates = new List<string>();
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            candidates.Add(Path.Combine(baseDir, "ReportAutomationAddin_dev.xlam"));
            candidates.Add(Path.Combine(baseDir, "..", "..", "report_automation_addin", "dev", "ReportAutomationAddin_dev.xlam"));
            candidates.Add(Path.Combine(Environment.CurrentDirectory, "report_automation_addin", "dev", "ReportAutomationAddin_dev.xlam"));

            foreach (string candidate in candidates)
            {
                string fullPath = Path.GetFullPath(candidate);
                if (File.Exists(fullPath))
                {
                    return fullPath;
                }
            }

            return Path.GetFullPath(candidates[1]);
        }

        public static string ResolvePythonPath()
        {
            if (StandaloneBundle.Root != null) return StandaloneBundle.Find("runtime/python.exe");
            string env = Environment.GetEnvironmentVariable("REPORT_AUTOMATION_PYTHON");
            if (!string.IsNullOrWhiteSpace(env) && File.Exists(env))
            {
                return env;
            }

            var candidates = new List<string>();
            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            candidates.Add(Path.Combine(baseDir, "..", "..", ".venv", "Scripts", "python.exe"));
            candidates.Add(Path.Combine(Environment.CurrentDirectory, ".venv", "Scripts", "python.exe"));
            candidates.Add(Path.Combine(baseDir, "python.exe"));
            candidates.Add(Path.Combine(userProfile, ".cache", "codex-runtimes", "codex-primary-runtime", "dependencies", "python", "python.exe"));
            candidates.Add(@"C:\Python312\python.exe");
            candidates.Add(@"C:\Python311\python.exe");

            foreach (string candidate in candidates)
            {
                string fullPath = Path.GetFullPath(candidate);
                if (File.Exists(fullPath))
                {
                    return fullPath;
                }
            }

            return "";
        }

        public static string ResolveExcelEnginePath()
        {
            if (StandaloneBundle.Root != null) return StandaloneBundle.Find("report_automation_engine/excel_report_generator.py");
            string env = Environment.GetEnvironmentVariable("REPORT_AUTOMATION_ENGINE");
            if (!string.IsNullOrWhiteSpace(env) && File.Exists(env))
            {
                return env;
            }

            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            var candidates = new List<string>();
            candidates.Add(Path.Combine(baseDir, "report_automation_engine", "excel_report_generator.py"));
            candidates.Add(Path.Combine(baseDir, "..", "..", "report_automation_engine", "excel_report_generator.py"));
            candidates.Add(Path.Combine(Environment.CurrentDirectory, "report_automation_engine", "excel_report_generator.py"));

            foreach (string candidate in candidates)
            {
                string fullPath = Path.GetFullPath(candidate);
                if (File.Exists(fullPath))
                {
                    return fullPath;
                }
            }

            return Path.GetFullPath(candidates[1]);
        }

        public static string ResolveEngineToolPath(string fileName)
        {
            string enginePath = ResolveExcelEnginePath();
            if (!string.IsNullOrWhiteSpace(enginePath))
            {
                string candidate = Path.Combine(Path.GetDirectoryName(enginePath), fileName);
                if (File.Exists(candidate))
                {
                    return candidate;
                }
            }

            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            var candidates = new List<string>();
            candidates.Add(Path.Combine(baseDir, "report_automation_engine", fileName));
            candidates.Add(Path.Combine(baseDir, "..", "..", "report_automation_engine", fileName));
            candidates.Add(Path.Combine(Environment.CurrentDirectory, "report_automation_engine", fileName));

            foreach (string candidate in candidates)
            {
                string fullPath = Path.GetFullPath(candidate);
                if (File.Exists(fullPath))
                {
                    return fullPath;
                }
            }

            return Path.GetFullPath(candidates[1]);
        }

        public static string ResolveDefaultStyleConfigPath()
        {
            string enginePath = ResolveExcelEnginePath();
            if (!string.IsNullOrWhiteSpace(enginePath))
            {
                string configPath = Path.Combine(Path.GetDirectoryName(enginePath), "config", "default_style_schema.json");
                if (File.Exists(configPath))
                {
                    return configPath;
                }
            }
            return "";
        }
    }
}
