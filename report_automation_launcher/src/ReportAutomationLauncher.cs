using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
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
                string generatedWorkbookPath = AutomationRunner.Run(options, Console.WriteLine);
                options.LastGeneratedWorkbookPath = generatedWorkbookPath;
                if (options.GenerateDraftText)
                {
                    options.LastDraftTextPath = EngineRunner.TryGenerateDraft(generatedWorkbookPath, Console.WriteLine);
                }
                EngineRunner.TryGenerateReportPackage(generatedWorkbookPath, options, Console.WriteLine);
                AutomationRunner.WriteLauncherConfig(generatedWorkbookPath, options);
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

    internal enum LauncherButtonKind
    {
        Primary,
        Secondary,
        Ghost
    }

    internal static class LauncherUi
    {
        public static readonly Color ColorBackground = Color.FromArgb(246, 248, 251);
        public static readonly Color ColorSurface = Color.White;
        public static readonly Color ColorSurfaceAlt = Color.FromArgb(240, 244, 248);
        public static readonly Color ColorSurfaceStrong = Color.FromArgb(232, 239, 247);
        public static readonly Color ColorBorder = Color.FromArgb(211, 219, 230);
        public static readonly Color ColorText = Color.FromArgb(31, 41, 55);
        public static readonly Color ColorMutedText = Color.FromArgb(89, 99, 112);
        public static readonly Color ColorPrimary = Color.FromArgb(39, 101, 173);
        public static readonly Color ColorPrimaryHover = Color.FromArgb(28, 83, 145);
        public static readonly Color ColorSuccess = Color.FromArgb(20, 112, 74);
        public static readonly Color ColorSuccessSurface = Color.FromArgb(226, 246, 236);
        public static readonly Color ColorWarning = Color.FromArgb(166, 94, 12);
        public static readonly Color ColorWarningSurface = Color.FromArgb(255, 246, 226);
        public static readonly Color ColorDanger = Color.FromArgb(175, 48, 48);
        public static readonly Color ColorDisabledSurface = Color.FromArgb(232, 236, 241);
        public static readonly Color ColorDisabledText = Color.FromArgb(133, 143, 156);
        public static readonly Color ColorRowAlt = Color.FromArgb(249, 251, 253);

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
            button.Font = SectionFont();
            button.Padding = new Padding(SpaceMd, 0, SpaceMd, 0);

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
                normalFore = ColorPrimary;
                border = Color.FromArgb(160, 184, 213);
                hoverBack = Color.FromArgb(231, 240, 250);
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

        public static void StyleStatusLabel(Label label, bool isReady)
        {
            label.ForeColor = isReady ? ColorSuccess : ColorWarning;
            label.BackColor = isReady ? ColorSuccessSurface : ColorWarningSurface;
            label.BorderStyle = BorderStyle.FixedSingle;
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
                textBox.BackColor = textBox.ReadOnly ? Color.FromArgb(250, 252, 255) : ColorSurface;
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
                if (button.Text == "실행" || button.Text.IndexOf("생성", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    StyleButton(button, LauncherButtonKind.Primary);
                }
                else if (button.Text == "닫기")
                {
                    StyleButton(button, LauncherButtonKind.Ghost);
                }
                else
                {
                    StyleButton(button, LauncherButtonKind.Secondary);
                }
                return;
            }

            if (control is CheckBox)
            {
                CheckBox checkBox = (CheckBox)control;
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

            Rectangle border = new Rectangle(0, 12, groupBox.Width - 1, groupBox.Height - 13);
            using (Pen borderPen = new Pen(ColorBorder))
            using (Pen accentPen = new Pen(ColorPrimary, 2))
            {
                e.Graphics.DrawRectangle(borderPen, border);
                e.Graphics.DrawLine(accentPen, SpaceLg, 12, SpaceLg + 42, 12);
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

    internal sealed class MainForm : Form
    {
        private readonly TextBox workbookPathText = new TextBox();
        private readonly TextBox addinPathText = new TextBox();
        private readonly ComboBox outputTypeCombo = new ComboBox();
        private readonly TextBox hwpTemplateText = new TextBox();
        private readonly TextBox pptTemplateText = new TextBox();
        private readonly TextBox hwpTableStyleProfileText = new TextBox();
        private readonly Button inspectTemplateButton = new Button();
        private readonly Button createHwpTemplateButton = new Button();
        private readonly Button createPptTemplateButton = new Button();
        private readonly Button createChartTemplateButton = new Button();
        private readonly Button autoFixTemplateButton = new Button();
        private readonly Button openTemplateGuideButton = new Button();
        private readonly Label templateStatusLabel = new Label();
        private readonly CheckBox hwpVisibleCheck = new CheckBox();
        private readonly CheckBox hwpKeepOpenOnErrorCheck = new CheckBox();
        private readonly ComboBox hwpMaxSectionsCombo = new ComboBox();
        private readonly ComboBox hwpDispatchModeCombo = new ComboBox();
        private readonly Button hwpEnvironmentCheckButton = new Button();
        private readonly TextBox bannerText = new TextBox();
        private readonly CheckedListBox bannerList = new CheckedListBox();
        private readonly TabControl workflowTabs = new TabControl();
        private readonly ListView tablePreviewList = new ListView();
        private readonly Label dataStatusLabel = new Label();
        private readonly Label fileStepStatusLabel = new Label();
        private readonly Label dataStepStatusLabel = new Label();
        private readonly Label bannerStepStatusLabel = new Label();
        private readonly Label runStepStatusLabel = new Label();
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
        private readonly Button closeButton = new Button();
        private readonly Button openWorkbookButton = new Button();
        private readonly Button openDraftButton = new Button();
        private readonly Button openHwpOutputButton = new Button();
        private readonly Button openHwpReportButton = new Button();
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

        public MainForm()
        {
            Text = "보고서 자동화 Alpha";
            MinimumSize = new Size(980, 720);
            Size = new Size(1120, 820);
            StartPosition = FormStartPosition.CenterScreen;
            LauncherUi.ApplyToForm(this);

            var root = new TableLayoutPanel();
            root.Dock = DockStyle.Fill;
            root.Padding = new Padding(LauncherUi.SpaceXl);
            root.ColumnCount = 1;
            root.RowCount = 4;
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(root);

            root.Controls.Add(BuildHeaderPanel(), 0, 0);
            root.Controls.Add(BuildWorkflowSummaryStrip(), 0, 1);

            workflowTabs.Dock = DockStyle.Fill;
            workflowTabs.DrawMode = TabDrawMode.OwnerDrawFixed;
            workflowTabs.SizeMode = TabSizeMode.Fixed;
            workflowTabs.ItemSize = new Size(148, 34);
            workflowTabs.DrawItem += WorkflowTabs_DrawItem;
            workflowTabs.TabPages.Add(CreateStepPage("1 파일 등록", BuildFileGroup()));
            workflowTabs.TabPages.Add(CreateStepPage("2 데이터 확인", BuildDataReviewGroup()));
            workflowTabs.TabPages.Add(CreateStepPage("3 작성 규칙", BuildRulesPage()));
            workflowTabs.TabPages.Add(CreateStepPage("4 실행/결과", BuildRunGroup()));
            workflowTabs.TabPages.Add(CreateStepPage("대시보드 PPT", BuildDashboardPage()));
            workflowTabs.SelectedIndexChanged += delegate { UpdateWorkflowStatus(); };
            root.Controls.Add(workflowTabs, 0, 2);

            var buttons = new FlowLayoutPanel();
            buttons.FlowDirection = FlowDirection.RightToLeft;
            buttons.Dock = DockStyle.Fill;
            buttons.AutoSize = true;
            buttons.BackColor = LauncherUi.ColorSurface;
            buttons.Padding = new Padding(LauncherUi.SpaceMd);
            buttons.Margin = new Padding(0, LauncherUi.SpaceMd, 0, 0);
            runButton.Text = "실행";
            runButton.Width = 104;
            runButton.Click += RunButton_Click;
            closeButton.Text = "닫기";
            closeButton.Width = 96;
            closeButton.Click += delegate { Close(); };
            buttons.Controls.Add(closeButton);
            buttons.Controls.Add(runButton);
            root.Controls.Add(buttons, 0, 3);

            workbookPathText.TextChanged += delegate { UpdateWorkflowStatus(); };
            addinPathText.TextChanged += delegate { UpdateWorkflowStatus(); };
            outputTypeCombo.SelectedIndexChanged += delegate { UpdateWorkflowStatus(); };
            hwpTemplateText.TextChanged += delegate { UpdateWorkflowStatus(); };
            pptTemplateText.TextChanged += delegate { UpdateWorkflowStatus(); };

            addinPathText.Text = PathResolver.ResolveDefaultAddinPath();
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
            LauncherUi.StyleButton(runButton, LauncherButtonKind.Primary);
            LauncherUi.StyleButton(closeButton, LauncherButtonKind.Ghost);
            UpdateWorkflowStatus();
            UpdateReadinessChecklist();
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
            panel.BackColor = LauncherUi.ColorSurface;
            panel.Padding = new Padding(LauncherUi.SpaceLg);
            panel.Margin = new Padding(0, 0, 0, LauncherUi.SpaceMd);
            panel.ColumnCount = 2;
            panel.RowCount = 2;
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 8));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var accent = new Panel();
            accent.BackColor = LauncherUi.ColorPrimary;
            accent.Dock = DockStyle.Fill;
            accent.Margin = new Padding(0, 0, LauncherUi.SpaceMd, 0);
            panel.Controls.Add(accent, 0, 0);
            panel.SetRowSpan(accent, 2);

            var title = new Label();
            title.Text = "보고서 자동화 Alpha";
            title.Font = LauncherUi.TitleFont();
            title.AutoSize = true;
            title.ForeColor = LauncherUi.ColorText;
            title.Margin = new Padding(0, 0, 0, LauncherUi.SpaceXs);
            panel.Controls.Add(title, 1, 0);

            var subtitle = new Label();
            subtitle.Text = "Excel 집계표 확인, 분석 배너 선택, 문장 검토, HWPX/PPTX 초본 생성을 한 곳에서 관리합니다.";
            subtitle.AutoSize = true;
            subtitle.ForeColor = LauncherUi.ColorMutedText;
            subtitle.Margin = new Padding(0);
            panel.Controls.Add(subtitle, 1, 1);
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
            using (SolidBrush textBrush = new SolidBrush(foreColor))
            using (StringFormat format = new StringFormat())
            {
                format.Alignment = StringAlignment.Center;
                format.LineAlignment = StringAlignment.Center;
                e.Graphics.FillRectangle(backBrush, bounds);
                if (selected)
                {
                    using (Pen accent = new Pen(LauncherUi.ColorPrimary, 3))
                    {
                        e.Graphics.DrawLine(accent, bounds.Left + 12, bounds.Bottom - 3, bounds.Right - 12, bounds.Bottom - 3);
                    }
                }
                e.Graphics.DrawString(page.Text, LauncherUi.SectionFont(), textBrush, bounds, format);
            }
        }

        private Control BuildWorkflowSummaryStrip()
        {
            var panel = new TableLayoutPanel();
            panel.Dock = DockStyle.Top;
            panel.AutoSize = true;
            panel.ColumnCount = 4;
            panel.RowCount = 1;
            panel.Margin = new Padding(0, 0, 0, 10);
            for (int i = 0; i < 4; i++)
            {
                panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            }

            panel.Controls.Add(CreateStepStatusLabel(fileStepStatusLabel, "파일"), 0, 0);
            panel.Controls.Add(CreateStepStatusLabel(dataStepStatusLabel, "데이터"), 1, 0);
            panel.Controls.Add(CreateStepStatusLabel(bannerStepStatusLabel, "배너"), 2, 0);
            panel.Controls.Add(CreateStepStatusLabel(runStepStatusLabel, "실행"), 3, 0);
            return panel;
        }

        private Label CreateStepStatusLabel(Label label, string title)
        {
            label.Text = title + ": 대기";
            label.Dock = DockStyle.Fill;
            label.AutoSize = false;
            label.Height = 34;
            label.TextAlign = ContentAlignment.MiddleLeft;
            label.Padding = new Padding(LauncherUi.SpaceMd, 0, LauncherUi.SpaceMd, 0);
            label.Margin = new Padding(0, 0, LauncherUi.SpaceSm, 0);
            label.BorderStyle = BorderStyle.FixedSingle;
            label.BackColor = LauncherUi.ColorSurface;
            label.ForeColor = LauncherUi.ColorMutedText;
            return label;
        }

        private void UpdateWorkflowStatus()
        {
            bool hasWorkbook = File.Exists(workbookPathText.Text.Trim());
            bool hasAddin = File.Exists(addinPathText.Text.Trim());
            int tableCount = tablePreviewList.Items.Count;
            int checkedBanners = CountCheckedBanners();

            SetStepStatus(fileStepStatusLabel, "파일", hasWorkbook && hasAddin ? "준비됨" : "확인 필요", hasWorkbook && hasAddin);
            SetStepStatus(dataStepStatusLabel, "데이터", tableCount > 0 ? tableCount + "개 표" : "미확인", tableCount > 0);
            SetStepStatus(bannerStepStatusLabel, "배너", checkedBanners > 0 ? checkedBanners + "개 선택" : "전체 기준", checkedBanners > 0 || bannerList.Items.Count == 0);
            SetStepStatus(runStepStatusLabel, "실행", workflowTabs.SelectedIndex == 3 ? "검토 중" : "대기", workflowTabs.SelectedIndex == 3);
            UpdateReadinessChecklist();
        }

        private void SetStepStatus(Label label, string title, string value, bool isReady)
        {
            label.Text = title + ": " + value;
            LauncherUi.StyleStatusLabel(label, isReady);
        }

        private void UpdateReadinessChecklist()
        {
            if (readinessList.Columns.Count == 0)
            {
                return;
            }

            bool workbookReady = File.Exists(workbookPathText.Text.Trim());
            bool addinReady = File.Exists(addinPathText.Text.Trim());
            int tableCount = tablePreviewList.Items.Count;
            int checkedBanners = CountCheckedBanners();
            bool bannerPreviewReady = bannerList.Items.Count > 0;
            bool outputReady = IsOutputReadyForSelectedOutput();

            readinessList.BeginUpdate();
            try
            {
                readinessList.Items.Clear();
                AddReadinessItem("집계표", workbookReady, workbookReady ? Path.GetFileName(workbookPathText.Text.Trim()) : "집계표 엑셀 파일을 선택하세요.");
                AddReadinessItem("추가기능", addinReady, addinReady ? Path.GetFileName(addinPathText.Text.Trim()) : "보고서 자동화 add-in 경로를 확인하세요.");
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
            item.ForeColor = ready ? LauncherUi.ColorSuccess : LauncherUi.ColorWarning;
            item.BackColor = ready ? LauncherUi.ColorSuccessSurface : LauncherUi.ColorWarningSurface;
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
            var group = new GroupBox();
            group.Text = "파일";
            group.Dock = DockStyle.Top;
            group.AutoSize = true;
            group.Padding = new Padding(10);

            var grid = CreateGrid(3);
            group.Controls.Add(grid);

            AddPathRow(grid, 0, "집계표 엑셀", workbookPathText, "찾기", BrowseWorkbook);
            AddPathRow(grid, 1, "자동화 추가기능", addinPathText, "찾기", BrowseAddin);
            addinPathText.ReadOnly = false;

            var note = new Label();
            note.Text = "선택한 파일은 직접 수정하지 않고, 기본값으로 복사본에 산출 시트를 생성합니다.";
            note.AutoSize = true;
            note.ForeColor = LauncherUi.ColorMutedText;
            note.Margin = new Padding(130, 4, 0, 0);
            grid.Controls.Add(note, 1, 2);
            grid.SetColumnSpan(note, 2);

            return group;
        }

        private Control BuildOutputGroup()
        {
            var group = new GroupBox();
            group.Text = "산출 방식";
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

            AddPathRow(grid, 1, "HWP/HWPX 템플릿", hwpTemplateText, "찾기", BrowseHwpTemplate);
            AddPathRow(grid, 2, "PPTX 템플릿", pptTemplateText, "찾기", BrowsePptTemplate);
            AddPathRow(grid, 3, "HWP 표 스타일", hwpTableStyleProfileText, "찾기", BrowseHwpTableStyleProfile);

            var templateButtons = new FlowLayoutPanel();
            templateButtons.Dock = DockStyle.Fill;
            templateButtons.AutoSize = true;
            inspectTemplateButton.Text = "검사";
            inspectTemplateButton.Width = 60;
            inspectTemplateButton.Click += InspectTemplateButton_Click;
            createHwpTemplateButton.Text = "기본 HWPX";
            createHwpTemplateButton.Width = 90;
            createHwpTemplateButton.Click += CreateHwpTemplateButton_Click;
            createPptTemplateButton.Text = "기본 PPTX";
            createPptTemplateButton.Width = 90;
            createPptTemplateButton.Click += CreatePptTemplateButton_Click;
            createChartTemplateButton.Text = "차트 PPTX";
            createChartTemplateButton.Width = 90;
            createChartTemplateButton.Click += CreateChartTemplateButton_Click;
            autoFixTemplateButton.Text = "자동 보정";
            autoFixTemplateButton.Width = 85;
            autoFixTemplateButton.Click += AutoFixTemplateButton_Click;
            openTemplateGuideButton.Text = "가이드";
            openTemplateGuideButton.Width = 70;
            openTemplateGuideButton.Click += OpenTemplateGuideButton_Click;
            templateButtons.Controls.Add(inspectTemplateButton);
            templateButtons.Controls.Add(createHwpTemplateButton);
            templateButtons.Controls.Add(createPptTemplateButton);
            templateButtons.Controls.Add(createChartTemplateButton);
            templateButtons.Controls.Add(autoFixTemplateButton);
            templateButtons.Controls.Add(openTemplateGuideButton);
            AddLabel(grid, 4, "템플릿 도구");
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
            hwpLimitLabel.Text = "초본 문항 수";
            hwpLimitLabel.AutoSize = true;
            hwpLimitLabel.Margin = new Padding(12, 7, 6, 0);
            hwpMaxSectionsCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            hwpMaxSectionsCombo.Width = 120;
            hwpMaxSectionsCombo.Items.Add("1개 검증");
            hwpMaxSectionsCombo.Items.Add("3개 검증");
            hwpMaxSectionsCombo.Items.Add("전체");
            var hwpDispatchLabel = new Label();
            hwpDispatchLabel.Text = "dispatch";
            hwpDispatchLabel.AutoSize = true;
            hwpDispatchLabel.Margin = new Padding(12, 7, 6, 0);
            hwpDispatchModeCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            hwpDispatchModeCombo.Width = 135;
            hwpDispatchModeCombo.Items.Add("ensure_dispatch");
            hwpDispatchModeCombo.Items.Add("dispatch");
            hwpDispatchModeCombo.Items.Add("dispatch_ex");
            hwpEnvironmentCheckButton.Text = "COM diag";
            hwpEnvironmentCheckButton.Width = 85;
            hwpEnvironmentCheckButton.Click += HwpEnvironmentCheckButton_Click;
            hwpOptions.Controls.Add(hwpVisibleCheck);
            hwpOptions.Controls.Add(hwpKeepOpenOnErrorCheck);
            hwpOptions.Controls.Add(hwpLimitLabel);
            hwpOptions.Controls.Add(hwpMaxSectionsCombo);
            hwpOptions.Controls.Add(hwpDispatchLabel);
            hwpOptions.Controls.Add(hwpDispatchModeCombo);
            hwpOptions.Controls.Add(hwpEnvironmentCheckButton);
            AddLabel(grid, 6, "HWPX 옵션");
            grid.Controls.Add(hwpOptions, 1, 6);
            grid.SetColumnSpan(hwpOptions, 2);

            var components = new FlowLayoutPanel();
            components.Dock = DockStyle.Fill;
            components.AutoSize = true;
            analysisCheck.Text = "분석문";
            chartCheck.Text = "차트 데이터";
            tableCheck.Text = "삽입용 집계표";
            qaCheck.Text = "QA/출처/수정이력";
            draftTextCheck.Text = "문장 초안 TXT(Python)";
            components.Controls.Add(analysisCheck);
            components.Controls.Add(chartCheck);
            components.Controls.Add(tableCheck);
            components.Controls.Add(qaCheck);
            components.Controls.Add(draftTextCheck);
            AddLabel(grid, 7, "구성요소");
            grid.Controls.Add(components, 1, 7);
            grid.SetColumnSpan(components, 2);

            var note = new Label();
            note.Text = "HWPX 보고서는 Excel 산출과 preflight 생성 후 아래한글 COM으로 새 HWPX 초본을 저장합니다. PPTX 보고서 직접 생성은 다음 단계입니다.";
            note.AutoSize = true;
            note.ForeColor = LauncherUi.ColorMutedText;
            note.Margin = new Padding(0, 4, 0, 0);
            grid.Controls.Add(note, 1, 8);
            grid.SetColumnSpan(note, 2);

            return group;
        }

        private Control BuildDataReviewGroup()
        {
            var panel = new TableLayoutPanel();
            panel.Dock = DockStyle.Fill;
            panel.ColumnCount = 1;
            panel.RowCount = 3;
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            dataStatusLabel.Text = "집계표 파일을 선택하면 표 목록과 배너 목록을 자동으로 확인합니다.";
            dataStatusLabel.AutoSize = true;
            dataStatusLabel.ForeColor = LauncherUi.ColorMutedText;
            dataStatusLabel.Margin = new Padding(0, 0, 0, 8);
            panel.Controls.Add(dataStatusLabel, 0, 0);

            tablePreviewList.Dock = DockStyle.Fill;
            tablePreviewList.View = View.Details;
            tablePreviewList.FullRowSelect = true;
            tablePreviewList.GridLines = true;
            tablePreviewList.Columns.Add("No", 48);
            tablePreviewList.Columns.Add("표번호", 90);
            tablePreviewList.Columns.Add("제목", 520);
            tablePreviewList.Columns.Add("시트", 130);
            tablePreviewList.Columns.Add("행", 70);
            panel.Controls.Add(tablePreviewList, 0, 1);

            var note = new Label();
            note.Text = "표 제목이 누락되거나 전체행(■전체■)이 없는 표는 산출 후 QA 시트에서 추가 확인합니다.";
            note.AutoSize = true;
            note.ForeColor = LauncherUi.ColorMutedText;
            note.Margin = new Padding(0, 8, 0, 0);
            panel.Controls.Add(note, 0, 2);
            return panel;
        }

        private Control BuildRulesPage()
        {
            var panel = new TableLayoutPanel();
            panel.Dock = DockStyle.Fill;
            panel.ColumnCount = 1;
            panel.RowCount = 4;
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            panel.Controls.Add(BuildOutputGroup(), 0, 0);
            panel.Controls.Add(BuildOptionsGroup(), 0, 1);
            panel.Controls.Add(BuildAdvancedOptionsGroup(), 0, 2);

            var note = new Label();
            note.Text = "현재 선택값은 실행 설정 파일과 Excel의 보고서_설정 시트에 남습니다. 아직 미구현된 출력도 설정 계약으로 먼저 저장합니다.";
            note.AutoSize = true;
            note.ForeColor = LauncherUi.ColorMutedText;
            note.Margin = new Padding(0, 10, 0, 0);
            panel.Controls.Add(note, 0, 3);
            return panel;
        }

        private Control BuildAdvancedOptionsGroup()
        {
            var group = new GroupBox();
            group.Text = "고급 옵션";
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

            llmEnabledCheck.Text = "LLM 문장 고도화 사용";
            llmEnabledCheck.AutoSize = true;
            AddLabel(grid, 3, "LLM");
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
            panel.RowCount = 6;
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 120));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 34));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 22));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 44));

            var guide = new Label();
            guide.Text = "실행 후 로그, 완료 요약, 문장 초안 미리보기를 확인합니다. 오류가 나면 이 로그를 기준으로 원본 표 구조나 Excel 신뢰 설정을 점검합니다.";
            guide.AutoSize = true;
            guide.ForeColor = LauncherUi.ColorMutedText;
            guide.Margin = new Padding(0, 0, 0, 8);
            panel.Controls.Add(guide, 0, 0);

            readinessList.Dock = DockStyle.Fill;
            readinessList.View = View.Details;
            readinessList.FullRowSelect = true;
            readinessList.GridLines = true;
            readinessList.HeaderStyle = ColumnHeaderStyle.Nonclickable;
            readinessList.Columns.Add("항목", 130);
            readinessList.Columns.Add("상태", 100);
            readinessList.Columns.Add("확인 내용", 650);
            panel.Controls.Add(readinessList, 0, 1);

            logText.Multiline = true;
            logText.ReadOnly = true;
            logText.ScrollBars = ScrollBars.Vertical;
            logText.Dock = DockStyle.Fill;
            logText.Margin = new Padding(0, 0, 0, 8);
            panel.Controls.Add(logText, 0, 2);

            resultSummaryText.Multiline = true;
            resultSummaryText.ReadOnly = true;
            resultSummaryText.ScrollBars = ScrollBars.Vertical;
            resultSummaryText.Dock = DockStyle.Fill;
            resultSummaryText.Text = "아직 실행 결과가 없습니다.";
            panel.Controls.Add(resultSummaryText, 0, 3);

            var resultButtons = new FlowLayoutPanel();
            resultButtons.Dock = DockStyle.Fill;
            resultButtons.AutoSize = true;
            resultButtons.BackColor = LauncherUi.ColorSurfaceAlt;
            resultButtons.Padding = new Padding(LauncherUi.SpaceSm);
            resultButtons.Margin = new Padding(0, LauncherUi.SpaceSm, 0, LauncherUi.SpaceSm);
            openWorkbookButton.Text = "산출 엑셀 열기";
            openWorkbookButton.Width = 110;
            openWorkbookButton.Enabled = false;
            openWorkbookButton.Click += OpenWorkbookButton_Click;
            openDraftButton.Text = "초안 TXT 열기";
            openDraftButton.Width = 105;
            openDraftButton.Enabled = false;
            openDraftButton.Click += OpenDraftButton_Click;
            openHwpOutputButton.Text = "HWPX 초본 열기";
            openHwpOutputButton.Width = 115;
            openHwpOutputButton.Enabled = false;
            openHwpOutputButton.Click += OpenHwpOutputButton_Click;
            openHwpReportButton.Text = "HWPX 리포트 열기";
            openHwpReportButton.Width = 125;
            openHwpReportButton.Enabled = false;
            openHwpReportButton.Click += OpenHwpReportButton_Click;
            copyDraftButton.Text = "초안 복사";
            copyDraftButton.Width = 85;
            copyDraftButton.Enabled = false;
            copyDraftButton.Click += CopyDraftButton_Click;
            draftPreviewStatusLabel.AutoSize = true;
            draftPreviewStatusLabel.ForeColor = LauncherUi.ColorMutedText;
            draftPreviewStatusLabel.Margin = new Padding(10, 7, 0, 0);
            draftPreviewStatusLabel.Text = "문장 초안이 생성되면 아래에 표시됩니다.";
            resultButtons.Controls.Add(openWorkbookButton);
            resultButtons.Controls.Add(openDraftButton);
            resultButtons.Controls.Add(openHwpOutputButton);
            resultButtons.Controls.Add(openHwpReportButton);
            resultButtons.Controls.Add(copyDraftButton);
            resultButtons.Controls.Add(draftPreviewStatusLabel);
            panel.Controls.Add(resultButtons, 0, 4);

            draftPreviewText.Multiline = true;
            draftPreviewText.ReadOnly = true;
            draftPreviewText.ScrollBars = ScrollBars.Vertical;
            draftPreviewText.Dock = DockStyle.Fill;
            draftPreviewText.Text = "아직 문장 초안이 없습니다.";

            draftReviewTabs.Dock = DockStyle.Fill;
            draftReviewTabs.TabPages.Add(CreateStepPage("전체 초안", draftPreviewText));
            draftReviewTabs.TabPages.Add(CreateStepPage("문장 리뷰", BuildSentenceReviewPanel()));
            draftReviewTabs.TabPages.Add(CreateStepPage("QA 경고", BuildQaReviewPanel()));
            panel.Controls.Add(draftReviewTabs, 0, 5);
            return panel;
        }

        private Control BuildDashboardPage()
        {
            var root = new TableLayoutPanel();
            root.Dock = DockStyle.Fill;
            root.ColumnCount = 1;
            root.RowCount = 5;
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 32));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 34));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 34));

            root.Controls.Add(BuildDashboardFileGroup(), 0, 0);
            root.Controls.Add(BuildDashboardSelectionGroup(), 0, 1);
            root.Controls.Add(BuildDashboardMappingGroup(), 0, 2);
            root.Controls.Add(BuildDashboardCommandGroup(), 0, 3);

            dashboardStatusText.Multiline = true;
            dashboardStatusText.ReadOnly = true;
            dashboardStatusText.ScrollBars = ScrollBars.Vertical;
            dashboardStatusText.Dock = DockStyle.Fill;
            dashboardStatusText.Text = "Excel을 읽은 뒤 사용할 sheet, 기관, 열을 선택하세요.";
            root.Controls.Add(dashboardStatusText, 0, 4);
            return root;
        }

        private Control BuildDashboardFileGroup()
        {
            var group = new GroupBox();
            group.Text = "대시보드 원자료";
            group.Dock = DockStyle.Top;
            group.AutoSize = true;
            group.Padding = new Padding(10);

            var grid = CreateGrid(4);
            group.Controls.Add(grid);
            AddLabel(grid, 0, "Excel 원자료");
            dashboardWorkbookText.Dock = DockStyle.Fill;
            grid.Controls.Add(dashboardWorkbookText, 1, 0);
            dashboardBrowseButton.Text = "찾기";
            dashboardBrowseButton.Click += DashboardBrowseButton_Click;
            grid.Controls.Add(dashboardBrowseButton, 2, 0);

            AddLabel(grid, 1, "sheet");
            dashboardSheetCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            dashboardSheetCombo.Dock = DockStyle.Fill;
            dashboardSheetCombo.SelectedIndexChanged += DashboardSheetCombo_SelectedIndexChanged;
            grid.Controls.Add(dashboardSheetCombo, 1, 1);
            dashboardInspectButton.Text = "데이터 읽기";
            dashboardInspectButton.Click += DashboardInspectButton_Click;
            grid.Controls.Add(dashboardInspectButton, 2, 1);

            AddLabel(grid, 2, "작업용 PPT 템플릿");
            dashboardTemplateText.Dock = DockStyle.Fill;
            grid.Controls.Add(dashboardTemplateText, 1, 2);
            dashboardTemplateBrowseButton.Text = "찾기";
            dashboardTemplateBrowseButton.Click += DashboardTemplateBrowseButton_Click;
            grid.Controls.Add(dashboardTemplateBrowseButton, 2, 2);

            var note = new Label();
            note.Text = "템플릿을 지정하면 첫 슬라이드의 RA_DASH_* 위치와 폰트를 유지하고 데이터만 교체합니다.";
            note.AutoSize = true;
            note.ForeColor = LauncherUi.ColorMutedText;
            grid.Controls.Add(note, 1, 3);
            grid.SetColumnSpan(note, 2);
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
            dashboardSelectAllEntitiesButton.Text = "전체 선택";
            dashboardSelectAllEntitiesButton.Click += delegate { SetChecked(dashboardEntityList, true); };
            dashboardClearEntitiesButton.Text = "선택 해제";
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
            dashboardSelectAllColumnsButton.Text = "전체 선택";
            dashboardSelectAllColumnsButton.Click += delegate { SetChecked(dashboardColumnList, true); RefreshDashboardMappingCandidates(); };
            dashboardClearColumnsButton.Text = "선택 해제";
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
            outputLabel.Text = "출력";
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
            fontLabel.Text = "폰트";
            fontLabel.AutoSize = true;
            fontLabel.Margin = new Padding(12, 7, 6, 0);
            dashboardFontCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            dashboardFontCombo.Width = 130;
            dashboardFontCombo.Items.Add("맑은 고딕");
            dashboardFontCombo.Items.Add("나눔고딕");
            dashboardFontCombo.Items.Add("Noto Sans CJK KR");
            dashboardFontCombo.Items.Add("Arial");
            dashboardGenerateButton.Text = "대시보드 PPT 생성";
            dashboardGenerateButton.Width = 140;
            dashboardGenerateButton.Click += DashboardGenerateButton_Click;
            dashboardOpenOutputButton.Text = "결과 열기";
            dashboardOpenOutputButton.Width = 90;
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
            applySentenceEditButton.Text = "수정 적용";
            applySentenceEditButton.Width = 90;
            applySentenceEditButton.Enabled = false;
            applySentenceEditButton.Click += ApplySentenceEditButton_Click;
            copySelectedSentenceButton.Text = "선택 문장 복사";
            copySelectedSentenceButton.Width = 110;
            copySelectedSentenceButton.Enabled = false;
            copySelectedSentenceButton.Click += CopySelectedSentenceButton_Click;
            exportReviewedDraftButton.Text = "검토본 저장";
            exportReviewedDraftButton.Width = 95;
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
            group.Text = "분석 옵션";
            group.Dock = DockStyle.Top;
            group.AutoSize = true;
            group.Padding = new Padding(10);

            var grid = CreateGrid(8);
            group.Controls.Add(grid);

            AddLabel(grid, 0, "보고서 유형");
            reportProfileCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            reportProfileCombo.Items.Add("인식도/만족도 조사형");
            reportProfileCombo.Items.Add("산업 실태조사형");
            reportProfileCombo.Items.Add("정책 수요조사형");
            reportProfileCombo.Items.Add("일반 빈도/교차표형");
            reportProfileCombo.Dock = DockStyle.Fill;
            grid.Controls.Add(reportProfileCombo, 1, 0);
            grid.SetColumnSpan(reportProfileCombo, 2);

            AddLabel(grid, 1, "문체");
            styleProfileCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            styleProfileCombo.Items.Add("공식 보고서체");
            styleProfileCombo.Items.Add("간결 요약체");
            styleProfileCombo.Items.Add("상세 해석체");
            styleProfileCombo.Items.Add("검토 메모체");
            styleProfileCombo.Dock = DockStyle.Fill;
            grid.Controls.Add(styleProfileCombo, 1, 1);
            grid.SetColumnSpan(styleProfileCombo, 2);

            AddLabel(grid, 2, "추출 배너 목록");
            bannerText.Dock = DockStyle.Fill;
            grid.Controls.Add(bannerText, 1, 2);
            grid.SetColumnSpan(bannerText, 2);

            AddLabel(grid, 3, "발견된 배너");
            bannerList.Dock = DockStyle.Fill;
            bannerList.CheckOnClick = true;
            bannerList.Height = 82;
            bannerList.ItemCheck += BannerList_ItemCheck;
            grid.Controls.Add(bannerList, 1, 3);
            grid.SetColumnSpan(bannerList, 2);

            var bannerButtons = new FlowLayoutPanel();
            bannerButtons.Dock = DockStyle.Fill;
            bannerButtons.AutoSize = true;
            reloadBannerButton.Text = "새로고침";
            reloadBannerButton.Width = 80;
            reloadBannerButton.Click += delegate { LoadBannerPreviewAsync(); };
            recommendedBannerButton.Text = "추천 선택";
            recommendedBannerButton.Width = 80;
            recommendedBannerButton.Click += delegate { SelectRecommendedBanners(); };
            selectAllBannerButton.Text = "전체 선택";
            selectAllBannerButton.Width = 80;
            selectAllBannerButton.Click += delegate { SetAllBannersChecked(true); };
            clearBannerButton.Text = "선택 해제";
            clearBannerButton.Width = 80;
            clearBannerButton.Click += delegate { SetAllBannersChecked(false); };
            moveBannerUpButton.Text = "위로";
            moveBannerUpButton.Width = 60;
            moveBannerUpButton.Click += delegate { MoveSelectedBanner(-1); };
            moveBannerDownButton.Text = "아래로";
            moveBannerDownButton.Width = 60;
            moveBannerDownButton.Click += delegate { MoveSelectedBanner(1); };
            deleteBannerButton.Text = "목록 삭제";
            deleteBannerButton.Width = 80;
            deleteBannerButton.Click += delegate { DeleteSelectedBanners(); };
            bannerStatusLabel.AutoSize = true;
            bannerStatusLabel.ForeColor = LauncherUi.ColorMutedText;
            bannerStatusLabel.Margin = new Padding(10, 8, 0, 0);
            bannerButtons.Controls.Add(reloadBannerButton);
            bannerButtons.Controls.Add(recommendedBannerButton);
            bannerButtons.Controls.Add(selectAllBannerButton);
            bannerButtons.Controls.Add(clearBannerButton);
            bannerButtons.Controls.Add(moveBannerUpButton);
            bannerButtons.Controls.Add(moveBannerDownButton);
            bannerButtons.Controls.Add(deleteBannerButton);
            bannerButtons.Controls.Add(bannerStatusLabel);
            grid.Controls.Add(bannerButtons, 1, 4);
            grid.SetColumnSpan(bannerButtons, 2);

            AddLabel(grid, 5, "제목 제거 접두어");
            titlePrefixesText.Dock = DockStyle.Fill;
            grid.Controls.Add(titlePrefixesText, 1, 5);
            grid.SetColumnSpan(titlePrefixesText, 2);

            copyWorkbookCheck.Text = "원본 옆에 작업 복사본 생성";
            copyWorkbookCheck.AutoSize = true;
            keepExcelOpenCheck.Text = "완료 후 Excel 열어두기";
            keepExcelOpenCheck.AutoSize = true;
            var checks = new FlowLayoutPanel();
            checks.Dock = DockStyle.Fill;
            checks.AutoSize = true;
            checks.Controls.Add(copyWorkbookCheck);
            checks.Controls.Add(keepExcelOpenCheck);
            AddLabel(grid, 6, "실행 방식");
            grid.Controls.Add(checks, 1, 6);
            grid.SetColumnSpan(checks, 2);

            var note = new Label();
            note.Text = "알파 기본값은 인식도/만족도 조사형, 공식 보고서체, 소수점 한 자리입니다.";
            note.AutoSize = true;
            note.ForeColor = LauncherUi.ColorMutedText;
            note.Margin = new Padding(0, 4, 0, 0);
            grid.Controls.Add(note, 1, 7);
            grid.SetColumnSpan(note, 2);

            return group;
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
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 96));
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
            tablePreviewList.BeginUpdate();
            try
            {
                tablePreviewList.Items.Clear();
                int index = 1;
                foreach (TablePreview table in tables)
                {
                    var item = new ListViewItem(index.ToString());
                    item.SubItems.Add(table.TableNo);
                    item.SubItems.Add(table.Title);
                    item.SubItems.Add(table.SheetName);
                    item.SubItems.Add(table.Row.ToString());
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

        private void BrowseAddin(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = "보고서 자동화 추가기능 선택";
                dialog.Filter = "Excel add-ins (*.xlam;*.xlsm)|*.xlam;*.xlsm|All files (*.*)|*.*";
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    addinPathText.Text = dialog.FileName;
                }
            }
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

        private void RunTemplateTool(string scriptName, string arguments, int timeoutMs)
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

            using (Process process = Process.Start(startInfo))
            {
                if (!process.WaitForExit(timeoutMs))
                {
                    try { process.Kill(); } catch { }
                    throw new TimeoutException("템플릿 도구 실행 시간이 초과되었습니다.");
                }
                string stdout = process.StandardOutput.ReadToEnd();
                string stderr = process.StandardError.ReadToEnd();
                if (process.ExitCode != 0)
                {
                    throw new InvalidOperationException(string.IsNullOrWhiteSpace(stderr) ? stdout.Trim() : stderr.Trim());
                }
            }
        }

        private static string QuoteArg(string value)
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
                   "- 한글 문서에서 본문이 들어갈 위치에 {{BODY}} 한 줄을 유지합니다." + Environment.NewLine +
                   "- 표지, 글꼴, 여백, 제목 스타일, 로고, 쪽번호는 자유롭게 수정할 수 있습니다." + Environment.NewLine +
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

            runButton.Text = "실행 중...";
            runButton.Enabled = false;
            closeButton.Enabled = false;
            workflowTabs.SelectedIndex = 3;
            resultSummaryText.Text = "실행 중입니다. Excel이 백그라운드에서 열리고 산출 시트를 생성합니다.";
            draftPreviewText.Text = "문장 초안을 기다리는 중입니다.";
            draftPreviewStatusLabel.Text = "실행 중...";
            openWorkbookButton.Enabled = false;
            openDraftButton.Enabled = false;
            openHwpOutputButton.Enabled = false;
            openHwpReportButton.Enabled = false;
            copyDraftButton.Enabled = false;
            Log("실행을 시작합니다.");

            var thread = new Thread(delegate()
            {
                try
                {
                    string generatedWorkbookPath = AutomationRunner.Run(options, Log);
                    options.LastGeneratedWorkbookPath = generatedWorkbookPath;
                    if (options.GenerateDraftText)
                    {
                        options.LastDraftTextPath = EngineRunner.TryGenerateDraft(generatedWorkbookPath, Log);
                    }
                    EngineRunner.TryGenerateReportPackage(generatedWorkbookPath, options, Log);
                    if (options.OutputType.IndexOf("HWP", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        EngineRunner.TryGenerateHwpDocument(options, Log);
                    }
                    AutomationRunner.WriteLauncherConfig(generatedWorkbookPath, options);
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
                    BeginInvoke(new Action(delegate()
                    {
                        resultSummaryText.Text = "실행 중 오류가 발생했습니다." + Environment.NewLine + ex.Message;
                        MessageBox.Show(this, ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }));
                }
                finally
                {
                    BeginInvoke(new Action(delegate()
                    {
                        runButton.Text = "실행";
                        runButton.Enabled = true;
                        closeButton.Enabled = true;
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

        private LauncherOptions ReadOptions()
        {
            var options = new LauncherOptions();
            options.WorkbookPath = workbookPathText.Text.Trim();
            options.AddinPath = addinPathText.Text.Trim();
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
            if (string.IsNullOrWhiteSpace(AddinPath))
            {
                throw new InvalidOperationException("보고서 자동화 추가기능 파일을 선택하세요.");
            }
            AddinPath = Path.GetFullPath(AddinPath);
            if (!File.Exists(AddinPath))
            {
                throw new FileNotFoundException("보고서 자동화 추가기능 파일을 찾을 수 없습니다.", AddinPath);
            }
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
            options.AddinPath = Get(values, "addin", PathResolver.ResolveDefaultAddinPath());
            options.OutputType = Get(values, "output", "Excel 산출 시트");
            options.ReportProfile = Get(values, "report-profile", "인식도/만족도 조사형");
            options.StyleProfile = Get(values, "style-profile", "공식 보고서체");
            options.HwpTemplatePath = Get(values, "hwp-template", "");
            options.PptTemplatePath = Get(values, "ppt-template", "");
            options.HwpTableStyleProfilePath = Get(values, "hwp-table-style-profile", "");
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
            string workbookToOpen = options.CopyWorkbook ? CreateWorkingCopy(options.WorkbookPath) : options.WorkbookPath;
            log("대상 파일: " + workbookToOpen);
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
                workbook = xl.Workbooks.Open(workbookToOpen);
                addin = xl.Workbooks.Open(options.AddinPath);
                dynamic wb = workbook;
                wb.Activate();

                string macroName = "'" + Path.GetFileName(options.AddinPath).Replace("'", "''") + "'!ReportAutomation_RunWithOptionsSilent";
                log("보고서 자동화 매크로를 실행합니다.");
                xl.Run(macroName, options.BannerSetting, options.TitlePrefixes);

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
                CloseComWorkbookForLauncher(workbook, true);
                QuitComExcelForLauncher(excel);
                ReleaseComForLauncher(addin);
                ReleaseComForLauncher(workbook);
                ReleaseComForLauncher(excel);
            }

            return workbookToOpen;
        }

        private static string CreateWorkingCopy(string sourcePath)
        {
            string directory = Path.GetDirectoryName(sourcePath);
            string name = Path.GetFileNameWithoutExtension(sourcePath);
            string extension = Path.GetExtension(sourcePath);
            string copyPath = Path.Combine(directory, name + "_report_alpha_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + extension);
            File.Copy(sourcePath, copyPath, true);
            return copyPath;
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
                writer.WriteLine("BannerSetting=" + options.BannerSetting);
                writer.WriteLine("TitlePrefixes=" + options.TitlePrefixes);
                writer.WriteLine("DecimalPlaces=" + options.DecimalPlaces);
                writer.WriteLine("ChartOutputMode=" + options.ChartOutputMode);
                writer.WriteLine("TableInsertMode=" + options.TableInsertMode);
                writer.WriteLine("HwpMaxSections=" + options.HwpMaxSections);
                writer.WriteLine("HwpDispatchMode=" + options.HwpDispatchMode);
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
                using (Process process = Process.Start(startInfo))
                {
                    if (!process.WaitForExit(120000))
                    {
                        try { process.Kill(); } catch { }
                        log("문장 초안 생성 시간이 길어져 이번 실행에서는 건너뜁니다.");
                        return "";
                    }

                    string stdout = process.StandardOutput.ReadToEnd();
                    string stderr = process.StandardError.ReadToEnd();

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
                                      " --hwp-template " + Quote(options.HwpTemplatePath) +
                                      " --ppt-template " + Quote(options.PptTemplatePath) +
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
                using (Process process = Process.Start(startInfo))
                {
                    if (!process.WaitForExit(120000))
                    {
                        try { process.Kill(); } catch { }
                        log("preflight 생성 시간이 길어져 이번 실행에서는 건너뜁니다.");
                        return;
                    }
                    string stdout = process.StandardOutput.ReadToEnd();
                    string stderr = process.StandardError.ReadToEnd();
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
                startInfo.Arguments = Quote(toolPath) +
                                      " --package " + Quote(options.LastReportPackagePath) +
                                      " --preflight " + Quote(options.LastPreflightReportPath) +
                                      " --template " + Quote(options.HwpTemplatePath) +
                                      " --output " + Quote(outputPath) +
                                      " --visible " + Quote(options.HwpVisible ? "true" : "false") +
                                     " --report-output " + Quote(reportPath) +
                                     " --render-plan-output " + Quote(renderPlanPath) +
                                     " --max-sections " + Quote(options.HwpMaxSections.ToString()) +
                                     " --dispatch-mode " + Quote(options.HwpDispatchMode) +
                                     OptionalArgument(" --table-style-profile ", options.HwpTableStyleProfilePath) +
                                     (options.HwpKeepOpenOnError ? " --keep-open-on-error" : "");
                startInfo.UseShellExecute = false;
                startInfo.CreateNoWindow = true;
                startInfo.RedirectStandardOutput = true;
                startInfo.RedirectStandardError = true;
                startInfo.StandardOutputEncoding = System.Text.Encoding.UTF8;
                startInfo.StandardErrorEncoding = System.Text.Encoding.UTF8;

                log("아래한글 COM으로 HWPX 초본을 생성합니다.");
                using (Process process = Process.Start(startInfo))
                {
                    if (!process.WaitForExit(180000))
                    {
                        try { process.Kill(); } catch { }
                        log("HWPX 생성 시간이 길어져 이번 실행에서는 중단했습니다.");
                        options.LastHwpWriterReportPath = File.Exists(reportPath) ? reportPath : "";
                        return;
                    }
                    string stdout = process.StandardOutput.ReadToEnd();
                    string stderr = process.StandardError.ReadToEnd();
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
            using (Process process = Process.Start(startInfo))
            {
                if (!process.WaitForExit(60000))
                {
                    try { process.Kill(); } catch { }
                    log("아래한글 COM 환경 확인 시간이 초과되었습니다.");
                    return false;
                }
                string stdout = process.StandardOutput.ReadToEnd();
                string stderr = process.StandardError.ReadToEnd();
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
        public string SheetName;
        public string TableNo;
        public string Title;
        public int Row;
    }

    internal static class BannerInspector
    {
        private const int MaxPreviewRows = 5000;
        private const int MaxPreviewColumns = 300;

        public static List<string> ReadBanners(string workbookPath)
        {
            WorkbookPreview preview = ReadPreview(workbookPath);
            if (preview.Banners.Count == 1)
            {
                throw new InvalidOperationException("집계표에서 배너 목록을 찾지 못했습니다. 표 헤더 구조를 확인하세요.");
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

                foreach (dynamic ws in wb.Worksheets)
                {
                    if (IsGeneratedSheet(Convert.ToString(ws.Name)))
                    {
                        continue;
                    }

                    ScanWorksheet(ws, found, seen, primary, primarySeen, preview.Tables);
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

            var tableStarts = new List<int>();
            for (int row = sheet.FirstRow; row <= sheet.LastRow; row++)
            {
                string text = sheet.Text(row, 1);
                if (IsTableTitle(text))
                {
                    tableStarts.Add(row);
                }
            }

            for (int i = 0; i < tableStarts.Count; i++)
            {
                int startRow = tableStarts[i];
                int endRow = (i + 1 < tableStarts.Count) ? tableStarts[i + 1] - 1 : sheet.LastRow;
                AddTablePreview(tables, Convert.ToString(ws.Name), sheet.Text(startRow, 1), startRow);
                int totalRow = FindTotalRow(sheet, startRow, endRow);
                if (totalRow == 0)
                {
                    continue;
                }

                foreach (string banner in FindBannerGroups(sheet, startRow, totalRow, sheet.LastCol))
                {
                    if (seen.Add(banner))
                    {
                        found.Add(banner);
                    }
                    if (IsPrimaryBannerCandidate(banner) && primarySeen.Add(banner))
                    {
                        primary.Add(banner);
                    }
                }
            }
        }

        private static bool IsTableTitle(string text)
        {
            text = (text ?? "").Trim();
            return text.StartsWith("[표", StringComparison.Ordinal) ||
                   text.StartsWith("[ 표", StringComparison.Ordinal) ||
                   text.StartsWith("<표", StringComparison.Ordinal);
        }

        private static void AddTablePreview(List<TablePreview> tables, string sheetName, string rawTitle, int row)
        {
            if (tables.Count >= 500)
            {
                return;
            }

            var preview = new TablePreview();
            preview.SheetName = sheetName;
            preview.Row = row;
            preview.TableNo = ParseTableNo(rawTitle);
            preview.Title = ParseTableTitle(rawTitle);
            tables.Add(preview);
        }

        private static string ParseTableNo(string rawTitle)
        {
            string text = (rawTitle ?? "").Trim();
            int open = text.IndexOf("[", StringComparison.Ordinal);
            int close = text.IndexOf("]", StringComparison.Ordinal);
            if (open >= 0 && close > open)
            {
                string token = text.Substring(open + 1, close - open - 1);
                token = token.Replace("표", "").Trim();
                return token;
            }
            return "";
        }

        private static string ParseTableTitle(string rawTitle)
        {
            string text = (rawTitle ?? "").Trim();
            int close = text.IndexOf("]", StringComparison.Ordinal);
            if (close >= 0 && close + 1 < text.Length)
            {
                text = text.Substring(close + 1).Trim();
            }
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
                string text = sheet.Text(row, 1);
                if (text.StartsWith("■", StringComparison.Ordinal))
                {
                    return row;
                }
            }
            return 0;
        }

        private static List<string> FindBannerGroups(SheetSnapshot sheet, int startRow, int totalRow, int lastCol)
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

            var colGroupName = new string[lastCol + 1];
            string lastSeenGroupName = "";
            for (int col = 1; col <= lastCol; col++)
            {
                for (int row = headerStart; row <= headerEnd; row++)
                {
                    string text = sheet.Text(row, col);
                    if (IsGroupNameCandidate(text))
                    {
                        colGroupName[col] = text;
                        lastSeenGroupName = text;
                        break;
                    }
                }
                if (string.IsNullOrEmpty(colGroupName[col]) && IsPercentMeasureColumn(sheet, startRow, totalRow, col))
                {
                    colGroupName[col] = lastSeenGroupName;
                }
            }

            string currentName = "";
            int percentCount = 0;
            for (int col = 1; col <= lastCol; col++)
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

        private static bool IsPercentMeasureColumn(SheetSnapshot sheet, int startRow, int totalRow, int col)
        {
            for (int row = startRow + 1; row <= totalRow - 1; row++)
            {
                if (sheet.Text(row, col) == "%")
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
            string normalized = text.Trim();
            return normalized != "%" &&
                   !string.Equals(normalized, "N", StringComparison.OrdinalIgnoreCase) &&
                   normalized != "사례수";
        }

        private static void AddBanner(List<string> groups, string name)
        {
            name = (name ?? "").Trim();
            if (name.Length == 0)
            {
                return;
            }
            if (!name.StartsWith("◐", StringComparison.Ordinal))
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
            string name = (banner ?? "").Trim();
            while (name.StartsWith("◐", StringComparison.Ordinal) || name.StartsWith("▣", StringComparison.Ordinal))
            {
                name = name.Substring(1).Trim();
            }
            return name;
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
