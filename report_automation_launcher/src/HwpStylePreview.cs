using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;

namespace ReportAutomationLauncher
{
    internal sealed class HwpStylePreview : Control
    {
        private const string MainSentence = "전체 응답자의 65.0%가 서비스를 이용한 경험이 있는 것으로 나타남.";
        private const string DetailSentence = "연령별로는 20대(72.0%)가 30대(58.0%)보다 높게 나타남.";
        private const string BaseLine = "[BASE : 전체 | 단위 : 명, %]";
        private static readonly string[,] Sample =
        {
            { "응답자 특성", "", "표본", "이용 경험", "" },
            { "구분", "", "사례수", "있음", "없음" },
            { "전체", "", "(100)", "65.0", "35.0" },
            { "연령", "20대", "(50)", "72.0", "28.0" },
            { "", "30대", "(50)", "58.0", "42.0" }
        };
        private Dictionary<string, object> config;

        internal HwpStylePreview()
        {
            DoubleBuffered = true;
            ResizeRedraw = true;
            BackColor = LauncherUi.ColorSurface;
            AccessibleName = "한글 보고서 서식 미리보기";
            AccessibleRole = AccessibleRole.Graphic;
            TabStop = false;
        }

        internal void UpdateStyle(Dictionary<string, object> value)
        {
            config = value;
            var paragraphs = Map(config, "paragraph_styles");
            string bullet = Convert.ToString(Map(paragraphs, "보고서 본문2")["bullet"]);
            AccessibleDescription = "예시 분석문: " + MainSentence + " " + bullet + " " + DetailSentence +
                " 예시 집계표: " + BaseLine + " 전체 100명, 이용 경험 있음 65.0%, 없음 35.0%.";
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            if (config == null || Width < 100 || Height < 60) return;
            var g = e.Graphics;
            float padding = 12 * g.DpiX / 96;
            float split = Width * 0.42f;
            using (var pen = new Pen(LauncherUi.ColorBorder))
            using (var heading = LauncherUi.SectionFont())
            {
                g.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);
                g.DrawLine(pen, split, padding, split, Height - padding);
                g.DrawString("예시 분석문", heading, Brushes.Black, padding, padding);
                g.DrawString("예시 집계표", heading, Brushes.Black, split + padding, padding);
                float top = padding + heading.GetHeight(g) + padding / 2;
                var body = new RectangleF(padding, top, split - 2 * padding, Height - top - padding);
                var table = new RectangleF(split + padding, top, Width - split - 2 * padding, Height - top - padding);
                if (body.Height <= 0) return;
                var paragraphs = Map(config, "paragraph_styles");
                var state = g.Save();
                g.SetClip(body, CombineMode.Intersect);
                float y = DrawParagraph(g, body, body.Top + 4, Map(paragraphs, "보고서 본문1"), MainSentence);
                DrawParagraph(g, body, y + 4, Map(paragraphs, "보고서 본문2"), DetailSentence);
                g.Restore(state);
                DrawTable(g, table, paragraphs);
            }
        }

        private static float DrawParagraph(Graphics g, RectangleF box, float y, Dictionary<string, object> style, string text)
        {
            string bullet = Convert.ToString(style["bullet"]);
            if (bullet.Length > 0) text = bullet + " " + text;
            using (Font font = StyleFont(style))
            {
                float step = font.GetHeight(g) * Number(style, "line_spacing_percent") / 100;
                float left = Number(style, "left_indent_mm");
                float indent = left + Number(style, "first_line_indent_mm");
                while (text.Length > 0 && y + font.GetHeight(g) <= box.Bottom)
                {
                    float x = Math.Max(box.Left, box.Left + Pixels(g, indent));
                    float width = box.Right - x;
                    if (width < 1)
                    {
                        if (indent == left) break;
                    }
                    else
                    {
                        int length = text.Length;
                        while (length > 1 && g.MeasureString(text.Substring(0, length), font).Width > width) length--;
                        if (length < text.Length && text[length] != ' ')
                        {
                            int space = text.LastIndexOf(' ', length - 1);
                            if (space >= length / 2 && space > 0) length = space;
                        }
                        DrawText(g, text.Substring(0, length), font, new RectangleF(x, y, width, font.GetHeight(g) + 2),
                            Convert.ToString(style["alignment"]), "top");
                        text = text.Substring(length).TrimStart();
                    }
                    y += step;
                    indent = left;
                }
                return y;
            }
        }

        private void DrawTable(Graphics g, RectangleF box, Dictionary<string, object> paragraphs)
        {
            var cells = Map(config, "table_cell_styles");
            var border = Map(config, "table_border");
            float textHeight = Font.GetHeight(g) + 4;
            DrawText(g, BaseLine, Font, new RectangleF(box.X, box.Y, box.Width, textHeight), "right", "center");
            box.Y += textHeight;
            box.Height -= textHeight + 2;
            if (box.Height <= 0) return;

            var heights = new float[Sample.GetLength(0)];
            for (int r = 0; r < heights.Length; r++)
                for (int c = 0; c < Sample.GetLength(1); c++)
                {
                    if (Sample[r, c].Length == 0) continue;
                    string name = CellStyleName(cells, r, c);
                    var style = Map(paragraphs, name);
                    var cell = Map(cells, name);
                    using (Font font = StyleFont(style))
                        heights[r] = Math.Max(heights[r], font.GetHeight(g) * Number(style, "line_spacing_percent") / 100 +
                            Pixels(g, CellMarginMm(cell, "top") + CellMarginMm(cell, "bottom")) + 4);
                }
            float scale = Math.Min(1, box.Height / heights.Sum());
            var ys = new float[heights.Length + 1];
            ys[0] = box.Top;
            for (int r = 0; r < heights.Length; r++) ys[r + 1] = ys[r] + heights[r] * scale;
            float column = box.Width / Sample.GetLength(1);
            for (int r = 0; r < heights.Length; r++)
                for (int c = 0; c < Sample.GetLength(1); c++)
                {
                    if (Sample[r, c].Length == 0) continue;
                    string name = CellStyleName(cells, r, c);
                    var style = Map(paragraphs, name);
                    var cell = Map(cells, name);
                    bool mergedAcross = (c == 0 && r < 3) || (r == 0 && c == 3);
                    var bounds = new RectangleF(box.Left + c * column, ys[r], mergedAcross ? 2 * column : column,
                        ys[r == 3 && c == 0 ? r + 2 : r + 1] - ys[r]);
                    using (var fill = new SolidBrush(ColorTranslator.FromHtml("#" + cell["fill_color"])))
                        g.FillRectangle(fill, bounds);
                    float left = Math.Max(0, Pixels(g, CellMarginMm(cell, "left") + Number(style, "left_indent_mm") + Number(style, "first_line_indent_mm")));
                    float right = Pixels(g, CellMarginMm(cell, "right"));
                    float top = Pixels(g, CellMarginMm(cell, "top"));
                    float bottom = Pixels(g, CellMarginMm(cell, "bottom"));
                    var area = RectangleF.FromLTRB(bounds.Left + left, bounds.Top + top, bounds.Right - right, bounds.Bottom - bottom);
                    if (area.Width <= 0 || area.Height <= 0) continue;
                    string bullet = Convert.ToString(style["bullet"]);
                    using (Font font = StyleFont(style))
                        DrawText(g, bullet.Length == 0 ? Sample[r, c] : bullet + " " + Sample[r, c],
                            font, area, Convert.ToString(style["alignment"]), Convert.ToString(cell["vertical_alignment"]));
                }

            for (int r = 0; r < ys.Length; r++)
                DrawEdge(g, border, r == 0 || r == heights.Length ? "outer" : r == 2 ? "header_bottom" : "inner_horizontal",
                    new PointF(box.Left + (r == 4 ? column : 0), ys[r]), new PointF(box.Right, ys[r]));
            for (int c = 0; c <= Sample.GetLength(1); c++)
            {
                bool outer = c == 0 || c == Sample.GetLength(1);
                if (outer && !(bool)border["show_side_borders"]) continue;
                float x = box.Left + c * column;
                DrawEdge(g, border, outer ? "outer" : "inner_vertical", new PointF(x, ys[c == 1 ? 3 : c == 4 ? 1 : 0]),
                    new PointF(x, ys[heights.Length]));
            }
        }

        private static string CellStyleName(Dictionary<string, object> cells, int row, int column)
        {
            string role = row == 0 ? "표가로배너" : row == 1 ? "표머리행" : row == 2 ? "표BASE" :
                column == 0 ? "표세로배너" : column == 1 ? "표보기" : "표숫자";
            return cells.ContainsKey(role) ? role : HwpStylePresetStore.RoleStyleDefaults[role];
        }

        private static void DrawEdge(Graphics g, Dictionary<string, object> border, string group, PointF a, PointF b)
        {
            Dictionary<string, object> line = null;
            if (border.ContainsKey("edges"))
            {
                var edges = Map(border, "edges");
                if (edges.ContainsKey(group)) line = Map(edges, group);
                else if (group == "header_bottom" && edges.ContainsKey("inner_horizontal")) line = Map(edges, "inner_horizontal");
            }
            string type = line == null ? "solid" : Convert.ToString(line["type"]);
            if (type == "none") return;
            float width = Pixels(g, line == null ? Number(border, group == "outer" ? "outer_width_mm" : "inner_width_mm") : Number(line, "width_mm"));
            using (var pen = new Pen(ColorTranslator.FromHtml("#" + (line ?? border)["color"]), Math.Max(1, width)))
            {
                if (type == "dot") pen.DashStyle = DashStyle.Dot;
                if (type == "dash") pen.DashStyle = DashStyle.Dash;
                if (type == "double_slim")
                {
                    pen.Width = Math.Max(1, width / 3);
                    float gap = Math.Max(1, width / 3);
                    float dx = a.X == b.X ? gap : 0, dy = a.Y == b.Y ? gap : 0;
                    g.DrawLine(pen, a.X - dx, a.Y - dy, b.X - dx, b.Y - dy);
                    g.DrawLine(pen, a.X + dx, a.Y + dy, b.X + dx, b.Y + dy);
                }
                else g.DrawLine(pen, a, b);
            }
        }

        private static void DrawText(Graphics g, string text, Font font, RectangleF area, string horizontal, string vertical)
        {
            using (var format = new StringFormat { FormatFlags = StringFormatFlags.NoWrap, Trimming = StringTrimming.EllipsisCharacter })
            {
                format.Alignment = horizontal == "right" ? StringAlignment.Far : horizontal == "center" ? StringAlignment.Center : StringAlignment.Near;
                format.LineAlignment = vertical == "bottom" ? StringAlignment.Far : vertical == "center" ? StringAlignment.Center : StringAlignment.Near;
                g.DrawString(text, font, Brushes.Black, area, format);
            }
        }

        private static Font StyleFont(Dictionary<string, object> style)
        {
            float size = Number(style, "font_size_pt");
            FontStyle weight = (bool)style["bold"] ? FontStyle.Bold : FontStyle.Regular;
            try { return new Font(Convert.ToString(style["font_family"]), size, weight); }
            catch (ArgumentException) { return new Font("맑은 고딕", size, weight); }
        }

        private static float CellMarginMm(Dictionary<string, object> cell, string side)
        {
            return cell.ContainsKey("margins_mm") ? Number(Map(cell, "margins_mm"), side) :
                side == "top" || side == "bottom" ? 0.3f : 1.8f;
        }

        private static float Pixels(Graphics g, float mm) { return mm * g.DpiX / 25.4f; }
        private static float Number(Dictionary<string, object> values, string key) { return Convert.ToSingle(values[key]); }
        private static Dictionary<string, object> Map(Dictionary<string, object> values, string key) { return (Dictionary<string, object>)values[key]; }
    }
}
