using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using System.Linq;
using System.Collections.Generic;

namespace VisioAutomationSamples
{
    public static class TextSamples
    {
        public static VA.Text.Markup.TextElement AddElementEx(this VA.Text.Markup.TextElement p, string text, string font, double? size, int? color, VA.Drawing.AlignmentHorizontal? halign, VA.Text.CharStyle? cs)
        {
            var el = p.AppendElement(text);
            if (font != null)
            {
                el.TextFormat.Font = font;
            }
            if (size.HasValue)
            {
                el.TextFormat.FontSize = size.Value;
            }
            if (color.HasValue)
            {
                el.TextFormat.Color = new VA.Drawing.ColorRGB(color.Value);
            }
            if (halign.HasValue)
            {
                el.TextFormat.HAlign = halign.Value;
            }

            if (cs.HasValue)
            {
                el.TextFormat.CharStyle = cs;
            }

            return el;
        }

        public static VA.Text.Markup.TextElement get_markup_1()
        {
            var e1 = new VA.Text.Markup.TextElement();
            e1.AppendText("E1Calibri 15pt red\n");
            e1.TextFormat.Color = new VA.Drawing.ColorRGB(0xff0000);
            e1.TextFormat.Font = "Calibri";
            e1.TextFormat.FontSize = 15;

            var e2 = e1.AddElementEx("Segoe UI 20 pt blue]\n", "Segoe UI", 20, 0x0000ff, null, null);
            var e3 = e2.AddElementEx("E3 left italic gray\n", null, null, 0xa0a0a0, VA.Drawing.AlignmentHorizontal.Left, null);
            var e4 = e3.AddElementEx("E4 bold italic right\n", null, null, null, VA.Drawing.AlignmentHorizontal.Right, VA.Text.CharStyle.Bold | VA.Text.CharStyle.Italic);
            var e5 = e3.AddElementEx("E5 nobold, noitalic, Center]\n", null, null, null, VA.Drawing.AlignmentHorizontal.Center, VA.Text.CharStyle.None);

            e3.AppendElement("E3 FFF\n");
            e2.AppendElement("E2 [left italic gray\n");
            e1.AppendElement("E1 [Calibri 15pt red]\n");

            return e1;

        }

        public static void NonRotatingText()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();
            var s0 = page.DrawRectangle(1, 1, 4, 4);
            s0.Text = "Hello World";

            s0.GetCell(VA.ShapeSheet.SRCConstants.TxtAngle).Formula = "-Angle";
        }

        public static void TextFields()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();
            var s0 = page.DrawRectangle(1, 1, 4, 4);

            VA.Text.TextHelper.SetText(s0, "{0} ({1} of {2})", 
                VA.Text.Markup.Fields.NumberOfPages,
                VA.Text.Markup.Fields.PageNumber,
                VA.Text.Markup.Fields.PageName);
        }

        public static void TextMarkup1()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            // Create the Shapes that will hold the text
            var s1 = page.DrawRectangle(0, 0, 8, 8);
            var s2 = page.DrawRectangle(8, 0, 16, 8);
            var s3 = page.DrawRectangle(0, 8, 8, 16);
            var s4 = page.DrawRectangle(8, 8, 16, 16);

            var m1 = get_markup_1();
            m1.SetText(s1);
        }

        public static void TextMarkup2()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();
            var s0 = page.DrawRectangle(1, 1, 4, 4);

            page.DrawRectangle(1, 1, 4, 4);

            var tokens = new[] {"The ", "Quick ", "Brown ", "Fox"};
            var e1 = new VA.Text.Markup.TextElement();
            foreach (var token in tokens)
            {
                e1.AppendText(token);
            }
            //vi.Text.Markup = e1;
        }

        public static void TextSizing()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            var s0 = page.DrawRectangle(0, 0, 4, 4);

            // Alignment Box fits to accomodate text
            s0.Text = "Alignment Box fits to accomodate text";

            s0.GetCell(VA.ShapeSheet.SRCConstants.Width).Formula = "2.0";
            s0.GetCell(VA.ShapeSheet.SRCConstants.Height).Formula = "GUARD(TxtHeight)";
            s0.GetCell(VA.ShapeSheet.SRCConstants.TxtWidth).Formula = "Width*1";
            s0.GetCell(VA.ShapeSheet.SRCConstants.TxtHeight).Formula = "TEXTHEIGHT(TheText,TxtWidth)";

            // Text Scales Proportional to Shape Height
            var s1 = page.DrawRectangle(0, 4, 8, 8);
            s1.Text = "Text Scales Proportional to Shape Height";
            s0.GetCell(VA.ShapeSheet.SRCConstants.Char_Size).Formula = "Height*0.25";

            // Text scales smaller to fit more text
            var s2 = page.DrawRectangle(4, 0, 8, 4);
            s2.Text = "Text scales smaller to fit more text";
            s2.GetCell(VA.ShapeSheet.SRCConstants.Char_Size).Formula =
                "11pt * 10/SQRT(LEN(SHAPETEXT(TheText)))";
        }

        public static void FontChart()
        {
            var stencil = SampleEnvironment.Application.Documents.OpenStencil("basic_u.vss");
            var master = stencil.Masters["Rectangle"];
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            var fonts = new[] {"Segoe UI", "Calibri", "Arial"};
            var sizes = new[] {"28.0pt", "18.0pt", "14.0pt", "12.0pt", "10.0pt"};
            var fontids = fonts.Select(f => page.Document.Fonts[f].ID).ToList();

            var layout = new VA.Layout.Grid.GridLayout(sizes.Length, fonts.Length, new VA.Drawing.Size(3.0, 0.5), master);
            layout.Origin = new VA.Drawing.Point(0, VA.Pages.PageHelper.GetSize(page).Height);
            layout.CellSpacing = new VA.Drawing.Size(0.5, 0.5);
            layout.RowDirection = VA.Layout.Grid.RowDirection.TopToBottom;
            
            layout.PerformLayout();
            
            layout.Render(page);

            page.ResizeToFitContents(1.0,1.0);
            var nodes = layout.Nodes.ToList();

            var items = from fi in Enumerable.Range(0, fonts.Count())
                        from size in sizes
                        select new {font = fonts[fi], size = size, fontid = fontids[fi]};

            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();

            var charcells = new VA.Text.CharacterFormatCells();
            var fmt = new VA.Format.ShapeFormatCells();
            int i = 0;
            foreach (var item in items)
            {
                var shape = nodes[i].Shape;
                shape.Text = item.font + " " + item.size;
                var shapeid = nodes[i].ShapeID;
                charcells.Size = item.size;
                charcells.Font = item.fontid;
                charcells.Apply(update, shapeid, 0);

                fmt.FillForegnd = "rgb(250,250,250)";
                fmt.LinePattern = 0;
                fmt.LineWeight = 0;
                fmt.Apply(update,shapeid);
                i++;
            }

            update.Execute(page);
        }
    }
}