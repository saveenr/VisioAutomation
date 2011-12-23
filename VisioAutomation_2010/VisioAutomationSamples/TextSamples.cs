using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using System.Linq;
using System.Collections.Generic;

namespace VisioAutomationSamples
{
    public static class TextSamples
    {
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

            var markup1 = new VA.Text.Markup.TextElement();
            markup1.AppendField(VA.Text.Markup.Fields.PageName);
            markup1.AppendText(" (");
            markup1.AppendField(VA.Text.Markup.Fields.PageNumber);
            markup1.AppendText(" of ");
            markup1.AppendField(VA.Text.Markup.Fields.NumberOfPages);
            markup1.AppendText(") ");
            markup1.SetText(s0);
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