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
            markup1.AppendField(VA.Text.Markup.FieldConstants.PageName);
            markup1.AppendText(" (");
            markup1.AppendField(VA.Text.Markup.FieldConstants.PageNumber);
            markup1.AppendText(" of ");
            markup1.AppendField(VA.Text.Markup.FieldConstants.NumberOfPages);
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

        public static void FontCompare()
        {
            var stencil = SampleEnvironment.Application.Documents.OpenStencil("basic_u.vss");
            var master = stencil.Masters["Rectangle"];
            var page1 = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            var text1 =
                @"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvqxyz1234567890!@#$%^&*()``_-+=[]{}\|;:'"",.<>/?";
            var fonts = new[] { "Calibri", "Arial" };
            var sizes = new[] { "8.pt", "10.0pt", "12.0pt", "14.0pt", "18.0pt", "28.0pt"};
            var fontids = fonts.Select(f => page1.Document.Fonts[f].ID).ToList();

            var r = new System.Random();
            double left=1;
            double vs = 0.25;
            double cell1_w = 2.0;
            double cell1_h = 2.0;
            double cell1_top = 8.5;
            double cell2_h = 0.5;
            double cell2_w = 4.0;
            double cell_sep = 0.5;

            var char1 = new VA.Text.CharacterFormatCells();
            char1.Font = 0;
            char1.Size = "30pt";

            var fmt1 = new VA.Format.ShapeFormatCells();
            fmt1.LineWeight = 0;
            fmt1.LinePattern = 0;

            var fmt2 = new VA.Format.ShapeFormatCells();
            fmt2.LineWeight = 0;
            fmt2.LinePattern = 0;

            var fmt3 = new VA.Format.ShapeFormatCells();
            fmt3.LineWeight = 0;
            fmt3.LinePattern = 0;


            foreach (var size in sizes)
            {
                var shape1 = page1.DrawRectangle(left, cell1_top - cell1_h, left + cell1_w, cell1_top);
                shape1.Text = string.Format("{0}", size);

                double cell2_top = cell1_top;
                for (int i = 0; i < fonts.Count(); i++)
                {
                    double cell2_bottom = cell2_top - cell2_h;
                    var fontname = fonts[i];
                    double cell2_left = left + cell1_w + cell_sep;
                    var shape2 = page1.DrawRectangle(cell2_left, cell2_bottom, cell2_left + cell2_w, cell2_top);
                    shape2.Text = string.Format("{0}", fontname);

                    double cell3_h = r.NextDouble()*3.0 + 0.5;
                    var cell3_top = cell2_bottom;
                    var cell3_bottom = cell3_top - cell3_h;

                    var shape_3 = page1.DrawRectangle(cell2_left, cell3_bottom, cell2_left + cell2_w, cell3_top);
                    shape_3.Text = text1;


                    var char3 = new VA.Text.CharacterFormatCells();
                    char3.Font = fontids[i];
                    char3.Size = size;

                    var para3 = new VA.Text.ParagraphFormatCells();
                    para3.HorizontalAlign = 0;
                    var tb3 = new VA.Text.TextBlockFormatCells();
                    tb3.VerticalAlign = 0;

                    var update3 = new VA.ShapeSheet.Update.SRCUpdate();
                    para3.Apply(update3,0);
                    tb3.Apply(update3);
                    char3.Apply(update3,0);
                    fmt3.Apply(update3);
                    update3.Execute(shape_3);

                    var update1 = new VA.ShapeSheet.Update.SRCUpdate();
                    //para1.Apply(update1, 0);
                    //tb1.Apply(update1);
                    char1.Apply(update1, 0);
                    fmt1.Apply(update1);
                    update1.Execute(shape1);

                    var update2 = new VA.ShapeSheet.Update.SRCUpdate();
                    //para1.Apply(update2, 0);
                    //tb1.Apply(update2);
                    //char2.Apply(update2, 0);
                    fmt2.Apply(update2);
                    update1.Execute(shape2);


                    shape_3.CellsU["Height"].FormulaU = "TEXTHEIGHT(TheText,TxtWidth)";
                    var cell3_real_size = new VA.Drawing.Size(shape_3.CellsU["Width"].get_Result(null),
                                                              shape_3.CellsU["Height"].get_Result(null));
                    shape_3.CellsU["PinY"].FormulaU = (cell2_bottom - (cell3_real_size.Height/2.0)).ToString();

                    cell2_top -= cell2_h + cell3_real_size.Height + vs;
                }
                cell1_top = cell2_top ;

            }


            page1.ResizeToFitContents(1.0, 1.0);

        }

    }
}