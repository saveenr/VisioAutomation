using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using VisioAutomation.Extensions;
using VA=VisioAutomation;

namespace VisioPowerTools2010
{
    public partial class VPTRibbon
    {
        private void VPTRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void buttonHelp_Click_1(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.MessageBox.Show("Hello World");

        }

        private void buttonImportColors_Click(object sender, RibbonControlEventArgs e)
        {
            var form = new FormImportColors();
            var result = form.ShowDialog();
            if (result == DialogResult.OK)
            {
                var colors = form.Colors;
                if (colors.Count < 1)
                {
                    return;
                }

                var app = Globals.ThisAddIn.Application;
                var docs = app.Documents;
                var doc = docs.Add("");
                var page = doc.Pages[1];

                var dom = new VA.DOM.Document();
                double y = 8;
                double col0_w = 3.0;
                double col0_x = 0;
                double sep = 0.5;
                double col1_x = col0_x + col0_w + sep;
                double cellwidth = 1.0;
                double col2_x = col1_x + cellwidth + sep;
                double col3_x = col2_x + cellwidth + sep;
                var sb = new System.Text.StringBuilder();
                foreach (var color in colors)
                {
                    var shape0 = dom.DrawRectangle(col0_x, y, col0_x + col0_w, y + cellwidth);
                    var shape1 = dom.DrawRectangle(col1_x, y, col1_x + cellwidth, y + cellwidth);
                    var shape2 = dom.DrawRectangle(col2_x, y, col2_x + cellwidth, y + cellwidth);
                    var shape3 = dom.DrawRectangle(col3_x, y, col3_x + cellwidth, y + cellwidth);
                    var fill = new VisioAutomation.Drawing.ColorRGB(color.R, color.G, color.B);
                    string color_formula = fill.ToFormula();
                    double trans = (color.A/255.0);
                    string transparency_formula = trans.ToString(System.Globalization.CultureInfo.InvariantCulture);



                    shape1.Cells.FillForegnd = color_formula;
                    shape1.Cells.LinePattern = "0";
                    shape1.Cells.LineWeight = "0";
                    shape2.Cells.LineColor = color_formula;
                    shape2.Cells.LineWeight= "0.25in";
                    shape2.Cells.FillPattern = "0";
                    shape3.Cells.CharColor= color_formula;
                    shape3.Cells.FillPattern = "0";
                    shape3.Cells.LinePattern= "0";
                    shape3.Cells.LineWeight = "0";
                    shape3.Text = new VA.Text.Markup.TextElement("ABC");
                    shape3.CharFontName = "Segoe UI";
                    shape3.Cells.CharSize = "24pt";

                    sb.Clear();
                    sb.AppendFormat("rgb({0},{1},{2})\n", color.R, color.G, color.B);
                    sb.AppendFormat("{0}\n", System.Drawing.ColorTranslator.ToHtml(color));

                    if (color.A != 255)
                    {
                        sb.AppendFormat("transparency={0:0.00}", trans);
                    }

                    shape0.Text = new VA.Text.Markup.TextElement(sb.ToString());
                    shape0.CharFontName = "Segoe UI";
                    shape0.Cells.HAlign = "0";
                   shape0.Cells.VerticalAlign = "0";
                    shape0.Cells.LinePattern = "0";
                    shape0.Cells.LineWeight = "0";
                    shape0.Cells.FillForegnd = "rgb(240,240,240)";

                    if (color.A != 255)
                    {
                        shape1.Cells.FillForegndTrans = transparency_formula;
                        shape2.Cells.LineColorTrans = transparency_formula;
                        shape3.Cells.CharTransparency = transparency_formula;
                    }


                    y -= cellwidth + sep;
                }

                dom.Render(page);
                page.ResizeToFitContents(cellwidth,cellwidth);
                var window = app.ActiveWindow;
                window.ShowPageBreaks = 0;
                window.ShowGuides = 0;
                window.DeselectAll();
            }

        }

        private void buttonCreateStencilCatalog_Click(object sender, RibbonControlEventArgs e)
        {
            var form = new FormGetMasterImages();
            form.ShowDialog();
        }

        private void buttonCreateStyle_Click(object sender, RibbonControlEventArgs e)
        {
            var form = new FormCreateStyle();
            var result = form.ShowDialog();
            if (result == DialogResult.OK)
            {
                string name = form.StyleName.Trim();
                if (name.Length < 1)
                {
                    MessageBox.Show("Must have non-empty name");
                    return;
                }

                var app = Globals.ThisAddIn.Application;
                var doc = app.ActiveDocument;
                var styles = doc.Styles;
                var names = styles.AsEnumerable().Select(s => s.NameU).ToList();
                if (names.Contains(name.ToLower()))
                {
                    MessageBox.Show("Style with that name already exists");                    
                    return;
                }

                var style = styles.Add(name, "", VA.Convert.BoolToShort(form.IncludesText), VA.Convert.BoolToShort(form.IncludesLine), VA.Convert.BoolToShort(form.IncludesFill));
                int x = 1;
            }
        }
    }
}
