using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using VisioAutomation.Extensions;
using VA=VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using DGMODEL = VisioAutomation.Models.DirectedGraph;

namespace VisioPowerTools2010
{
    public partial class VPTRibbon
    {
        private VisioAutomation.Scripting.Client client;
 
        private void VPTRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            try
            {
                this.client = new VisioAutomation.Scripting.Client(Globals.ThisAddIn.Application);
            }
            catch (System.Exception)
            {
                string msg = "Failed to load Visio Power Tools";
                MessageBox.Show(msg);
            }
        }

        private void execute_cmd(System.Action func)
        {
            try
            {
                func();
            }
            catch (System.Exception)
            {
                string msg = "Failed to execute command";
                MessageBox.Show(msg);
            }
        }

        private void buttonHelp_Click_1(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("Hello World");

        }

        private void buttonImportColors_Click(object sender, RibbonControlEventArgs e)
        {
            this.execute_cmd(this.cmd_import_colors );
        }

        private void buttonCreateStencilCatalog_Click(object sender, RibbonControlEventArgs e)
        {
            this.execute_cmd(this.cmd_create_stencil_catalog);
        }

        private void buttonCreateStyle_Click(object sender, RibbonControlEventArgs e)
        {
            this.execute_cmd(this.cmd_create_style);
        }

        private void buttonToggleTextCase_Click(object sender, RibbonControlEventArgs e)
        {
            this.execute_cmd(this.cmd_toggle_text_case);
        }

        private void buttonCopyText_Click(object sender, RibbonControlEventArgs e)
        {
            this.execute_cmd(this.cmd_copy_text);
        }

        // -----------------------------------------------------------------------------------------------'
        private void cmd_create_style()
        {
            var app = Globals.ThisAddIn.Application;
            var doc = app.ActiveDocument;

            if (doc == null)
            {
                MessageBox.Show("Must have a document open");
                return;
            }

            if (doc.Type != IVisio.VisDocumentTypes.visTypeDrawing)
            {
                MessageBox.Show("Must have a drawing open");
                return;
            }

            var styles = doc.Styles;

            var form = new FormCreateStyle();
            var result = form.ShowDialog();

            if (result != DialogResult.OK)
            {
                return;
            }

            string name = form.StyleName.Trim();

            if (name.Length < 1)
            {
                MessageBox.Show("Must have non-empty name");
                return;
            }

            var names = styles.AsEnumerable().Select(s => s.NameU).ToList();
            var names_lc = names.Select(s => s.ToLower()).ToList();

            if (names_lc.Contains(name.ToLower()))
            {
                string msg = string.Format("Style with name \"{0}\" already exists", name);
                MessageBox.Show(msg);
                return;
            }

            short fIncludesText = VA.Convert.BoolToShort(form.IncludesText);
            short fIncludesLine = VA.Convert.BoolToShort(form.IncludesLine);
            short fIncludesFill = VA.Convert.BoolToShort(form.IncludesFill);
            var style = styles.Add(name, "", fIncludesText, fIncludesLine, fIncludesFill);

        }

        private void cmd_copy_text()
        {
            var shape_text = this.client.Text.Get(null);
            var text = string.Join("\r\n", shape_text) + "\r\n";
            Clipboard.SetText(text);
        }
        
        private void cmd_toggle_text_case()
        {
            this.client.Text.ToogleCase(null);
        }

        private void cmd_import_colors()
        {
            var form = new FormImportColors();
            var result = form.ShowDialog();
            if (result == DialogResult.OK)
            {
                var colors = form.Colors;
                VPTRibbon.draw_colors(colors);
            }

        }

        private void cmd_create_stencil_catalog()
        {
            var form = new FormGetMasterImages();
            form.ShowDialog();
        }

        private static void draw_colors(List<System.Drawing.Color> colors)
        {
            if (colors.Count < 1)
            {
                return;
            }

            var app = Globals.ThisAddIn.Application;

            var domdoc = new VA.DOM.Document();
            var page_node = new VA.DOM.Page();
            domdoc.Pages.Add(page_node);
          

            double y = 8;
            double col0_w = 3.0;
            double col0_x = 0;
            double sep = 0.5;
            double col1_x = col0_x + col0_w + sep;
            double cellwidth = 1.0;
            double col2_x = col1_x + cellwidth + sep;
            double col3_x = col2_x + cellwidth + sep;
            var sb = new StringBuilder();
            foreach (var color in colors)
            {
                var shape0 = page_node.Shapes.DrawRectangle(col0_x, y, col0_x + col0_w, y + cellwidth);
                var shape1 = page_node.Shapes.DrawRectangle(col1_x, y, col1_x + cellwidth, y + cellwidth);
                var shape2 = page_node.Shapes.DrawRectangle(col2_x, y, col2_x + cellwidth, y + cellwidth);
                var shape3 = page_node.Shapes.DrawRectangle(col3_x, y, col3_x + cellwidth, y + cellwidth);
                var fill = new VisioAutomation.Drawing.ColorRGB(color.R, color.G, color.B);
                string color_formula = fill.ToFormula();
                double trans = (color.A / 255.0);
                string transparency_formula = trans.ToString(System.Globalization.CultureInfo.InvariantCulture);


                shape1.Cells.FillForegnd = color_formula;
                shape1.Cells.LinePattern = "0";
                shape1.Cells.LineWeight = "0";
                shape2.Cells.LineColor = color_formula;
                shape2.Cells.LineWeight = "0.25in";
                shape2.Cells.FillPattern = "0";
                shape3.Cells.CharColor = color_formula;
                shape3.Cells.FillPattern = "0";
                shape3.Cells.LinePattern = "0";
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
                shape0.Cells.ParaHorizontalAlign = "0";
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

            page_node.ResizeToFit = true;
            page_node.ResizeToFitMargin = new VA.Drawing.Size(cellwidth, cellwidth);

            domdoc.Render(app);

            var window = app.ActiveWindow;
            window.ShowPageBreaks = 0;
            window.ShowGuides = 0;
            window.DeselectAll();
        }

        private void buttonDeveloper_Click(object sender, RibbonControlEventArgs e)
        {
            var form = new FormDeveloper();
            form.ShowDialog();
        }

        private void buttonGraph_Click(object sender, RibbonControlEventArgs e)
        {
            var form = new FormDirectedGraph();
            var result = form.ShowDialog();

            if (result != DialogResult.OK)
            {
                return;
            }

            var text = form.GraphText.Trim();
            var lines = text.Split('\n').Select(s => s.Trim()).Where( s=>s.Length>0).ToList();

            var model = new VA.Models.DirectedGraph.Drawing();

            int cn = 0;
            var dic = new Dictionary<string, VA.Models.DirectedGraph.Shape>();
            foreach (var line in lines)
            {
                var tokens = line.Split(new[] {"->"}, System.StringSplitOptions.RemoveEmptyEntries);
                if (tokens.Length==0)
                {
                    // do nothing
                }
                else if (tokens.Length==1)
                {
                    string from = tokens[0];
                    if (dic.ContainsKey(from))
                    {
                        
                    }
                    else
                    {
                    }
                }
                else if (tokens.Length >=2 )
                {
                    string from = tokens[0];
                    string to = tokens[1];

                    VA.Models.DirectedGraph.Shape fromnode;
                    VA.Models.DirectedGraph.Shape tonode;
                    if (!dic.ContainsKey(from))
                    {
                        fromnode = model.AddShape(from, from, "basic_u.vss", "rectangle");
                        fromnode.Label = from;
                        dic[from] = fromnode;
                    }
                    else
                    {
                        fromnode = dic[from];
                    }

                    if (!dic.ContainsKey(to))
                    {
                        tonode= model.AddShape(to, to, "basic_u.vss", "rectangle");
                        tonode.Label = to;
                        dic[to] = tonode;
                    }
                    else
                    {
                        tonode = dic[to];
                    }

                    model.AddConnection("C" + cn.ToString(), fromnode, tonode);
                    cn +=1;

                }
            }


            var app = Globals.ThisAddIn.Application;
            var doc = Globals.ThisAddIn.Application.ActiveDocument;
            IVisio.Page page;
            if (doc==null)
            {
                var docs = app.Documents;
                doc = docs.Add("");
                var pages = doc.Pages;
                page = pages[1];
            }
            else
            {
                page = doc.Pages.Add();                
            }

            var visio_options = new DGMODEL.VisioLayoutOptions();
            model.Render(page,visio_options);

            var pl = new VA.Pages.PageLayout.FlowchartLayout();
            pl.ConnectorStyle = VisioAutomation.Pages.PageLayout.ConnectorStyle.Flowchart;
            pl.ConnectorAppearance = VisioAutomation.Pages.PageLayout.ConnectorAppearance.Curved;
            pl.Apply(page);

            page.ResizeToFitContents();

        }

        private void buttonExportSelection_Click(object sender, RibbonControlEventArgs e)
        {
            var form = new FormExportSelectionAsFormat( FormExportSelectionAsFormat.enumExportFormat.ExportXAML);
            form.ShowDialog();
            
        }

        private void buttonSelectionXHTML_Click(object sender, RibbonControlEventArgs e)
        {
            var form = new FormExportSelectionAsFormat(FormExportSelectionAsFormat.enumExportFormat.ExportSVGXHTML);
            form.ShowDialog();
        }

        private void buttonScrambleText_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            var doc = app.ActiveDocument;
            if (doc == null)
            {
                return;
            }

            if (doc.Type != IVisio.VisDocumentTypes.visTypeDrawing )
            {
                MessageBox.Show("Currently Active Document is not a Drawing");
                return;
            }

            var activewindow = app.ActiveWindow;
            var sb = new StringBuilder();
            var pages = doc.Pages.AsEnumerable().ToList();

            using (var scope = new VA.Application.UndoScope(app, "Scramble Text"))
            {
                // Begin Undo Scope
                doc.Company = VPTRibbon.Scramble(sb, doc.Company);
                doc.Category = VPTRibbon.Scramble(sb,doc.Category);
                doc.Title = VPTRibbon.Scramble(sb, doc.Title);
                doc.Subject = VPTRibbon.Scramble(sb, doc.Subject);
                doc.Creator = VPTRibbon.Scramble(sb, doc.Creator);
                doc.Manager = VPTRibbon.Scramble(sb, doc.Manager);
                doc.Keywords = VPTRibbon.Scramble(sb, doc.Keywords);
                foreach (var page in pages)
                {
                    activewindow.Page = page;
                    var shapes = page.Shapes.AsEnumerable().ToList();
                    foreach (var shape in shapes)
                    {
                        VPTRibbon.Scramble(sb, shape);

                        var shape_shapes = shape.Shapes;
                        if (shape_shapes!=null && shape_shapes.Count>0)
                        {
                            foreach (var nested_shape in VA.Shapes.ShapeHelper.GetNestedShapes(shape))
                            {
                                VPTRibbon.Scramble(sb,nested_shape);                                
                            }
                        }
                    }
                }
                // End Undo Scope
            }
        }

        private static void Scramble(StringBuilder sb, IVisio.Shape shape)
        {
            sb.Clear();
            string text = shape.Text;
            string text_trimmed = text.Trim();
            if (text_trimmed.Length >= 1)
            {
                shape.Text = VPTRibbon.Scramble(sb,shape.Text);
            }
        }

        private static string Scramble(StringBuilder sb, string text)
        {
            string text_trimmed = text.Trim();
            if (text_trimmed.Length >= 1)
            {
                for (int i = 0; i < text_trimmed.Length; i++)
                {
                    sb.Append("X");
                }
                return sb.ToString();
            }
            return text;
        }

        private void buttomResetPageOrigin_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            var doc = app.ActiveDocument;
            if (doc == null)
            {
                return;
            }

            if (doc.Type != IVisio.VisDocumentTypes.visTypeDrawing)
            {
                MessageBox.Show("Currently Active Document is not a Drawing");
                return;
            }

            var page = app.ActivePage;
            if (page==null)
            {
                return;
            }
            this.client.Page.ResetOrigin(null);
        }

        private void buttonResizePageToFit_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            var doc = app.ActiveDocument;
            if (doc == null)
            {
                return;
            }

            if (doc.Type != IVisio.VisDocumentTypes.visTypeDrawing)
            {
                MessageBox.Show("Currently Active Document is not a Drawing");
                return;
            }

            var page = app.ActivePage;
            if (page == null)
            {
                return;
            }
            page.ResizeToFitContents();

        }
    }
}
