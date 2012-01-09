using VisioAutomation.DOM;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using System.Linq;
using System.Collections.Generic;
using BoxL = VisioAutomation.Layout.BoxLayout;

namespace VisioAutomationSamples
{
    public static class BoxLayoutSamples
    {
        public class NodeData
        {
            public IVisio.Shape VisioShape;
            public string Text;
            public bool Render;
            public VA.DOM.ShapeCells Cells;
            public string Font;

            public NodeData()
            {
                this.Render = true;
                this.Cells = new VA.DOM.ShapeCells();
            }
        }

        public static BoxL.Node AddNodeEx(this BoxL.Node p, double w, double h, string s)
        {
            var box = p.AddBox(w, h);
            var node_data = new NodeData();
            node_data.Text = s;
            box.Data = node_data;
            return box;
        }

        public static void BoxLayout()
        {
            // Create a layout
            var layout = BoxLayoutShared.CreateSampleLayout();

            // Ask the Layout to place the nodes
            var origin = new VA.Drawing.Point(0, 0);
            layout.LayoutOptions.Origin = origin;

            layout.PerformLayout();

            // Create a blank canvas in Visio 
            var app = SampleEnvironment.Application;
            var documents = app.Documents;
            var doc = documents.Add(string.Empty);
            var page1 = doc.Pages[1];

            // and tinker with it
            // render
            foreach (var node in layout.Nodes)
            {
                BoxLayoutShared.DrawNode(node, node.Rectangle, page1);
            }

            var src_linepat = VA.ShapeSheet.SRCConstants.LinePattern;
            var root_shape = (IVisio.Shape) layout.Root.Data;
            var cell_linepat = root_shape.GetCell(src_linepat);
            cell_linepat.FormulaU = "7";

            // Make the page big enough to fit what was drawn + a small border
            var margin = new VA.Drawing.Size(0.5, 0.5);
            page1.ResizeToFitContents(margin);
        }

        public static void FontCompare()
        {
            var visapp = new IVisio.Application();
            var doc = visapp.Documents.Add("");

            var fontnames = new[] {"Arial", "Roboto"};

            var sampletext = "ABCDEFGHIJKLMNOPQRSTUVWXYZ" + "abcdefghijklmnopqrstuvwxyz" +
                             "<>[](),./|\\:;\'\"1234567890!@#$%^&*()`~";
            var samplechars = sampletext.Select(c => new string(new char[] {c})).ToList();

            FontGlyphComparision(doc, fontnames, samplechars);
            FontGlyphComparision2(doc, fontnames, samplechars);
            FontGlyphComparision3(doc, fontnames, samplechars);
        }

        public static void FontGlyphComparision(IVisio.Document doc, string[] fontnames, List<string> samplechars)
        {
            var layout = new BoxL.BoxLayout();
            layout.LayoutOptions.DirectionVertical = VA.Layout.BoxLayout.DirectionVertical.TopToBottom;

            var root = layout.Root;
            root.Direction = BoxL.LayoutDirection.Vertical;
            root.ChildSeparation = 0.5;

            var nodedata = new NodeData();
            nodedata.Render = false;
            root.Data = nodedata;

            var fontname_cells = new VA.DOM.ShapeCells();
            fontname_cells.FillPattern = 0;
            fontname_cells.LinePattern = 0;
            fontname_cells.LineWeight = 0.0;
            fontname_cells.HAlign = 0;
            fontname_cells.CharSize = VA.Convert.PointsToInches(36.0);

            var charbox_cells = new VA.DOM.ShapeCells();
            charbox_cells.FillPattern = 0;
            charbox_cells.LinePattern = 1;
            charbox_cells.LineWeight = 0.0;
            charbox_cells.LineColor = "rgb(150,150,150)";
            charbox_cells.HAlign = 1;
            charbox_cells.CharSize = VA.Convert.PointsToInches(24.0);

            foreach (string fontname in fontnames)
            {
                var fontname_box = root.AddNodeEx(5, 0.5, fontname);
                var fontname_box_data = (NodeData) fontname_box.Data;
                fontname_box_data.Cells = fontname_cells;

                var font_box = root.AddBox(1.0,1.0);
                font_box.Direction = BoxL.LayoutDirection.Vertical;
                font_box.ChildSeparation = 0.25;
                var font_vox_data = (NodeData) font_box.Data;
                if (font_vox_data != null)
                {
                    font_vox_data.Render = false;                   
                }

                int numcols = 17;
                int numrows = 5;
                int numcells = numcols*numrows;


                foreach (int row in Enumerable.Range(0, numrows))
                {
                    var row_box = font_box.AddNodeEx(1.0,1.0,null);
                    row_box.Direction = BoxL.LayoutDirection.Horizontal;
                    row_box.ChildSeparation = 0.25;
                    var row_box_data = (NodeData) row_box.Data;
                    row_box_data.Render = false;

                    foreach (int col in Enumerable.Range(0, numcols))
                    {
                        int charindex = (col + (numcols*row))%numcells;
                        string curchar = samplechars[charindex];
                        var cell_box = row_box.AddNodeEx(0.50, 0.50, curchar);
                        var cell_box_data = (NodeData) cell_box.Data;
                        cell_box_data.Font = fontname;
                        cell_box_data.Cells = charbox_cells;
                    }
                }
            }

            layout.PerformLayout();

            var page = doc.Pages.Add();

            var dom = new VA.DOM.Document();

            foreach (var node in layout.Nodes)
            {
                if (node.Data == null)
                {
                    continue;
                }
                var node_data = (NodeData)node.Data;

                if (node_data.Render == false)
                {
                    continue;
                }

                var dom_shape = dom.Drop("Rectangle", "basic_u.vss", node.Rectangle);
       
                var cells = node_data.Cells;
                if (cells == null)
                {
                    cells = new VA.DOM.ShapeCells();
                }
                else
                {
                    cells = node_data.Cells.ShallowCopy();
                }

                if (node_data.Font != null)
                {
                    dom_shape.CharFontName = node_data.Font;
                }

                dom_shape.ShapeCells = cells;
                dom_shape.Text = new VA.Text.Markup.TextElement( node_data.Text );
            }

            dom.Render(page);
            page.ResizeToFitContents(0.5, 0.5);
        }

        public static void FontGlyphComparision2(IVisio.Document doc, string[] fontnames, List<string> samplechars)
        {
            double w = 2.0;
            double h = 1;
            double th = 1;

            int chunksize = 12;
            var chunks = LinqUtil.Split(samplechars, chunksize);

            foreach (var chunk in chunks)
            {
                var dom = new VA.DOM.Document();

                for (int j = 0; j < fontnames.Count(); j++)
                {
                    string fontname = fontnames[j];
                    double x0 = j*w;

                    var r = new VA.Drawing.Rectangle(x0, 0 - th, x0 + w, 0);
                    var n1 = dom.Drop("Rectangle", "basic_u.vss", r);
                    n1.Text = new VA.Text.Markup.TextElement( fontname.ToUpper() ) ;
                    n1.ShapeCells.FillForegnd = "rgb(255,255,255)";
                    n1.ShapeCells.LineWeight = 0.0;
                    n1.ShapeCells.LinePattern = 0;
                    n1.ShapeCells.CharSize = VA.Convert.PointsToInches(16);
                }


                for (int j = 0; j < fontnames.Count(); j++)
                {
                    for (int i = 0; i < chunksize; i++)
                    {
                        double x0 = j*w;
                        double y0 = i*h*-1 - th - h;

                        var r = new VA.Drawing.Rectangle(x0, y0, x0 + w, y0 + h);
                        var n1 = dom.Drop("Rectangle", "basic_u.vss", r);
                        if (i < chunk.Count)
                        {
                            n1.Text = new VA.Text.Markup.TextElement(chunk[i]);
                        }
                        else
                        {
                            // empty
                        }
                        n1.CharFontName = fontnames[j];
                        n1.ShapeCells.CharSize = VA.Convert.PointsToInches(36);
                        n1.ShapeCells.FillForegnd = "rgb(255,255,255)";
                        n1.ShapeCells.LineWeight = 0.0;
                        n1.ShapeCells.LinePattern = 0;
                    }
                }

                var page = doc.Pages.Add();
                dom.Render(page);
                page.ResizeToFitContents(0.5, 0.5);
            }
        }

        public static void FontGlyphComparision3(IVisio.Document doc, string[] fontnames, List<string> samplechars)
        {
            var colors = new[] {"rgb(0,0,255)", "rgb(255,0,0)"};

            double w = 2.0;
            double h = 1;
            double th = 1;

            int chunksize = 12;
            var chunks = LinqUtil.Split(samplechars, chunksize);


            foreach (var chunk in chunks)
            {
                var dom = new VA.DOM.Document();

                for (int j = 0; j < fontnames.Count(); j++)
                {
                    for (int i = 0; i < chunksize; i++)
                    {
                        double x0 = 0;
                        double y0 = i*h*-1;

                        var r = new VA.Drawing.Rectangle(x0, y0, x0 + w, y0 + h);
                        var n1 = dom.Drop("Rectangle", "basic_u.vss", r);
                        if (i < chunk.Count)
                        {
                            n1.Text = new VA.Text.Markup.TextElement(chunk[i]);
                        }
                        else
                        {
                            // empty
                        }
                        n1.CharFontName = fontnames[j];
                        n1.ShapeCells.CharColor = colors[j];
                        n1.ShapeCells.CharTransparency = 0.7;
                        n1.ShapeCells.CharSize = VA.Convert.PointsToInches(36);
                        n1.ShapeCells.FillPattern = 0;
                        n1.ShapeCells.LineWeight = 0.0;
                        n1.ShapeCells.LinePattern = 0;
                    }
                }

                var page = doc.Pages.Add();

                dom.Render(page);

                page.ResizeToFitContents(0.5, 0.5);
            }
        }
    }
}