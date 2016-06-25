using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using System.Linq;
using System.Collections.Generic;
using BoxL = VisioAutomation.Models.BoxLayout;

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

        public static BoxL.Box AddNodeEx(this BoxL.Container p, double w, double h, string s)
        {
            var box = p.AddBox(w, h);
            var node_data = new NodeData();
            node_data.Text = s;
            box.Data = node_data;
            return box;
        }

        public static void FontCompare()
        {
            var visapp = new IVisio.Application();
            var doc = visapp.Documents.Add("");

            var fontnames = new[] {"Arial", "Calibri"};

            var sampletext = "ABCDEFGHIJKLMNOPQRSTUVWXYZ" + "abcdefghijklmnopqrstuvwxyz" +
                             "<>[](),./|\\:;\'\"1234567890!@#$%^&*()`~";
            var samplechars = sampletext.Select(c => new string(new[] {c})).ToList();

            BoxLayoutSamples.FontGlyphComparision(doc, fontnames, samplechars);
            BoxLayoutSamples.FontGlyphComparision2(doc, fontnames, samplechars);
            BoxLayoutSamples.FontGlyphComparision3(doc, fontnames, samplechars);
        }

        public static void FontGlyphComparision(IVisio.Document doc, string[] fontnames, List<string> samplechars)
        {
            var layout = new BoxL.BoxLayout();

            var root = new BoxL.Container( BoxL.Direction.TopToBottom);

            layout.Root = root;
            root.ChildSpacing = 0.5;

            var nodedata = new NodeData();
            nodedata.Render = false;
            root.Data = nodedata;

            var fontname_cells = new VA.DOM.ShapeCells();
            fontname_cells.FillPattern = 0;
            fontname_cells.LinePattern = 0;
            fontname_cells.LineWeight = 0.0;
            fontname_cells.ParaHorizontalAlign = 0;
            fontname_cells.CharSize = "36pt";

            var charbox_cells = new VA.DOM.ShapeCells();
            charbox_cells.FillPattern = 0;
            charbox_cells.LinePattern = 1;
            charbox_cells.LineWeight = 0.0;
            charbox_cells.LineColor = "rgb(150,150,150)";
            charbox_cells.ParaHorizontalAlign = 1;
            charbox_cells.CharSize = "24pt";

            foreach (string fontname in fontnames)
            {
                var fontname_box = root.AddNodeEx(5, 0.5, fontname);
                var fontname_box_data = (NodeData) fontname_box.Data;
                fontname_box_data.Cells = fontname_cells;

                var font_box = root.AddContainer(BoxL.Direction.TopToBottom);
                font_box.ChildSpacing = 0.25;
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
                    var row_box = font_box.AddContainer(BoxL.Direction.LeftToRight);
                    row_box.ChildSpacing = 0.25;
                    var row_box_data = new NodeData();
                    row_box_data.Render = false;

                    row_box.Data = row_box_data;

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

            var domshapescol = new VA.DOM.ShapeList();

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

                var shape_node = domshapescol.Drop("Rectangle", "basic_u.vss", node.Rectangle);
       
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
                    shape_node.CharFontName = node_data.Font;
                }

                shape_node.Cells = cells;
                shape_node.Text = new VA.Text.Markup.TextElement( node_data.Text );
            }

            domshapescol.Render(page);

            var bordersize = new VA.Drawing.Size(0.5, 0.5);
            page.ResizeToFitContents(bordersize);
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
                var domshapescol = new VA.DOM.ShapeList();

                for (int j = 0; j < fontnames.Count(); j++)
                {
                    string fontname = fontnames[j];
                    double x0 = j*w;

                    var r = new VA.Drawing.Rectangle(x0, 0 - th, x0 + w, 0);
                    var n1 = domshapescol.Drop("Rectangle", "basic_u.vss", r);
                    n1.Text = new VA.Text.Markup.TextElement( fontname.ToUpper() ) ;
                    n1.Cells.FillForegnd = "rgb(255,255,255)";
                    n1.Cells.LineWeight = 0.0;
                    n1.Cells.LinePattern = 0;
                    n1.Cells.CharSize = "16pt";
                }


                for (int j = 0; j < fontnames.Count(); j++)
                {
                    for (int i = 0; i < chunksize; i++)
                    {
                        double x0 = j*w;
                        double y0 = i*h*-1 - th - h;

                        var r = new VA.Drawing.Rectangle(x0, y0, x0 + w, y0 + h);
                        var n1 = domshapescol.Drop("Rectangle", "basic_u.vss", r);
                        if (i < chunk.Count)
                        {
                            n1.Text = new VA.Text.Markup.TextElement(chunk[i]);
                        }
                        n1.CharFontName = fontnames[j];
                        n1.Cells.CharSize = "36pt";
                        n1.Cells.FillForegnd = "rgb(255,255,255)";
                        n1.Cells.LineWeight = 0.0;
                        n1.Cells.LinePattern = 0;
                    }
                }

                var page = doc.Pages.Add();
                domshapescol.Render(page);

                var bordersize = new VA.Drawing.Size(0.5, 0.5);
                page.ResizeToFitContents(bordersize);
            }
        }

        public static void FontGlyphComparision3(IVisio.Document doc, string[] fontnames, List<string> samplechars)
        {
            var colors = new[] {"rgb(0,0,255)", "rgb(255,0,0)"};

            double w = 2.0;
            double h = 1;

            int chunksize = 12;
            var chunks = LinqUtil.Split(samplechars, chunksize);


            foreach (var chunk in chunks)
            {
                var domshapescol = new VA.DOM.ShapeList();

                for (int j = 0; j < fontnames.Count(); j++)
                {
                    for (int i = 0; i < chunksize; i++)
                    {
                        double x0 = 0;
                        double y0 = i*h*-1;

                        var r = new VA.Drawing.Rectangle(x0, y0, x0 + w, y0 + h);
                        var n1 = domshapescol.Drop("Rectangle", "basic_u.vss", r);
                        if (i < chunk.Count)
                        {
                            n1.Text = new VA.Text.Markup.TextElement(chunk[i]);
                            n1.Text.CharacterCells.Color = colors[j];

                        }
                        n1.CharFontName = fontnames[j];
    
                        //n1.Cells.CharColor = "=RGB(255,0,0)";// colors[j];
                        n1.Cells.CharTransparency = 0.7;
                        n1.Cells.CharSize = "36pt";
                        n1.Cells.FillPattern = 0;
                        n1.Cells.LineWeight = 0.0;
                        n1.Cells.LinePattern = 0;
                    }
                }

                var page = doc.Pages.Add();

                domshapescol.Render(page);

                var bordersize = new VA.Drawing.Size(0.5, 0.5);
                page.ResizeToFitContents(bordersize);
            }
        }
    }
}