using VisioAutomation.DOM;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using System.Linq;
using System.Collections.Generic;
using BH = VisioAutomation.Layout.BoxLayout;

namespace VisioAutomationSamples
{
    public static class LayoutSamples
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

        private static BH.Node<NodeData> new_node()
        {
            return new_node(null);
        }

        public static BH.Node<NodeData> new_node(string s)
        {
            var box = new BH.Node<NodeData>();
            box.Data = new NodeData();
            box.Data.Text = s;
            return box;
        }

        public static BH.Node<NodeData> new_node(double w, double h, string s)
        {
            var box = new BH.Node<NodeData>();
            box.Width = w;
            box.Height = h;
            box.Data = new NodeData();
            box.Data.Text = s;
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
            var root_shape = layout.Root.Data;
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

            var fontnames = new[] {"Consolas", "Ubuntu Mono"};

            var sampletext = "ABCDEFGHIJKLMNOPQRSTUVWXYZ" + "abcdefghijklmnopqrstuvwxyz" +
                             "<>[](),./|\\:;\'\"1234567890!@#$%^&*()`~";
            var samplechars = sampletext.Select(c => new string(new char[] {c})).ToList();

            FontGlyphComparision(doc, fontnames, samplechars);
            FontGlyphComparision2(doc, fontnames, samplechars);
            FontGlyphComparision3(doc, fontnames, samplechars);
        }

        public static void FontGlyphComparision(IVisio.Document doc, string[] fontnames, List<string> samplechars)
        {
            var layout = new BH.BoxLayout<NodeData>();
            layout.LayoutOptions.DirectionVertical = VA.Layout.BoxLayout.DirectionVertical.TopToBottom;

            var root = layout.Root;
            root.Direction = BH.LayoutDirection.Vertical;
            root.ChildSeparation = 0.5;
            root.Data = new NodeData();
            root.Data.Render = false;

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
                var fontname_box = new_node(5, 0.5, fontname);
                fontname_box.Data.Cells = fontname_cells;
                root.AddNode(fontname_box);

                var font_box = new_node();
                font_box.Direction = BH.LayoutDirection.Vertical;
                font_box.ChildSeparation = 0.25;
                font_box.Data.Render = false;
                root.AddNode(font_box);

                int numcols = 17;
                int numrows = 5;
                int numcells = numcols*numrows;


                foreach (int row in Enumerable.Range(0, numrows))
                {
                    var row_box = new_node();
                    row_box.Direction = BH.LayoutDirection.Horizonal;
                    row_box.ChildSeparation = 0.25;
                    row_box.Data.Render = false;
                    font_box.AddNode(row_box);

                    foreach (int col in Enumerable.Range(0, numcols))
                    {
                        int charindex = (col + (numcols*row))%numcells;
                        string curchar = samplechars[charindex];

                        var cell_box = new_node(0.50, 0.50, curchar);
                        cell_box.Data.Font = fontname;
                        cell_box.Data.Cells = charbox_cells;
                        row_box.AddNode(cell_box);
                    }
                }
            }

            layout.PerformLayout();

            var page = doc.Pages.Add();

            var nodes = layout.Nodes.Where(n => n.Data.Render).ToList();
            var dom = new VA.DOM.Document();
            dom.ResolveVisioShapeObjects = true;

            foreach (var node in nodes)
            {
                var dom_shape = dom.Drop("Rectangle", "basic_u.vss", node.Rectangle.Center);
                var cells = node.Data.Cells;
                if (cells == null)
                {
                    cells = new VA.DOM.ShapeCells();
                }
                else
                {
                    cells = node.Data.Cells.ShallowCopy();
                }

                cells.Width = node.Rectangle.Width;
                cells.Height = node.Rectangle.Height;

                if (node.Data.Font != null)
                {
                    dom_shape.CharFontName = node.Data.Font;
                }

                dom_shape.ShapeCells = cells;
                dom_shape.Text = node.Data.Text;
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
                    var n1 = dom.Drop("Rectangle", "basic_u.vss", r.Center);
                    n1.ShapeCells.Width = r.Width;
                    n1.ShapeCells.Height = th;
                    n1.Text = fontname.ToUpper();
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
                        var n1 = dom.Drop("Rectangle", "basic_u.vss", r.Center);
                        n1.ShapeCells.Width = r.Width;
                        n1.ShapeCells.Height = r.Height;
                        if (i < chunk.Count)
                        {
                            n1.Text = chunk[i];
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
                        var n1 = dom.Drop("Rectangle", "basic_u.vss", r.Center);
                        n1.ShapeCells.Width = r.Width;
                        n1.ShapeCells.Height = r.Height;
                        if (i < chunk.Count)
                        {
                            n1.Text = chunk[i];
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

        public static void DirectedGraphViaMSAGL()
        {
            var page1 = SampleEnvironment.Application.ActiveDocument.Pages.Add();
            var directed_graph_drawing = new VA.Layout.DirectedGraph.Drawing();

            // Create a Node 0
            var n0 = directed_graph_drawing.AddShape("n0", "N0 Untitled Node", "basflo_u.vss", "Decision");

            // Format Node 0
            n0.Size = new VA.Drawing.Size(3, 2);

            // Create Node 1
            var n1 = directed_graph_drawing.AddShape("n1", "N1", "basflo_u.vss", "Decision");

            // Format Node 1
            n1.ShapeCells = new VA.DOM.ShapeCells();
            n1.ShapeCells.FillForegnd = "rgb(255,0,0)";
            n1.ShapeCells.FillBkgnd = "rgb(255,255,0)";
            n1.ShapeCells.FillPattern = 40;

            // Create Node 2
            var n2 = directed_graph_drawing.AddShape("n2", "N2 MailServer", "server_u.vss", "Server");

            // Create Node 3

            var n3 = directed_graph_drawing.AddShape("n3", "N3", "basflo_u.vss", "Data");

            // Create Node 4
            var n4 = directed_graph_drawing.AddShape("n4", "N4", "basflo_u.vss", "Data");

            // Create the connectors to join the nodes
            // Note that Node 4 is deliberately not connected to any other node
            var c0 = directed_graph_drawing.Connect("c0", n0, n1, null, VA.Connections.ConnectorType.Curved);
            var c1 = directed_graph_drawing.Connect("c1", n1, n2, "YES", VA.Connections.ConnectorType.RightAngle);
            var c2 = directed_graph_drawing.Connect("c2", n3, n4, "NO", VA.Connections.ConnectorType.Curved);
            var c3 = directed_graph_drawing.Connect("c3", n0, n2, null, VA.Connections.ConnectorType.Straight);
            var c4 = directed_graph_drawing.Connect("c4", n2, n3, null, VA.Connections.ConnectorType.Curved);
            var c5 = directed_graph_drawing.Connect("c5", n3, n0, null, VA.Connections.ConnectorType.Curved);

            // Format connector 0 to point "back" 
            c0.ShapeCells = new VA.DOM.ShapeCells();
            c0.ShapeCells.BeginArrow = 1;
            c0.ShapeCells.LineWeight = 0.10;

            // Format connector 1 to point "forward" 
            c1.ShapeCells = new VA.DOM.ShapeCells();
            c1.ShapeCells.EndArrow = 1;
            c1.ShapeCells.LineWeight = 0.10;

            // Format connector 2 to point "back" and "forward"  
            c2.ShapeCells = new VA.DOM.ShapeCells();
            c2.ShapeCells.EndArrow = 1;
            c2.ShapeCells.BeginArrow = 1;
            c2.ShapeCells.LineWeight = 0.10;

            // Perform the rendering
            var options = new VA.Layout.MSAGL.LayoutOptions();
            options.UseDynamicConnectors = false;

            VA.Layout.MSAGL.MSAGLRenderer.Render(page1, directed_graph_drawing, options);
        }

        public static void DirectedGraphViaVisio()
        {
            var page1 = SampleEnvironment.Application.ActiveDocument.Pages.Add();
            var directed_graph_drawing = new VA.Layout.DirectedGraph.Drawing();

            // Create a Node 0
            var n0 = directed_graph_drawing.AddShape("n0", "N0 Untitled Node", "basflo_u.vss", "Decision");

            // Format Node 0
            n0.Size = new VA.Drawing.Size(3, 2);

            // Create Node 1
            var n1 = directed_graph_drawing.AddShape("n1", "N1", "basflo_u.vss", "Decision");

            // Format Node 1
            n1.ShapeCells = new VA.DOM.ShapeCells();
            n1.ShapeCells.FillForegnd = "rgb(255,0,0)";
            n1.ShapeCells.FillBkgnd = "rgb(255,255,0)";
            n1.ShapeCells.FillPattern = 40;

            // Create Node 2
            var n2 = directed_graph_drawing.AddShape("n2", "N2 MailServer", "server_u.vss", "Server");

            // Create Node 3

            var n3 = directed_graph_drawing.AddShape("n3", "N3", "basflo_u.vss", "Data");

            // Create Node 4
            var n4 = directed_graph_drawing.AddShape("n4", "N4", "basflo_u.vss", "Data");

            // Create the connectors to join the nodes
            // Note that Node 4 is deliberately not connected to any other node
            var c0 = directed_graph_drawing.Connect("c0", n0, n1, null, VA.Connections.ConnectorType.Curved);
            var c1 = directed_graph_drawing.Connect("c1", n1, n2, "YES", VA.Connections.ConnectorType.RightAngle);
            var c2 = directed_graph_drawing.Connect("c2", n3, n4, "NO", VA.Connections.ConnectorType.Curved);
            var c3 = directed_graph_drawing.Connect("c3", n0, n2, null, VA.Connections.ConnectorType.Straight);
            var c4 = directed_graph_drawing.Connect("c4", n2, n3, null, VA.Connections.ConnectorType.Curved);
            var c5 = directed_graph_drawing.Connect("c5", n3, n0, null, VA.Connections.ConnectorType.Curved);

            // Format connector 0 to point "back" 
            c0.ShapeCells = new VA.DOM.ShapeCells();
            c0.ShapeCells.BeginArrow = 1;
            c0.ShapeCells.LineWeight = 0.10;

            // Format connector 1 to point "forward" 
            c1.ShapeCells = new VA.DOM.ShapeCells();
            c1.ShapeCells.EndArrow = 1;
            c1.ShapeCells.LineWeight = 0.10;

            // Format connector 2 to point "back" and "forward"  
            c2.ShapeCells = new VA.DOM.ShapeCells();
            c2.ShapeCells.EndArrow = 1;
            c2.ShapeCells.BeginArrow = 1;
            c2.ShapeCells.LineWeight = 0.10;

            // Perform the rendering
            var options = new VA.Layout.MSAGL.LayoutOptions();
            options.UseDynamicConnectors = false;

            directed_graph_drawing.Render(page1);

            var layout_config = new VA.ShapeLayout.HierarchyLayout();
            layout_config.Direction = VA.ShapeLayout.Direction.BottomToTop;
            layout_config.HorizontalAlignment = VA.ShapeLayout.HorizontalAlignment.Center;
            layout_config.ResizePageToFit = true;
            layout_config.AvenueSize = new VA.Drawing.Size(1, 1);
            layout_config.ConnectorAppearance = VA.ShapeLayout.ConnectorAppearance.Curved;
            layout_config.Apply(page1);
        }


        public static void TreeWithTwoPassLayoutAndFormatting()
        {
            var doc = SampleEnvironment.Application.ActiveDocument;
            var page1 = doc.Pages.Add();

            var t = new VA.Layout.Tree.Drawing();

            t.Root = new VA.Layout.Tree.Node("Root");

            var na = new VA.Layout.Tree.Node("A");
            var nb = new VA.Layout.Tree.Node("B");

            var na1 = new VA.Layout.Tree.Node("A1");
            var na2 = new VA.Layout.Tree.Node("A2");

            var nb1 = new VA.Layout.Tree.Node("B1");
            var nb2 = new VA.Layout.Tree.Node("B2");

            t.Root.Children.Add(na);
            t.Root.Children.Add(nb);

            na.Children.Add(na1);
            na.Children.Add(na2);

            nb.Children.Add(nb1);
            nb1.Children.Add(nb2);

            var fontname = "Segoe UI";
            var font = doc.Fonts[fontname];

            foreach (var tn in t.Nodes)
            {
                var cells = new ShapeCells();
                tn.ShapeCells = cells;

                cells.HAlign = 0; // align text to left
                cells.VerticalAlign = 0; // align text block to top
                cells.CharFont = font.ID;
                cells.CharSize = "10pt";
                cells.FillForegnd = "rgb(255,250,200)";
                cells.CharColor = "rgb(255,0,0)";
            }
        }
    }

    internal static class LinqUtil
    {
        public static List<List<T>> Split<T>(List<T> source, int chunksize)
        {
            return source
                .Select((x, i) => new {Index = i, Value = x})
                .GroupBy(x => x.Index/chunksize)
                .Select(x => x.Select(v => v.Value).ToList())
                .ToList();
        }
    }
}