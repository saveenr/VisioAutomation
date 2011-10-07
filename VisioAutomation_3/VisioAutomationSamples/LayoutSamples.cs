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

        public static void FontGlyphComparision()
        {
            var sampletext = "ABCDEFGHIJKLMNOPQRSTUVWXYZ" + "abcdefghijklmnopqrstuvwxyz" +
                             "<>[](),./|\\:;\'\"1234567890!@#$%^&*()`~";

            var samplechars = sampletext.Select(c => new string(new char[] {c})).ToList();

            var fontnames = new[] {"Segoe", "Calibri", "Impact"};

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

            var visapp = new IVisio.Application();
            var doc = visapp.Documents.Add("");
            var page = visapp.ActivePage;
            var docs = visapp.Documents;
            var stencil = docs.OpenStencil("basic_u.vss");
            var rectmaster = stencil.Masters["Rectangle"];
            
            var nodes = layout.Nodes.Where(n => n.Data.Render).ToList();
            var dom = new VA.DOM.Document();
            dom.ResolveVisioShapeObjects = true;

            var font_to_id = doc.Fonts.AsEnumerable().ToDictionary(f => f.Name, f => f.ID);
            var unique_fonts = new HashSet<string>();
            var unique_fontids = new HashSet<int>();
            
            foreach (var node in nodes)
            {
                var dom_shape = dom.Drop(rectmaster, node.Rectangle.Center);
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
                    unique_fonts.Add(node.Data.Font);
                    int font = font_to_id[node.Data.Font];
                    unique_fontids.Add(font);
                    cells.CharFont = font;
                }

                dom_shape.ShapeCells = cells;
                dom_shape.Text = node.Data.Text;
            }

            dom.Render(page);


            page.ResizeToFitContents(0.5, 0.5);
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
                
            }
            t.LayoutOptions.DefaultNodeSize = new VA.Drawing.Size(2.0, 0.25);
            t.LayoutOptions.UseDynamicConnectors = false;
            t.Render(page1);
        }
    }
}