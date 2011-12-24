using VisioAutomation.DOM;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using System.Linq;
using System.Collections.Generic;
using BH = VisioAutomation.Layout.BoxLayout;

namespace VisioAutomationSamples
{
    public static class DirectGraphLayoutSamples
    {
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
}