using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using System.Linq;
using System.Collections.Generic;

namespace VisioAutomationSamples
{
    public static class LayoutSamples
    {
        public static void BoxHierarchy()
        {
            // Create a layout
            var layout = BoxHierarchyShared.CreateSampleBoxHierarchyLayout();

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
                BoxHierarchyShared.DrawBoxHierarchyDrawNode(node, node.Rectangle, page1);
            }

            var src_linepat = new VA.ShapeSheet.SRC(
                IVisio.VisSectionIndices.visSectionObject, IVisio.VisRowIndices.visRowLine,
                IVisio.VisCellIndices.visLinePattern);
            var root_shape = layout.Root.Data;
            var cell_linepat = root_shape.GetCell(src_linepat);
            cell_linepat.FormulaU = "7";

            // Make the page big enough to fit what was drawn + a small border
            var margin = new VA.Drawing.Size(0.5, 0.5);
            page1.ResizeToFitContents(margin);
        }

        public static void MSAGL()
        {
            var page1 = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            var vdom = new VisioAutomation.Layout.MSAGL.Drawing();

            // Create a Node 0
            var n0 = vdom.AddShape("n0", "N0 Untitled Node", "basflo_u.vss", "Decision");

            // Format Node 0
            n0.Size = new VA.Drawing.Size(3, 2);

            // Create Node 1

            var n1 = vdom.AddShape("n1", "N1", "basflo_u.vss", "Decision");

            // Format Node 1
            n1.ShapeCells = new VA.DOM.ShapeCells();
            n1.ShapeCells.FillForegnd = "rgb(255,0,0)";
            n1.ShapeCells.FillBkgnd = "rgb(255,255,0)";
            n1.ShapeCells.FillPattern = 40;

            // Create Node 2
            var n2 = vdom.AddShape("n2", "N2 MailServer", "server_u.vss", "Server");

            // Create Node 3

            var n3 = vdom.AddShape("n3", "N3", "basflo_u.vss", "Data");

            // Create Node 4
            var n4 = vdom.AddShape("n4", "N4", "basflo_u.vss", "Data");

            // Create the connectors to join the nodes
            // Note that Node 4 is deliberately not connected to any other node
            var c0 = vdom.Connect("c0", n0, n1, null, VA.Connections.ConnectorType.Curved);
            var c1 = vdom.Connect("c1", n1, n2, "YES", VA.Connections.ConnectorType.RightAngle);
            var c2 = vdom.Connect("c2", n3, n4, "NO", VA.Connections.ConnectorType.Curved);
            var c3 = vdom.Connect("c3", n0, n2, null, VA.Connections.ConnectorType.Straight);
            var c4 = vdom.Connect("c4", n2, n3, null, VA.Connections.ConnectorType.Curved);
            var c5 = vdom.Connect("c5", n3, n0, null, VA.Connections.ConnectorType.Curved);

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
            var renderer = new VisioAutomation.Layout.MSAGL.DirectedGraphLayout();
            renderer.LayoutOptions.UseDynamicConnectors = true;

            var options = new VisioAutomation.Layout.MSAGL.LayoutOptions();
            options.UseDynamicConnectors = false;

            vdom.Render(page1, options);
        }
    }
}