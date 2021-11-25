﻿using VA = VisioAutomation;
using VisioAutomation.Extensions;
using VisioAutomation.Models.Dom;
using VisioAutomation.Models.Layouts.DirectedGraph;

namespace VisioAutomationSamples
{
    public static class DirectedGraphLayoutSamples
    {
        public static void DirectedGraphViaMsagl()
        {
            var page1 = SampleEnvironment.Application.ActiveDocument.Pages.Add();
            var directed_graph_drawing = DirectedGraphLayoutSamples.get_dg_drawing();

            var renderer = new MsaglRenderer();
            renderer.LayoutOptions.UseDynamicConnectors = false;
            renderer.Render(page1, directed_graph_drawing);
        }

        public static void DirectedGraphViaVisio()
        {
            var page1 = SampleEnvironment.Application.ActiveDocument.Pages.Add();
            var directed_graph_drawing = DirectedGraphLayoutSamples.get_dg_drawing();

            var vr = new VisioLayoutRenderer();
            vr.Render(page1, directed_graph_drawing);

            var padding = new VA.Geometry.Size(0.5, 0.5);
            page1.ResizeToFitContents(padding);
        }

        private static DirectedGraphLayout get_dg_drawing()
        {

            var ver = VA.Application.ApplicationHelper.GetVersion(SampleEnvironment.Application);

            string server_stencil = (ver.Major >= 15) ? "server_u.vssx" : "server_u.vss";
            string basflo_stencil = (ver.Major >= 15) ? "basflo_u.vssx" : "basflo_u.vss";

            var directed_graph_drawing = new DirectedGraphLayout();

            // Create a Node 0
            var n0 = directed_graph_drawing.AddNode("n0", "N0 Untitled Node", basflo_stencil, "Decision");

            // Format Node 0
            n0.Size = new VA.Geometry.Size(3, 2);

            // Create Node 1
            var n1 = directed_graph_drawing.AddNode("n1", "N1", basflo_stencil, "Decision");

            // Format Node 1
            n1.Cells = new ShapeCells();
            n1.Cells.FillForeground = "rgb(255,0,0)";
            n1.Cells.FillBackground = "rgb(255,255,0)";
            n1.Cells.FillPattern = 40;

            // Create Node 2
            var n2 = directed_graph_drawing.AddNode("n2", "N2 MailServer", server_stencil, "Server");

            // Create Node 3

            var n3 = directed_graph_drawing.AddNode("n3", "N3", basflo_stencil, "Data");

            // Create Node 4
            var n4 = directed_graph_drawing.AddNode("n4", "N4", basflo_stencil, "Data");

            // Create the connectors to join the nodes
            // Note that Node 4 is deliberately not connected to any other node

            var curved = VisioAutomation.Models.ConnectorType.Curved;
            var rightangle = VisioAutomation.Models.ConnectorType.RightAngle;

            var c0 = directed_graph_drawing.AddEdge("c0", n0, n1, null, curved);
            var c1 = directed_graph_drawing.AddEdge("c1", n1, n2, "YES", rightangle);
            var c2 = directed_graph_drawing.AddEdge("c2", n3, n4, "NO", curved);
            var c3 = directed_graph_drawing.AddEdge("c3", n0, n2, null, rightangle);
            var c4 = directed_graph_drawing.AddEdge("c4", n2, n3, null, curved);
            var c5 = directed_graph_drawing.AddEdge("c5", n3, n0, null, curved);

            // Format connector 0 to point "back" 
            c0.Cells = new ShapeCells();
            c0.Cells.LineBeginArrow = 1;
            c0.Cells.LineWeight = 0.10;

            // Format connector 1 to point "forward" 
            c1.Cells = new ShapeCells();
            c1.Cells.LineEndArrow = 1;
            c1.Cells.LineWeight = 0.10;

            // Format connector 2 to point "back" and "forward"  
            c2.Cells = new ShapeCells();
            c2.Cells.LineEndArrow = 1;
            c2.Cells.LineBeginArrow = 1;
            c2.Cells.LineWeight = 0.10;
            return directed_graph_drawing;
        }
    }
}