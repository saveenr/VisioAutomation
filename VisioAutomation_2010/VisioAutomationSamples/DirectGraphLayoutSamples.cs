using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using DGMODEL = VisioAutomation.Models.DirectedGraph;
using VisioAutomation.Extensions;

namespace VisioAutomationSamples
{
    public static class DirectGraphLayoutSamples
    {
        public static void DirectedGraphViaMSAGL()
        {
            var page1 = SampleEnvironment.Application.ActiveDocument.Pages.Add();
            var directed_graph_drawing = get_dg_drawing();
            var options = new DGMODEL.MSAGLLayoutOptions();
            options.UseDynamicConnectors = false;
            directed_graph_drawing.Render(page1, options);
        }

        public static void DirectedGraphViaVisio()
        {
            var page1 = SampleEnvironment.Application.ActiveDocument.Pages.Add();
            var directed_graph_drawing = get_dg_drawing();

            var visio_options = new DGMODEL.VisioLayoutOptions();
            directed_graph_drawing.Render(page1, visio_options);

            page1.ResizeToFitContents(new VA.Drawing.Size(0.5, 0.5));
        }

        private static DGMODEL.Drawing get_dg_drawing()
        {

            var appinfo = VA.Application.ApplicationHelper.GetInformation(SampleEnvironment.Application);
            var ver = appinfo.Version;
            string server_stencil = (ver.Major >= 15) ? "server_u.vssx" : "server_u.vss";
            string basflo_stencil = (ver.Major >= 15) ? "basflo_u.vssx" : "basflo_u.vss";

            var directed_graph_drawing = new DGMODEL.Drawing();

            // Create a Node 0
            var n0 = directed_graph_drawing.AddShape("n0", "N0 Untitled Node", basflo_stencil, "Decision");

            // Format Node 0
            n0.Size = new VA.Drawing.Size(3, 2);

            // Create Node 1
            var n1 = directed_graph_drawing.AddShape("n1", "N1", basflo_stencil, "Decision");

            // Format Node 1
            n1.Cells = new VA.DOM.ShapeCells();
            n1.Cells.FillForegnd = "rgb(255,0,0)";
            n1.Cells.FillBkgnd = "rgb(255,255,0)";
            n1.Cells.FillPattern = 40;

            // Create Node 2
            var n2 = directed_graph_drawing.AddShape("n2", "N2 MailServer", server_stencil, "Server");

            // Create Node 3

            var n3 = directed_graph_drawing.AddShape("n3", "N3", basflo_stencil, "Data");

            // Create Node 4
            var n4 = directed_graph_drawing.AddShape("n4", "N4", basflo_stencil, "Data");

            // Create the connectors to join the nodes
            // Note that Node 4 is deliberately not connected to any other node

            var curved = VA.Shapes.Connections.ConnectorType.Curved;
            var rightangle = VA.Shapes.Connections.ConnectorType.RightAngle;

            var c0 = directed_graph_drawing.AddConnection("c0", n0, n1, null, curved);
            var c1 = directed_graph_drawing.AddConnection("c1", n1, n2, "YES", rightangle);
            var c2 = directed_graph_drawing.AddConnection("c2", n3, n4, "NO", curved);
            var c3 = directed_graph_drawing.AddConnection("c3", n0, n2, null, rightangle);
            var c4 = directed_graph_drawing.AddConnection("c4", n2, n3, null, curved);
            var c5 = directed_graph_drawing.AddConnection("c5", n3, n0, null, curved);

            // Format connector 0 to point "back" 
            c0.Cells = new VA.DOM.ShapeCells();
            c0.Cells.BeginArrow = 1;
            c0.Cells.LineWeight = 0.10;

            // Format connector 1 to point "forward" 
            c1.Cells = new VA.DOM.ShapeCells();
            c1.Cells.EndArrow = 1;
            c1.Cells.LineWeight = 0.10;

            // Format connector 2 to point "back" and "forward"  
            c2.Cells = new VA.DOM.ShapeCells();
            c2.Cells.EndArrow = 1;
            c2.Cells.BeginArrow = 1;
            c2.Cells.LineWeight = 0.10;
            return directed_graph_drawing;
        }
    }
}