using VisioAutomation.Shapes.Connections;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomationSamples
{
    public static class PathAnalysisSamples
    {
        public static void PathAnalysis()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();
            page.DrawRectangle(0, 0, 1, 1);

            var s0 = page.DrawRectangle(0, 0, 1, 1);
            var s1 = page.DrawRectangle(3, 0, 4, 1);
            var s2 = page.DrawRectangle(1.5, 1.5, 2.5, 2.5);
            var s3 = page.DrawRectangle(0, 3, 1, 4);
            var s4 = page.DrawRectangle(3, 3, 4, 4);
            var s5 = page.DrawRectangle(3, 6, 4, 7);
            var s6 = page.DrawRectangle(7, 8, 8, 9);

            s0.Text = "s0";
            s1.Text = "s1";
            s2.Text = "s2";
            s3.Text = "s3";
            s4.Text = "s4";
            s5.Text = "s5";
            s6.Text = "s6";

            var stencil = page.Application.Documents.OpenStencil("basic_u.vss");
            var connector = stencil.Masters["Dynamic Connector"];

            // connect shapes - but leave s0 alone
            var dir = IVisio.VisAutoConnectDir.visAutoConnectDirNone;
            s1.AutoConnect(s2, dir, null);
            s2.AutoConnect(s3, dir, null);
            s3.AutoConnect(s2, dir, null);
            s3.AutoConnect(s4, dir, null);
            s5.AutoConnect(s6, dir, null);

            var normal_edges = VisioAutomation.Shapes.Connections.PathAnalysis.GetDirectedEdges(page, VA.Shapes.Connections.DirectedEdgeHandling.Raw);
            var edge_handling_0 = DirectedEdgeHandling.EdgesWithoutArrowsAreExcluded;
            var edge_handling_1 = DirectedEdgeHandling.EdgesWithoutArrowsAreBidirectional;
            var tc_edges_0 = VisioAutomation.Shapes.Connections.PathAnalysis.GetTransitiveClosure(page, edge_handling_0);
            var tc_edges_1 = VisioAutomation.Shapes.Connections.PathAnalysis.GetTransitiveClosure(page, edge_handling_1);

            var legend0 = page.DrawRectangle(5, 0, 6.5, 6);
            var sb0 = new System.Text.StringBuilder();
            sb0.AppendLine("Connections");
            
            foreach (var e in normal_edges)
            {
                string s = string.Format("{0} - {1}", e.From.Text, e.To.Text);
                sb0.AppendLine(s);
            }

            legend0.Text = sb0.ToString();

            var legend1 = page.DrawRectangle(6.5, 0, 8.5, 6);
            var sb1 = new System.Text.StringBuilder();
            sb1.AppendLine("Transitive closure (treat edges as bidirectional)");
            
            foreach (var e in tc_edges_1)
            {
                string s = string.Format("{0} -> {1}", e.From.Text, e.To.Text);
                sb1.AppendLine(s);
            }

            legend1.Text = sb1.ToString();
        }
    }
}