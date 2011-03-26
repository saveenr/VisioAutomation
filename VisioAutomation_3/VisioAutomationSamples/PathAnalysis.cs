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
            var s2 = page.DrawRectangle(1, 1, 3, 3);
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
            s1.AutoConnect(s2, IVisio.VisAutoConnectDir.visAutoConnectDirNone, null);
            s2.AutoConnect(s3, IVisio.VisAutoConnectDir.visAutoConnectDirNone, null);
            s3.AutoConnect(s2, IVisio.VisAutoConnectDir.visAutoConnectDirNone, null);
            s3.AutoConnect(s4, IVisio.VisAutoConnectDir.visAutoConnectDirNone, null);
            s5.AutoConnect(s6, IVisio.VisAutoConnectDir.visAutoConnectDirNone, null);

            var normal_edges = VA.Connections.PathAnalysis.GetEdges(page);
            var tc_edges_0 = VA.Connections.PathAnalysis.GetTransitiveClosure(page,
                                                                  VisioAutomation.Connections.ConnectorArrowEdgeHandling.ExcludeNoArrowEdges);
            var tc_edges_1 = VA.Connections.PathAnalysis.GetTransitiveClosure(page,
                                                                  VisioAutomation.Connections.ConnectorArrowEdgeHandling.TreatNoArrowEdgesAsBidirectional);
        }
    }
}