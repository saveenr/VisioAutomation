using Microsoft.VisualStudio.TestTools.UnitTesting;
using VAS = VisioAutomation.Scripting;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace TestVisioAutomation
{
    [TestClass]
    public class ScriptingConnectTests : VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_Connects_Scenario_0()
        {
            var ss = GetScriptingSession();
            ss.Document.NewDocument();
            var pagesize = new VA.Drawing.Size(4, 4);
            ss.Page.NewPage(pagesize, false);

            var s1 = ss.Draw.DrawRectangle(1, 1, 1.25, 1.5);

            var s2 = ss.Draw.DrawRectangle(2, 3, 2.5, 3.5);

            var s3 = ss.Draw.DrawRectangle(4.5, 2.5, 6, 3.5);

            ss.Selection.SelectNone();
            ss.Selection.SelectShape(s1);
            ss.Selection.SelectShape(s2);
            ss.Selection.SelectShape(s3);

            ss.Document.OpenStencil("basic_u.vss");
            var master = ss.Master.GetMaster("Dynamic Connector", "basic_u.vss");
            var directed_connectors = ss.Connection.ConnectShapes(master);
            ss.Selection.SelectNone();
            ss.Selection.SelectShapes(directed_connectors);

            IVisio.VisGetSetArgs flags = 0;
            ss.ShapeSheet.SetFormula("EndArrow", "13", flags);

            var undirected_edges0 = ss.Connection.GetEdges();
            Assert.AreEqual(2, undirected_edges0.Count);

            var directed_edges0 = ss.Connection.GetDirectedEdges(VisioAutomation.Connections.ConnectorArrowEdgeHandling.ExcludeNoArrowEdges);
            Assert.AreEqual(2, directed_edges0.Count);

            var directed_edges1 =
                ss.Connection.GetDirectedEdges(VisioAutomation.Connections.ConnectorArrowEdgeHandling.TreatNoArrowEdgesAsBidirectional);
            Assert.AreEqual(2, directed_edges1.Count);

            ss.Document.CloseDocument(true);
        }

        [TestMethod]
        public void Scripting_Connects_Scenario_1()
        {
            var ss = GetScriptingSession();
            ss.Document.NewDocument();
            var pagesize = new VA.Drawing.Size(4, 4);
            ss.Page.NewPage(pagesize, false);

            var s1 = ss.Draw.DrawRectangle(1, 1, 1.25, 1.5);

            var s2 = ss.Draw.DrawRectangle(2, 3, 2.5, 3.5);

            var s3 = ss.Draw.DrawRectangle(4.5, 2.5, 6, 3.5);

            ss.Selection.SelectNone();
            ss.Selection.SelectShape(s1);
            ss.Selection.SelectShape(s2);
            ss.Selection.SelectShape(s3);

            ss.Document.OpenStencil("basic_u.vss");
            var master = ss.Master.GetMaster("Dynamic Connector", "basic_u.vss");
            var undirected_connectors = ss.Connection.ConnectShapes(master);

            var directed_edges0 = ss.Connection.GetDirectedEdges(VisioAutomation.Connections.ConnectorArrowEdgeHandling.ExcludeNoArrowEdges);
            Assert.AreEqual(0, directed_edges0.Count);

            var directed_edges1 =
                ss.Connection.GetDirectedEdges(VisioAutomation.Connections.ConnectorArrowEdgeHandling.TreatNoArrowEdgesAsBidirectional);
            Assert.AreEqual(4, directed_edges1.Count);

            var undirected_edges0 = ss.Connection.GetEdges();
            Assert.AreEqual(2, undirected_edges0.Count);

            ss.Document.CloseDocument(true);
        }
    }
}