using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace TestVisioAutomation
{
    [TestClass]
    public class ScriptingConnectTests : VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_Connects_Scenarios()
        {
            this.Scripting_Connects_Scenario_0();
            this.Scripting_Connects_Scenario_1();
        }

        public void Scripting_Connects_Scenario_0()
        {
            var ss = GetScriptingSession();
            ss.Document.New();
            var pagesize = new VA.Drawing.Size(4, 4);
            ss.Page.New(pagesize, false);

            var s1 = ss.Draw.Rectangle(1, 1, 1.25, 1.5);

            var s2 = ss.Draw.Rectangle(2, 3, 2.5, 3.5);

            var s3 = ss.Draw.Rectangle(4.5, 2.5, 6, 3.5);

            ss.Selection.SelectNone();
            ss.Selection.Select(s1);
            ss.Selection.Select(s2);
            ss.Selection.Select(s3);

            ss.Document.OpenStencil("basic_u.vss");
            var connec_stencil = ss.Document.OpenStencil("connec_u.vss");
            var master = ss.Master.Get("Dynamic Connector", connec_stencil);
            var directed_connectors = ss.Connection.Connect(master,new [] { s1,s2},new [] { s2,s3});
            ss.Selection.SelectNone();
            ss.Selection.Select(directed_connectors);

            IVisio.VisGetSetArgs flags = 0;
            ss.ShapeSheet.SetFormula(null,new[] { VA.ShapeSheet.SRCConstants.EndArrow }, new [] {"13"}, flags);

            var undirected_edges0 = ss.Connection.GetEdges();
            Assert.AreEqual(2, undirected_edges0.Count);

            var directed_edges0 = ss.Connection.GetDirectedEdges(VisioAutomation.Connections.ConnectorArrowEdgeHandling.ExcludeNoArrowEdges);
            Assert.AreEqual(2, directed_edges0.Count);

            var directed_edges1 =
                ss.Connection.GetDirectedEdges(VisioAutomation.Connections.ConnectorArrowEdgeHandling.TreatNoArrowEdgesAsBidirectional);
            Assert.AreEqual(2, directed_edges1.Count);

            ss.Document.Close(true);
        }

        public void Scripting_Connects_Scenario_1()
        {
            var ss = GetScriptingSession();
            ss.Document.New();
            var pagesize = new VA.Drawing.Size(4, 4);
            ss.Page.New(pagesize, false);

            var s1 = ss.Draw.Rectangle(1, 1, 1.25, 1.5);

            var s2 = ss.Draw.Rectangle(2, 3, 2.5, 3.5);

            var s3 = ss.Draw.Rectangle(4.5, 2.5, 6, 3.5);

            ss.Selection.SelectNone();
            ss.Selection.Select(s1);
            ss.Selection.Select(s2);
            ss.Selection.Select(s3);

            ss.Document.OpenStencil("basic_u.vss");
            var connec_stencil = ss.Document.OpenStencil("connec_u.vss");
            var master = ss.Master.Get("Dynamic Connector", connec_stencil);
            var undirected_connectors = ss.Connection.Connect(master, new [] { s1,s2},new [] { s2,s3});

            var directed_edges0 = ss.Connection.GetDirectedEdges(VisioAutomation.Connections.ConnectorArrowEdgeHandling.ExcludeNoArrowEdges);
            Assert.AreEqual(0, directed_edges0.Count);

            var directed_edges1 =
                ss.Connection.GetDirectedEdges(VisioAutomation.Connections.ConnectorArrowEdgeHandling.TreatNoArrowEdgesAsBidirectional);
            Assert.AreEqual(4, directed_edges1.Count);

            var undirected_edges0 = ss.Connection.GetEdges();
            Assert.AreEqual(2, undirected_edges0.Count);

            ss.Document.Close(true);
        }
    }
}