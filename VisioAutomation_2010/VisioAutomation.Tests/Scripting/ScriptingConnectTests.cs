using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.DocumentAnalysis;
using VACONNECT = VisioAutomation.Shapes.Connections;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation_Tests.Scripting
{
    [TestClass]
    public class ScriptingConnectTests : VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_Connects_Scenario_0()
        {
            var client = this.GetScriptingClient();
            client.Document.New();
            var pagesize = new VA.Drawing.Size(4, 4);
            client.Page.New(pagesize, false);

            var s1 = client.Draw.Rectangle(1, 1, 1.25, 1.5);

            var s2 = client.Draw.Rectangle(2, 3, 2.5, 3.5);

            var s3 = client.Draw.Rectangle(4.5, 2.5, 6, 3.5);

            client.Selection.None();
            client.Selection.Select(s1);
            client.Selection.Select(s2);
            client.Selection.Select(s3);

            client.Document.OpenStencil("basic_u.vss");
            var connec_stencil = client.Document.OpenStencil("connec_u.vss");
            var master = client.Master.Get("Dynamic Connector", connec_stencil);
            var fromshapes = new [] { s1,s2};
            var toshapes = new [] { s2,s3};
            var directed_connectors = client.Connection.Connect(fromshapes,toshapes, master);
            client.Selection.None();
            client.Selection.Select(directed_connectors);

            IVisio.VisGetSetArgs flags = 0;
            client.ShapeSheet.SetFormula(null,new[] { VA.ShapeSheet.SRCConstants.EndArrow }, new [] {"13"}, flags);

            var h7 = new VA.DocumentAnalysis.ConnectorEdgeHandling();
            h7.AR = ArrowHandling.Raw;
            var undirected_edges0 = client.Connection.GetDirectedEdges(h7);
            Assert.AreEqual(2, undirected_edges0.Count);

            var h0 = new VisioAutomation.DocumentAnalysis.ConnectorEdgeHandling();
            h0.ADR = VisioAutomation.DocumentAnalysis.ArrowDirectionHandling.NoArrows_Exclude;

            var h1 = new VisioAutomation.DocumentAnalysis.ConnectorEdgeHandling();
            h1.ADR = VisioAutomation.DocumentAnalysis.ArrowDirectionHandling.NoArrows_Bidirectional;

            var directed_edges0 = client.Connection.GetDirectedEdges(h0);
            Assert.AreEqual(2, directed_edges0.Count);

            var directed_edges1 = client.Connection.GetDirectedEdges(h1);
            Assert.AreEqual(2, directed_edges1.Count);

            client.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_Connects_Scenario_1()
        {
            var client = this.GetScriptingClient();
            client.Document.New();
            var pagesize = new VA.Drawing.Size(4, 4);
            client.Page.New(pagesize, false);

            var s1 = client.Draw.Rectangle(1, 1, 1.25, 1.5);

            var s2 = client.Draw.Rectangle(2, 3, 2.5, 3.5);

            var s3 = client.Draw.Rectangle(4.5, 2.5, 6, 3.5);

            client.Selection.None();
            client.Selection.Select(s1);
            client.Selection.Select(s2);
            client.Selection.Select(s3);

            client.Document.OpenStencil("basic_u.vss");
            var connec_stencil = client.Document.OpenStencil("connec_u.vss");
            var master = client.Master.Get("Dynamic Connector", connec_stencil);
            var undirected_connectors = client.Connection.Connect(new [] { s1,s2},new [] { s2,s3}, master);

            var h1 = new VisioAutomation.DocumentAnalysis.ConnectorEdgeHandling();
            h1.ADR = VisioAutomation.DocumentAnalysis.ArrowDirectionHandling.NoArrows_Exclude;

            var directed_edges0 = client.Connection.GetDirectedEdges(h1);
            Assert.AreEqual(0, directed_edges0.Count);

            var h7 = new VA.DocumentAnalysis.ConnectorEdgeHandling();
            h7.ADR = ArrowDirectionHandling.NoArrows_Bidirectional;

            var directed_edges1 = client.Connection.GetDirectedEdges(h7);
            Assert.AreEqual(4, directed_edges1.Count);

            var h8 = new VA.DocumentAnalysis.ConnectorEdgeHandling();
            h8.AR = ArrowHandling.Raw;

            var undirected_edges0 = client.Connection.GetDirectedEdges(h8);
            Assert.AreEqual(2, undirected_edges0.Count);

            client.Document.Close(true);
        }


        [TestMethod]
        public void Scripting_Connects_Scenario_3()
        {
            var client = this.GetScriptingClient();
            client.Document.New();
            var s1 = client.Draw.Rectangle(1, 1, 2,2);
            var s2 = client.Draw.Rectangle(4, 4, 5, 5);

            client.Connection.Connect(new[] {s1}, new[] {s2}, null);
            
            client.Document.Close(true);
        }
    }
}