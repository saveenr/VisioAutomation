using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.DocumentAnalysis;
using VA = VisioAutomation;

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
            var pagesize = new VA.Geometry.Size(4, 4);
            client.Page.NewPage(pagesize, false);

            var s1 = client.Draw.DrawRectangle(1, 1, 1.25, 1.5);
            var s2 = client.Draw.DrawRectangle(2, 3, 2.5, 3.5);
            var s3 = client.Draw.DrawRectangle(4.5, 2.5, 6, 3.5);

            client.Selection.SelectNone();
            client.Selection.SelectShapesById(s1);
            client.Selection.SelectShapesById(s2);
            client.Selection.SelectShapesById(s3);

            client.Document.OpenStencil("basic_u.vss");
            var connec_stencil = client.Document.OpenStencil("connec_u.vss");
            var master = client.Master.Get("Dynamic Connector", connec_stencil);
            var fromshapes = new [] { s1,s2};
            var toshapes = new [] { s2,s3};
            var directed_connectors = client.Connection.Connect(fromshapes,toshapes, master);
            client.Selection.SelectNone();
            client.Selection.SelectShapes(directed_connectors);


            var page = client.Page.GetActivePage();
            var writer = client.ShapeSheet.GetWriter(page);

            var shapes = client.Selection.GetShapesInSelection();
            foreach (var shape in shapes)
            {
                writer.SetFormula( shape.ID16, VA.ShapeSheet.SrcConstants.LineEndArrow, "13");
            }
            writer.Commit();

            var ch = new VA.DocumentAnalysis.ConnectorHandling();
            ch.DirectionSource = DirectionSource.UseConnectionOrder;
            var undirected_edges0 = client.Connection.GetDirectedEdges(ch);
            Assert.AreEqual(2, undirected_edges0.Count);

            var h0 = new VisioAutomation.DocumentAnalysis.ConnectorHandling();
            h0.NoArrowsHandling = VisioAutomation.DocumentAnalysis.NoArrowsHandling.ExcludeEdge;

            var h1 = new VisioAutomation.DocumentAnalysis.ConnectorHandling();
            h1.NoArrowsHandling = VisioAutomation.DocumentAnalysis.NoArrowsHandling.TreatEdgeAsBidirectional;

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
            var pagesize = new VA.Geometry.Size(4, 4);
            client.Page.NewPage(pagesize, false);

            var s1 = client.Draw.DrawRectangle(1, 1, 1.25, 1.5);

            var s2 = client.Draw.DrawRectangle(2, 3, 2.5, 3.5);

            var s3 = client.Draw.DrawRectangle(4.5, 2.5, 6, 3.5);

            client.Selection.SelectNone();
            client.Selection.SelectShapesById(s1);
            client.Selection.SelectShapesById(s2);
            client.Selection.SelectShapesById(s3);

            client.Document.OpenStencil("basic_u.vss");
            var connec_stencil = client.Document.OpenStencil("connec_u.vss");
            var master = client.Master.Get("Dynamic Connector", connec_stencil);
            var undirected_connectors = client.Connection.Connect(new [] { s1,s2},new [] { s2,s3}, master);

            var h1 = new VisioAutomation.DocumentAnalysis.ConnectorHandling();
            h1.NoArrowsHandling = VisioAutomation.DocumentAnalysis.NoArrowsHandling.ExcludeEdge;

            var directed_edges0 = client.Connection.GetDirectedEdges(h1);
            Assert.AreEqual(0, directed_edges0.Count);

            var h7 = new VA.DocumentAnalysis.ConnectorHandling();
            h7.NoArrowsHandling = NoArrowsHandling.TreatEdgeAsBidirectional;

            var directed_edges1 = client.Connection.GetDirectedEdges(h7);
            Assert.AreEqual(4, directed_edges1.Count);

            var h8 = new VA.DocumentAnalysis.ConnectorHandling();
            h8.DirectionSource = DirectionSource.UseConnectionOrder;

            var undirected_edges0 = client.Connection.GetDirectedEdges(h8);
            Assert.AreEqual(2, undirected_edges0.Count);

            client.Document.Close(true);
        }


        [TestMethod]
        public void Scripting_Connects_Scenario_3()
        {
            var client = this.GetScriptingClient();
            client.Document.New();
            var s1 = client.Draw.DrawRectangle(1, 1, 2,2);
            var s2 = client.Draw.DrawRectangle(4, 4, 5, 5);

            client.Connection.Connect(new[] {s1}, new[] {s2}, null);
            
            client.Document.Close(true);
        }
    }
}