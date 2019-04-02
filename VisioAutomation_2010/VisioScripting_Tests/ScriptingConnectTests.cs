using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.DocumentAnalysis;
using VisioScripting.Models;
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
            client.Document.NewDocument();
            var pagesize = new VA.Geometry.Size(4, 4);
            client.Page.NewPage(pagesize, false);

            var s1 = client.Draw.DrawRectangle(1, 1, 1.25, 1.5);
            var s2 = client.Draw.DrawRectangle(2, 3, 2.5, 3.5);
            var s3 = client.Draw.DrawRectangle(4.5, 2.5, 6, 3.5);

            client.Selection.SelectNone();
            client.Selection.SelectShapesById(s1);
            client.Selection.SelectShapesById(s2);
            client.Selection.SelectShapesById(s3);

            client.Document.OpenStencilDocument("basic_u.vss");
            var connec_stencil = client.Document.OpenStencilDocument("connec_u.vss");

            var tdoc = new VisioScripting.TargetDocument(connec_stencil);
            var master = client.Master.GetMasterWithNameInDocument(tdoc, "Dynamic Connector");
            var fromshapes = new [] { s1,s2};
            var toshapes = new [] { s2,s3};
            var directed_connectors = client.Connection.ConnectShapes(fromshapes,toshapes, master);
            client.Selection.SelectNone();
            client.Selection.SelectShapes(directed_connectors);

            var page = new VisioScripting.TargetPage();
            var writer = client.ShapeSheet.GetWriterForPage(page);

            var shapes = client.Selection.GetShapesInSelection();
            foreach (var shape in shapes)
            {
                writer.SetFormula( shape.ID16, VA.ShapeSheet.SrcConstants.LineEndArrow, "13");
            }
            writer.Commit();

            var options0 = new VA.DocumentAnalysis.ConnectionAnalyzerOptions();
            options0.DirectionSource = DirectionSource.UseConnectionOrder;
            var undirected_edges0 = client.Connection.GetDirectedEdgesOnPage( new VisioScripting.TargetPage(), options0);
            Assert.AreEqual(2, undirected_edges0.Count);

            var options1 = new VA.DocumentAnalysis.ConnectionAnalyzerOptions();
            options1.NoArrowsHandling = NoArrowsHandling.ExcludeEdge;

            var options2 = new VA.DocumentAnalysis.ConnectionAnalyzerOptions();
            options2.NoArrowsHandling = NoArrowsHandling.TreatEdgeAsBidirectional;

            var directed_edges0 = client.Connection.GetDirectedEdgesOnPage(new VisioScripting.TargetPage(), options1);
            Assert.AreEqual(2, directed_edges0.Count);

            var directed_edges1 = client.Connection.GetDirectedEdgesOnPage(new VisioScripting.TargetPage(), options2);
            Assert.AreEqual(2, directed_edges1.Count);

            client.Document.CloseActiveDocument(true);
        }

        [TestMethod]
        public void Scripting_Connects_Scenario_1()
        {
            var client = this.GetScriptingClient();
            client.Document.NewDocument();
            var pagesize = new VA.Geometry.Size(4, 4);
            client.Page.NewPage(pagesize, false);

            var s1 = client.Draw.DrawRectangle(1, 1, 1.25, 1.5);

            var s2 = client.Draw.DrawRectangle(2, 3, 2.5, 3.5);

            var s3 = client.Draw.DrawRectangle(4.5, 2.5, 6, 3.5);

            client.Selection.SelectNone();
            client.Selection.SelectShapesById(s1);
            client.Selection.SelectShapesById(s2);
            client.Selection.SelectShapesById(s3);

            client.Document.OpenStencilDocument("basic_u.vss");
            var connec_stencil = client.Document.OpenStencilDocument("connec_u.vss");
            var connec_tdoc = new VisioScripting.TargetDocument(connec_stencil);
            var master = client.Master.GetMasterWithNameInDocument(connec_tdoc, "Dynamic Connector");
            var undirected_connectors = client.Connection.ConnectShapes(new [] { s1,s2},new [] { s2,s3}, master);

            var options1 = new VisioAutomation.DocumentAnalysis.ConnectionAnalyzerOptions();
            options1.NoArrowsHandling = NoArrowsHandling.ExcludeEdge;

            var directed_edges0 = client.Connection.GetDirectedEdgesOnPage(new VisioScripting.TargetPage(), options1);
            Assert.AreEqual(0, directed_edges0.Count);

            var options2 = new VA.DocumentAnalysis.ConnectionAnalyzerOptions();
            options2.NoArrowsHandling = NoArrowsHandling.TreatEdgeAsBidirectional;

            var directed_edges1 = client.Connection.GetDirectedEdgesOnPage(new VisioScripting.TargetPage(), options2);
            Assert.AreEqual(4, directed_edges1.Count);

            var options3 = new VA.DocumentAnalysis.ConnectionAnalyzerOptions();
            options3.DirectionSource = DirectionSource.UseConnectionOrder;

            var undirected_edges0 = client.Connection.GetDirectedEdgesOnPage(new VisioScripting.TargetPage(), options3);
            Assert.AreEqual(2, undirected_edges0.Count);

            client.Document.CloseActiveDocument(true);
        }


        [TestMethod]
        public void Scripting_Connects_Scenario_3()
        {
            var client = this.GetScriptingClient();
            client.Document.NewDocument();
            var s1 = client.Draw.DrawRectangle(1, 1, 2,2);
            var s2 = client.Draw.DrawRectangle(4, 4, 5, 5);

            client.Connection.ConnectShapes(new[] {s1}, new[] {s2}, null);
            
            client.Document.CloseActiveDocument(true);
        }
    }
}