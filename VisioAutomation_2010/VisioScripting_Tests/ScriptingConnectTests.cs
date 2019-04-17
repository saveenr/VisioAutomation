using Microsoft.Office.Interop.Visio;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.DocumentAnalysis;
using VisioScripting;
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

            client.Page.NewPage(VisioScripting.TargetDocument.Active, pagesize, false);

            var s1 = client.Draw.DrawRectangle(1, 1, 1.25, 1.5);
            var s2 = client.Draw.DrawRectangle(2, 3, 2.5, 3.5);
            var s3 = client.Draw.DrawRectangle(4.5, 2.5, 6, 3.5);

            client.Selection.SelectNone(VisioScripting.TargetWindow.Active);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Active, s1);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Active, s2);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Active, s3);

            client.Document.OpenStencilDocument("basic_u.vss");
            var connec_stencil = client.Document.OpenStencilDocument("connec_u.vss");

            var page = VisioScripting.TargetPage.Active;

            var tdoc = new VisioScripting.TargetDocument(connec_stencil);
            var master = client.Master.GetMaster(tdoc, "Dynamic Connector");
            var fromshapes = new [] { s1,s2};
            var toshapes = new [] { s2,s3};
            var directed_connectors = client.Connection.ConnectShapes(page, fromshapes,toshapes, master);
            client.Selection.SelectNone(VisioScripting.TargetWindow.Active);

            client.Selection.SelectShapes(VisioScripting.TargetWindow.Active, directed_connectors);

            var writer = client.ShapeSheet.GetWriterForPage(page);


            var shapes = client.Selection.GetSelectedShapes(VisioScripting.TargetWindow.Active);
            foreach (var shape in shapes)
            {
                writer.SetFormula( shape.ID16, VA.ShapeSheet.SrcConstants.LineEndArrow, "13");
            }
            writer.Commit();

            var options0 = new VA.DocumentAnalysis.ConnectionAnalyzerOptions();
            options0.DirectionSource = DirectionSource.UseConnectionOrder;
            var undirected_edges0 = client.Connection.GetDirectedEdgesOnPage(VisioScripting.TargetPage.Active, options0);
            Assert.AreEqual(2, undirected_edges0.Count);

            var options1 = new VA.DocumentAnalysis.ConnectionAnalyzerOptions();
            options1.NoArrowsHandling = NoArrowsHandling.ExcludeEdge;

            var options2 = new VA.DocumentAnalysis.ConnectionAnalyzerOptions();
            options2.NoArrowsHandling = NoArrowsHandling.TreatEdgeAsBidirectional;

            var directed_edges0 = client.Connection.GetDirectedEdgesOnPage(VisioScripting.TargetPage.Active, options1);
            Assert.AreEqual(2, directed_edges0.Count);

            var directed_edges1 = client.Connection.GetDirectedEdgesOnPage(VisioScripting.TargetPage.Active, options2);
            Assert.AreEqual(2, directed_edges1.Count);

            client.Document.CloseDocument(VisioScripting.TargetDocument.Active, true);
        }

        [TestMethod]
        public void Scripting_Connects_Scenario_1()
        {
            var client = this.GetScriptingClient();

            var targetdoc = new VisioScripting.TargetDocument();

            client.Document.NewDocument();
            var pagesize = new VA.Geometry.Size(4, 4);
            client.Page.NewPage(targetdoc, pagesize, false);

            var s1 = client.Draw.DrawRectangle(1, 1, 1.25, 1.5);
            var s2 = client.Draw.DrawRectangle(2, 3, 2.5, 3.5);
            var s3 = client.Draw.DrawRectangle(4.5, 2.5, 6, 3.5);

            client.Selection.SelectNone(VisioScripting.TargetWindow.Active);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Active, s1);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Active, s2);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Active, s3);

            client.Document.OpenStencilDocument("basic_u.vss");

            var targetpage = VisioScripting.TargetPage.Active;

            var connec_stencil = client.Document.OpenStencilDocument("connec_u.vss");
            var connec_tdoc = new VisioScripting.TargetDocument(connec_stencil);
            var master = client.Master.GetMaster(connec_tdoc, "Dynamic Connector");
            var undirected_connectors = client.Connection.ConnectShapes(targetpage, new [] { s1,s2},new [] { s2,s3}, master);

            var options1 = new VisioAutomation.DocumentAnalysis.ConnectionAnalyzerOptions();
            options1.NoArrowsHandling = NoArrowsHandling.ExcludeEdge;

            var directed_edges0 = client.Connection.GetDirectedEdgesOnPage(targetpage, options1);
            Assert.AreEqual(0, directed_edges0.Count);

            var options2 = new VA.DocumentAnalysis.ConnectionAnalyzerOptions();
            options2.NoArrowsHandling = NoArrowsHandling.TreatEdgeAsBidirectional;

            var directed_edges1 = client.Connection.GetDirectedEdgesOnPage(targetpage, options2);
            Assert.AreEqual(4, directed_edges1.Count);

            var options3 = new VA.DocumentAnalysis.ConnectionAnalyzerOptions();
            options3.DirectionSource = DirectionSource.UseConnectionOrder;

            var undirected_edges0 = client.Connection.GetDirectedEdgesOnPage(targetpage, options3);
            Assert.AreEqual(2, undirected_edges0.Count);

            client.Document.CloseDocument(targetdoc, true);
        }


        [TestMethod]
        public void Scripting_Connects_Scenario_3()
        {
            var client = this.GetScriptingClient();
            client.Document.NewDocument();
            var s1 = client.Draw.DrawRectangle(1, 1, 2,2);
            var s2 = client.Draw.DrawRectangle(4, 4, 5, 5);

            var tagetpage = VisioScripting.TargetPage.Active;
            var fromshapes = new[] {s1};
            var toshapes = new[] {s2};
            Master master = null;
            client.Connection.ConnectShapes(tagetpage, fromshapes, toshapes, master);

            client.Document.CloseDocument(VisioScripting.TargetDocument.Active, true);
        }
    }
}