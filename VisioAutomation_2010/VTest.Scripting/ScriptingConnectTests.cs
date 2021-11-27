using Microsoft.Office.Interop.Visio;
using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.DocumentAnalysis;
using VA = VisioAutomation;

namespace VisioScripting_Tests
{
    [MUT.TestClass]
    public class ScriptingConnectTests : VTest.VisioAutomationTest
    {
        [MUT.TestMethod]
        public void Scripting_Connects_Scenario_0()
        {
            var client = this.GetScriptingClient();


            client.Document.NewDocument();
            var pagesize = new VA.Core.Size(4, 4);

            client.Page.NewPage(VisioScripting.TargetDocument.Auto, pagesize, false);

            var s1 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 1, 1, 1.25, 1.5);
            var s2 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 2, 3, 2.5, 3.5);
            var s3 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 4.5, 2.5, 6, 3.5);

            client.Selection.SelectNone(VisioScripting.TargetWindow.Auto);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s1);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s2);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s3);

            client.Document.OpenStencilDocument("basic_u.vss");
            var connec_stencil = client.Document.OpenStencilDocument("connec_u.vss");

            var page = VisioScripting.TargetPage.Auto;

            var tdoc = new VisioScripting.TargetDocument(connec_stencil);
            var master = client.Master.GetMaster(tdoc, "Dynamic Connector");
            var fromshapes = new [] { s1,s2};
            var toshapes = new [] { s2,s3};
            var directed_connectors = client.Connection.ConnectShapes(page, fromshapes,toshapes, master);
            client.Selection.SelectNone(VisioScripting.TargetWindow.Auto);

            client.Selection.SelectShapes(VisioScripting.TargetWindow.Auto, directed_connectors);

            var writer = client.ShapeSheet.GetWriterForPage(page);


            var shapes = client.Selection.GetSelectedShapes(VisioScripting.TargetWindow.Auto);
            foreach (var shape in shapes)
            {
                writer.SetFormula( shape.ID16, VA.Core.SrcConstants.LineEndArrow, "13");
            }
            writer.Commit();

            var options0 = new VA.DocumentAnalysis.ConnectionAnalyzerOptions();
            options0.DirectionSource = DirectionSource.UseConnectionOrder;
            var undirected_edges0 = client.Connection.GetDirectedEdgesOnPage(VisioScripting.TargetPage.Auto, options0);
            MUT.Assert.AreEqual(2, undirected_edges0.Count);

            var options1 = new VA.DocumentAnalysis.ConnectionAnalyzerOptions();
            options1.NoArrowsHandling = NoArrowsHandling.ExcludeEdge;

            var options2 = new VA.DocumentAnalysis.ConnectionAnalyzerOptions();
            options2.NoArrowsHandling = NoArrowsHandling.TreatEdgeAsBidirectional;

            var directed_edges0 = client.Connection.GetDirectedEdgesOnPage(VisioScripting.TargetPage.Auto, options1);
            MUT.Assert.AreEqual(2, directed_edges0.Count);

            var directed_edges1 = client.Connection.GetDirectedEdgesOnPage(VisioScripting.TargetPage.Auto, options2);
            MUT.Assert.AreEqual(2, directed_edges1.Count);

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

        [MUT.TestMethod]
        public void Scripting_Connects_Scenario_1()
        {
            var client = this.GetScriptingClient();

            client.Document.NewDocument();
            var pagesize = new VA.Core.Size(4, 4);
            client.Page.NewPage(VisioScripting.TargetDocument.Auto, pagesize, false);

            var s1 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 1, 1, 1.25, 1.5);
            var s2 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 2, 3, 2.5, 3.5);
            var s3 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 4.5, 2.5, 6, 3.5);

            client.Selection.SelectNone(VisioScripting.TargetWindow.Auto);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s1);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s2);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s3);

            client.Document.OpenStencilDocument("basic_u.vss");

            var targetpage = VisioScripting.TargetPage.Auto;

            var connec_stencil = client.Document.OpenStencilDocument("connec_u.vss");
            var connec_tdoc = new VisioScripting.TargetDocument(connec_stencil);
            var master = client.Master.GetMaster(connec_tdoc, "Dynamic Connector");
            var undirected_connectors = client.Connection.ConnectShapes(targetpage, new [] { s1,s2},new [] { s2,s3}, master);

            var options1 = new VisioAutomation.DocumentAnalysis.ConnectionAnalyzerOptions();
            options1.NoArrowsHandling = NoArrowsHandling.ExcludeEdge;

            var directed_edges0 = client.Connection.GetDirectedEdgesOnPage(targetpage, options1);
            MUT.Assert.AreEqual(0, directed_edges0.Count);

            var options2 = new VA.DocumentAnalysis.ConnectionAnalyzerOptions();
            options2.NoArrowsHandling = NoArrowsHandling.TreatEdgeAsBidirectional;

            var directed_edges1 = client.Connection.GetDirectedEdgesOnPage(targetpage, options2);
            MUT.Assert.AreEqual(4, directed_edges1.Count);

            var options3 = new VA.DocumentAnalysis.ConnectionAnalyzerOptions();
            options3.DirectionSource = DirectionSource.UseConnectionOrder;

            var undirected_edges0 = client.Connection.GetDirectedEdgesOnPage(targetpage, options3);
            MUT.Assert.AreEqual(2, undirected_edges0.Count);

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }


        [MUT.TestMethod]
        public void Scripting_Connects_Scenario_3()
        {
            var client = this.GetScriptingClient();
            client.Document.NewDocument();

            var s1 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 1, 1, 2,2);
            var s2 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 4, 4, 5, 5);

            var tagetpage = VisioScripting.TargetPage.Auto;
            var fromshapes = new[] {s1};
            var toshapes = new[] {s2};
            Master master = null;
            client.Connection.ConnectShapes(tagetpage, fromshapes, toshapes, master);

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }
    }
}