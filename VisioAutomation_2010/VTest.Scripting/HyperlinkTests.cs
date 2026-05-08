using MUT = Microsoft.VisualStudio.TestTools.UnitTesting;
using VA = VisioAutomation;

namespace VTest.Scripting
{
    [MUT.TestClass]
    public class HyperlinkTests : Framework.VTest
    {
        [MUT.TestMethod]
        public void GetHyperlinks_OnFreshShapes_ReturnsZeroPerShape()
        {
            var client = this.GetScriptingClient();
            client.Document.NewDocument();
            client.Page.NewPage(VisioScripting.TargetDocument.Auto, new VisioAutomation.Core.Size(4, 4), false);

            var s1 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 1, 1, 1.5, 1.5);
            var s2 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 2, 3, 2.5, 3.5);
            var s3 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 1.5, 3.5, 2, 4.0);

            client.Selection.SelectNone(VisioScripting.TargetWindow.Auto);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s1);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s2);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s3);

            var hyperlinks = client.Hyperlink.GetHyperlinks(VisioScripting.TargetShapes.Auto, VisioAutomation.Core.CellValueType.Formula);

            MUT.Assert.AreEqual(3, hyperlinks.Count);
            MUT.Assert.AreEqual(0, hyperlinks[s1].Count);
            MUT.Assert.AreEqual(0, hyperlinks[s2].Count);
            MUT.Assert.AreEqual(0, hyperlinks[s3].Count);

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

        [MUT.TestMethod]
        public void AddHyperlink_OnSelectedShapes_AppendsOneHyperlinkPerShape()
        {
            var client = this.GetScriptingClient();
            client.Document.NewDocument();
            client.Page.NewPage(VisioScripting.TargetDocument.Auto, new VisioAutomation.Core.Size(4, 4), false);

            var s1 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 1, 1, 1.5, 1.5);
            var s2 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 2, 3, 2.5, 3.5);
            var s3 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 1.5, 3.5, 2, 4.0);

            client.Selection.SelectNone(VisioScripting.TargetWindow.Auto);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s1);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s2);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s3);

            var hyperlink = new VA.Shapes.HyperlinkCells();
            hyperlink.Address = "http://www.microsoft.com";
            client.Hyperlink.AddHyperlink(VisioScripting.TargetShapes.Auto, hyperlink);

            var hyperlinks = client.Hyperlink.GetHyperlinks(VisioScripting.TargetShapes.Auto, VisioAutomation.Core.CellValueType.Formula);
            MUT.Assert.AreEqual(3, hyperlinks.Count);
            MUT.Assert.AreEqual(1, hyperlinks[s1].Count);
            MUT.Assert.AreEqual(1, hyperlinks[s2].Count);
            MUT.Assert.AreEqual(1, hyperlinks[s3].Count);

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

        [MUT.TestMethod]
        public void DeleteHyperlinkAtIndex_OnShapesWithOneHyperlink_RemovesAllHyperlinks()
        {
            var client = this.GetScriptingClient();
            client.Document.NewDocument();
            client.Page.NewPage(VisioScripting.TargetDocument.Auto, new VisioAutomation.Core.Size(4, 4), false);

            var s1 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 1, 1, 1.5, 1.5);
            var s2 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 2, 3, 2.5, 3.5);
            var s3 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 1.5, 3.5, 2, 4.0);

            client.Selection.SelectNone(VisioScripting.TargetWindow.Auto);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s1);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s2);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s3);

            var hyperlink = new VA.Shapes.HyperlinkCells();
            hyperlink.Address = "http://www.microsoft.com";
            client.Hyperlink.AddHyperlink(VisioScripting.TargetShapes.Auto, hyperlink);

            client.Hyperlink.DeleteHyperlinkAtIndex(VisioScripting.TargetShapes.Auto, 0);
            var hyperlinks = client.Hyperlink.GetHyperlinks(VisioScripting.TargetShapes.Auto, VisioAutomation.Core.CellValueType.Formula);
            MUT.Assert.AreEqual(3, hyperlinks.Count);
            MUT.Assert.AreEqual(0, hyperlinks[s1].Count);
            MUT.Assert.AreEqual(0, hyperlinks[s2].Count);
            MUT.Assert.AreEqual(0, hyperlinks[s3].Count);

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }
    }
}
