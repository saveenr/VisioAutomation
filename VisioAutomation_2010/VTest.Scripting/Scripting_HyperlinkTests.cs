using MUT = Microsoft.VisualStudio.TestTools.UnitTesting;
using VA = VisioAutomation;

namespace VTest.Scripting
{
    [MUT.TestClass]
    public class Scripting_HyperlinkTests : Framework.VTest
    {


        [MUT.TestMethod]
        public void Scripting_Hyperlinks_Scenarios()
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


            var hyperlinks0 = client.Hyperlink.GetHyperlinks(VisioScripting.TargetShapes.Auto, VisioAutomation.Core.CellValueType.Formula);

            MUT.Assert.AreEqual(3, hyperlinks0.Count);
            MUT.Assert.AreEqual(0, hyperlinks0[s1].Count);
            MUT.Assert.AreEqual(0, hyperlinks0[s2].Count);
            MUT.Assert.AreEqual(0, hyperlinks0[s3].Count);

            var hyperlink = new VA.Shapes.HyperlinkCells();
            hyperlink.Address = "http://www.microsoft.com";
            client.Hyperlink.AddHyperlink(VisioScripting.TargetShapes.Auto, hyperlink);

            var hyperlinks1 = client.Hyperlink.GetHyperlinks(VisioScripting.TargetShapes.Auto, VisioAutomation.Core.CellValueType.Formula);
            MUT.Assert.AreEqual(3, hyperlinks1.Count);
            MUT.Assert.AreEqual(1, hyperlinks1[s1].Count);
            MUT.Assert.AreEqual(1, hyperlinks1[s2].Count);
            MUT.Assert.AreEqual(1, hyperlinks1[s3].Count);

            client.Hyperlink.DeleteHyperlinkAtIndex(VisioScripting.TargetShapes.Auto, 0);
            var hyperlinks2 = client.Hyperlink.GetHyperlinks(VisioScripting.TargetShapes.Auto, VisioAutomation.Core.CellValueType.Formula);
            MUT.Assert.AreEqual(3, hyperlinks0.Count);
            MUT.Assert.AreEqual(0, hyperlinks2[s1].Count);
            MUT.Assert.AreEqual(0, hyperlinks2[s2].Count);
            MUT.Assert.AreEqual(0, hyperlinks2[s3].Count);

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }
    }
}