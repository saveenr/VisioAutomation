using Microsoft.VisualStudio.TestTools.UnitTesting;
using VA=VisioAutomation;
using VASS=VisioAutomation.ShapeSheet;

namespace VisioAutomation_Tests.Scripting
{
    [TestClass]
    public class ScriptingHyperlinkTests : VisioAutomationTest
    {


        [TestMethod]
        public void Scripting_Hyperlinks_Scenarios()
        {
            var client = this.GetScriptingClient();

            var targetdoc = new VisioScripting.TargetDocument();


            client.Document.NewDocument();
            client.Page.NewPage(targetdoc, new VisioAutomation.Geometry.Size(4, 4), false);

            var s1 = client.Draw.DrawRectangle(1, 1, 1.5, 1.5);
            var s2 = client.Draw.DrawRectangle(2, 3, 2.5, 3.5);
            var s3 = client.Draw.DrawRectangle(1.5, 3.5, 2, 4.0);
            var targetwindow = new VisioScripting.TargetWindow();
            client.Selection.SelectNone(targetwindow);
            client.Selection.SelectShapesById(targetwindow, s1);
            client.Selection.SelectShapesById(targetwindow, s2);
            client.Selection.SelectShapesById(targetwindow, s3);

            var targetshapes = new VisioScripting.TargetShapes();

            var hyperlinks0 = client.Hyperlink.GetHyperlinks(targetshapes, VASS.CellValueType.Formula);

            Assert.AreEqual(3, hyperlinks0.Count);
            Assert.AreEqual(0, hyperlinks0[s1].Count);
            Assert.AreEqual(0, hyperlinks0[s2].Count);
            Assert.AreEqual(0, hyperlinks0[s3].Count);

            var hyperlink = new VA.Shapes.HyperlinkCells();
            hyperlink.Address = "http://www.microsoft.com";
            client.Hyperlink.AddHyperlink(targetshapes, hyperlink);

            var hyperlinks1 = client.Hyperlink.GetHyperlinks(targetshapes, VASS.CellValueType.Formula);
            Assert.AreEqual(3, hyperlinks1.Count);
            Assert.AreEqual(1, hyperlinks1[s1].Count);
            Assert.AreEqual(1, hyperlinks1[s2].Count);
            Assert.AreEqual(1, hyperlinks1[s3].Count);

            client.Hyperlink.DeleteHyperlinkAtIndex(targetshapes, 0);
            var hyperlinks2 = client.Hyperlink.GetHyperlinks(targetshapes, VASS.CellValueType.Formula);
            Assert.AreEqual(3, hyperlinks0.Count);
            Assert.AreEqual(0, hyperlinks2[s1].Count);
            Assert.AreEqual(0, hyperlinks2[s2].Count);
            Assert.AreEqual(0, hyperlinks2[s3].Count);

            client.Document.CloseDocument(targetdoc, true);
        }
    }
}