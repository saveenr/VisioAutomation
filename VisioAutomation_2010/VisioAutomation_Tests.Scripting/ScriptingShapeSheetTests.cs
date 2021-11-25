using VisioAutomation.Extensions;

namespace VisioAutomation_Tests.Scripting
{
    [TestClass]
    public class ScriptingShapeSheetTests : VisioAutomationTest
    {
        [TestMethod]
        public void QueryPage()
        {
            var client = this.GetScriptingClient();
            var doc = client.Document.NewDocument();


            client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 0, 0, 1, 1);
            client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 1, 1, 2, 2);

            var shapes = client.Page.GetShapesOnPage(VisioScripting.TargetPage.Auto);
            var shapeids = shapes.Select(s => s.ID16).ToList();
            
            var srcs = new[] { VisioAutomation.ShapeSheet.SrcConstants.XFormPinX };


            var reader = client.ShapeSheet.GetReaderForPage(VisioScripting.TargetPage.Auto);
            foreach (var shapeid in shapeids)
            {
                foreach (var src in srcs)
                {
                    reader.AddCell(shapeid,src);
                }
            }

            var formulas = reader.GetFormulas();
            Assert.AreEqual("0.5 in", formulas[0]);
            Assert.AreEqual("1.5 in", formulas[1]);
            doc.Close(true);
        }
    }
}