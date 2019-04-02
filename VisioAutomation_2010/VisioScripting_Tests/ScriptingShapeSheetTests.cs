using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using VisioScripting.Models;

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
            client.Draw.DrawRectangle(0, 0, 1, 1);
            client.Draw.DrawRectangle(1, 1, 2, 2);


            var targetpage = new TargetPage();
            var shapes = client.Page.GetShapesOnPage(targetpage);
            var shapeids = shapes.Select(s => s.ID16).ToList();
            
            var srcs = new[] { VisioAutomation.ShapeSheet.SrcConstants.XFormPinX };


            var reader = client.ShapeSheet.GetReaderForPage(targetpage);
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