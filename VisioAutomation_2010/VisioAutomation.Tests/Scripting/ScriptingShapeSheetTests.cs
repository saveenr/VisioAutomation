using System.Collections.Generic;
using System.Linq;
using IVisio=Microsoft.Office.Interop.Visio;
using Microsoft.VisualStudio.TestTools.UnitTesting;
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
            var doc = client.Document.New();
            client.Draw.Rectangle(0, 0, 1, 1);
            client.Draw.Rectangle(1, 1, 2, 2);


            var shapes = client.Page.GetShapes();
            var shapeids = shapes.Select(s => s.ID16).ToList();
            
            var srcs = new[] { VisioAutomation.ShapeSheet.SRCConstants.PinX };

            var page = client.Page.Get();

            var reader = client.ShapeSheet.GetReader(page);
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