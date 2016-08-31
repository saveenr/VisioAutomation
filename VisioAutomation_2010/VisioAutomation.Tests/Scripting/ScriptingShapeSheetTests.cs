using System.Collections.Generic;
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

            var targets = new VisioAutomation.Scripting.TargetShapes();

            var srcs = new[] { VisioAutomation.ShapeSheet.SRCConstants.PinX };

            var formulas = client.ShapeSheet.QueryFormulas(targets, srcs);
            Assert.AreEqual("1.5 in", formulas[0].Cells[0]);

            client.Selection.SelectAll();
            formulas = client.ShapeSheet.QueryFormulas(targets, srcs);
            Assert.AreEqual("1.5 in", formulas[0].Cells[0]);
            Assert.AreEqual("0.5 in", formulas[1].Cells[0]);


            var m = client.Master.New(doc,"MasterX");

            try
            {
                client.Master.OpenForEdit(m);
                var s1 = client.Draw.Oval(0, 0, 1, 1);
                var s2 = client.Draw.Oval(1, 1, 2, 2);
                var s3 = client.Draw.Oval(2, 2, 3, 3);
                
                client.Selection.SelectAll();

                var targets2 = new VisioAutomation.Scripting.TargetShapes( s1,s2,s3);
                formulas = client.ShapeSheet.QueryFormulas(targets2, srcs);
                Assert.AreEqual("0.5 in", formulas[0].Cells[0]);
                Assert.AreEqual("1.5 in", formulas[1].Cells[0]);
                Assert.AreEqual("2.5 in", formulas[2].Cells[0]);
            }
            finally
            {
                client.Master.CloseMasterEditing();
                
            }

            doc.Close(true);
        }

    }
}