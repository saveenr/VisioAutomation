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

            var formulas = client.ShapeSheet.QueryFormulas(targets, new[] {VisioAutomation.ShapeSheet.SRCConstants.PinX});
            Assert.AreEqual("1.5 in", formulas[0].Cells[0]);

            client.Selection.SelectAll();
            formulas = client.ShapeSheet.QueryFormulas(targets, new[] { VisioAutomation.ShapeSheet.SRCConstants.PinX });
            Assert.AreEqual("1.5 in", formulas[0].Cells[0]);
            Assert.AreEqual("0.5 in", formulas[1].Cells[0]);


            var m = client.Master.New(doc,"MasterX");

            try
            {
                client.Master.OpenForEdit(m);
                client.Draw.Oval(0, 0, 1, 1);
                client.Draw.Oval(1, 1, 2, 2);
                client.Draw.Oval(2, 2, 3, 3);


                client.Selection.SelectAll();


                formulas = client.ShapeSheet.QueryFormulas(targets, new[] { VisioAutomation.ShapeSheet.SRCConstants.PinX });
                //Assert.AreEqual("1.5 in", formulas[0][0]);
                //Assert.AreEqual("0.5 in", formulas[1][0]);


            }
            finally
            {
                client.Master.CloseMasterEditing();
                
            }

            doc.Close(true);
        }

    }
}