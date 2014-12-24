using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

namespace TestVisioAutomation
{
    [TestClass]
    public class ScriptingShapeSheetTests : VisioAutomationTest
    {
        [TestMethod]
        public void QueryPage()
        {
            var client = GetScriptingClient();
            var doc = client.Document.New();
            client.Draw.Rectangle(0, 0, 1, 1);
            client.Draw.Rectangle(1, 1, 2, 2);

            var formulas = client.ShapeSheet.QueryFormulas(null, new[] {VA.ShapeSheet.SRCConstants.PinX});
            Assert.AreEqual("1.5 in", formulas[0][0]);

            client.Selection.All();
            formulas = client.ShapeSheet.QueryFormulas(null, new[] { VA.ShapeSheet.SRCConstants.PinX });
            Assert.AreEqual("1.5 in", formulas[0][0]);
            Assert.AreEqual("0.5 in", formulas[1][0]);


            var m = client.Master.New(doc,"MasterX");

            try
            {
                client.Master.OpenForEdit(m);
                client.Draw.Oval(0, 0, 1, 1);
                client.Draw.Oval(1, 1, 2, 2);
                client.Draw.Oval(2, 2, 3, 3);


                client.Selection.All();

                formulas = client.ShapeSheet.QueryFormulas(null, new[] { VA.ShapeSheet.SRCConstants.PinX });
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