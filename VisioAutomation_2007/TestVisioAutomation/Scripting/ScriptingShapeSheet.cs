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
            var ss = GetScriptingClient();
            var doc = ss.Document.New();
            ss.Draw.Rectangle(0, 0, 1, 1);
            ss.Draw.Rectangle(1, 1, 2, 2);

            var formulas = ss.ShapeSheet.QueryFormulas(null, new[] {VA.ShapeSheet.SRCConstants.PinX});
            Assert.AreEqual("1.5 in", formulas[0][0]);

            ss.Selection.All();
            formulas = ss.ShapeSheet.QueryFormulas(null, new[] { VA.ShapeSheet.SRCConstants.PinX });
            Assert.AreEqual("1.5 in", formulas[0][0]);
            Assert.AreEqual("0.5 in", formulas[1][0]);


            var m = ss.Master.New(doc,"MasterX");

            try
            {
                ss.Master.OpenForEdit(m);
                ss.Draw.Oval(0, 0, 1, 1);
                ss.Draw.Oval(1, 1, 2, 2);
                ss.Draw.Oval(2, 2, 3, 3);


                ss.Selection.All();

                formulas = ss.ShapeSheet.QueryFormulas(null, new[] { VA.ShapeSheet.SRCConstants.PinX });
                //Assert.AreEqual("1.5 in", formulas[0][0]);
                //Assert.AreEqual("0.5 in", formulas[1][0]);


            }
            finally
            {
                ss.Master.CloseMasterEditing();
                
            }

            doc.Close(true);
        }

    }
}