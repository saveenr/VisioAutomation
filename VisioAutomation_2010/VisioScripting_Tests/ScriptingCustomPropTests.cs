using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Shapes;
using VisioScripting.Models;
using VA = VisioAutomation;

namespace VisioAutomation_Tests.Scripting
{
    [TestClass]
    public class ScriptingCustomPropTests : VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_ShapeSheet_SetNoShapes()
        {

            var client = this.GetScriptingClient();


            client.Document.NewDocument();
            client.Page.NewPage(VisioScripting.TargetDocument.Auto, new VA.Geometry.Size(4, 4), false);

            var s1 = client.Draw.DrawRectangle(1, 1, 1.25, 1.5);
            var s2 = client.Draw.DrawRectangle(2, 3, 2.5, 3.5);
            var s3 = client.Draw.DrawRectangle(4.5, 2.5, 6, 3.5);

            client.Selection.SelectNone(VisioScripting.TargetWindow.Auto);

            var targetshapes = new VisioScripting.TargetShapes(s1,s2,s3);
            var targetshapeids = targetshapes.ToShapeIDs();
            var writer = client.ShapeSheet.GetWriterForPage(VisioScripting.TargetPage.Auto);

            foreach (var shapeid in targetshapeids)
            {
                writer.SetFormula( (short) shapeid, VA.ShapeSheet.SrcConstants.XFormPinX, "1.0");
            }

            writer.Commit();

            client.Document.CloseDocument(VisioScripting.TargetDocument.Auto, true);
        }

        [TestMethod]
        public void Scripting_CustomProps_Scenarios()
        {
            var client = this.GetScriptingClient();
           
            client.Document.NewDocument();
            client.Page.NewPage(VisioScripting.TargetDocument.Auto, new VA.Geometry.Size(4, 4), false);

            var s1 = client.Draw.DrawRectangle(1, 1, 1.25, 1.5);
            var s2 = client.Draw.DrawRectangle(2, 3, 2.5, 3.5);
            var s3 = client.Draw.DrawRectangle(4.5, 2.5, 6, 3.5);

            client.Selection.SelectNone(VisioScripting.TargetWindow.Auto);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s1);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s2);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s3);

            var prop_dic0 = client.CustomProperty.GetCustomPropertiesAsShapeDictionary(VisioScripting.TargetShapes.Auto, VisioAutomation.ShapeSheet.CellValueType.Formula);
            Assert.AreEqual(3, prop_dic0.Count);
            Assert.AreEqual(0, prop_dic0[s1].Count);
            Assert.AreEqual(0, prop_dic0[s2].Count);
            Assert.AreEqual(0, prop_dic0[s3].Count);

            var cp = new CustomPropertyCells();
            cp.Value = "\"BAR\"";
            client.CustomProperty.SetCustomProperty(VisioScripting.TargetShapes.Auto, "FOO",cp);

            var prop_dic1 = client.CustomProperty.GetCustomPropertiesAsShapeDictionary(VisioScripting.TargetShapes.Auto, VisioAutomation.ShapeSheet.CellValueType.Formula);
            Assert.AreEqual(3, prop_dic1.Count);
            Assert.AreEqual(1, prop_dic1[s1].Count);
            Assert.AreEqual(1, prop_dic1[s2].Count);
            Assert.AreEqual(1, prop_dic1[s3].Count);

            var cp1 = prop_dic1[s1]["FOO"];
            var cp2 = prop_dic1[s2]["FOO"];
            var cp3 = prop_dic1[s3]["FOO"];
            Assert.AreEqual("\"BAR\"", cp1.Value.Value);
            Assert.AreEqual("\"BAR\"", cp2.Value.Value);
            Assert.AreEqual("\"BAR\"", cp3.Value.Value);
            

            var hasprops0 = client.CustomProperty.ContainCustomPropertyWithName(VisioScripting.TargetShapes.Auto, "FOO");
            Assert.IsTrue(hasprops0.All(v => v == true));

            client.CustomProperty.DeleteCustomPropertyWithName(VisioScripting.TargetShapes.Auto, "FOO");

            var prop_dic2 = client.CustomProperty.GetCustomPropertiesAsShapeDictionary(VisioScripting.TargetShapes.Auto, VisioAutomation.ShapeSheet.CellValueType.Formula);
            Assert.AreEqual(3, prop_dic2.Count);
            Assert.AreEqual(0, prop_dic2[s1].Count);
            Assert.AreEqual(0, prop_dic2[s2].Count);
            Assert.AreEqual(0, prop_dic2[s3].Count);

            var hasprops1 = client.CustomProperty.ContainCustomPropertyWithName(VisioScripting.TargetShapes.Auto, "FOO");
            Assert.IsTrue(hasprops1.All(v => v == false));

            client.Document.CloseDocument(VisioScripting.TargetDocument.Auto, true);
        }
    }
}