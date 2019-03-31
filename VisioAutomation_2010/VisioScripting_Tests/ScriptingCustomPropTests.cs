using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Shapes;
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
            client.Page.NewPage(new VA.Geometry.Size(4, 4), false);

            var s1 = client.Draw.DrawRectangle(1, 1, 1.25, 1.5);
            var s2 = client.Draw.DrawRectangle(2, 3, 2.5, 3.5);
            var s3 = client.Draw.DrawRectangle(4.5, 2.5, 6, 3.5);

            client.Selection.SelectNone();

            var targetshapes = new VisioScripting.Models.TargetShapes(s1,s2,s3);
            var targetshapeids = targetshapes.ToShapeIDs();
            var page = client.Page.GetActivePage();
            var writer = client.ShapeSheet.GetWriterForPage(page);

            foreach (var shapeid in targetshapeids)
            {
                writer.SetFormula( (short) shapeid, VA.ShapeSheet.SrcConstants.XFormPinX, "1.0");
            }

            writer.Commit();
            
            client.Document.CloseActiveDocument(true);
        }

        [TestMethod]
        public void Scripting_CustomProps_Scenarios()
        {
            var client = this.GetScriptingClient();
            client.Document.NewDocument();
            client.Page.NewPage(new VA.Geometry.Size(4, 4), false);

            var s1 = client.Draw.DrawRectangle(1, 1, 1.25, 1.5);
            var s2 = client.Draw.DrawRectangle(2, 3, 2.5, 3.5);
            var s3 = client.Draw.DrawRectangle(4.5, 2.5, 6, 3.5);

            client.Selection.SelectNone();
            client.Selection.SelectShapesById(s1);
            client.Selection.SelectShapesById(s2);
            client.Selection.SelectShapesById(s3);

            var targetshapes = new VisioScripting.Models.TargetShapes();
            var prop_dic0 = client.CustomProperty.GetCustomProperties(targetshapes);
            Assert.AreEqual(3, prop_dic0.Count);
            Assert.AreEqual(0, prop_dic0[s1].Count);
            Assert.AreEqual(0, prop_dic0[s2].Count);
            Assert.AreEqual(0, prop_dic0[s3].Count);

            var cp = new CustomPropertyCells();
            cp.Value = "\"BAR\"";
            client.CustomProperty.SetCustomProperty(targetshapes, "FOO",cp);

            var prop_dic1 = client.CustomProperty.GetCustomProperties(targetshapes);
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
            

            var hasprops0 = client.CustomProperty.ContainCustomPropertyWithName(targetshapes,"FOO");
            Assert.IsTrue(hasprops0.All(v => v == true));

            client.CustomProperty.DeleteCustomPropertyWithName(targetshapes,"FOO");

            var prop_dic2 = client.CustomProperty.GetCustomProperties(targetshapes);
            Assert.AreEqual(3, prop_dic2.Count);
            Assert.AreEqual(0, prop_dic2[s1].Count);
            Assert.AreEqual(0, prop_dic2[s2].Count);
            Assert.AreEqual(0, prop_dic2[s3].Count);

            var hasprops1 = client.CustomProperty.ContainCustomPropertyWithName(targetshapes,"FOO");
            Assert.IsTrue(hasprops1.All(v => v == false));

            client.Document.CloseActiveDocument(true);
        }
    }
}