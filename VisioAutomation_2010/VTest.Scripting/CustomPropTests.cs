using System.Linq;
using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Shapes;
using VA = VisioAutomation;

namespace VTest.Scripting
{
    [MUT.TestClass]
    public class CustomPropTests : Framework.VTest
    {
        [MUT.TestMethod]
        public void ShapeSheet_SetNoShapes()
        {

            var client = this.GetScriptingClient();

            client.Document.NewDocument();
            client.Page.NewPage(VisioScripting.TargetDocument.Auto, new VA.Core.Size(4, 4), false);

            var s1 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 1, 1, 1.25, 1.5);
            var s2 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 2, 3, 2.5, 3.5);
            var s3 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 4.5, 2.5, 6, 3.5);

            client.Selection.SelectNone(VisioScripting.TargetWindow.Auto);

            var targetshapes = new VisioScripting.TargetShapes(s1,s2,s3);
            var targetshapeids = targetshapes.ToShapeIDs();
            var writer = client.ShapeSheet.GetWriterForPage(VisioScripting.TargetPage.Auto);

            foreach (var shapeid in targetshapeids)
            {
                writer.SetFormula( (short) shapeid, VA.Core.SrcConstants.XFormPinX, "1.0");
            }

            writer.Commit();

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

        [MUT.TestMethod]
        public void GetCustomPropertiesAsShapeDictionary_OnFreshShapes_ReturnsZeroPerShape()
        {
            var client = this.GetScriptingClient();
            client.Document.NewDocument();
            client.Page.NewPage(VisioScripting.TargetDocument.Auto, new VA.Core.Size(4, 4), false);

            var s1 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 1, 1, 1.25, 1.5);
            var s2 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 2, 3, 2.5, 3.5);
            var s3 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 4.5, 2.5, 6, 3.5);

            client.Selection.SelectNone(VisioScripting.TargetWindow.Auto);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s1);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s2);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s3);

            var props = client.CustomProperty.GetCustomPropertiesAsShapeDictionary(VisioScripting.TargetShapes.Auto, VisioAutomation.Core.CellValueType.Formula);
            MUT.Assert.AreEqual(3, props.Count);
            MUT.Assert.AreEqual(0, props[s1].Count);
            MUT.Assert.AreEqual(0, props[s2].Count);
            MUT.Assert.AreEqual(0, props[s3].Count);

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

        [MUT.TestMethod]
        public void SetCustomProperty_OnSelectedShapes_AppendsOnePropertyPerShapeWithMatchingValue()
        {
            var client = this.GetScriptingClient();
            client.Document.NewDocument();
            client.Page.NewPage(VisioScripting.TargetDocument.Auto, new VA.Core.Size(4, 4), false);

            var s1 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 1, 1, 1.25, 1.5);
            var s2 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 2, 3, 2.5, 3.5);
            var s3 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 4.5, 2.5, 6, 3.5);

            client.Selection.SelectNone(VisioScripting.TargetWindow.Auto);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s1);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s2);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s3);

            var cp = new CustomPropertyCells();
            cp.Value = "\"BAR\"";
            client.CustomProperty.SetCustomProperty(VisioScripting.TargetShapes.Auto, "FOO", cp);

            var props = client.CustomProperty.GetCustomPropertiesAsShapeDictionary(VisioScripting.TargetShapes.Auto, VisioAutomation.Core.CellValueType.Formula);
            MUT.Assert.AreEqual(3, props.Count);
            MUT.Assert.AreEqual(1, props[s1].Count);
            MUT.Assert.AreEqual(1, props[s2].Count);
            MUT.Assert.AreEqual(1, props[s3].Count);

            MUT.Assert.AreEqual("\"BAR\"", props[s1]["FOO"].Value.Value);
            MUT.Assert.AreEqual("\"BAR\"", props[s2]["FOO"].Value.Value);
            MUT.Assert.AreEqual("\"BAR\"", props[s3]["FOO"].Value.Value);

            var has_props = client.CustomProperty.ContainCustomPropertyWithName(VisioScripting.TargetShapes.Auto, "FOO");
            MUT.Assert.IsTrue(has_props.All(v => v == true));

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

        [MUT.TestMethod]
        public void DeleteCustomPropertyWithName_OnShapesWithThatProperty_RemovesItFromAllShapes()
        {
            var client = this.GetScriptingClient();
            client.Document.NewDocument();
            client.Page.NewPage(VisioScripting.TargetDocument.Auto, new VA.Core.Size(4, 4), false);

            var s1 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 1, 1, 1.25, 1.5);
            var s2 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 2, 3, 2.5, 3.5);
            var s3 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 4.5, 2.5, 6, 3.5);

            client.Selection.SelectNone(VisioScripting.TargetWindow.Auto);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s1);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s2);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s3);

            var cp = new CustomPropertyCells();
            cp.Value = "\"BAR\"";
            client.CustomProperty.SetCustomProperty(VisioScripting.TargetShapes.Auto, "FOO", cp);

            client.CustomProperty.DeleteCustomPropertyWithName(VisioScripting.TargetShapes.Auto, "FOO");

            var props = client.CustomProperty.GetCustomPropertiesAsShapeDictionary(VisioScripting.TargetShapes.Auto, VisioAutomation.Core.CellValueType.Formula);
            MUT.Assert.AreEqual(3, props.Count);
            MUT.Assert.AreEqual(0, props[s1].Count);
            MUT.Assert.AreEqual(0, props[s2].Count);
            MUT.Assert.AreEqual(0, props[s3].Count);

            var has_props = client.CustomProperty.ContainCustomPropertyWithName(VisioScripting.TargetShapes.Auto, "FOO");
            MUT.Assert.IsTrue(has_props.All(v => v == false));

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }
    }
}
