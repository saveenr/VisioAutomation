using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Shapes;
using VisioAutomation.ShapeSheet;

namespace VisioAutomation_Tests.Scripting
{
    [TestClass]
    public class ScriptingControlTests : VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_Controls_Scenarios()
        {
            var client = this.GetScriptingClient();

            client.Document.NewDocument();
            client.Page.NewPage(VisioScripting.TargetDocument.Active, new VisioAutomation.Geometry.Size(4, 4), false);

            var s1 = client.Draw.DrawRectangle(1, 1, 1.5, 1.5);
            var s2 = client.Draw.DrawRectangle(2, 3, 2.5, 3.5);
            var s3 = client.Draw.DrawRectangle(1.5, 3.5, 2, 4.0);

            client.Selection.SelectNone(VisioScripting.TargetWindow.Active);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Active, s1);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Active, s2);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Active, s3);

            var targetshapes = new VisioScripting.TargetShapes();

            var controls0 = client.Control.GetControls(targetshapes, CellValueType.Formula);
            int found_controls = controls0.Count;
            Assert.AreEqual(3, controls0.Count);
            Assert.AreEqual(0, controls0[s1].Count);
            Assert.AreEqual(0, controls0[s2].Count);
            Assert.AreEqual(0, controls0[s3].Count);

            var ctrl = new ControlCells();
            ctrl.X = "Width*0.5";
            ctrl.Y = "0";
            client.Control.AddControlToShapes(targetshapes, ctrl);

            var controls1 = client.Control.GetControls(targetshapes, CellValueType.Formula);
            Assert.AreEqual(3, controls1.Count);
            Assert.AreEqual(1, controls1[s1].Count);
            Assert.AreEqual(1, controls1[s2].Count);
            Assert.AreEqual(1, controls1[s3].Count);

            client.Control.DeleteControlWithIndex(targetshapes, 0);
            var controls2 = client.Control.GetControls(targetshapes, CellValueType.Formula);
            Assert.AreEqual(3, controls0.Count);
            Assert.AreEqual(0, controls2[s1].Count);
            Assert.AreEqual(0, controls2[s2].Count);
            Assert.AreEqual(0, controls2[s3].Count);

            client.Document.CloseDocument(VisioScripting.TargetDocument.Active, true);
        }
    }
}