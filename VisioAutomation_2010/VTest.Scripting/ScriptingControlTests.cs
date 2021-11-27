using UT = Microsoft.VisualStudio.TestTools.UnitTesting;
using VA = VisioAutomation;

namespace VisioScripting_Tests
{
    [MUT.TestClass]
    public class ScriptingControlTests : VTest.VisioAutomationTest
    {
        [MUT.TestMethod]
        public void Scripting_Controls_Scenarios()
        {
            var client = this.GetScriptingClient();

            client.Document.NewDocument();
            client.Page.NewPage(VisioScripting.TargetDocument.Auto, new VisioAutomation.Core.Size(4, 4), false);

            var s1 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 1, 1, 1.5, 1.5);
            var s2 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 2, 3, 2.5, 3.5);
            var s3 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 1.5, 3.5, 2, 4.0);

            client.Selection.SelectNone(VisioScripting.TargetWindow.Auto);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s1);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s2);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s3);

            var controls0 = client.Control.GetControls(VisioScripting.TargetShapes.Auto, VisioAutomation.Core.CellValueType.Formula);
            int found_controls = controls0.Count;
            MUT.Assert.AreEqual(3, controls0.Count);
            MUT.Assert.AreEqual(0, controls0[s1].Count);
            MUT.Assert.AreEqual(0, controls0[s2].Count);
            MUT.Assert.AreEqual(0, controls0[s3].Count);

            var ctrl = new VA.Shapes.ControlCells();
            ctrl.X = "Width*0.5";
            ctrl.Y = "0";
            client.Control.AddControlToShapes(VisioScripting.TargetShapes.Auto, ctrl);

            var controls1 = client.Control.GetControls(VisioScripting.TargetShapes.Auto, VisioAutomation.Core.CellValueType.Formula);
            MUT.Assert.AreEqual(3, controls1.Count);
            MUT.Assert.AreEqual(1, controls1[s1].Count);
            MUT.Assert.AreEqual(1, controls1[s2].Count);
            MUT.Assert.AreEqual(1, controls1[s3].Count);

            client.Control.DeleteControlWithIndex(VisioScripting.TargetShapes.Auto, 0);
            var controls2 = client.Control.GetControls(VisioScripting.TargetShapes.Auto, VisioAutomation.Core.CellValueType.Formula);
            MUT.Assert.AreEqual(3, controls0.Count);
            MUT.Assert.AreEqual(0, controls2[s1].Count);
            MUT.Assert.AreEqual(0, controls2[s2].Count);
            MUT.Assert.AreEqual(0, controls2[s3].Count);

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }
    }
}