using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.ShapeSheet;
using VAS = VisioAutomation.Scripting;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class ScriptingControlTests : VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_Controls_Scenario_0()
        {
            var ss = GetScriptingSession();
            ss.Document.NewDocument();
            ss.Page.NewPage(new VA.Drawing.Size(4, 4), false);

            var s1 = ss.Draw.DrawRectangle(1, 1, 1.5, 1.5);

            var s2 = ss.Draw.DrawRectangle(2, 3, 2.5, 3.5);

            var s3 = ss.Draw.DrawRectangle(1.5, 3.5, 2, 4.0);

            ss.Selection.SelectNone();
            ss.Selection.SelectShape(s1);
            ss.Selection.SelectShape(s2);
            ss.Selection.SelectShape(s3);

            var controls0 = ss.Control.GetControls();
            Assert.AreEqual(3, controls0.Count);
            Assert.AreEqual(0, controls0[s1].Count);
            Assert.AreEqual(0, controls0[s2].Count);
            Assert.AreEqual(0, controls0[s3].Count);

            var ctrl = new VA.Controls.ControlCells();
            ctrl.X = "Width*0.5";
            ctrl.Y = "0";
            ss.Control.AddControl(ctrl);

            var controls1 = ss.Control.GetControls();
            Assert.AreEqual(3, controls1.Count);
            Assert.AreEqual(1, controls1[s1].Count);
            Assert.AreEqual(1, controls1[s2].Count);
            Assert.AreEqual(1, controls1[s3].Count);

            ss.Control.DeleteControl(0);
            var controls2 = ss.Control.GetControls();
            Assert.AreEqual(3, controls0.Count);
            Assert.AreEqual(0, controls2[s1].Count);
            Assert.AreEqual(0, controls2[s2].Count);
            Assert.AreEqual(0, controls2[s3].Count);

            ss.Document.CloseDocument(true);
        }
    }
}