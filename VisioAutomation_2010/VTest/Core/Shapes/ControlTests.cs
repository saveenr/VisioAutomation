using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Shapes;

namespace VTest.Core.Shapes
{
    [TestClass]
    public class ControlTests : VisioAutomationTest
    {
        [TestMethod]
        public void Controls_AddRemove()
        {
            var page1 = this.GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 4, 1);

            // Ensure we start with 0 controls
            Assert.AreEqual(0, ControlHelper.GetCount(s1));

            // Add the first control
            int ci1 = ControlHelper.Add(s1);
            Assert.AreEqual(1, ControlHelper.GetCount(s1));

            // Add the second control
            int ci2 = ControlHelper.Add(s1);
            Assert.AreEqual(2, ControlHelper.GetCount(s1));
            
            // retrieve the control information
            var controls = ControlCells.GetCells(s1, VisioAutomation.Core.CellValueType.Formula);

            // verify that the controls were set propery
            Assert.AreEqual(2, controls.Count);
            Assert.AreEqual("Width*0", controls[0].X.Value);
            Assert.AreEqual("Controls.Row_1", controls[0].XDynamics.Value);
            Assert.AreEqual("Width*0", controls[1].X.Value);
            Assert.AreEqual("Controls.Row_2", controls[1].XDynamics.Value);

            // Delete both controls
            ControlHelper.Delete(s1, 0);
            Assert.AreEqual(1, ControlHelper.GetCount(s1));
            ControlHelper.Delete(s1, 0);
            Assert.AreEqual(0, ControlHelper.GetCount(s1));

            page1.Delete(0);
        }
    }
}