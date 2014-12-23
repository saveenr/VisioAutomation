using Microsoft.VisualStudio.TestTools.UnitTesting;
using VACONT = VisioAutomation.Shapes.Controls;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class Control_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void Controls_AddRemove()
        {
            var page1 = GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 4, 1);

            // Ensure we start with 0 controls
            Assert.AreEqual(0, VACONT.ControlHelper.GetCount(s1));

            // Add the first control
            int ci1 = VACONT.ControlHelper.Add(s1);
            Assert.AreEqual(1, VACONT.ControlHelper.GetCount(s1));

            // Add the second control
            int ci2 = VACONT.ControlHelper.Add(s1);
            Assert.AreEqual(2, VACONT.ControlHelper.GetCount(s1));
            
            // retrieve the control information
            var controls = VACONT.ControlCells.GetCells(s1);

            // verify that the controls were set propery
            Assert.AreEqual(2, controls.Count);
            Assert.AreEqual("Width*0", controls[0].X.Formula);
            Assert.AreEqual("Controls.Row_1", controls[0].XDynamics.Formula);
            Assert.AreEqual("Width*0", controls[1].X.Formula);
            Assert.AreEqual("Controls.Row_2", controls[1].XDynamics.Formula);

            // Delete both controls
            VACONT.ControlHelper.Delete(s1, 0);
            Assert.AreEqual(1, VACONT.ControlHelper.GetCount(s1));
            VACONT.ControlHelper.Delete(s1, 0);
            Assert.AreEqual(0, VACONT.ControlHelper.GetCount(s1));

            page1.Delete(0);
        }
    }
}