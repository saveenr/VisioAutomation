using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Controls;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class ControlHelper_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void Test_New_Shapes_Have_No_Controls()
        {
            var page1 = GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 4, 1);
            Assert.AreEqual(0, ControlHelper.GetControlsCount(s1));
            page1.Delete(0);
        }

        [TestMethod]
        public void AddControls()
        {
            var page1 = GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 4, 1);

            Assert.AreEqual(0, ControlHelper.GetControlsCount(s1));
            int ci1 = ControlHelper.AddControl(s1);
            Assert.AreEqual(1, ControlHelper.GetControlsCount(s1));
            int ci2 = ControlHelper.AddControl(s1);
            Assert.AreEqual(2, ControlHelper.GetControlsCount(s1));

            var controls = ControlHelper.GetControls(s1);
            Assert.IsNotNull(controls);
            Assert.AreEqual(2, controls.Count);
            Assert.AreEqual("Width*0", controls[0].X.Formula);
            Assert.AreEqual("Controls.Row_1", controls[0].XDynamics.Formula);
            ControlHelper.DeleteControl(s1, 0);
            Assert.AreEqual(1, ControlHelper.GetControlsCount(s1));
            ControlHelper.DeleteControl(s1, 0);
            Assert.AreEqual(0, ControlHelper.GetControlsCount(s1));

            page1.Delete(0);
        }

        [TestMethod]
        public void DeleteControls()
        {
            var page1 = GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 4, 1);

            Assert.AreEqual(0, ControlHelper.GetControlsCount(s1));
            int ci1 = ControlHelper.AddControl(s1);
            Assert.AreEqual(1, ControlHelper.GetControlsCount(s1));
            int ci2 = ControlHelper.AddControl(s1);
            Assert.AreEqual(2, ControlHelper.GetControlsCount(s1));

            var controls = ControlHelper.GetControls(s1);
            Assert.IsNotNull(controls);
            Assert.AreEqual(2, controls.Count);
            Assert.AreEqual(0.0, controls[0].X.Result);
            ControlHelper.DeleteControl(s1, 0);
            Assert.AreEqual(1, ControlHelper.GetControlsCount(s1));
            ControlHelper.DeleteControl(s1, 0);
            Assert.AreEqual(0, ControlHelper.GetControlsCount(s1));
            page1.Delete(0);
        }

        [TestMethod]
        public void CountControls()
        {
            var page1 = GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 4, 1);

            Assert.AreEqual(0, ControlHelper.GetControlsCount(s1));
            int ci1 = ControlHelper.AddControl(s1);
            Assert.AreEqual(1, ControlHelper.GetControlsCount(s1));
            int ci2 = ControlHelper.AddControl(s1);
            Assert.AreEqual(2, ControlHelper.GetControlsCount(s1));
            page1.Delete(0);
        }
    }
}