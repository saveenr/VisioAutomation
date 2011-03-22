using Microsoft.VisualStudio.TestTools.UnitTesting;
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
            Assert.AreEqual(0, VA.Controls.ControlHelper.GetControlsCount(s1));
            page1.Delete(0);
        }

        [TestMethod]
        public void AddControls()
        {
            var page1 = GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 4, 1);

            Assert.AreEqual(0, VA.Controls.ControlHelper.GetControlsCount(s1));
            int ci1 = VA.Controls.ControlHelper.AddControl(s1);
            Assert.AreEqual(1, VA.Controls.ControlHelper.GetControlsCount(s1));
            int ci2 = VA.Controls.ControlHelper.AddControl(s1);
            Assert.AreEqual(2, VA.Controls.ControlHelper.GetControlsCount(s1));

            var controls = VA.Controls.ControlHelper.GetControls(s1);
            Assert.IsNotNull(controls);
            Assert.AreEqual(2, controls.Count);
            Assert.AreEqual("Width*0", controls[0].X.Formula);
            Assert.AreEqual("Controls.Row_1", controls[0].XDynamics.Formula);
            VA.Controls.ControlHelper.DeleteControl(s1, 0);
            Assert.AreEqual(1, VA.Controls.ControlHelper.GetControlsCount(s1));
            VA.Controls.ControlHelper.DeleteControl(s1, 0);
            Assert.AreEqual(0, VA.Controls.ControlHelper.GetControlsCount(s1));

            page1.Delete(0);
        }

        [TestMethod]
        public void DeleteControls()
        {
            var page1 = GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 4, 1);

            Assert.AreEqual(0, VA.Controls.ControlHelper.GetControlsCount(s1));
            int ci1 = VA.Controls.ControlHelper.AddControl(s1);
            Assert.AreEqual(1, VA.Controls.ControlHelper.GetControlsCount(s1));
            int ci2 = VA.Controls.ControlHelper.AddControl(s1);
            Assert.AreEqual(2, VA.Controls.ControlHelper.GetControlsCount(s1));

            var controls = VA.Controls.ControlHelper.GetControls(s1);
            Assert.IsNotNull(controls);
            Assert.AreEqual(2, controls.Count);
            Assert.AreEqual(0.0, controls[0].X.Result);
            VA.Controls.ControlHelper.DeleteControl(s1, 0);
            Assert.AreEqual(1, VA.Controls.ControlHelper.GetControlsCount(s1));
            VA.Controls.ControlHelper.DeleteControl(s1, 0);
            Assert.AreEqual(0, VA.Controls.ControlHelper.GetControlsCount(s1));
            page1.Delete(0);
        }

        [TestMethod]
        public void CountControls()
        {
            var page1 = GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 4, 1);

            Assert.AreEqual(0, VA.Controls.ControlHelper.GetControlsCount(s1));
            int ci1 = VA.Controls.ControlHelper.AddControl(s1);
            Assert.AreEqual(1, VA.Controls.ControlHelper.GetControlsCount(s1));
            int ci2 = VA.Controls.ControlHelper.AddControl(s1);
            Assert.AreEqual(2, VA.Controls.ControlHelper.GetControlsCount(s1));
            page1.Delete(0);
        }
    }
}