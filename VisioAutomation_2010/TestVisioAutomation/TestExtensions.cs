using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VA=VisioAutomation;

namespace TestVisioAutomation
{
    public static class TestExtensions
    {
        public static VisioAutomation.Drawing.Point Pin(this VisioAutomation.Shapes.XFormCells xform)
        {
            return new VisioAutomation.Drawing.Point(xform.PinX.Result, xform.PinY.Result);
        }
    }


    [TestClass]
    public class StencilHelperTests : VisioAutomationTest
    {

        [TestMethod]
        public void TestStencilLocation()
        {
            var app = this.GetVisioApplication();
            string path = VA.Application.ApplicationHelper.GetContentLocation(app);

            Assert.IsTrue(System.IO.Directory.Exists(path));

            var files1 = System.IO.Directory.GetFiles(path, "*.vs?");
            var files2 = System.IO.Directory.GetFiles(path, "*.vss?");

            Assert.IsTrue( files1.Count()>10 || files2.Count()>10);

        }
    }
}