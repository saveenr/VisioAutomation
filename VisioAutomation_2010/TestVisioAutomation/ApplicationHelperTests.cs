using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TestVisioAutomation
{
    [TestClass]
    public class ApplicationHelperTests : VisioAutomationTest
    {
        [TestMethod]
        public void TestStencilLocation()
        {
            var app = this.GetVisioApplication();
            string path = VisioAutomation.Application.ApplicationHelper.GetContentLocation(app);

            Assert.IsTrue(System.IO.Directory.Exists(path));

            var files1 = System.IO.Directory.GetFiles(path, "*.vs?");
            var files2 = System.IO.Directory.GetFiles(path, "*.vss?");

            Assert.IsTrue( files1.Count()>100 || files2.Count()>100);
        }
    }
}