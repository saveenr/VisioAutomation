using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VisioAutomation_Tests.Core.Application
{
    [TestClass]
    public class ApplicationHelperTests : VisioAutomationTest
    {
        [TestMethod]
        public void TestStencilLocation()
        {
            var app = this.GetVisioApplication();
            string path = VisioAutomation.Application.ApplicationHelper.GetContentLocation(app);

            Assert.IsTrue(Directory.Exists(path));

            var files1 = Directory.GetFiles(path, "*.vs?");
            var files2 = Directory.GetFiles(path, "*.vss?");

            Assert.IsTrue( files1.Length>100 || files2.Length>100);
        }
    }
}