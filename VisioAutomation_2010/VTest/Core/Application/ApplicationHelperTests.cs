using System.IO;
using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VTest.Core.Application
{
    [MUT.TestClass]
    public class ApplicationHelperTests : VisioAutomationTest
    {
        [MUT.TestMethod]
        public void TestStencilLocation()
        {
            var app = this.GetVisioApplication();
            string path = VisioAutomation.Application.ApplicationHelper.GetContentLocation(app);

            MUT.Assert.IsTrue(Directory.Exists(path));

            var files1 = Directory.GetFiles(path, "*.vs?");
            var files2 = Directory.GetFiles(path, "*.vss?");

            MUT.Assert.IsTrue( files1.Length>100 || files2.Length>100);
        }
    }
}