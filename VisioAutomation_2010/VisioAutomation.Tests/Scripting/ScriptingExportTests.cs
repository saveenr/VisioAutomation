using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VisioAutomation_Tests.Scripting
{
    [TestClass]
    public class ScriptingExportTests : VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_Test_ExportSVGHTML()
        {
            var client = this.GetScriptingClient();

            var doc = client.Document.New(10, 5);

            var page1 = doc.Pages[1];
            var app = page1.Application;

            var s1 = page1.DrawRectangle(0, 0, 1, 1);
            var s2 = page1.DrawRectangle(1, 0, 2, 1);
            var s3 = page1.DrawRectangle(0, 1, 1, 2);
            var s4 = page1.DrawRectangle(1, 1, 2, 2);

            client.Selection.All();

            string output_filename = TestGlobals.TestHelper.GetTestMethodOutputFilename(".html");

            if (File.Exists(output_filename))
            {
                File.Delete(output_filename);
            }
            client.Export.SelectionToSVGXHTML(output_filename);

            Assert.IsTrue( File.Exists(output_filename));
            client.Document.Close(true);
        }
    }
}