using Microsoft.VisualStudio.TestTools.UnitTesting;
using VA=VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class ScriptingExportTests : VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_Test_ExportSVGHTML()
        {
            var client = GetScriptingClient();

            var doc = client.Document.New(10, 5);

            var page1 = doc.Pages[1];
            var app = page1.Application;

            var s1 = page1.DrawRectangle(0, 0, 1, 1);
            var s2 = page1.DrawRectangle(1, 0, 2, 1);
            var s3 = page1.DrawRectangle(0, 1, 1, 2);
            var s4 = page1.DrawRectangle(1, 1, 2, 2);

            client.Selection.All();

            string output_filename = Common.Globals.Helper.GetTestMethodOutputFilename(".html");

            if (System.IO.File.Exists(output_filename))
            {
                System.IO.File.Delete(output_filename);
            }
            client.Export.SelectionToSVGXHTML(output_filename);

            Assert.IsTrue( System.IO.File.Exists(output_filename));
            client.Document.Close(true);
        }
    }
}