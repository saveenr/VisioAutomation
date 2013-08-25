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
            var ss = GetScriptingSession();

            var doc = ss.Document.New(10, 5);

            var page1 = doc.Pages[1];
            var app = page1.Application;

            var s1 = page1.DrawRectangle(0, 0, 1, 1);
            var s2 = page1.DrawRectangle(1, 0, 2, 1);
            var s3 = page1.DrawRectangle(0, 1, 1, 2);
            var s4 = page1.DrawRectangle(1, 1, 2, 2);

            ss.Selection.SelectAll();

            string output_filename = TestCommon.Globals.Helper.GetTestMethodOutputFilename(".html");

            if (System.IO.File.Exists(output_filename))
            {
                System.IO.File.Delete(output_filename);
            }
            ss.Export.SelectionToSVGXHTML(output_filename);

            Assert.IsTrue( System.IO.File.Exists(output_filename));
            ss.Document.Close(true);
        }
    }
}