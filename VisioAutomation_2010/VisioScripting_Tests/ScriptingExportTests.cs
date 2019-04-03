using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VisioAutomation_Tests.Scripting
{
    [TestClass]
    public class ScriptingExportTests : VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_Test_Export_Selection_SVGHTML()
        {
            var client = this.GetScriptingClient();
            var page_size = new VisioAutomation.Geometry.Size(10,5);

            var doc = client.Document.NewDocument(page_size);

            var page1 = doc.Pages[1];

            var s1 = page1.DrawRectangle(0, 0, 1, 1);
            var s2 = page1.DrawRectangle(1, 0, 2, 1);
            var s3 = page1.DrawRectangle(0, 1, 1, 2);
            var s4 = page1.DrawRectangle(1, 1, 2, 2);

            client.Selection.SelectAllShapes();

            string output_filename = TestGlobals.TestHelper.GetOutputFilename(nameof(Scripting_Test_Export_Selection_SVGHTML),".html");

            if (File.Exists(output_filename))
            {
                File.Delete(output_filename);
            }

            client.ExportSelection.ExportSelectionToHtml(output_filename);

            AssertUtil.FileExists(output_filename);
            var targetdoc = new VisioScripting.TargetDocument();
            client.Document.CloseDocument(targetdoc, true);
        }
    }
}