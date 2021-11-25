using System.IO;
using UT=Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VisioScripting_Tests
{
    [UT.TestClass]
    public class ScriptingExportTests : VisioAutomation_Tests.VisioAutomationTest
    {
        [UT.TestMethod]
        public void Scripting_Test_Export_Selection_SVGHTML()
        {
            var client = this.GetScriptingClient();
            var page_size = new VisioAutomation.Core.Size(10,5);

            var doc = client.Document.NewDocument(page_size);

            var page1 = doc.Pages[1];

            var s1 = page1.DrawRectangle(0, 0, 1, 1);
            var s2 = page1.DrawRectangle(1, 0, 2, 1);
            var s3 = page1.DrawRectangle(0, 1, 1, 2);
            var s4 = page1.DrawRectangle(1, 1, 2, 2);

            client.Selection.SelectAllShapes(VisioScripting.TargetWindow.Auto);

            string output_filename = VisioAutomation_Tests.TestGlobals.TestHelper.GetOutputFilename(nameof(Scripting_Test_Export_Selection_SVGHTML),".html");

            if (File.Exists(output_filename))
            {
                File.Delete(output_filename);
            }


            client.Export.ExportSelectionToHtml(VisioScripting.TargetSelection.Auto, output_filename);

            VisioAutomation_Tests.AssertUtil.FileExists(output_filename);
            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }
    }
}