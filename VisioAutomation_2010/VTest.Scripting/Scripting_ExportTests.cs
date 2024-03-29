using System.IO;
using VTest.Framework;
using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VTest.Scripting
{
    [MUT.TestClass]
    public class Scripting_ExportTests : Framework.VTest
    {
        [MUT.TestMethod]
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

            string output_filename = VTestGlobals.VTestHelper.GetOutputFilename(nameof(Scripting_Test_Export_Selection_SVGHTML),".html");

            if (File.Exists(output_filename))
            {
                File.Delete(output_filename);
            }


            client.Export.ExportSelectionToHtml(VisioScripting.TargetSelection.Auto, output_filename);

            AssertUtil.FileExists(output_filename);
            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }
    }
}