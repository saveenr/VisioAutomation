using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VTest.Scripting
{
    [MUT.TestClass]
    public class DevTests : Framework.VTest
    {
        [MUT.TestMethod]
        public void Dev_ScriptingDocumentation()
        {
            var client = this.GetScriptingClient();
            client.Developer.DrawScriptingDocumentation();

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

        [MUT.TestMethod]
        public void DrawInteropEnumDocumentation_NewDocument_HasPagesWithShapes()
        {
            // The only working caller of the internal form page model (VisioScripting.Helpers.Forms).
            var client = this.GetScriptingClient();
            var doc = client.Developer.DrawInteropEnumDocumentation();

            MUT.Assert.IsNotNull(doc);
            MUT.Assert.IsTrue(doc.Pages.Count > 0, "The documentation document should have at least one page");
            MUT.Assert.IsTrue(doc.Pages[1].Shapes.Count > 0, "The first page should contain the title and body shapes");

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }
    }
}