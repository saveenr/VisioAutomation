using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VisioAutomation_Tests.Scripting
{
    [TestClass]
    public class ScriptingClientTests : VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_DevDocumentationScenarios
            ()
        {
            var client = this.GetScriptingClient();
            this.DrawVAScriptingAPIDiagram();
            this.DrawVANamespaceDiagram();
        }

        public void DrawVAScriptingAPIDiagram()
        {
            var client = this.GetScriptingClient();
            var doc = client.Developer.DrawScriptingDocumentation();

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

        public void DrawVANamespaceDiagram()
        {
            var client = this.GetScriptingClient();
            var doc = client.Developer.DrawNamespaces();

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }

        [TestMethod]
        public void Scripting_CanCloseUnsavedDrawings()
        {
            var client = this.GetScriptingClient();
            client.Document.CloseAllDocumentsWithoutSaving();

            Assert.IsFalse(client.Document.HasActiveDocument);

            var doc1 = client.Document.NewDocument();
            Assert.IsTrue(client.Document.HasActiveDocument);
            Assert.IsFalse(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Auto));

            client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 0, 0, 1, 1);
            Assert.IsTrue(client.Document.HasActiveDocument);
            Assert.IsTrue(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Auto));
            Assert.IsTrue(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Auto, 1));
            Assert.IsFalse(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Auto, 2));

            client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 2, 2, 3, 3);
            client.Selection.SelectAllShapes(VisioScripting.TargetWindow.Auto);
            Assert.IsTrue(client.Document.HasActiveDocument);
            Assert.IsTrue(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Auto));
            Assert.IsTrue(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Auto, 1));
            Assert.IsTrue(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Auto, 2));

            client.Selection.SelectNone(VisioScripting.TargetWindow.Auto);
            Assert.IsTrue(client.Document.HasActiveDocument);
            Assert.IsFalse(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Auto));

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }
    }
}