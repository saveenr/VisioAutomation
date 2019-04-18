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

            client.Document.CloseDocument(VisioScripting.TargetDocument.Active, true);
        }

        public void DrawVANamespaceDiagram()
        {
            var client = this.GetScriptingClient();
            var doc = client.Developer.DrawNamespaces();

            client.Document.CloseDocument(VisioScripting.TargetDocument.Active, true);
        }

        [TestMethod]
        public void Scripting_CanCloseUnsavedDrawings()
        {
            var client = this.GetScriptingClient();
            client.Document.CloseAllDocumentsWithoutSaving();

            Assert.IsFalse(client.Document.HasActiveDocument);

            var doc1 = client.Document.NewDocument();
            Assert.IsTrue(client.Document.HasActiveDocument);
            Assert.IsFalse(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Active));

            client.Draw.DrawRectangle(0, 0, 1, 1);
            Assert.IsTrue(client.Document.HasActiveDocument);
            Assert.IsTrue(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Active));
            Assert.IsTrue(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Active, 1));
            Assert.IsFalse(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Active, 2));

            client.Draw.DrawRectangle(2, 2, 3, 3);
            client.Selection.SelectAllShapes(VisioScripting.TargetWindow.Active);
            Assert.IsTrue(client.Document.HasActiveDocument);
            Assert.IsTrue(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Active));
            Assert.IsTrue(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Active, 1));
            Assert.IsTrue(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Active, 2));

            client.Selection.SelectNone(VisioScripting.TargetWindow.Active);
            Assert.IsTrue(client.Document.HasActiveDocument);
            Assert.IsFalse(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Active));

            client.Document.CloseDocument(VisioScripting.TargetDocument.Active, true);
        }
    }
}