using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VisioScripting_Tests
{
    [MUT.TestClass]
    public class ScriptingClientTests : VTest.VisioAutomationTest
    {
        [MUT.TestMethod]
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

        [MUT.TestMethod]
        public void Scripting_CanCloseUnsavedDrawings()
        {
            var client = this.GetScriptingClient();
            client.Document.CloseAllDocumentsWithoutSaving();

            MUT.Assert.IsFalse(client.Document.HasActiveDocument);

            var doc1 = client.Document.NewDocument();
            MUT.Assert.IsTrue(client.Document.HasActiveDocument);
            MUT.Assert.IsFalse(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Auto));

            client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 0, 0, 1, 1);
            MUT.Assert.IsTrue(client.Document.HasActiveDocument);
            MUT.Assert.IsTrue(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Auto));
            MUT.Assert.IsTrue(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Auto, 1));
            MUT.Assert.IsFalse(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Auto, 2));

            client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 2, 2, 3, 3);
            client.Selection.SelectAllShapes(VisioScripting.TargetWindow.Auto);
            MUT.Assert.IsTrue(client.Document.HasActiveDocument);
            MUT.Assert.IsTrue(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Auto));
            MUT.Assert.IsTrue(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Auto, 1));
            MUT.Assert.IsTrue(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Auto, 2));

            client.Selection.SelectNone(VisioScripting.TargetWindow.Auto);
            MUT.Assert.IsTrue(client.Document.HasActiveDocument);
            MUT.Assert.IsFalse(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Auto));

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }
    }
}