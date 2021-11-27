using UT=Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VisioScripting_Tests
{
    [UT.TestClass]
    public class ScriptingClientTests : VTest.VisioAutomationTest
    {
        [UT.TestMethod]
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

        [UT.TestMethod]
        public void Scripting_CanCloseUnsavedDrawings()
        {
            var client = this.GetScriptingClient();
            client.Document.CloseAllDocumentsWithoutSaving();

            UT.Assert.IsFalse(client.Document.HasActiveDocument);

            var doc1 = client.Document.NewDocument();
            UT.Assert.IsTrue(client.Document.HasActiveDocument);
            UT.Assert.IsFalse(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Auto));

            client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 0, 0, 1, 1);
            UT.Assert.IsTrue(client.Document.HasActiveDocument);
            UT.Assert.IsTrue(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Auto));
            UT.Assert.IsTrue(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Auto, 1));
            UT.Assert.IsFalse(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Auto, 2));

            client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 2, 2, 3, 3);
            client.Selection.SelectAllShapes(VisioScripting.TargetWindow.Auto);
            UT.Assert.IsTrue(client.Document.HasActiveDocument);
            UT.Assert.IsTrue(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Auto));
            UT.Assert.IsTrue(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Auto, 1));
            UT.Assert.IsTrue(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Auto, 2));

            client.Selection.SelectNone(VisioScripting.TargetWindow.Auto);
            UT.Assert.IsTrue(client.Document.HasActiveDocument);
            UT.Assert.IsFalse(client.Selection.ContainsShapes(VisioScripting.TargetSelection.Auto));

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }
    }
}