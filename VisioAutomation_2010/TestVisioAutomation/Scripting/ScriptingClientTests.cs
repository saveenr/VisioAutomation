using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TestVisioAutomation.Scripting
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
            client.Document.Close(true);
        }

        public void DrawVANamespaceDiagram()
        {
            var client = this.GetScriptingClient();
            var doc = client.Developer.DrawNamespaces();
            client.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_CanCloseUnsavedDrawings()
        {
            var client = this.GetScriptingClient();
            client.Document.CloseAllWithoutSaving();

            Assert.IsFalse(client.Document.HasActiveDocument);

            var doc1 = client.Document.New();
            Assert.IsTrue(client.Document.HasActiveDocument);
            Assert.IsFalse(client.Selection.HasShapes());

            client.Draw.Rectangle(0, 0, 1, 1);
            Assert.IsTrue(client.Document.HasActiveDocument);
            Assert.IsTrue(client.Selection.HasShapes());
            Assert.IsTrue(client.Selection.HasShapes(1));
            Assert.IsFalse(client.Selection.HasShapes(2));

            client.Draw.Rectangle(2, 2, 3, 3);
            client.Selection.All();
            Assert.IsTrue(client.Document.HasActiveDocument);
            Assert.IsTrue(client.Selection.HasShapes());
            Assert.IsTrue(client.Selection.HasShapes(1));
            Assert.IsTrue(client.Selection.HasShapes(2));

            client.Selection.None();
            Assert.IsTrue(client.Document.HasActiveDocument);
            Assert.IsFalse(client.Selection.HasShapes());

            client.Document.Close(true);
        }
    }
}