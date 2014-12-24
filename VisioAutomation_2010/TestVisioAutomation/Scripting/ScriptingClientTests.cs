using Microsoft.VisualStudio.TestTools.UnitTesting;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

namespace TestVisioAutomation
{
    [TestClass]
    public class ScriptingClientTests : VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_DevDocumentationScenarios
            ()
        {
            var client = GetScriptingClient();
            this.DrawVAScriptingAPIDiagram();
            this.DrawVANamespaceDiagram();
        }

        public void DrawVAScriptingAPIDiagram()
        {
            var client = GetScriptingClient();
            var doc = client.Developer.DrawScriptingDocumentation();
            client.Document.Close(true);
        }

        public void DrawVANamespaceDiagram()
        {
            var client = GetScriptingClient();
            var doc = client.Developer.DrawNamespaces();
            client.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_CanCloseUnsavedDrawings()
        {
            var client = GetScriptingClient();
            client.Document.CloseAllWithoutSaving();

            Assert.IsFalse(client.HasActiveDocument);

            var doc1 = client.Document.New();
            Assert.IsTrue(client.HasActiveDocument);
            Assert.IsFalse(client.Selection.HasShapes());

            client.Draw.Rectangle(0, 0, 1, 1);
            Assert.IsTrue(client.HasActiveDocument);
            Assert.IsTrue(client.Selection.HasShapes());
            Assert.IsTrue(client.Selection.HasShapes(1));
            Assert.IsFalse(client.Selection.HasShapes(2));

            client.Draw.Rectangle(2, 2, 3, 3);
            client.Selection.All();
            Assert.IsTrue(client.HasActiveDocument);
            Assert.IsTrue(client.Selection.HasShapes());
            Assert.IsTrue(client.Selection.HasShapes(1));
            Assert.IsTrue(client.Selection.HasShapes(2));

            client.Selection.None();
            Assert.IsTrue(client.HasActiveDocument);
            Assert.IsFalse(client.Selection.HasShapes());

            client.Document.Close(true);
        }
    }
}