using Microsoft.VisualStudio.TestTools.UnitTesting;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

namespace TestVisioAutomation
{
    [TestClass]
    public class ScriptingSessionTests : VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_DevDocumentationScenarios
            ()
        {
            var ss = GetScriptingSession();
            this.DrawVAScriptingAPIDiagram();
            this.DrawVANamespaceDiagram();
        }

        public void DrawVAScriptingAPIDiagram()
        {
            var ss = GetScriptingSession();
            var doc = ss.Developer.DrawScriptingDocumentation();
            ss.Document.Close(true);
        }

        public void DrawVANamespaceDiagram()
        {
            var ss = GetScriptingSession();
            var doc = ss.Developer.DrawNamespaces();
            ss.Document.Close(true);
        }

        [TestMethod]
        public void Scripting_CanCloseUnsavedDrawings()
        {
            var ss = GetScriptingSession();
            ss.Document.CloseAllWithoutSaving();

            Assert.IsFalse(ss.HasActiveDocument);

            var doc1 = ss.Document.New();
            Assert.IsTrue(ss.HasActiveDocument);
            Assert.IsFalse(ss.Selection.HasShapes());

            ss.Draw.Rectangle(0, 0, 1, 1);
            Assert.IsTrue(ss.HasActiveDocument);
            Assert.IsTrue(ss.Selection.HasShapes());
            Assert.IsTrue(ss.Selection.HasShapes(1));
            Assert.IsFalse(ss.Selection.HasShapes(2));

            ss.Draw.Rectangle(2, 2, 3, 3);
            ss.Selection.All();
            Assert.IsTrue(ss.HasActiveDocument);
            Assert.IsTrue(ss.Selection.HasShapes());
            Assert.IsTrue(ss.Selection.HasShapes(1));
            Assert.IsTrue(ss.Selection.HasShapes(2));

            ss.Selection.None();
            Assert.IsTrue(ss.HasActiveDocument);
            Assert.IsFalse(ss.Selection.HasShapes());

            ss.Document.Close(true);
        }
    }
}