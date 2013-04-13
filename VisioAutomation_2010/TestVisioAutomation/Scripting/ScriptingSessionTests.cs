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
        public void Scripting_DevDocumentation()
        {
            var ss = GetScriptingSession();
            var doc= ss.Developer.DrawScriptingDocumentation();
            doc.Close(true);
        }

        [TestMethod]
        public void Scripting_DevDocumentation2()
        {
            var ss = GetScriptingSession();
            var doc = ss.Developer.DrawNamespaces();
            //doc.Close(true);
        }

        [TestMethod]
        public void Scripting_Test_Has_Star()
        {
            var ss = GetScriptingSession();
            ss.Document.CloseAllWithoutSaving();

            Assert.IsFalse(ss.HasActiveDrawing);

            var doc1 = ss.Document.New();
            Assert.IsTrue(ss.HasActiveDrawing);
            Assert.IsFalse(ss.Selection.HasShapes());

            ss.Draw.Rectangle(0, 0, 1, 1);
            Assert.IsTrue(ss.HasActiveDrawing);
            Assert.IsTrue(ss.Selection.HasShapes());
            Assert.IsTrue(ss.Selection.HasShapes(1));
            Assert.IsFalse(ss.Selection.HasShapes(2));

            ss.Draw.Rectangle(2, 2, 3, 3);
            ss.Selection.SelectAll();
            Assert.IsTrue(ss.HasActiveDrawing);
            Assert.IsTrue(ss.Selection.HasShapes());
            Assert.IsTrue(ss.Selection.HasShapes(1));
            Assert.IsTrue(ss.Selection.HasShapes(2));

            ss.Selection.SelectNone();
            Assert.IsTrue(ss.HasActiveDrawing);
            Assert.IsFalse(ss.Selection.HasShapes());

            ss.Document.CloseAllWithoutSaving();
        }
    }
}