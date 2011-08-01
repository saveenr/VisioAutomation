using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TestVisioAutomation
{
    [TestClass]
    public class ScriptingSessionTest : VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_Test_Has_Star()
        {
            var ss = GetScriptingSession();
            ss.Document.CloseAllDocumentsWithoutSaving();

            Assert.IsFalse(ss.HasActiveDrawing());
            Assert.IsFalse(ss.Selection.HasSelectedShapes());

            var doc1 = ss.Document.NewDocument();
            Assert.IsTrue(ss.HasActiveDrawing());
            Assert.IsFalse(ss.Selection.HasSelectedShapes());

            ss.Draw.DrawRectangle(0, 0, 1, 1);
            Assert.IsTrue(ss.HasActiveDrawing());
            Assert.IsTrue(ss.Selection.HasSelectedShapes());
            Assert.IsTrue(ss.Selection.HasSelectedShapes(1));
            Assert.IsFalse(ss.Selection.HasSelectedShapes(2));

            ss.Draw.DrawRectangle(2, 2, 3, 3);
            ss.Selection.SelectAll();
            Assert.IsTrue(ss.HasActiveDrawing());
            Assert.IsTrue(ss.Selection.HasSelectedShapes());
            Assert.IsTrue(ss.Selection.HasSelectedShapes(1));
            Assert.IsTrue(ss.Selection.HasSelectedShapes(2));

            ss.Selection.SelectNone();
            Assert.IsTrue(ss.HasActiveDrawing());
            Assert.IsFalse(ss.Selection.HasSelectedShapes());

            ss.Document.CloseAllDocumentsWithoutSaving();
        }
    }
}