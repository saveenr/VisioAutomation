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
            ss.Document.CloseAllWithoutSaving();

            Assert.IsFalse(ss.HasActiveDrawing);
            Assert.IsFalse(ss.Selection.HasShapes());

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