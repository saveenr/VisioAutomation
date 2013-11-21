using Microsoft.VisualStudio.TestTools.UnitTesting;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class ScriptingGroupTests : VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_Grouping()
        {
            var ss = GetScriptingSession();
            ss.Document.New();
            ss.Page.New(new VA.Drawing.Size(4, 4), false);

            var shape_rect = ss.Draw.Rectangle(1, 1, 3, 3);
            var shape_line = ss.Draw.Line(0.5, 0.5, 3.5, 3.5);
            var shape_oval1 = ss.Draw.Oval(0.2, 1, 3.8, 2);
            var shape_oval2 = ss.Draw.Oval(new VA.Drawing.Point(2, 2), 0.5);

            ss.Selection.All();
            var s0 = ss.Selection.GetShapes();
            Assert.AreEqual(4, s0.Count);

            var g = ss.Arrange.Group();
            ss.Selection.None();
            ss.Selection.All();

            var s1 = ss.Selection.GetShapes();
            Assert.AreEqual(1, s1.Count);

            ss.Arrange.Ungroup(null);
            ss.Selection.All();
            var s2 = ss.Selection.GetShapes();
            Assert.AreEqual(4, s2.Count);
            ss.Document.Close(true);
        }
    }
}