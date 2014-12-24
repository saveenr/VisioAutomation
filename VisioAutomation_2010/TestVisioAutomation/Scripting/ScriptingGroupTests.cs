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
            var client = GetScriptingClient();
            client.Document.New();
            client.Page.New(new VA.Drawing.Size(4, 4), false);

            var shape_rect = client.Draw.Rectangle(1, 1, 3, 3);
            var shape_line = client.Draw.Line(0.5, 0.5, 3.5, 3.5);
            var shape_oval1 = client.Draw.Oval(0.2, 1, 3.8, 2);
            var shape_oval2 = client.Draw.Oval(new VA.Drawing.Point(2, 2), 0.5);

            client.Selection.All();
            var s0 = client.Selection.GetShapes();
            Assert.AreEqual(4, s0.Count);

            var g = client.Arrange.Group();
            client.Selection.None();
            client.Selection.All();

            var s1 = client.Selection.GetShapes();
            Assert.AreEqual(1, s1.Count);

            client.Arrange.Ungroup(null);
            client.Selection.All();
            var s2 = client.Selection.GetShapes();
            Assert.AreEqual(4, s2.Count);
            client.Document.Close(true);
        }
    }
}