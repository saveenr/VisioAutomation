using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VisioAutomation_Tests.Scripting
{
    [TestClass]
    public class ScriptingGroupTests : VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_Grouping()
        {
            var client = this.GetScriptingClient();
            client.Document.New();
            client.Page.NewPage(new VisioAutomation.Geometry.Size(4, 4), false);

            var shape_rect = client.Draw.Rectangle(1, 1, 3, 3);
            var shape_line = client.Draw.Line(0.5, 0.5, 3.5, 3.5);
            var shape_oval1 = client.Draw.Oval(0.2, 1, 3.8, 2);
            var shape_oval2 = client.Draw.Oval(1.5,1.5, 2.5,2.5);

            client.Selection.SelectAllShapes();
            var s0 = client.Selection.GetShapesInSelection();
            Assert.AreEqual(4, s0.Count);

            var g = client.Grouping.Group();
            client.Selection.SelectNone();
            client.Selection.SelectAllShapes();

            var s1 = client.Selection.GetShapesInSelection();
            Assert.AreEqual(1, s1.Count);

            var targets = new VisioScripting.Models.TargetShapes();

            client.Grouping.Ungroup(targets);
            client.Selection.SelectAllShapes();
            var s2 = client.Selection.GetShapesInSelection();
            Assert.AreEqual(4, s2.Count);
            client.Document.Close(true);
        }
    }
}