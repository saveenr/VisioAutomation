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


            client.Document.NewDocument();
            client.Page.NewPage(VisioScripting.TargetDocument.Auto, new VisioAutomation.Geometry.Size(4, 4), false);

            var shape_rect = client.Draw.DrawRectangle(1, 1, 3, 3);
            var shape_line = client.Draw.DrawLine(0.5, 0.5, 3.5, 3.5);
            var shape_oval1 = client.Draw.DrawOval(0.2, 1, 3.8, 2);
            var shape_oval2 = client.Draw.DrawOval(1.5,1.5, 2.5,2.5);

            client.Selection.SelectAllShapes(VisioScripting.TargetWindow.Auto);
            var s0 = client.Selection.GetSelectedShapes(VisioScripting.TargetWindow.Auto);
            Assert.AreEqual(4, s0.Count);

            var g = client.Grouping.Group(VisioScripting.TargetSelection.Auto);
            client.Selection.SelectNone(VisioScripting.TargetWindow.Auto);
            client.Selection.SelectAllShapes(VisioScripting.TargetWindow.Auto);

            var s1 = client.Selection.GetSelectedShapes(VisioScripting.TargetWindow.Auto);
            Assert.AreEqual(1, s1.Count);

            client.Grouping.Ungroup(VisioScripting.TargetShapes.Auto);
            client.Selection.SelectAllShapes(VisioScripting.TargetWindow.Auto);
            var s2 = client.Selection.GetSelectedShapes(VisioScripting.TargetWindow.Auto);
            Assert.AreEqual(4, s2.Count);

            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }
    }
}