using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VisioAutomation_Tests.Scripting
{
    [TestClass]
    public class ScriptingConnectionPointTests : VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_ConnectionPoints_Scenarios()
        {
            var client = this.GetScriptingClient();
            client.Document.NewDocument();
            client.Page.NewPage(new VisioAutomation.Geometry.Size(4, 4), false);

            var s1 = client.Draw.DrawRectangle(1, 1, 1.25, 1.5);
            var s2 = client.Draw.DrawRectangle(2, 3, 2.5, 3.5);
            var s3 = client.Draw.DrawRectangle(4.5, 2.5, 6, 3.5);

            client.Selection.SelectNone();
            client.Selection.SelectShapesById(s1);
            client.Selection.SelectShapesById(s2);
            client.Selection.SelectShapesById(s3);

            var indices0 = client.ConnectionPoint.AddConnectionPoint("0", "Width*0.67", VisioScripting.Models.ConnectionPointType.Outward);
            Assert.AreEqual(3, indices0.Count);
            Assert.AreEqual(0, indices0[0]);
            Assert.AreEqual(0, indices0[1]);
            Assert.AreEqual(0, indices0[2]);

            var targets = new VisioScripting.Models.TargetShapes();
            var dic = client.ConnectionPoint.GetConnectionPoints(targets);
            Assert.AreEqual(3, dic.Count);
            Assert.AreEqual("Width*0.67", dic[s1][0].Y.Value);
            Assert.AreEqual("Width*0.67", dic[s2][0].Y.Value);
            Assert.AreEqual("Width*0.67", dic[s2][0].Y.Value);

            client.ConnectionPoint.DeleteConnectionPointAtIndex(targets,0);
            client.Document.CloseActiveDocument(true);
        }
    }
}