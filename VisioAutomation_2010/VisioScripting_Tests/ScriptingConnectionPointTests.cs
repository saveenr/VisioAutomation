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
            client.Page.NewPage(VisioScripting.TargetDocument.Auto, new VisioAutomation.Core.Size(4, 4), false);

            var s1 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 1, 1, 1.25, 1.5);
            var s2 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 2, 3, 2.5, 3.5);
            var s3 = client.Draw.DrawRectangle(VisioScripting.TargetPage.Auto, 4.5, 2.5, 6, 3.5);

            client.Selection.SelectNone(VisioScripting.TargetWindow.Auto);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s1);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s2);
            client.Selection.SelectShapesById(VisioScripting.TargetWindow.Auto, s3);

            var indices0 = client.ConnectionPoint.AddConnectionPoint(VisioScripting.TargetShapes.Auto, "0", "Width*0.67", VisioScripting.Models.ConnectionPointType.Outward);
            Assert.AreEqual(3, indices0.Count);
            Assert.AreEqual(0, indices0[0]);
            Assert.AreEqual(0, indices0[1]);
            Assert.AreEqual(0, indices0[2]);

            var dic = client.ConnectionPoint.GetConnectionPoints(VisioScripting.TargetShapes.Auto);
            Assert.AreEqual(3, dic.Count);
            Assert.AreEqual("Width*0.67", dic[s1][0].Y.Value);
            Assert.AreEqual("Width*0.67", dic[s2][0].Y.Value);
            Assert.AreEqual("Width*0.67", dic[s2][0].Y.Value);

            client.ConnectionPoint.DeleteConnectionPointAtIndex(VisioScripting.TargetShapes.Auto, 0);
            client.Document.CloseDocument(VisioScripting.TargetDocuments.Auto);
        }
    }
}