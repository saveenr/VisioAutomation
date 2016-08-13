using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Shapes.ConnectionPoints;

namespace VisioAutomation_Tests.Scripting
{
    [TestClass]
    public class ScriptingConnectionPointTests : VisioAutomationTest
    {
        [TestMethod]
        public void Scripting_ConnectionPoints_Scenarios()
        {
            var client = this.GetScriptingClient();
            client.Document.New();
            client.Page.New(new VisioAutomation.Drawing.Size(4, 4), false);

            var s1 = client.Draw.Rectangle(1, 1, 1.25, 1.5);

            var s2 = client.Draw.Rectangle(2, 3, 2.5, 3.5);

            var s3 = client.Draw.Rectangle(4.5, 2.5, 6, 3.5);

            client.Selection.None();
            client.Selection.Select(s1);
            client.Selection.Select(s2);
            client.Selection.Select(s3);

            var indices0 = client.ConnectionPoint.Add("0", "Width*0.67",
                                                 ConnectionPointType.Outward);
            Assert.AreEqual(3, indices0.Count);
            Assert.AreEqual(0, indices0[0]);
            Assert.AreEqual(0, indices0[1]);
            Assert.AreEqual(0, indices0[2]);

            var dic = client.ConnectionPoint.Get(null);
            Assert.AreEqual(3, dic.Count);
            Assert.AreEqual("Width*0.67", dic[s1][0].Y.Formula);
            Assert.AreEqual("Width*0.67", dic[s2][0].Y.Formula);
            Assert.AreEqual("Width*0.67", dic[s2][0].Y.Formula);

            client.ConnectionPoint.Delete(null,0);
            client.Document.Close(true);
        }
    }
}