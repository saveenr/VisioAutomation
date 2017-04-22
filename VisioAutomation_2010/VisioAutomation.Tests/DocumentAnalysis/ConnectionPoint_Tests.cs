using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioScripting.Models;
using VisioAutomation.Shapes;

namespace VisioAutomation_Tests.Core.Connections
{
    [TestClass]
    public class ConnectionPoint_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void ConnectionPoints_AddRemove()
        {
            var page1 = this.GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 4, 1);
            Assert.AreEqual(0, ConnectionPointHelper.GetCount(s1));

            var cp_type = ConnectionPointType.Inward;

            var cpd1 = new ConnectionPointCells();
            cpd1.X = "Width*0.25";
            cpd1.Y = "Height*0";
            cpd1.Type = (int) cp_type;

            var cpd2 = new ConnectionPointCells();
            cpd2.X = "Width*0.75";
            cpd2.Y = "Height*0";
            cpd2.Type = (int) cp_type;

            ConnectionPointHelper.Add(s1, cpd1);
            Assert.AreEqual(1, ConnectionPointHelper.GetCount(s1));

            ConnectionPointHelper.Add(s1, cpd2);
            Assert.AreEqual(2, ConnectionPointHelper.GetCount(s1));

            var controlpoints = ConnectionPointCells.GetCells(s1);
            Assert.AreEqual(2, controlpoints.Count);
            var cp_0 = controlpoints[0];
            Assert.AreEqual("0 in", cp_0.DirX.Formula);
            Assert.AreEqual("0 in", cp_0.DirY.Formula);
            Assert.AreEqual("0", cp_0.Type.Formula);
            Assert.AreEqual("Width*0.25", cp_0.X.Formula);
            Assert.AreEqual("Height*0", cp_0.Y.Formula);

            Assert.AreEqual("0.0000 in.", cp_0.DirX.Result);
            Assert.AreEqual("0.0000 in.", cp_0.DirY.Result);
            Assert.AreEqual("0", cp_0.Type.Result);
            Assert.AreEqual("1.0000 in.", cp_0.X.Result);
            Assert.AreEqual("0.0000 in.", cp_0.Y.Result);

            var cp_1 = controlpoints[1];
            Assert.AreEqual("0 in", cp_1.DirX.Formula);
            Assert.AreEqual("0 in", cp_1.DirY.Formula);
            Assert.AreEqual("0", cp_1.Type.Formula);
            Assert.AreEqual("Width*0.75", cp_1.X.Formula);
            Assert.AreEqual("Height*0", cp_1.Y.Formula);

            Assert.AreEqual("0.0000 in.", cp_1.DirX.Result);
            Assert.AreEqual("0.0000 in.", cp_1.DirY.Result);
            Assert.AreEqual("0", cp_1.Type.Result);
            Assert.AreEqual("3.0000 in.", cp_1.X.Result);
            Assert.AreEqual("0.0000 in.", cp_1.Y.Result);


            ConnectionPointHelper.Delete(s1, 1);
            Assert.AreEqual(1, ConnectionPointHelper.GetCount(s1));
            ConnectionPointHelper.Delete(s1, 0);
            Assert.AreEqual(0, ConnectionPointHelper.GetCount(s1));

            page1.Delete(0);
        }

        [TestMethod]
        public void ConnectionPoints_DeleteAll()
        {
            var page1 = this.GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 4, 1);
            Assert.AreEqual(0, ConnectionPointHelper.GetCount(s1));

            var cp_type = ConnectionPointType.Inward;

            var xpositions = new[] {"Width*0.25", "Width*0.30", "Width*0.75", "Width*0.90"};
            var ypos = "Height*0";

            foreach (var xpos in xpositions)
            {
                var cp = new ConnectionPointCells();
                cp.X = xpos;
                cp.Y = ypos;
                cp.DirX = 0;
                cp.DirY = 0;
                cp.Type = (int) cp_type;

                ConnectionPointHelper.Add(s1, cp);
            }

            Assert.AreEqual(4, ConnectionPointHelper.GetCount(s1));

            int num_deleted = ConnectionPointHelper.Delete(s1);
            Assert.AreEqual(4, num_deleted);
            Assert.AreEqual(0, ConnectionPointHelper.GetCount(s1));

            page1.Delete(0);
        }
    }
}