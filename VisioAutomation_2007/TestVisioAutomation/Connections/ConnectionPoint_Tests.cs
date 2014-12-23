using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using VA = VisioAutomation;
using VACXN = VisioAutomation.Shapes.Connections;

namespace TestVisioAutomation
{
    [TestClass]
    public class ConnectionPoint_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void ConnectionPoints_AddRemove()
        {
            var page1 = GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 4, 1);
            Assert.AreEqual(0, VACXN.ConnectionPointHelper.GetCount(s1));

            var cp_type = VACXN.ConnectionPointType.Inward;

            var cpd1 = new VACXN.ConnectionPointCells();
            cpd1.X = "Width*0.25";
            cpd1.Y = "Height*0";
            cpd1.Type = (int) cp_type;

            var cpd2 = new VACXN.ConnectionPointCells();
            cpd2.X = "Width*0.75";
            cpd2.Y = "Height*0";
            cpd2.Type = (int) cp_type;

            VACXN.ConnectionPointHelper.Add(s1, cpd1);
            Assert.AreEqual(1, VACXN.ConnectionPointHelper.GetCount(s1));

            VACXN.ConnectionPointHelper.Add(s1, cpd2);
            Assert.AreEqual(2, VACXN.ConnectionPointHelper.GetCount(s1));

            var controlpoints = VACXN.ConnectionPointCells.GetCells(s1);
            Assert.AreEqual(2, controlpoints.Count);
            var cp_0 = controlpoints[0];
            AssertVA.AreEqual("0 in", 0, cp_0.DirX);
            AssertVA.AreEqual("0 in", 0, cp_0.DirY);
            AssertVA.AreEqual("0", 0, cp_0.Type);
            AssertVA.AreEqual("Width*0.25", 1, cp_0.X);
            AssertVA.AreEqual("Height*0", 0, cp_0.Y);

            var cp_1 = controlpoints[1];
            AssertVA.AreEqual("0 in", 0, cp_1.DirX);
            AssertVA.AreEqual("0 in", 0, cp_1.DirY);
            AssertVA.AreEqual("0", 0, cp_1.Type);
            AssertVA.AreEqual("Width*0.75", 3, cp_1.X);
            AssertVA.AreEqual("Height*0", 0, cp_1.Y);

            VACXN.ConnectionPointHelper.Delete(s1, 1);
            Assert.AreEqual(1, VACXN.ConnectionPointHelper.GetCount(s1));
            VACXN.ConnectionPointHelper.Delete(s1, 0);
            Assert.AreEqual(0, VACXN.ConnectionPointHelper.GetCount(s1));

            page1.Delete(0);
        }

        [TestMethod]
        public void ConnectionPoints_DeleteAll()
        {
            var page1 = GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 4, 1);
            Assert.AreEqual(0, VACXN.ConnectionPointHelper.GetCount(s1));

            var cp_type = VACXN.ConnectionPointType.Inward;

            var xpositions = new[] {"Width*0.25", "Width*0.30", "Width*0.75", "Width*0.90"};
            var ypos = "Height*0";

            foreach (var xpos in xpositions)
            {
                var cp = new VACXN.ConnectionPointCells();
                cp.X = xpos;
                cp.Y = ypos;
                cp.DirX = 0;
                cp.DirY = 0;
                cp.Type = (int) cp_type;

                VACXN.ConnectionPointHelper.Add(s1, cp);
            }

            Assert.AreEqual(4, VACXN.ConnectionPointHelper.GetCount(s1));

            int num_deleted = VACXN.ConnectionPointHelper.Delete(s1);
            Assert.AreEqual(4, num_deleted);
            Assert.AreEqual(0, VACXN.ConnectionPointHelper.GetCount(s1));

            page1.Delete(0);
        }
    }
}