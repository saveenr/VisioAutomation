using Microsoft.VisualStudio.TestTools.UnitTesting;
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

            var cp_type = VisioScripting.Models.ConnectionPointType.Inward;

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

            var controlpoints_f = ConnectionPointCells.GetCells(s1, VisioAutomation.ShapeSheet.CellValueType.Formula);
            var controlpoints_r = ConnectionPointCells.GetCells(s1, VisioAutomation.ShapeSheet.CellValueType.Result);
            Assert.AreEqual(2, controlpoints_f.Count);
            Assert.AreEqual(2, controlpoints_r.Count);
            var cp_f0 = controlpoints_f[0];
            var cp_r0 = controlpoints_r[0];
            Assert.AreEqual("0 in", cp_f0.DirX.Value);
            Assert.AreEqual("0 in", cp_f0.DirY.Value);
            Assert.AreEqual("0", cp_f0.Type.Value);
            Assert.AreEqual("Width*0.25", cp_f0.X.Value);
            Assert.AreEqual("Height*0", cp_f0.Y.Value);

            Assert.AreEqual("0.0000 in.", cp_r0.DirX.Value);
            Assert.AreEqual("0.0000 in.", cp_r0.DirY.Value);
            Assert.AreEqual("0", cp_r0.Type.Value);
            Assert.AreEqual("1.0000 in.", cp_r0.X.Value);
            Assert.AreEqual("0.0000 in.", cp_f0.Y.Value);

            var cp_f1 = controlpoints_f[1];
            var cp_r1 = controlpoints_r[1];
            Assert.AreEqual("0 in", cp_f1.DirX.Value);
            Assert.AreEqual("0 in", cp_f1.DirY.Value);
            Assert.AreEqual("0", cp_f1.Type.Value);
            Assert.AreEqual("Width*0.75", cp_f1.X.Value);
            Assert.AreEqual("Height*0", cp_f1.Y.Value);

            Assert.AreEqual("0.0000 in.", cp_r1.DirX.Value);
            Assert.AreEqual("0.0000 in.", cp_r1.DirY.Value);
            Assert.AreEqual("0", cp_r1.Type.Value);
            Assert.AreEqual("3.0000 in.", cp_r1.X.Value);
            Assert.AreEqual("0.0000 in.", cp_r1.Y.Value);


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

            var cp_type = VisioScripting.Models.ConnectionPointType.Inward;

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