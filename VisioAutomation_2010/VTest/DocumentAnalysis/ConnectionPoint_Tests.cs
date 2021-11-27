using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using VisioAutomation.Shapes;

namespace VTest.Core.Connections
{
    [MUT.TestClass]
    public class ConnectionPoint_Tests : VisioAutomationTest
    {
        [MUT.TestMethod]
        public void ConnectionPoints_AddRemove()
        {
            var page1 = this.GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 4, 1);
            MUT.Assert.AreEqual(0, ConnectionPointHelper.GetCount(s1));

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
            MUT.Assert.AreEqual(1, ConnectionPointHelper.GetCount(s1));

            ConnectionPointHelper.Add(s1, cpd2);
            MUT.Assert.AreEqual(2, ConnectionPointHelper.GetCount(s1));

            var controlpoints_f = ConnectionPointCells.GetCells(s1, VisioAutomation.Core.CellValueType.Formula);
            var controlpoints_r = ConnectionPointCells.GetCells(s1, VisioAutomation.Core.CellValueType.Result);
            MUT.Assert.AreEqual(2, controlpoints_f.Count);
            MUT.Assert.AreEqual(2, controlpoints_r.Count);
            var cp_f0 = controlpoints_f[0];
            var cp_r0 = controlpoints_r[0];
            MUT.Assert.AreEqual("0 in", cp_f0.DirX.Value);
            MUT.Assert.AreEqual("0 in", cp_f0.DirY.Value);
            MUT.Assert.AreEqual("0", cp_f0.Type.Value);
            MUT.Assert.AreEqual("Width*0.25", cp_f0.X.Value);
            MUT.Assert.AreEqual("Height*0", cp_f0.Y.Value);

            MUT.Assert.AreEqual("0.0000 in.", cp_r0.DirX.Value);
            MUT.Assert.AreEqual("0.0000 in.", cp_r0.DirY.Value);
            MUT.Assert.AreEqual("0", cp_r0.Type.Value);
            MUT.Assert.AreEqual("1.0000 in.", cp_r0.X.Value);
            MUT.Assert.AreEqual("0.0000 in.", cp_r0.Y.Value);

            var cp_f1 = controlpoints_f[1];
            var cp_r1 = controlpoints_r[1];
            MUT.Assert.AreEqual("0 in", cp_f1.DirX.Value);
            MUT.Assert.AreEqual("0 in", cp_f1.DirY.Value);
            MUT.Assert.AreEqual("0", cp_f1.Type.Value);
            MUT.Assert.AreEqual("Width*0.75", cp_f1.X.Value);
            MUT.Assert.AreEqual("Height*0", cp_f1.Y.Value);

            MUT.Assert.AreEqual("0.0000 in.", cp_r1.DirX.Value);
            MUT.Assert.AreEqual("0.0000 in.", cp_r1.DirY.Value);
            MUT.Assert.AreEqual("0", cp_r1.Type.Value);
            MUT.Assert.AreEqual("3.0000 in.", cp_r1.X.Value);
            MUT.Assert.AreEqual("0.0000 in.", cp_r1.Y.Value);


            ConnectionPointHelper.Delete(s1, 1);
            MUT.Assert.AreEqual(1, ConnectionPointHelper.GetCount(s1));
            ConnectionPointHelper.Delete(s1, 0);
            MUT.Assert.AreEqual(0, ConnectionPointHelper.GetCount(s1));

            page1.Delete(0);
        }

        [MUT.TestMethod]
        public void ConnectionPoints_DeleteAll()
        {
            var page1 = this.GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 4, 1);
            MUT.Assert.AreEqual(0, ConnectionPointHelper.GetCount(s1));

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

            MUT.Assert.AreEqual(4, ConnectionPointHelper.GetCount(s1));

            int num_deleted = ConnectionPointHelper.Delete(s1);
            MUT.Assert.AreEqual(4, num_deleted);
            MUT.Assert.AreEqual(0, ConnectionPointHelper.GetCount(s1));

            page1.Delete(0);
        }

        [MUT.TestMethod]
        public void ConnectionPoints_Set()
        {
            var page1 = this.GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 4, 1);
            MUT.Assert.AreEqual(0, ConnectionPointHelper.GetCount(s1));

            var cp_type = VisioScripting.Models.ConnectionPointType.Inward;

            var xpositions = new[] { "Width*0.25", "Width*0.30", "Width*0.75", "Width*0.90" };
            var ypositions = new[] { "Height*0.15", "Height*0.21", "Height*0.65", "Height*0.89" };

            foreach (int i in Enumerable.Range(0,xpositions.Length))
            {
                var xpos = xpositions[i];
                var ypos = ypositions[i];
                var cp = new ConnectionPointCells();
                cp.X = xpos;
                cp.Y = ypos;
                cp.DirX = 0;
                cp.DirY = 0;
                cp.Type = (int)cp_type;

                ConnectionPointHelper.Add(s1, cp);
            }

            MUT.Assert.AreEqual(4, ConnectionPointHelper.GetCount(s1));

            var desired_cp0 = new ConnectionPointCells();
            desired_cp0.X = "Width*0.025";
            desired_cp0.Y = "Height*0.015";

            var desired_cp1 = new ConnectionPointCells();
            desired_cp1.X = "Width*0.0025";
            desired_cp1.Y = "Height*0.0015";

            ConnectionPointHelper.Set(s1, 0, desired_cp0);
            ConnectionPointHelper.Set(s1, 1, desired_cp1);

            var actual_cp = ConnectionPointCells.GetCells(s1, VisioAutomation.Core.CellValueType.Formula);

            MUT.Assert.AreEqual(desired_cp0.X, actual_cp[0].X);
            MUT.Assert.AreEqual(desired_cp0.Y, actual_cp[0].Y);


            MUT.Assert.AreEqual(desired_cp1.X, actual_cp[1].X);
            MUT.Assert.AreEqual(desired_cp1.Y, actual_cp[1].Y);
            page1.Delete(0);
        }
    }
}