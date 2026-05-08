using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Shapes;

namespace VTest.Core.Shapes
{
    [MUT.TestClass]
    public class ControlTests : Framework.VTest
    {
        [MUT.TestMethod]
        public void GetCount_OnFreshShape_ReturnsZero()
        {
            var page1 = this.GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 4, 1);

            MUT.Assert.AreEqual(0, ControlHelper.GetCount(s1));

            page1.Delete(0);
        }

        [MUT.TestMethod]
        public void Add_TwoControlsToShape_GetCellsReturnsBothWithRowDynamics()
        {
            var page1 = this.GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 4, 1);

            ControlHelper.Add(s1);
            MUT.Assert.AreEqual(1, ControlHelper.GetCount(s1));

            ControlHelper.Add(s1);
            MUT.Assert.AreEqual(2, ControlHelper.GetCount(s1));

            var controls = ControlCells.GetCells(s1, VisioAutomation.Core.CellValueType.Formula);
            MUT.Assert.AreEqual(2, controls.Count);
            MUT.Assert.AreEqual("Width*0", controls[0].X.Value);
            MUT.Assert.AreEqual("Controls.Row_1", controls[0].XDynamics.Value);
            MUT.Assert.AreEqual("Width*0", controls[1].X.Value);
            MUT.Assert.AreEqual("Controls.Row_2", controls[1].XDynamics.Value);

            page1.Delete(0);
        }

        [MUT.TestMethod]
        public void Delete_RemovesControlsOneAtATime_CountDropsToZero()
        {
            var page1 = this.GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 4, 1);

            ControlHelper.Add(s1);
            ControlHelper.Add(s1);
            MUT.Assert.AreEqual(2, ControlHelper.GetCount(s1));

            ControlHelper.Delete(s1, 0);
            MUT.Assert.AreEqual(1, ControlHelper.GetCount(s1));
            ControlHelper.Delete(s1, 0);
            MUT.Assert.AreEqual(0, ControlHelper.GetCount(s1));

            page1.Delete(0);
        }
    }
}
