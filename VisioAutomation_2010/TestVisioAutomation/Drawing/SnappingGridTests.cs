using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class SnappingGridTests : VisioAutomationTest
    {
        [TestMethod]
        public void Internal_ValidateSnappingGrid()
        {
            double delta = 0.000000001;

            var g1 = new VA.Drawing.SnappingGrid(1.0, 1.0);

            AssertVA.AssertSnap(0.0, 0.0, g1, 0.0, 0.0, delta);
            AssertVA.AssertSnap(0.0, 0.0, g1, 0.3, 0.3, delta);
            AssertVA.AssertSnap(0.0, 0.0, g1, 0.49999, 0.49999, delta);
            AssertVA.AssertSnap(1.0, 1.0, g1, 0.5, 0.5, delta);
            AssertVA.AssertSnap(1.0, 1.0, g1, 0.500001, 0.500001, delta);
            AssertVA.AssertSnap(1.0, 1.0, g1, 1.0, 1.0, delta);
            AssertVA.AssertSnap(1.0, 1.0, g1, 1.3, 1.3, delta);
            AssertVA.AssertSnap(1.0, 1.0, g1, 1.49999, 1.49999, delta);
            AssertVA.AssertSnap(2.0, 2.0, g1, 1.5, 1.5, delta);
            AssertVA.AssertSnap(2.0, 2.0, g1, 1.500001, 1.500001, delta);

            var g2 = new VA.Drawing.SnappingGrid(1.0, 0.3);

            AssertVA.AssertSnap(0.0, 0.0, g2, 0.0, 0.0, delta);
            AssertVA.AssertSnap(0.0, 0.0, g2, 0.3, 0.1, delta);
            AssertVA.AssertSnap(0.0, 0.0, g2, 0.49999, 0.149, delta);
            AssertVA.AssertSnap(1.0, 0.3, g2, 0.5, 0.3, delta);
            AssertVA.AssertSnap(1.0, 0.3, g2, 0.500001, 0.30001, delta);
        }
    }
}