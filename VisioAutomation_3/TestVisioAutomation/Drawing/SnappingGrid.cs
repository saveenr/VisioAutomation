using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class SnappingGraidTests : VisioAutomationTest
    {

        [TestMethod]
        public void Snap1()
        {
            var g = new VA.Drawing.SnappingGrid(1.0, 1.0);

            double delta = 0.000000001;

            AssertX.AreEqual(0.0, 0.0, g.Snap(0.0, 0.0), delta);
            AssertX.AreEqual(0.0, 0.0, g.Snap(0.3,0.3), delta);
            AssertX.AreEqual(0.0, 0.0, g.Snap(0.49999, 0.49999), delta);
            AssertX.AreEqual(1.0, 1.0, g.Snap(0.5, 0.5), delta);
            AssertX.AreEqual(1.0, 1.0, g.Snap(0.500001, 0.500001), delta);
            AssertX.AreEqual(1.0, 1.0, g.Snap(1.0, 1.0), delta);
            AssertX.AreEqual(1.0, 1.0, g.Snap(1.3,1.3), delta);
            AssertX.AreEqual(1.0, 1.0, g.Snap(1.49999, 1.49999), delta);
            AssertX.AreEqual(2.0, 2.0, g.Snap(1.5, 1.5), delta);
            AssertX.AreEqual(2.0, 2.0, g.Snap(1.500001, 1.500001), delta);
        }
    }
}