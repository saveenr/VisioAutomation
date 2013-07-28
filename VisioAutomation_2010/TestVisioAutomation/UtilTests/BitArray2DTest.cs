using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class BitArray2DTest
    {
        [TestMethod]
        public void Internal_Construct2DBitArray()
        {
            // check that cols and rows must be > 0
            bool caught = false;
            try
            {
                var ba = new VA.Internal.BitArray2D(0, 1);
            }
            catch (System.ArgumentOutOfRangeException)
            {
                caught = true;
            }

            if (caught == false)
            {
                Assert.Fail("Did not catch expected exception");
            }

            caught = false;
            try
            {
                var ba = new VA.Internal.BitArray2D(1, 0);
            }
            catch (System.ArgumentOutOfRangeException)
            {
                caught = true;
            }

            if (caught == false)
            {
                Assert.Fail("Did not catch expected exception");
            }

            // Create a 1x1 BitArray
            var ba2 = new VA.Internal.BitArray2D(1, 1);
            Assert.AreEqual(false, ba2[0, 0]);
            ba2[0, 0] = true;
            Assert.AreEqual(true, ba2[0, 0]);
            ba2[0, 0] = false;
            Assert.AreEqual(false, ba2[0, 0]);
        }
    }
}