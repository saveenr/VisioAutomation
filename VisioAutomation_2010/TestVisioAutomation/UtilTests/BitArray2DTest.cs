using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using VA = VisioAutomation;

namespace TestVisioAutomation
{

    [TestClass]
    public class BitArray2DTest
    {
        [TestMethod]
        public void CheckInvalidConstruction()
        {
            bool caught = false;
            try { var ba = new VA.Internal.BitArray2D(0, 1); }
            catch (System.ArgumentOutOfRangeException)
            {
                caught = true;
            }

            if (caught == false)
            {
                Assert.Fail("Did not catch expected exception");
            }

            caught = false;
            try { var ba = new VA.Internal.BitArray2D(1, 0); }
            catch (System.ArgumentOutOfRangeException)
            {
                caught = true;
            }

            if (caught == false)
            {
                Assert.Fail("Did not catch expected exception");
            }

        }

        [TestMethod]
        public void Create_1x1_BitArray()
        {
            var ba = new VA.Internal.BitArray2D(1, 1);
            Assert.AreEqual(false, ba[0, 0]);
            ba[0, 0] = true;
            Assert.AreEqual(true, ba[0, 0]);
            ba[0, 0] = false;
            Assert.AreEqual(false, ba[0, 0]);
        }
    }
}
