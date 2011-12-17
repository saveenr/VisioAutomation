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

        [TestMethod]
        public void Check_4x2_bitarray_is_exactly_1_byte()
        {
            var ba = new VA.Internal.BitArray2D(4, 2);

            Assert.AreEqual(8, ba.BitArray.Count);

            var bytes1 = ba.ToBytes();
            Assert.AreEqual(1, bytes1.Length);
            Assert.AreEqual(0x0, bytes1[0]);

            ba.SetAll(true);
            var bytes2 = ba.ToBytes();
            Assert.AreEqual(1, bytes2.Length);
            Assert.AreEqual(0xff, bytes2[0]);

            ba.SetAll(false);
            var bytes3 = ba.ToBytes();
            Assert.AreEqual(1, bytes3.Length);
            Assert.AreEqual(0x0, bytes3[0]);

            ba[3, 1] = true;
            var bytes4 = ba.ToBytes();
            Assert.AreEqual(1, bytes4.Length);
            Assert.AreEqual(0x80, bytes4[0]);

            ba.SetAll(true);
            var bytes5 = ba.ToBytes();
            Assert.AreEqual(0xff, bytes5[0]);
        }

        [TestMethod]
        public void Check_4x3_bitarray_is_2_bytes()
        {
            // 12 bits should reserve 16 bits (2 bytes)
            // only the first 12 bits are usable. The remaining bits will be kept at false.
            var ba = new VA.Internal.BitArray2D(4, 3);
            Assert.AreEqual(12, ba.BitArray.Count);

            var bytes1 = ba.ToBytes();
            Assert.AreEqual(2, bytes1.Length);
            Assert.AreEqual(0x00, bytes1[0]);
            Assert.AreEqual(0x00, bytes1[1]);

            ba[3, 2] = true;
            var bytes4 = ba.ToBytes();
            Assert.AreEqual(2, bytes4.Length);
            Assert.AreEqual(0x0, bytes4[0]);
            Assert.AreEqual(0x8, bytes4[1]);
        }
    }

}
