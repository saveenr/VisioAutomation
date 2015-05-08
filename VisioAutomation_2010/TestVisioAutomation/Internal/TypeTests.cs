using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TestVisioAutomation.Internal
{
    [TestClass]
    public class TypeTests : VisioAutomationTest
    {
        [TestMethod]
        public void VerifySRCSize()
        {
            // SRCs must be 6 bytes
            var c1 = new VisioAutomation.ShapeSheet.SRC();
            int actual_size = System.Runtime.InteropServices.Marshal.SizeOf(c1);
            Assert.AreEqual(6, actual_size);

            this.VerifyFormulaLiteralSize();
        }

        public void VerifyFormulaLiteralSize()
        {
            // A FormulaLiteral only has a reference to a string
            // so it should be as big as a reference to a string

            var instance = new VisioAutomation.ShapeSheet.FormulaLiteral();
            int actual_size = System.Runtime.InteropServices.Marshal.SizeOf(instance);
            Assert.AreEqual(4, actual_size);
        }

        [TestMethod]
        public void Construct2DBitArray()
        {
            // check that cols and rows must be > 0
            bool caught = false;
            try
            {
                var ba = new VisioAutomation.Internal.BitArray2D(0, 1);
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
                var ba = new VisioAutomation.Internal.BitArray2D(1, 0);
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
            var ba2 = new VisioAutomation.Internal.BitArray2D(1, 1);
            Assert.AreEqual(false, ba2[0, 0]);
            ba2[0, 0] = true;
            Assert.AreEqual(true, ba2[0, 0]);
            ba2[0, 0] = false;
            Assert.AreEqual(false, ba2[0, 0]);
        }
    }
}