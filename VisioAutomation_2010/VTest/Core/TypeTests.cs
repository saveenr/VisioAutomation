using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VTest.Core
{
    [MUT.TestClass]
    public class TypeTests : Framework.VTest
    {
        [MUT.TestMethod]
        public void VerifySrcSize()
        {
            // Srcs must be 6 bytes
            var c1 = new VisioAutomation.Core.Src();
            int actual_size = System.Runtime.InteropServices.Marshal.SizeOf(c1);
            MUT.Assert.AreEqual(6, actual_size);

            this.VerifyFormulaLiteralSize();
        }

        public void VerifyFormulaLiteralSize()
        {
            // CellValue is a struct holding one managed-string reference, so its
            // marshalled size should equal the pointer size for the running
            // process: 4 on 32-bit, 8 on 64-bit. Hardcoding 4 was a 32-bit-only
            // assumption that broke once the testhost ran 64-bit.

            var instance = new VisioAutomation.Core.CellValue();
            int actual_size = System.Runtime.InteropServices.Marshal.SizeOf(instance);
            MUT.Assert.AreEqual(System.IntPtr.Size, actual_size);
        }

        [MUT.TestMethod]
        public void Construct2DBitArray()
        {
            // check that cols and rows must be > 0
            bool caught = false;
            try
            {
                var ba = new VisioAutomation.Analyzers.BitArray2D(0, 1);
            }
            catch (System.ArgumentOutOfRangeException)
            {
                caught = true;
            }

            if (caught == false)
            {
                MUT.Assert.Fail("Did not catch expected exception");
            }

            caught = false;
            try
            {
                var ba = new VisioAutomation.Analyzers.BitArray2D(1, 0);
            }
            catch (System.ArgumentOutOfRangeException)
            {
                caught = true;
            }

            if (caught == false)
            {
                MUT.Assert.Fail("Did not catch expected exception");
            }

            // Create a 1x1 BitArray
            var ba2 = new VisioAutomation.Analyzers.BitArray2D(1, 1);
            MUT.Assert.AreEqual(false, ba2[0, 0]);
            ba2[0, 0] = true;
            MUT.Assert.AreEqual(true, ba2[0, 0]);
            ba2[0, 0] = false;
            MUT.Assert.AreEqual(false, ba2[0, 0]);
        }
    }
}