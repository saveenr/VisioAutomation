using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VTest.Core.Internal
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
            // A FormulaLiteral only has a reference to a string
            // so it should be as big as a reference to a string

            var instance = new VisioAutomation.Core.CellValue();
            int actual_size = System.Runtime.InteropServices.Marshal.SizeOf(instance);
            MUT.Assert.AreEqual(4, actual_size);
        }

        [MUT.TestMethod]
        public void Construct2DBitArray()
        {
            // check that cols and rows must be > 0
            bool caught = false;
            try
            {
                var ba = new VisioAutomation.DocumentAnalysis.BitArray2D(0, 1);
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
                var ba = new VisioAutomation.DocumentAnalysis.BitArray2D(1, 0);
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
            var ba2 = new VisioAutomation.DocumentAnalysis.BitArray2D(1, 1);
            MUT.Assert.AreEqual(false, ba2[0, 0]);
            ba2[0, 0] = true;
            MUT.Assert.AreEqual(true, ba2[0, 0]);
            ba2[0, 0] = false;
            MUT.Assert.AreEqual(false, ba2[0, 0]);
        }
    }
}