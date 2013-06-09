using Microsoft.VisualStudio.TestTools.UnitTesting;
using VA=VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class Struct_Sizes
    {
        [TestMethod]
        public void VerifySRCLayout()
        {
            this.SRCSizeIs6Bytes();
            this.Verify_Size_of_instance();
        }

        public void SRCSizeIs6Bytes()
        {
            var c1 = new VA.ShapeSheet.SRC();
            Assert.AreEqual(6, System.Runtime.InteropServices.Marshal.SizeOf(c1));
        }

        public void Verify_Size_of_instance()
        {
            var instance = new VA.ShapeSheet.FormulaLiteral();
            Assert.AreEqual(4, System.Runtime.InteropServices.Marshal.SizeOf(instance));
        }
    }
}