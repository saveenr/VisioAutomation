using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.ShapeSheet;
using VA = VisioAutomation;

namespace VisioAutomation_Tests.Core.ShapeSheet
{
    [TestClass]
    public class CellValueLiteralTests
    {

        [TestMethod]
        public void CellValueLiteral_Creation()
        {
            // unitialized means it has no value
            CellValueLiteral c0;
            Assert.IsNull(c0.Value);
            Assert.IsFalse(c0.HasValue);

            var c1 = new CellValueLiteral("FOO");
            Assert.AreEqual("FOO",c1.Value);
            Assert.IsTrue(c1.HasValue);

            var c2 = new CellValueLiteral(123.45);
            Assert.AreEqual("123.45", c2.Value);
            Assert.IsTrue(c2.HasValue);

            var c3 = new CellValueLiteral(12345);
            Assert.AreEqual("12345", c3.Value);
            Assert.IsTrue(c3.HasValue);

            var c4 = new CellValueLiteral(true);
            Assert.AreEqual("TRUE", c4.Value);
            Assert.IsTrue(c4.HasValue);

            var c5 = new CellValueLiteral(false);
            Assert.AreEqual("FALSE", c5.Value);
            Assert.IsTrue(c5.HasValue);

        }

    }
}