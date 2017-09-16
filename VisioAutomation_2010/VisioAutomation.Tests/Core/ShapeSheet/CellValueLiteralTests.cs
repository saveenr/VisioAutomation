using System.Linq;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.ShapeSheet;

namespace VisioAutomation_Tests.Core.ShapeSheet
{
    [TestClass]
    public class CellValueLiteralTests
    {

        [TestMethod]
        public void CellValueLiteral_Equivalence()
        {
            // unitialized CVTs are equal
            CellValueLiteral c0;
            CellValueLiteral c1;

            Assert.AreEqual(c0,c1);

            // initialized CVTs set to null are equal
            var c2 = new CellValueLiteral(null);
            var c3 = new CellValueLiteral(null);
            Assert.AreEqual(c2, c3);

            // initialized CVTs set to empty strings are equal
            var c4 = new CellValueLiteral(string.Empty);
            var c5 = new CellValueLiteral("");
            Assert.AreEqual(c4, c5);

            // initialized CVTs set to the same strings are equal
            var c6 = new CellValueLiteral("FOO");
            var sb = new StringBuilder();
            sb.Append("F");
            sb.Append("O");
            sb.Append("O");
            var c7 = new CellValueLiteral(sb.ToString());
            Assert.AreEqual(c6, c7);

            // itialized CVTs to different values are not considered equal
            var c8 = new CellValueLiteral("FOO");
            var c9 = new CellValueLiteral("BAR");
            Assert.AreNotEqual(c8, c9);

            var c10 = new CellValueLiteral(null);
            var c11 = new CellValueLiteral("BAR");
            Assert.AreNotEqual(c10, c11);
        }

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