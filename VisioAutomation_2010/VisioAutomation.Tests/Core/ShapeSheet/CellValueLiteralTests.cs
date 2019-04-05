using Microsoft.VisualStudio.TestTools.UnitTesting;
using VASS=VisioAutomation.ShapeSheet;

namespace VisioAutomation_Tests.Core.ShapeSheet
{
    [TestClass]
    public class CellValueLiteralTests
    {

        [TestMethod]
        public void CellValueLiteral_Equivalence()
        {
            // uninitialized CVTs are equal
            VASS.CellValueLiteral c0;
            VASS.CellValueLiteral c1;

            Assert.AreEqual(c0,c1);

            // initialized CVTs set to null are equal
            var c2 = new VASS.CellValueLiteral(null);
            var c3 = new VASS.CellValueLiteral(null);
            Assert.AreEqual(c2, c3);

            // initialized CVTs set to empty strings are equal
            var c4 = new VASS.CellValueLiteral(string.Empty);
            var c5 = new VASS.CellValueLiteral("");
            Assert.AreEqual(c4, c5);

            // initialized CVTs set to the same strings are equal (when the strings aren't interned)
            var c6 = new VASS.CellValueLiteral("FOO");
            var c7 = new VASS.CellValueLiteral(string.Copy("FOO")); // string.Copy avoids string interning
            Assert.AreEqual(c6, c7);

            // initialized CVTs to different values are not considered equal
            var c8 = new VASS.CellValueLiteral("FOO");
            var c9 = new VASS.CellValueLiteral("BAR");
            Assert.AreNotEqual(c8, c9);

            var c10 = new VASS.CellValueLiteral(null);
            var c11 = new VASS.CellValueLiteral("BAR");
            Assert.AreNotEqual(c10, c11);
        }

        [TestMethod]
        public void CellValueLiteral_Creation()
        {
            // unitialized means it has no value
            VASS.CellValueLiteral c0;
            Assert.IsNull(c0.Value);
            Assert.IsFalse(c0.HasValue);

            var c1 = new VASS.CellValueLiteral("FOO");
            Assert.AreEqual("FOO",c1.Value);
            Assert.IsTrue(c1.HasValue);

            var c2 = new VASS.CellValueLiteral(123.45);
            Assert.AreEqual("123.45", c2.Value);
            Assert.IsTrue(c2.HasValue);

            var c3 = new VASS.CellValueLiteral(12345);
            Assert.AreEqual("12345", c3.Value);
            Assert.IsTrue(c3.HasValue);

            var c4 = new VASS.CellValueLiteral(true);
            Assert.AreEqual("TRUE", c4.Value);
            Assert.IsTrue(c4.HasValue);

            var c5 = new VASS.CellValueLiteral(false);
            Assert.AreEqual("FALSE", c5.Value);
            Assert.IsTrue(c5.HasValue);
        }

        [TestMethod]
        public void CellValueLiteral_EncodeValue()
        {
            Assert.AreEqual(null, VASS.CellValueLiteral.EncodeValue(null));
            Assert.AreEqual("", VASS.CellValueLiteral.EncodeValue(""));
            Assert.AreEqual("=1", VASS.CellValueLiteral.EncodeValue("=1"));
            Assert.AreEqual("\"A\"", VASS.CellValueLiteral.EncodeValue("\"A\""));
            Assert.AreEqual("\"A\"", VASS.CellValueLiteral.EncodeValue("A"));
            Assert.AreEqual("\"A\"\"", VASS.CellValueLiteral.EncodeValue("\"A\"\""));
            Assert.AreEqual("\"A\"\"\"", VASS.CellValueLiteral.EncodeValue("A\""));

            Assert.AreEqual(null, VASS.CellValueLiteral.EncodeValue(null,false));
            Assert.AreEqual("", VASS.CellValueLiteral.EncodeValue("", false));
            Assert.AreEqual("=1", VASS.CellValueLiteral.EncodeValue("=1", false));
            Assert.AreEqual("\"A\"", VASS.CellValueLiteral.EncodeValue("\"A\"", false));
            Assert.AreEqual("A", VASS.CellValueLiteral.EncodeValue("A", false));
            Assert.AreEqual("\"A\"\"", VASS.CellValueLiteral.EncodeValue("\"A\"\"", false));
            Assert.AreEqual("A\"", VASS.CellValueLiteral.EncodeValue("A\"", false));
        }
    }
}