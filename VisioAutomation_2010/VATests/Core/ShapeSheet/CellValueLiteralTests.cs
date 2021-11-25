
namespace VisioAutomation_Tests.Core.ShapeSheet
{
    [TestClass]
    public class CellValueLiteralTests
    {

        [TestMethod]
        public void CellValueLiteral_Equivalence()
        {
            // uninitialized CVTs are equal
            VASS.CellValue c0;
            VASS.CellValue c1;

            Assert.AreEqual(c0,c1);

            // initialized CVTs set to null are equal
            var c2 = new VASS.CellValue(null);
            var c3 = new VASS.CellValue(null);
            Assert.AreEqual(c2, c3);

            // initialized CVTs set to empty strings are equal
            var c4 = new VASS.CellValue(string.Empty);
            var c5 = new VASS.CellValue("");
            Assert.AreEqual(c4, c5);

            // initialized CVTs set to the same strings are equal (when the strings aren't interned)
            var c6 = new VASS.CellValue("FOO");
            var c7 = new VASS.CellValue(string.Copy("FOO")); // string.Copy avoids string interning
            Assert.AreEqual(c6, c7);

            // initialized CVTs to different values are not considered equal
            var c8 = new VASS.CellValue("FOO");
            var c9 = new VASS.CellValue("BAR");
            Assert.AreNotEqual(c8, c9);

            var c10 = new VASS.CellValue(null);
            var c11 = new VASS.CellValue("BAR");
            Assert.AreNotEqual(c10, c11);
        }

        [TestMethod]
        public void CellValueLiteral_Creation()
        {
            // unitialized means it has no value
            VASS.CellValue c0;
            Assert.IsNull(c0.Value);
            Assert.IsFalse(c0.HasValue);

            var c1 = new VASS.CellValue("FOO");
            Assert.AreEqual("FOO",c1.Value);
            Assert.IsTrue(c1.HasValue);

            var c2 = new VASS.CellValue(123.45);
            Assert.AreEqual("123.45", c2.Value);
            Assert.IsTrue(c2.HasValue);

            var c3 = new VASS.CellValue(12345);
            Assert.AreEqual("12345", c3.Value);
            Assert.IsTrue(c3.HasValue);

            var c4 = new VASS.CellValue(true);
            Assert.AreEqual("TRUE", c4.Value);
            Assert.IsTrue(c4.HasValue);

            var c5 = new VASS.CellValue(false);
            Assert.AreEqual("FALSE", c5.Value);
            Assert.IsTrue(c5.HasValue);
        }

        [TestMethod]
        public void CellValueLiteral_EncodeValue()
        {
            Assert.AreEqual(null, VASS.CellValue.EncodeValue(null));
            Assert.AreEqual("", VASS.CellValue.EncodeValue(""));
            Assert.AreEqual("=1", VASS.CellValue.EncodeValue("=1"));
            Assert.AreEqual("\"A\"", VASS.CellValue.EncodeValue("\"A\""));
            Assert.AreEqual("\"A\"", VASS.CellValue.EncodeValue("A"));
            Assert.AreEqual("\"A\"\"", VASS.CellValue.EncodeValue("\"A\"\""));
            Assert.AreEqual("\"A\"\"\"", VASS.CellValue.EncodeValue("A\""));

            Assert.AreEqual(null, VASS.CellValue.EncodeValue(null,false));
            Assert.AreEqual("", VASS.CellValue.EncodeValue("", false));
            Assert.AreEqual("=1", VASS.CellValue.EncodeValue("=1", false));
            Assert.AreEqual("\"A\"", VASS.CellValue.EncodeValue("\"A\"", false));
            Assert.AreEqual("A", VASS.CellValue.EncodeValue("A", false));
            Assert.AreEqual("\"A\"\"", VASS.CellValue.EncodeValue("\"A\"\"", false));
            Assert.AreEqual("A\"", VASS.CellValue.EncodeValue("A\"", false));
        }
    }
}