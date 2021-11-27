using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VTest.Core.ShapeSheet
{
    [MUT.TestClass]
    public class CellValueLiteralTests
    {

        [MUT.TestMethod]
        public void CellValueLiteral_Equivalence()
        {
            // uninitialized CVTs are equal
            VisioAutomation.Core.CellValue c0;
            VisioAutomation.Core.CellValue c1;

            MUT.Assert.AreEqual(c0,c1);

            // initialized CVTs set to null are equal
            var c2 = new VisioAutomation.Core.CellValue(null);
            var c3 = new VisioAutomation.Core.CellValue(null);
            MUT.Assert.AreEqual(c2, c3);

            // initialized CVTs set to empty strings are equal
            var c4 = new VisioAutomation.Core.CellValue(string.Empty);
            var c5 = new VisioAutomation.Core.CellValue("");
            MUT.Assert.AreEqual(c4, c5);

            // initialized CVTs set to the same strings are equal (when the strings aren't interned)
            var c6 = new VisioAutomation.Core.CellValue("FOO");
            var c7 = new VisioAutomation.Core.CellValue(string.Copy("FOO")); // string.Copy avoids string interning
            MUT.Assert.AreEqual(c6, c7);

            // initialized CVTs to different values are not considered equal
            var c8 = new VisioAutomation.Core.CellValue("FOO");
            var c9 = new VisioAutomation.Core.CellValue("BAR");
            MUT.Assert.AreNotEqual(c8, c9);

            var c10 = new VisioAutomation.Core.CellValue(null);
            var c11 = new VisioAutomation.Core.CellValue("BAR");
            MUT.Assert.AreNotEqual(c10, c11);
        }

        [MUT.TestMethod]
        public void CellValueLiteral_Creation()
        {
            // unitialized means it has no value
            VisioAutomation.Core.CellValue c0;
            MUT.Assert.IsNull(c0.Value);
            MUT.Assert.IsFalse(c0.HasValue);

            var c1 = new VisioAutomation.Core.CellValue("FOO");
            MUT.Assert.AreEqual("FOO",c1.Value);
            MUT.Assert.IsTrue(c1.HasValue);

            var c2 = new VisioAutomation.Core.CellValue(123.45);
            MUT.Assert.AreEqual("123.45", c2.Value);
            MUT.Assert.IsTrue(c2.HasValue);

            var c3 = new VisioAutomation.Core.CellValue(12345);
            MUT.Assert.AreEqual("12345", c3.Value);
            MUT.Assert.IsTrue(c3.HasValue);

            var c4 = new VisioAutomation.Core.CellValue(true);
            MUT.Assert.AreEqual("TRUE", c4.Value);
            MUT.Assert.IsTrue(c4.HasValue);

            var c5 = new VisioAutomation.Core.CellValue(false);
            MUT.Assert.AreEqual("FALSE", c5.Value);
            MUT.Assert.IsTrue(c5.HasValue);
        }

        [MUT.TestMethod]
        public void CellValueLiteral_EncodeValue()
        {
            MUT.Assert.AreEqual(null, VisioAutomation.Core.CellValue.EncodeValue(null));
            MUT.Assert.AreEqual("", VisioAutomation.Core.CellValue.EncodeValue(""));
            MUT.Assert.AreEqual("=1", VisioAutomation.Core.CellValue.EncodeValue("=1"));
            MUT.Assert.AreEqual("\"A\"", VisioAutomation.Core.CellValue.EncodeValue("\"A\""));
            MUT.Assert.AreEqual("\"A\"", VisioAutomation.Core.CellValue.EncodeValue("A"));
            MUT.Assert.AreEqual("\"A\"\"", VisioAutomation.Core.CellValue.EncodeValue("\"A\"\""));
            MUT.Assert.AreEqual("\"A\"\"\"", VisioAutomation.Core.CellValue.EncodeValue("A\""));

            MUT.Assert.AreEqual(null, VisioAutomation.Core.CellValue.EncodeValue(null,false));
            MUT.Assert.AreEqual("", VisioAutomation.Core.CellValue.EncodeValue("", false));
            MUT.Assert.AreEqual("=1", VisioAutomation.Core.CellValue.EncodeValue("=1", false));
            MUT.Assert.AreEqual("\"A\"", VisioAutomation.Core.CellValue.EncodeValue("\"A\"", false));
            MUT.Assert.AreEqual("A", VisioAutomation.Core.CellValue.EncodeValue("A", false));
            MUT.Assert.AreEqual("\"A\"\"", VisioAutomation.Core.CellValue.EncodeValue("\"A\"\"", false));
            MUT.Assert.AreEqual("A\"", VisioAutomation.Core.CellValue.EncodeValue("A\"", false));
        }
    }
}