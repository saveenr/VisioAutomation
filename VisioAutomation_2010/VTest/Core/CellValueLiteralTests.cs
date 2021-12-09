namespace VTest.Core
{
    [Microsoft.VisualStudio.TestTools.UnitTesting.TestClass]
    public class CellValueLiteralTests
    {

        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void CellValueLiteral_Equivalence()
        {
            // uninitialized CVTs are equal
            VisioAutomation.Core.CellValue c0;
            VisioAutomation.Core.CellValue c1;

            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(c0,c1);

            // initialized CVTs set to null are equal
            var c2 = new VisioAutomation.Core.CellValue(null);
            var c3 = new VisioAutomation.Core.CellValue(null);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(c2, c3);

            // initialized CVTs set to empty strings are equal
            var c4 = new VisioAutomation.Core.CellValue(string.Empty);
            var c5 = new VisioAutomation.Core.CellValue("");
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(c4, c5);

            // initialized CVTs set to the same strings are equal (when the strings aren't interned)
            var c6 = new VisioAutomation.Core.CellValue("FOO");
            var c7 = new VisioAutomation.Core.CellValue(string.Copy("FOO")); // string.Copy avoids string interning
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(c6, c7);

            // initialized CVTs to different values are not considered equal
            var c8 = new VisioAutomation.Core.CellValue("FOO");
            var c9 = new VisioAutomation.Core.CellValue("BAR");
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreNotEqual(c8, c9);

            var c10 = new VisioAutomation.Core.CellValue(null);
            var c11 = new VisioAutomation.Core.CellValue("BAR");
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreNotEqual(c10, c11);
        }

        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void CellValueLiteral_Creation()
        {
            // unitialized means it has no value
            VisioAutomation.Core.CellValue c0;
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsNull(c0.Value);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsFalse(c0.HasValue);

            var c1 = new VisioAutomation.Core.CellValue("FOO");
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("FOO",c1.Value);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsTrue(c1.HasValue);

            var c2 = new VisioAutomation.Core.CellValue(123.45);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("123.45", c2.Value);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsTrue(c2.HasValue);

            var c3 = new VisioAutomation.Core.CellValue(12345);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("12345", c3.Value);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsTrue(c3.HasValue);

            var c4 = new VisioAutomation.Core.CellValue(true);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("TRUE", c4.Value);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsTrue(c4.HasValue);

            var c5 = new VisioAutomation.Core.CellValue(false);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("FALSE", c5.Value);
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.IsTrue(c5.HasValue);
        }

        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void CellValueLiteral_EncodeValue()
        {
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(null, VisioAutomation.Core.CellValue.EncodeValue(null));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("", VisioAutomation.Core.CellValue.EncodeValue(""));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("=1", VisioAutomation.Core.CellValue.EncodeValue("=1"));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("\"A\"", VisioAutomation.Core.CellValue.EncodeValue("\"A\""));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("\"A\"", VisioAutomation.Core.CellValue.EncodeValue("A"));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("\"A\"\"", VisioAutomation.Core.CellValue.EncodeValue("\"A\"\""));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("\"A\"\"\"", VisioAutomation.Core.CellValue.EncodeValue("A\""));

            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual(null, VisioAutomation.Core.CellValue.EncodeValue(null,false));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("", VisioAutomation.Core.CellValue.EncodeValue("", false));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("=1", VisioAutomation.Core.CellValue.EncodeValue("=1", false));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("\"A\"", VisioAutomation.Core.CellValue.EncodeValue("\"A\"", false));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("A", VisioAutomation.Core.CellValue.EncodeValue("A", false));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("\"A\"\"", VisioAutomation.Core.CellValue.EncodeValue("\"A\"\"", false));
            Microsoft.VisualStudio.TestTools.UnitTesting.Assert.AreEqual("A\"", VisioAutomation.Core.CellValue.EncodeValue("A\"", false));
        }
    }
}