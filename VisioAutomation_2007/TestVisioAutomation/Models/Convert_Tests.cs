using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using IVisio=Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class Convert_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void Convert_TestConversions()
        {
            this.TestBoolToShortConversion();
            this.Test_FormulaStringToString();
            this.Test_StringToFormulaString();
        }

        public void TestBoolToShortConversion()
        {
            Assert.AreEqual(1, VisioAutomation.Convert.BoolToShort(true));
            Assert.AreEqual(0, VisioAutomation.Convert.BoolToShort(false));
            Assert.AreEqual(true, VisioAutomation.Convert.ShortToBool(-1));
            Assert.AreEqual(true, VisioAutomation.Convert.ShortToBool(1));
            Assert.AreEqual(false, VisioAutomation.Convert.ShortToBool(0));
        }

        public void Test_StringToFormulaString()
        {
            bool caught = false;
            try
            {
                var t = VisioAutomation.Convert.StringToFormulaString(null);
            }
            catch (System.ArgumentNullException)
            {
                // this is expected
                caught = true;
            }

            if (!caught)
            {
                Assert.Fail("Did not throw expected exception");
            }
            
            Assert.AreEqual("", VisioAutomation.Convert.StringToFormulaString(string.Empty));
            Assert.AreEqual("\" \"", VisioAutomation.Convert.StringToFormulaString(" "));
            Assert.AreEqual("\" \"\"foo\"\" \"", VisioAutomation.Convert.StringToFormulaString(" \"foo\" "));
        }

        public void Test_FormulaStringToString()
        {
            bool caught = false;
            try
            {
                var t = VisioAutomation.Convert.FormulaStringToString(null);
            }
            catch (System.ArgumentNullException)
            {
                // this is expected
                caught = true;
            }

            if (!caught)
            {
                Assert.Fail("Did not throw expected exception");
            }

            Assert.AreEqual("", VisioAutomation.Convert.FormulaStringToString(string.Empty));
            Assert.AreEqual(" ", VisioAutomation.Convert.FormulaStringToString(" "));
            Assert.AreEqual(" \"foo\" ", VisioAutomation.Convert.FormulaStringToString(" \"foo\" "));

            Assert.AreEqual("", VisioAutomation.Convert.FormulaStringToString("\"\""));
            Assert.AreEqual(" ", VisioAutomation.Convert.FormulaStringToString("\" \""));
            Assert.AreEqual(" \"foo\" ", VisioAutomation.Convert.FormulaStringToString("\" \"\"foo\"\" \""));

            Assert.AreEqual("=", VisioAutomation.Convert.FormulaStringToString("="));
            Assert.AreEqual("=1", VisioAutomation.Convert.FormulaStringToString("=1"));
            Assert.AreEqual("=\"1\"", VisioAutomation.Convert.FormulaStringToString("=\"1\""));

        }
    }
}