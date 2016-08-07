using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VA = VisioAutomation;

namespace VisioAutomation_Tests.Models
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
            Assert.AreEqual(1, VA.Convert.BoolToShort(true));
            Assert.AreEqual(0, VA.Convert.BoolToShort(false));
            Assert.AreEqual(true, VA.Convert.ShortToBool(-1));
            Assert.AreEqual(true, VA.Convert.ShortToBool(1));
            Assert.AreEqual(false, VA.Convert.ShortToBool(0));
        }

        public void Test_StringToFormulaString()
        {
            bool caught = false;
            try
            {
                var t = VA.Convert.StringToFormulaString(null);
            }
            catch (ArgumentNullException)
            {
                // this is expected
                caught = true;
            }

            if (!caught)
            {
                Assert.Fail("Did not throw expected exception");
            }
            
            Assert.AreEqual("", VA.Convert.StringToFormulaString(string.Empty));
            Assert.AreEqual("\" \"", VA.Convert.StringToFormulaString(" "));
            Assert.AreEqual("\" \"\"foo\"\" \"", VA.Convert.StringToFormulaString(" \"foo\" "));
        }

        public void Test_FormulaStringToString()
        {
            bool caught = false;
            try
            {
                var t = VA.Convert.FormulaStringToString(null);
            }
            catch (ArgumentNullException)
            {
                // this is expected
                caught = true;
            }

            if (!caught)
            {
                Assert.Fail("Did not throw expected exception");
            }

            Assert.AreEqual("", VA.Convert.FormulaStringToString(string.Empty));
            Assert.AreEqual(" ", VA.Convert.FormulaStringToString(" "));
            Assert.AreEqual(" \"foo\" ", VA.Convert.FormulaStringToString(" \"foo\" "));

            Assert.AreEqual("", VA.Convert.FormulaStringToString("\"\""));
            Assert.AreEqual(" ", VA.Convert.FormulaStringToString("\" \""));
            Assert.AreEqual(" \"foo\" ", VA.Convert.FormulaStringToString("\" \"\"foo\"\" \""));

            Assert.AreEqual("=", VA.Convert.FormulaStringToString("="));
            Assert.AreEqual("=1", VA.Convert.FormulaStringToString("=1"));
            Assert.AreEqual("=\"1\"", VA.Convert.FormulaStringToString("=\"1\""));

        }
    }
}