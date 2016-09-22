using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Convert = VisioAutomation.Utilities.Convert;

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
            Assert.AreEqual(1, Convert.BoolToShort(true));
            Assert.AreEqual(0, Convert.BoolToShort(false));
            Assert.AreEqual(true, Convert.ShortToBool(-1));
            Assert.AreEqual(true, Convert.ShortToBool(1));
            Assert.AreEqual(false, Convert.ShortToBool(0));
        }

        public void Test_StringToFormulaString()
        {
            bool caught = false;
            try
            {
                var t = Convert.StringToFormulaString(null);
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
            
            Assert.AreEqual("", Convert.StringToFormulaString(string.Empty));
            Assert.AreEqual("\" \"", Convert.StringToFormulaString(" "));
            Assert.AreEqual("\" \"\"foo\"\" \"", Convert.StringToFormulaString(" \"foo\" "));
        }

        public void Test_FormulaStringToString()
        {
            bool caught = false;
            try
            {
                var t = Convert.FormulaStringToString(null);
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

            Assert.AreEqual("", Convert.FormulaStringToString(string.Empty));
            Assert.AreEqual(" ", Convert.FormulaStringToString(" "));
            Assert.AreEqual(" \"foo\" ", Convert.FormulaStringToString(" \"foo\" "));

            Assert.AreEqual("", Convert.FormulaStringToString("\"\""));
            Assert.AreEqual(" ", Convert.FormulaStringToString("\" \""));
            Assert.AreEqual(" \"foo\" ", Convert.FormulaStringToString("\" \"\"foo\"\" \""));

            Assert.AreEqual("=", Convert.FormulaStringToString("="));
            Assert.AreEqual("=1", Convert.FormulaStringToString("=1"));
            Assert.AreEqual("=\"1\"", Convert.FormulaStringToString("=\"1\""));

        }
    }
}