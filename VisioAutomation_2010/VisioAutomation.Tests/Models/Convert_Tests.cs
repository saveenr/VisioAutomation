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
            this.Test_FormulaStringToString();
            this.Test_StringToFormulaString();
        }

        public void Test_StringToFormulaString()
        {
            bool caught = false;
            try
            {
                var t = Convert.FormulaEncodeSmart(null);
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
            
            Assert.AreEqual("", Convert.FormulaEncodeSmart(string.Empty));
            Assert.AreEqual("\" \"", Convert.FormulaEncodeSmart(" "));
            Assert.AreEqual("\" \"\"foo\"\" \"", Convert.FormulaEncodeSmart(" \"foo\" "));
        }

        public void Test_FormulaStringToString()
        {
            bool caught = false;
            try
            {
                var t = Convert.FormulaDecode(null);
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

            Assert.AreEqual("", Convert.FormulaDecode(string.Empty));
            Assert.AreEqual(" ", Convert.FormulaDecode(" "));
            Assert.AreEqual(" \"foo\" ", Convert.FormulaDecode(" \"foo\" "));

            Assert.AreEqual("", Convert.FormulaDecode("\"\""));
            Assert.AreEqual(" ", Convert.FormulaDecode("\" \""));
            Assert.AreEqual(" \"foo\" ", Convert.FormulaDecode("\" \"\"foo\"\" \""));

            Assert.AreEqual("=", Convert.FormulaDecode("="));
            Assert.AreEqual("=1", Convert.FormulaDecode("=1"));
            Assert.AreEqual("=\"1\"", Convert.FormulaDecode("=\"1\""));

        }
    }
}