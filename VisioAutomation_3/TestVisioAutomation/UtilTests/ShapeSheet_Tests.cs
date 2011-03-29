using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using IVisio=Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class ShapeSheet_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void TestBoolToShortConversion()
        {
            Assert.AreEqual(1, VisioAutomation.Convert.BoolToShort(true));
            Assert.AreEqual(0, VisioAutomation.Convert.BoolToShort(false));
            Assert.AreEqual(true, VisioAutomation.Convert.ShortToBool(-1));
            Assert.AreEqual(true, VisioAutomation.Convert.ShortToBool(1));
            Assert.AreEqual(false, VisioAutomation.Convert.ShortToBool(0));
        }

        [TestMethod]
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
            
            Assert.AreEqual("\"\"", VisioAutomation.Convert.StringToFormulaString(string.Empty));
            Assert.AreEqual("\" \"", VisioAutomation.Convert.StringToFormulaString(" "));
            Assert.AreEqual("\" \"\"foo\"\" \"", VisioAutomation.Convert.StringToFormulaString(" \"foo\" "));
        }
    }
}