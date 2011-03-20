using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.CustomProperties;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class CustomPropertiesTest
    {
        [TestMethod]
        public void ValidNames()
        {
            Assert.IsFalse(CustomPropertyHelper.IsValidCustomPropertyName(null));
            Assert.IsFalse(CustomPropertyHelper.IsValidCustomPropertyName(""));
            Assert.IsFalse(CustomPropertyHelper.IsValidCustomPropertyName(" foo "));
            Assert.IsFalse(CustomPropertyHelper.IsValidCustomPropertyName("foo "));
            Assert.IsFalse(CustomPropertyHelper.IsValidCustomPropertyName("foo\t"));
            Assert.IsFalse(CustomPropertyHelper.IsValidCustomPropertyName("fo bar"));
            Assert.IsTrue(CustomPropertyHelper.IsValidCustomPropertyName("foobar"));
        }
    }
}