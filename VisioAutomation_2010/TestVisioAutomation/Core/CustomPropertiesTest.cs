using Microsoft.VisualStudio.TestTools.UnitTesting;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class CustomPropertiesTest
    {
        [TestMethod]
        public void ValidNames()
        {
            Assert.IsFalse(VA.CustomProperties.CustomPropertyHelper.IsValidCustomPropertyName(null));
            Assert.IsFalse(VA.CustomProperties.CustomPropertyHelper.IsValidCustomPropertyName(""));
            Assert.IsFalse(VA.CustomProperties.CustomPropertyHelper.IsValidCustomPropertyName(" foo "));
            Assert.IsFalse(VA.CustomProperties.CustomPropertyHelper.IsValidCustomPropertyName("foo "));
            Assert.IsFalse(VA.CustomProperties.CustomPropertyHelper.IsValidCustomPropertyName("foo\t"));
            Assert.IsFalse(VA.CustomProperties.CustomPropertyHelper.IsValidCustomPropertyName("fo bar"));
            Assert.IsTrue(VA.CustomProperties.CustomPropertyHelper.IsValidCustomPropertyName("foobar"));
        }
    }
}