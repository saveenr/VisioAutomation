using Microsoft.VisualStudio.TestTools.UnitTesting;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class CustomPropertiesTest : VisioAutomationTest
    {
        [TestMethod]
        public void CheckCustomPropertyNames()
        {
            Assert.IsFalse(VA.CustomProperties.CustomPropertyHelper.IsValidCustomPropertyName(null));
            Assert.IsFalse(VA.CustomProperties.CustomPropertyHelper.IsValidCustomPropertyName(""));
            Assert.IsFalse(VA.CustomProperties.CustomPropertyHelper.IsValidCustomPropertyName(" foo "));
            Assert.IsFalse(VA.CustomProperties.CustomPropertyHelper.IsValidCustomPropertyName("foo "));
            Assert.IsFalse(VA.CustomProperties.CustomPropertyHelper.IsValidCustomPropertyName("foo\t"));
            Assert.IsFalse(VA.CustomProperties.CustomPropertyHelper.IsValidCustomPropertyName("fo bar"));
            Assert.IsTrue(VA.CustomProperties.CustomPropertyHelper.IsValidCustomPropertyName("foobar"));
        }

        [TestMethod]
        public void GetSetCustomProps()
        {
            var page1 = GetNewPage();

            var s1 = page1.DrawRectangle(0,0,1,1);
            s1.Text = "Checking for Custom Properties";

            // A new rectangle should have zero props
            var c0 = VA.CustomProperties.CustomPropertyHelper.GetCustomProperties(s1);
            Assert.AreEqual(0,c0.Count);

            // Set one property
            // Notice that the properties some back double-quoted
            VA.CustomProperties.CustomPropertyHelper.SetCustomProperty(s1,"PROP1","VAL1");
            var c1 = VA.CustomProperties.CustomPropertyHelper.GetCustomProperties(s1);
            Assert.AreEqual(1, c1.Count);
            Assert.IsTrue(c1.ContainsKey("PROP1"));
            Assert.AreEqual("\"VAL1\"",c1["PROP1"].Value.Formula);

            // Add another property
            VA.CustomProperties.CustomPropertyHelper.SetCustomProperty(s1, "PROP2", "VAL 2");
            var c2 = VA.CustomProperties.CustomPropertyHelper.GetCustomProperties(s1);
            Assert.AreEqual(2, c2.Count);
            Assert.IsTrue(c2.ContainsKey("PROP1"));
            Assert.AreEqual("\"VAL1\"", c2["PROP1"].Value.Formula);
            Assert.IsTrue(c2.ContainsKey("PROP2"));
            Assert.AreEqual("\"VAL 2\"", c2["PROP2"].Value.Formula);

            // Modify the value of the second property
            VA.CustomProperties.CustomPropertyHelper.UpdateCustomProperty(s1, "PROP2", "\"VAL 2 MOD\"");
            var c3 = VA.CustomProperties.CustomPropertyHelper.GetCustomProperties(s1);
            Assert.AreEqual(2, c3.Count);
            Assert.IsTrue(c3.ContainsKey("PROP1"));
            Assert.AreEqual("\"VAL1\"", c3["PROP1"].Value.Formula);
            Assert.IsTrue(c3.ContainsKey("PROP2"));
            Assert.AreEqual("\"VAL 2 MOD\"", c3["PROP2"].Value.Formula);
            
            // Now delete all the custom properties
            foreach (string name in c3.Keys)
            {
                VA.CustomProperties.CustomPropertyHelper.DeleteCustomProperty(s1,name);
            }
            var c4 = VA.CustomProperties.CustomPropertyHelper.GetCustomProperties(s1);
            Assert.AreEqual(0, c4.Count);           
        }
    }
}