using System.Linq;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Shapes.CustomProperties;
using VA = VisioAutomation;
using VisioAutomation.Extensions;

namespace TestVisioAutomation
{
    [TestClass]
    public class CustomPropertiesTest : VisioAutomationTest
    {
        [TestMethod]
        public void CustomProps_Names()
        {
            Assert.IsFalse(CustomPropertyHelper.IsValidName(null));
            Assert.IsFalse(CustomPropertyHelper.IsValidName(""));
            Assert.IsFalse(CustomPropertyHelper.IsValidName(" foo "));
            Assert.IsFalse(CustomPropertyHelper.IsValidName("foo "));
            Assert.IsFalse(CustomPropertyHelper.IsValidName("foo\t"));
            Assert.IsFalse(CustomPropertyHelper.IsValidName("fo bar"));
            Assert.IsTrue(CustomPropertyHelper.IsValidName("foobar"));
        }

        [TestMethod]
        public void CustomProps_GetSet()
        {
            var page1 = GetNewPage();

            var s1 = page1.DrawRectangle(0,0,1,1);
            s1.Text = "Checking for Custom Properties";

            // A new rectangle should have zero props
            var c0 = CustomPropertyHelper.Get(s1);
            Assert.AreEqual(0,c0.Count);

            // Set one property
            // Notice that the properties some back double-quoted
            CustomPropertyHelper.Set(s1,"PROP1","VAL1");
            var c1 = CustomPropertyHelper.Get(s1);
            Assert.AreEqual(1, c1.Count);
            Assert.IsTrue(c1.ContainsKey("PROP1"));
            Assert.AreEqual("\"VAL1\"",c1["PROP1"].Value.Formula);

            // Add another property
            CustomPropertyHelper.Set(s1, "PROP2", "VAL 2");
            var c2 = CustomPropertyHelper.Get(s1);
            Assert.AreEqual(2, c2.Count);
            Assert.IsTrue(c2.ContainsKey("PROP1"));
            Assert.AreEqual("\"VAL1\"", c2["PROP1"].Value.Formula);
            Assert.IsTrue(c2.ContainsKey("PROP2"));
            Assert.AreEqual("\"VAL 2\"", c2["PROP2"].Value.Formula);

            // Modify the value of the second property
            CustomPropertyHelper.Set(s1, "PROP2", "\"VAL 2 MOD\"");
            var c3 = CustomPropertyHelper.Get(s1);
            Assert.AreEqual(2, c3.Count);
            Assert.IsTrue(c3.ContainsKey("PROP1"));
            Assert.AreEqual("\"VAL1\"", c3["PROP1"].Value.Formula);
            Assert.IsTrue(c3.ContainsKey("PROP2"));
            Assert.AreEqual("\"VAL 2 MOD\"", c3["PROP2"].Value.Formula);
            
            // Now delete all the custom properties
            foreach (string name in c3.Keys)
            {
                CustomPropertyHelper.Delete(s1,name);
            }
            var c4 = CustomPropertyHelper.Get(s1);
            Assert.AreEqual(0, c4.Count);

            var app = this.GetVisioApplication();
            var doc = app.ActiveDocument;
            if (doc != null)
            {
                VA.Documents.DocumentHelper.Close(doc, true);
            }
        }
    }
}