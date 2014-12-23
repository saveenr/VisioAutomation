using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VACUSTPROP = VisioAutomation.Shapes.CustomProperties;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class CustomPropertiesTest : VisioAutomationTest
    {
        [TestMethod]
        public void CustomProps_Names()
        {
            Assert.IsFalse(VACUSTPROP.CustomPropertyHelper.IsValidName(null));
            Assert.IsFalse(VACUSTPROP.CustomPropertyHelper.IsValidName(""));
            Assert.IsFalse(VACUSTPROP.CustomPropertyHelper.IsValidName(" foo "));
            Assert.IsFalse(VACUSTPROP.CustomPropertyHelper.IsValidName("foo "));
            Assert.IsFalse(VACUSTPROP.CustomPropertyHelper.IsValidName("foo\t"));
            Assert.IsFalse(VACUSTPROP.CustomPropertyHelper.IsValidName("fo bar"));
            Assert.IsTrue(VACUSTPROP.CustomPropertyHelper.IsValidName("foobar"));
        }

        [TestMethod]
        public void CustomProps_GetSet()
        {
            var page1 = GetNewPage();

            var s1 = page1.DrawRectangle(0,0,1,1);
            s1.Text = "Checking for Custom Properties";

            // A new rectangle should have zero props
            var c0 = VACUSTPROP.CustomPropertyHelper.Get(s1);
            Assert.AreEqual(0,c0.Count);

            // Set one property
            // Notice that the properties some back double-quoted
            VACUSTPROP.CustomPropertyHelper.Set(s1,"PROP1","VAL1");
            var c1 = VACUSTPROP.CustomPropertyHelper.Get(s1);
            Assert.AreEqual(1, c1.Count);
            Assert.IsTrue(c1.ContainsKey("PROP1"));
            Assert.AreEqual("\"VAL1\"",c1["PROP1"].Value.Formula);

            // Add another property
            VACUSTPROP.CustomPropertyHelper.Set(s1, "PROP2", "VAL 2");
            var c2 = VACUSTPROP.CustomPropertyHelper.Get(s1);
            Assert.AreEqual(2, c2.Count);
            Assert.IsTrue(c2.ContainsKey("PROP1"));
            Assert.AreEqual("\"VAL1\"", c2["PROP1"].Value.Formula);
            Assert.IsTrue(c2.ContainsKey("PROP2"));
            Assert.AreEqual("\"VAL 2\"", c2["PROP2"].Value.Formula);

            // Modify the value of the second property
            VACUSTPROP.CustomPropertyHelper.Set(s1, "PROP2", "\"VAL 2 MOD\"");
            var c3 = VACUSTPROP.CustomPropertyHelper.Get(s1);
            Assert.AreEqual(2, c3.Count);
            Assert.IsTrue(c3.ContainsKey("PROP1"));
            Assert.AreEqual("\"VAL1\"", c3["PROP1"].Value.Formula);
            Assert.IsTrue(c3.ContainsKey("PROP2"));
            Assert.AreEqual("\"VAL 2 MOD\"", c3["PROP2"].Value.Formula);
            
            // Now delete all the custom properties
            foreach (string name in c3.Keys)
            {
                VACUSTPROP.CustomPropertyHelper.Delete(s1,name);
            }
            var c4 = VACUSTPROP.CustomPropertyHelper.Get(s1);
            Assert.AreEqual(0, c4.Count);

            var app = this.GetVisioApplication();
            var doc = app.ActiveDocument;
            if (doc != null)
            {
                VA.Documents.DocumentHelper.Close(doc, true);
            }
        }

        [TestMethod]
        public void CustomProps_GetSet2()
        {
            var page1 = GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 1, 1);
            s1.Text = "Checking for Custom Properties";


            var cp1 = new VACUSTPROP.CustomPropertyCells();
            cp1.Ask = "1";
            cp1.Calendar = "0";
            cp1.Format= "1";
            cp1.Invisible = "0";
            cp1.Label= "1";
            cp1.LangId= "0";
            cp1.Prompt= "1";
            cp1.SortKey= "0";
            cp1.Type= "0";
            cp1.Value= "1";

            VACUSTPROP.CustomPropertyHelper.Set(s1, "PROP1", cp1);

            var props1 = VACUSTPROP.CustomPropertyHelper.Get(s1);

            var cp2 = props1["PROP1"];
            Assert.AreEqual("TRUE", cp2.Ask.Formula.Value);
            Assert.AreEqual("0", cp2.Calendar.Formula.Value);
            Assert.AreEqual("\"1\"", cp2.Format.Formula.Value);
            Assert.AreEqual("FALSE", cp2.Invisible.Formula.Value);
            Assert.AreEqual("\"1\"", cp2.Label.Formula.Value);

            Assert.AreEqual("0", cp2.LangId.Formula.Value);
            Assert.AreEqual("\"1\"", cp2.Prompt.Formula.Value);
            Assert.AreEqual("0", cp2.SortKey.Formula.Value);
            Assert.AreEqual("0", cp2.Type.Formula.Value);

            Assert.AreEqual("\"1\"", cp2.Value.Formula.Value);

            var cp3 = new VACUSTPROP.CustomPropertyCells();
            cp3.Ask = "0";
            cp3.Calendar = "2";
            cp3.Format = "0";
            cp3.Invisible = "TRUE";
            cp3.Label = "3";
            cp3.LangId = "2";
            cp3.Prompt = "3";
            cp3.SortKey = "2";
            cp3.Type = "3";
            cp3.Value = "2";

            VACUSTPROP.CustomPropertyHelper.Set(s1,"PROP1",cp3);
            var props2 = VACUSTPROP.CustomPropertyHelper.Get(s1);

            var cp4 = props2["PROP1"];
            Assert.AreEqual("FALSE", cp4.Ask.Formula.Value);
            Assert.AreEqual("2", cp4.Calendar.Formula.Value);
            Assert.AreEqual("\"0\"", cp4.Format.Formula.Value);
            Assert.AreEqual("TRUE", cp4.Invisible.Formula.Value);
            Assert.AreEqual("\"3\"", cp4.Label.Formula.Value);
                                   
            Assert.AreEqual("2", cp4.LangId.Formula.Value);
            Assert.AreEqual("\"3\"", cp4.Prompt.Formula.Value);
            Assert.AreEqual("2", cp4.SortKey.Formula.Value);
            Assert.AreEqual("3", cp4.Type.Formula.Value);
                                   
            Assert.AreEqual("2", cp4.Value.Formula.Value);

            var app = this.GetVisioApplication();
            var doc = app.ActiveDocument;
            if (doc != null)
            {
                VA.Documents.DocumentHelper.Close(doc, true);
            }
        }

    }
}