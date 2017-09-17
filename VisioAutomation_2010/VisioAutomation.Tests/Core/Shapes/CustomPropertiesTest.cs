using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using VisioAutomation.Shapes;
using VisioAutomation.ShapeSheet;

namespace VisioAutomation_Tests.Core.Shapes
{
    [TestClass]
    public class CustomPropertiesTest : VisioAutomationTest
    {
        [TestMethod]
        public void CustomProps_Names()
        {
            Assert.IsFalse(CustomPropertyHelper.IsValidName(null));
            Assert.IsFalse(CustomPropertyHelper.IsValidName(string.Empty));
            Assert.IsFalse(CustomPropertyHelper.IsValidName(" foo "));
            Assert.IsFalse(CustomPropertyHelper.IsValidName("foo "));
            Assert.IsFalse(CustomPropertyHelper.IsValidName("foo\t"));
            Assert.IsFalse(CustomPropertyHelper.IsValidName("fo bar"));
            Assert.IsTrue(CustomPropertyHelper.IsValidName("foobar"));
        }

        [TestMethod]
        public void CustomProps_GetSet()
        {
            var page1 = this.GetNewPage();

            var s1 = page1.DrawRectangle(0,0,1,1);
            s1.Text = "Checking for Custom Properties";

            // A new rectangle should have zero props
            var c0 = CustomPropertyHelper.GetCells(s1, CellValueType.Formula);
            Assert.AreEqual(0,c0.Count);

            // Set one property
            // Notice that the properties some back double-quoted
            CustomPropertyHelper.Set(s1,"PROP1","VAL1");
            var c1 = CustomPropertyHelper.GetCells(s1, CellValueType.Formula);
            Assert.AreEqual(1, c1.Count);
            Assert.IsTrue(c1.ContainsKey("PROP1"));
            Assert.AreEqual("\"VAL1\"",c1["PROP1"].Value.Value);

            // Add another property
            CustomPropertyHelper.Set(s1, "PROP2", "VAL 2");
            var c2 = CustomPropertyHelper.GetCells(s1, CellValueType.Formula);
            Assert.AreEqual(2, c2.Count);
            Assert.IsTrue(c2.ContainsKey("PROP1"));
            Assert.AreEqual("\"VAL1\"", c2["PROP1"].Value.Value);
            Assert.IsTrue(c2.ContainsKey("PROP2"));
            Assert.AreEqual("\"VAL 2\"", c2["PROP2"].Value.Value);

            // Modify the value of the second property
            CustomPropertyHelper.Set(s1, "PROP2", "\"VAL 2 MOD\"");
            var c3 = CustomPropertyHelper.GetCells(s1, CellValueType.Formula);
            Assert.AreEqual(2, c3.Count);
            Assert.IsTrue(c3.ContainsKey("PROP1"));
            Assert.AreEqual("\"VAL1\"", c3["PROP1"].Value.Value);
            Assert.IsTrue(c3.ContainsKey("PROP2"));
            Assert.AreEqual("\"VAL 2 MOD\"", c3["PROP2"].Value.Value);
            
            // Now delete all the custom properties
            foreach (string name in c3.Keys)
            {
                CustomPropertyHelper.Delete(s1,name);
            }
            var c4 = CustomPropertyHelper.GetCells(s1, CellValueType.Formula);
            Assert.AreEqual(0, c4.Count);

            var app = this.GetVisioApplication();
            var doc = app.ActiveDocument;
            if (doc != null)
            {
               doc.Close(true);
            }
        }

        [TestMethod]
        public void CustomProps_GetSet2()
        {
            var page1 = this.GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 1, 1);
            s1.Text = "Checking for Custom Properties";


            var cp1 = new CustomPropertyCells();
            cp1.Ask = "1";
            cp1.Calendar = "0";
            cp1.Format= CustomPropertyHelper.EncodeFormula("1",true);
            cp1.Invisible = "0";
            cp1.Label= CustomPropertyHelper.EncodeFormula("1", true);
            cp1.LangID= "0";
            cp1.Prompt= "1";
            cp1.SortKey= "0";
            cp1.Type= "0";
            cp1.Value= "1";

            CustomPropertyHelper.Set(s1, "PROP1", cp1);

            var props1 = CustomPropertyHelper.GetCells(s1, CellValueType.Formula);

            var cp2 = props1["PROP1"];
            Assert.AreEqual("TRUE", cp2.Ask.Value);
            Assert.AreEqual("0", cp2.Calendar.Value);
            Assert.AreEqual("\"1\"", cp2.Format.Value);
            Assert.AreEqual("FALSE", cp2.Invisible.Value);
            Assert.AreEqual("\"1\"", cp2.Label.Value);

            Assert.AreEqual("0", cp2.LangID.Value);
            Assert.AreEqual("\"1\"", cp2.Prompt.Value);
            Assert.AreEqual("0", cp2.SortKey.Value);
            Assert.AreEqual("0", cp2.Type.Value);

            Assert.AreEqual("\"1\"", cp2.Value.Value);

            var cp3 = new CustomPropertyCells();
            cp3.Ask = "0";
            cp3.Calendar = "2";
            cp3.Format = CustomPropertyHelper.EncodeFormula("0",true);
            cp3.Invisible = "TRUE";
            cp3.Label = CustomPropertyHelper.EncodeFormula("3",true);
            cp3.LangID = "2";
            cp3.Prompt = "3";
            cp3.SortKey = "2";
            cp3.Type = "3";
            cp3.Value = "2";

            CustomPropertyHelper.Set(s1,"PROP1",cp3);
            var props2 = CustomPropertyHelper.GetCells(s1, CellValueType.Formula);

            var cp4 = props2["PROP1"];
            Assert.AreEqual("FALSE", cp4.Ask.Value);
            Assert.AreEqual("2", cp4.Calendar.Value);
            Assert.AreEqual("\"0\"", cp4.Format.Value);
            Assert.AreEqual("TRUE", cp4.Invisible.Value);
            Assert.AreEqual("\"3\"", cp4.Label.Value);
                                   
            Assert.AreEqual("2", cp4.LangID.Value);
            Assert.AreEqual("\"3\"", cp4.Prompt.Value);
            Assert.AreEqual("2", cp4.SortKey.Value);
            Assert.AreEqual("3", cp4.Type.Value);
                                   
            Assert.AreEqual("2", cp4.Value.Value);

            var app = this.GetVisioApplication();
            var doc = app.ActiveDocument;
            if (doc != null)
            {
                doc.Close(true);
            }
        }

    }
}