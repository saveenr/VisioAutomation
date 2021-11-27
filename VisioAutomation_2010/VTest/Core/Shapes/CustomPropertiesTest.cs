using System.Globalization;
using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using VisioAutomation.Shapes;

namespace VTest.Core.Shapes
{
    [MUT.TestClass]
    public class CustomPropertiesTest : VisioAutomationTest
    {
        [MUT.TestMethod]
        public void CustomProps_Names()
        {
            MUT.Assert.IsFalse(CustomPropertyHelper.IsValidName(null));
            MUT.Assert.IsFalse(CustomPropertyHelper.IsValidName(string.Empty));
            MUT.Assert.IsFalse(CustomPropertyHelper.IsValidName(" foo "));
            MUT.Assert.IsFalse(CustomPropertyHelper.IsValidName("foo "));
            MUT.Assert.IsFalse(CustomPropertyHelper.IsValidName("foo\t"));
            MUT.Assert.IsFalse(CustomPropertyHelper.IsValidName("fo bar"));
            MUT.Assert.IsTrue(CustomPropertyHelper.IsValidName("foobar"));
        }

        [MUT.TestMethod]
        public void SimpleCP()
        {
            var page1 = this.GetNewPage();

            // Draw a shape
            var s1 = page1.DrawRectangle(1, 1, 4, 3);

            int cp_type = 0; // string type

            // Set some properties on it
            CustomPropertyHelper.Set(s1, "FOO1", "\"BAR1\"", cp_type);
            CustomPropertyHelper.Set(s1, "FOO2", "\"BAR2\"", cp_type);
            CustomPropertyHelper.Set(s1, "FOO3", "\"BAR3\"", cp_type);

            // Delete one of those properties
            CustomPropertyHelper.Delete(s1, "FOO2");

            // Set the value of an existing properties
            CustomPropertyHelper.Set(s1, "FOO3", "\"BAR3updated\"", cp_type);

            // retrieve all the properties
            var props = CustomPropertyHelper.GetDictionary(s1, VisioAutomation.Core.CellValueType.Formula);

            var cp_foo1 = props["FOO1"];
            // var cp_foo2 = props["FOO2"]; there is no prop called FOO2
            var cp_foo3 = props["FOO3"];

            var app = this.GetVisioApplication();
            var doc = app.ActiveDocument;
            if (doc != null)
            {
                doc.Close(true);
            }
        }

        [MUT.TestMethod]
        public void CustomProps_CRUD()
        {
            var page1 = this.GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 1, 1);
            s1.Text = "Checking for Custom Properties";

            // A new rectangle should have zero props
            var c0 = CustomPropertyHelper.GetDictionary(s1, VisioAutomation.Core.CellValueType.Formula);
            MUT.Assert.AreEqual(0, c0.Count);


            int cp_type = 0; // 0 for string

            // Set one property
            // Notice that the properties some back double-quoted
            CustomPropertyHelper.Set(s1, "PROP1", "\"VAL1\"", cp_type);

            var c1 = CustomPropertyHelper.GetDictionary(s1, VisioAutomation.Core.CellValueType.Formula);

            MUT.Assert.AreEqual(1, c1.Count);
            MUT.Assert.IsTrue(c1.ContainsKey("PROP1"));
            MUT.Assert.AreEqual("\"VAL1\"", c1["PROP1"].Value.Value);

            // Add another property
            CustomPropertyHelper.Set(s1, "PROP2", "\"VAL 2\"", cp_type);
            var c2 = CustomPropertyHelper.GetDictionary(s1, VisioAutomation.Core.CellValueType.Formula);

            MUT.Assert.AreEqual(2, c2.Count);
            MUT.Assert.IsTrue(c2.ContainsKey("PROP1"));
            MUT.Assert.AreEqual("\"VAL1\"", c2["PROP1"].Value.Value);
            MUT.Assert.IsTrue(c2.ContainsKey("PROP2"));
            MUT.Assert.AreEqual("\"VAL 2\"", c2["PROP2"].Value.Value);

            // Modify the value of the second property
            CustomPropertyHelper.Set(s1, "PROP2", "\"VAL 2 MOD\"", cp_type);
            var c3 = CustomPropertyHelper.GetDictionary(s1, VisioAutomation.Core.CellValueType.Formula);
  
            MUT.Assert.AreEqual(2, c3.Count);
            MUT.Assert.IsTrue(c3.ContainsKey("PROP1"));
            MUT.Assert.AreEqual("\"VAL1\"", c3["PROP1"].Value.Value);
            MUT.Assert.IsTrue(c3.ContainsKey("PROP2"));
            MUT.Assert.AreEqual("\"VAL 2 MOD\"", c3["PROP2"].Value.Value);

            // Now delete all the custom properties
            foreach (string name in c3.Keys)
            {
                CustomPropertyHelper.Delete(s1, name);
            }

            var c4 = CustomPropertyHelper.GetDictionary(s1, VisioAutomation.Core.CellValueType.Formula);


            MUT.Assert.AreEqual(0, c4.Count);

            var app = this.GetVisioApplication();
            var doc = app.ActiveDocument;
            if (doc != null)
            {
                doc.Close(true);
            }
        }

        [MUT.TestMethod]
        public void CustomProps_AllTypes()
        {
            var page1 = this.GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 1, 1);
            s1.Text = "Checking for Custom Properties";
            
            // String Custom Property
            var prop_string_in = new CustomPropertyCells();
            prop_string_in.Format = "\"Format\"";
            prop_string_in.Label = "\"Label\"";
            prop_string_in.Prompt = "\"Prompt\"";
            prop_string_in.Type = CustomPropertyCells.CustomPropertyTypeToInt(CustomPropertyType.String);
            prop_string_in.Value = "1";

            // Boolean
            var prop_bool_in = new CustomPropertyCells();
            prop_bool_in.Format = "\"Format\"";
            prop_bool_in.Label = "\"Label\"";
            prop_bool_in.Prompt = "\"Prompt\"";
            prop_bool_in.Type = CustomPropertyCells.CustomPropertyTypeToInt(CustomPropertyType.Boolean);
            prop_bool_in.Value = true;

            // Date
            var dt = new System.DateTime(2017,3,31,14,5,6);
            var st = dt.ToString(CultureInfo.InvariantCulture);
            var prop_date_in = new CustomPropertyCells();
            prop_date_in.Format = "\"Format\"";
            prop_date_in.Label = "\"Label\"";
            prop_date_in.Prompt = "\"Prompt\"";
            prop_date_in.Type = CustomPropertyCells.CustomPropertyTypeToInt(CustomPropertyType.Date);
            prop_date_in.Value = string.Format("DATETIME(\"{0}\")", st); ;

            // Boolean
            var prop_number_in = new CustomPropertyCells();
            prop_number_in.Format = "\"Format\"";
            prop_number_in.Label = "\"Label\"";
            prop_number_in.Prompt = "\"Prompt\"";
            prop_number_in.Type = CustomPropertyCells.CustomPropertyTypeToInt(CustomPropertyType.Number);
            prop_number_in.Value = "3.14";

            CustomPropertyHelper.Set(s1, "PROP_STRING", prop_string_in);
            CustomPropertyHelper.Set(s1, "PROP_BOOLEAN", prop_bool_in);
            CustomPropertyHelper.Set(s1, "PROP_DATE", prop_date_in);
            CustomPropertyHelper.Set(s1, "PROP_NUMBER", prop_number_in);

            var props_dic = CustomPropertyHelper.GetDictionary(s1, VisioAutomation.Core.CellValueType.Formula);


            var prop_string_out = props_dic["PROP_STRING"];

            MUT.Assert.AreEqual("\"Format\"", prop_string_out.Format.Value);
            MUT.Assert.AreEqual("\"Label\"", prop_string_out.Label.Value);
            MUT.Assert.AreEqual("\"Prompt\"", prop_string_out.Prompt.Value);
            MUT.Assert.AreEqual("0", prop_string_out.Type.Value);
            MUT.Assert.AreEqual("1", prop_string_out.Value.Value);

            var prop_bool_out = props_dic["PROP_BOOLEAN"];
            MUT.Assert.AreEqual("\"Format\"", prop_bool_out.Format.Value);
            MUT.Assert.AreEqual("\"Label\"", prop_bool_out.Label.Value);
            MUT.Assert.AreEqual("\"Prompt\"", prop_bool_out.Prompt.Value);
            MUT.Assert.AreEqual("3", prop_bool_out.Type.Value);
            MUT.Assert.AreEqual("TRUE", prop_bool_out.Value.Value);

            var prop_date_out = props_dic["PROP_DATE"];
            MUT.Assert.AreEqual("\"Format\"", prop_date_out.Format.Value);
            MUT.Assert.AreEqual("\"Label\"", prop_date_out.Label.Value);
            MUT.Assert.AreEqual("\"Prompt\"", prop_date_out.Prompt.Value);
            MUT.Assert.AreEqual("5", prop_date_out.Type.Value);
            MUT.Assert.AreEqual("DATETIME(\"03/31/2017 14:05:06\")", prop_date_out.Value.Value);

            var prop_number_out = props_dic["PROP_NUMBER"];
            MUT.Assert.AreEqual("\"Format\"", prop_number_out.Format.Value);
            MUT.Assert.AreEqual("\"Label\"", prop_number_out.Label.Value);
            MUT.Assert.AreEqual("\"Prompt\"", prop_number_out.Prompt.Value);
            MUT.Assert.AreEqual("2", prop_number_out.Type.Value);
            MUT.Assert.AreEqual("3.14", prop_number_out.Value.Value);

            var app = this.GetVisioApplication();
            var doc = app.ActiveDocument;
            if (doc != null)
            {
                doc.Close(true);
            }
        }
    }
}