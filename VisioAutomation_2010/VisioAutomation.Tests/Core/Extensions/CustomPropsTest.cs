using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation;
using VisioAutomation.Shapes;
using VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation_Tests.Core.Extensions
{
    [TestClass]
    public class CustomPropsTest : VisioAutomationTest
    {

        [TestMethod]
        public void CustomProps_SetCustomProps1()
        {
            var page1 = this.GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 2, 2);

            // By default a shape has ZERO custom Properties
            Assert.AreEqual(0, CustomPropertyHelper.GetCount(s1));

            // Add a Custom Property
            var cp = new CustomPropertyCells();
            cp.Value = "\"BAR1\"";
            CustomPropertyHelper.Set(s1, "FOO1", cp);
            // Asset that now we have ONE CustomProperty
            Assert.AreEqual(1, CustomPropertyHelper.GetCount(s1));
            // Check that it is called FOO1
            Assert.AreEqual(true, CustomPropertyHelper.Contains(s1, "FOO1"));

            // Check that non-existent properties can't be found
            Assert.AreEqual(false, CustomPropertyHelper.Contains(s1, "FOOX"));

            // Delete that custom property
            CustomPropertyHelper.Delete(s1, "FOO1");
            // Verify that we have zero Custom Properties
            Assert.AreEqual(0, CustomPropertyHelper.GetCount(s1));

            page1.Delete(0);
        }

        [TestMethod]
        public void CustomProps_SetSamePropMultipleTimes()
        {
            var page1 = this.GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 2, 2);

            // By default a shape has ZERO custom Properties
            Assert.AreEqual(0, CustomPropertyHelper.GetCount(s1));

            int cp_type = 0; // string type

            // Add the same one multiple times Custom Property
            CustomPropertyHelper.Set(s1, "FOO1", "\"BAR1\"", cp_type);
            // Asset that now we have ONE CustomProperty
            Assert.AreEqual(1, CustomPropertyHelper.GetCount(s1));
            // Check that it is called FOO1
            Assert.AreEqual(true, CustomPropertyHelper.Contains(s1, "FOO1"));

            // Try to SET the same property again many times
            CustomPropertyHelper.Set(s1, "FOO1", "\"BAR2\"", cp_type);
            CustomPropertyHelper.Set(s1, "FOO1", "\"BAR3\"", cp_type);
            CustomPropertyHelper.Set(s1, "FOO1", "\"BAR4\"", cp_type);

            // Asset that now we have ONE CustomProperty
            Assert.AreEqual(1, CustomPropertyHelper.GetCount(s1));
            // Check that it is called FOO1
            Assert.AreEqual(true, CustomPropertyHelper.Contains(s1, "FOO1"));

            page1.Delete(0);
        }

        [TestMethod]
        public void CustomProps_InvalidPropName()
        {
            bool caught = false;
            var page1 = this.GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 2, 2);

            Assert.AreEqual(0, CustomPropertyHelper.GetCellsAsDictionary(s1, CellValueType.Formula).Count);

            int cp_type = 0; // 0 for string

            try
            {
                CustomPropertyHelper.Set(s1, "FOO 1", "BAR1", cp_type);
            }
            catch (System.ArgumentException)
            {
                page1.Delete(0);
                caught = true;
            }

            if (!caught)
            {
                Assert.Fail("Did not catch expected exception");
            }
        }

        [TestMethod]
        public void CustomProps_VerifyCustomPropAttributes()
        {
            var page1 = this.GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 2, 2);

            var in_cp = new CustomPropertyCells();
            in_cp.Label = "\"The Foo property\"";
            in_cp.Value = "\"Some value\"";
            in_cp.Prompt = "\"Some Prompt\"";
            in_cp.LangID = 1034;
            in_cp.Type = 0; // 0 = string. see: http://msdn.microsoft.com/en-us/library/aa200980(v=office.10).aspx
            in_cp.Calendar = (int) IVisio.VisCellVals.visCalWestern;
            in_cp.Invisible = 0;
            CustomPropertyHelper.Set(s1, "foo", in_cp);

            var out_cp = CustomPropertyHelper.GetCellsAsDictionary(s1, CellValueType.Formula);

            Assert.AreEqual(1, out_cp.Count);
            page1.Delete(0);
        }

        [TestMethod]
        public void CustomProps_PropertyNames()
        {
            var page1 = this.GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 2, 2);

            int cp_type = 0; // 0 for string

            Assert.AreEqual(0, CustomPropertyHelper.GetCount(s1));
            CustomPropertyHelper.Set(s1, "FOO1", "\"BAR1\"", cp_type);
            Assert.AreEqual(1, CustomPropertyHelper.GetCount(s1));
            CustomPropertyHelper.Set(s1, "FOO1", "\"BAR2\"", cp_type);
            Assert.AreEqual(1, CustomPropertyHelper.GetCount(s1));
            CustomPropertyHelper.Set(s1, "FOO2", "\"BAR3\"", cp_type);

            var names1 = CustomPropertyHelper.GetNames(s1);
            Assert.AreEqual(2,names1.Count);
            Assert.IsTrue(names1.Contains("FOO1"));
            Assert.IsTrue(names1.Contains("FOO2"));

            Assert.AreEqual(2, CustomPropertyHelper.GetCount(s1));
            CustomPropertyHelper.Delete(s1, "FOO1");

            var names2 = CustomPropertyHelper.GetNames(s1);
            Assert.AreEqual(1, names2.Count);
            Assert.IsTrue(names2.Contains("FOO2"));

            CustomPropertyHelper.Set(s1, "FOO3", "\"BAR1\"", cp_type);
            var names3 = CustomPropertyHelper.GetNames(s1);
            Assert.AreEqual(2, names3.Count);
            Assert.IsTrue(names3.Contains("FOO3"));
            Assert.IsTrue(names3.Contains("FOO2"));

            CustomPropertyHelper.Delete(s1, "FOO3");

            Assert.AreEqual(1, CustomPropertyHelper.GetCount(s1));
            CustomPropertyHelper.Delete(s1, "FOO2");

            Assert.AreEqual(0, CustomPropertyHelper.GetCount(s1));
            page1.Delete(0);
        }

        [TestMethod]
        public void CustomProps_GetFromMultipleShapes()
        {
            var page1 = this.GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 2, 2);
            var s2 = page1.DrawRectangle(0, 0, 2, 2);
            var s3 = page1.DrawRectangle(0, 0, 2, 2);
            var s4 = page1.DrawRectangle(0, 0, 2, 2);

            int cp_type = 0; // 0 for string

            CustomPropertyHelper.Set(s1, "FOO1", "1", cp_type);
            CustomPropertyHelper.Set(s2, "FOO2", "2", cp_type);
            CustomPropertyHelper.Set(s2, "FOO3", "3", cp_type);
            CustomPropertyHelper.Set(s4, "FOO4", "4", cp_type);
            CustomPropertyHelper.Set(s4, "FOO5", "5", cp_type);
            CustomPropertyHelper.Set(s4, "FOO6", "6", cp_type);

            var shapes = new[] {s1, s2, s3, s4};
            var shapeidpairs = ShapeIDPairs.FromShapes(shapes);
            var allprops = CustomPropertyHelper.GetCellsAsDictionary(page1, shapeidpairs, CellValueType.Formula);


            Assert.AreEqual(4, allprops.Count);
            Assert.AreEqual(1, allprops[0].Count);
            Assert.AreEqual(2, allprops[1].Count);
            Assert.AreEqual(0, allprops[2].Count);
            Assert.AreEqual(3, allprops[3].Count);

            Assert.AreEqual("1", allprops[0]["FOO1"].Value.Value);
            Assert.AreEqual("2", allprops[1]["FOO2"].Value.Value);
            Assert.AreEqual("3", allprops[1]["FOO3"].Value.Value);
            Assert.AreEqual("4", allprops[3]["FOO4"].Value.Value);
            Assert.AreEqual("5", allprops[3]["FOO5"].Value.Value);
            Assert.AreEqual("6", allprops[3]["FOO6"].Value.Value);

            page1.Delete(0);
        }

        [TestMethod]
        public void CustomProps_TryAllTypes()
        {
            var page1 = this.GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 2, 2);

            // string
            var cp_string = new CustomPropertyCells();
            cp_string.Value = "\"Hello World\"";
            cp_string.Type = CustomPropertyCells.CustomPropertyTypeToInt(CustomPropertyType.String);

            var cp_int = new CustomPropertyCells();
            cp_int.Value = 1024;
            cp_int.Type = CustomPropertyCells.CustomPropertyTypeToInt(CustomPropertyType.Number);

            var cp_dt = new CustomPropertyCells();
            cp_dt.Value = "DATETIME(\"03/31/1979\")";
            cp_dt.Type = CustomPropertyCells.CustomPropertyTypeToInt(CustomPropertyType.Date);

            var cp_bool = new CustomPropertyCells();
            cp_bool.Value = "TRUE";
            cp_bool.Type = CustomPropertyCells.CustomPropertyTypeToInt(CustomPropertyType.Boolean);

            var cp_float = new CustomPropertyCells();
            cp_float.Value = 3.14;
            cp_float.Type = CustomPropertyCells.CustomPropertyTypeToInt(CustomPropertyType.Number);

            CustomPropertyHelper.Set(s1, "PropertyString", cp_string);
            CustomPropertyHelper.Set(s1, "PropertyInt", cp_int);
            CustomPropertyHelper.Set(s1, "PropertyFloat", cp_float);
            CustomPropertyHelper.Set(s1, "PropertyDateTime", cp_dt);
            CustomPropertyHelper.Set(s1, "PropertyBool", cp_bool);

            var cpdic = CustomPropertyHelper.GetCellsAsDictionary(s1, CellValueType.Formula);

            var out_cpstring = cpdic["PropertyString"];
            var out_cpint = cpdic["PropertyInt"];
            var out_cpfloat = cpdic["PropertyFloat"];
            var out_cpdatetime = cpdic["PropertyDateTime"];
            var out_cpbool = cpdic["PropertyBool"];

            Assert.AreEqual("\"Hello World\"", out_cpstring.Value.Value);
            Assert.AreEqual("0", out_cpstring.Type.Value);

            Assert.AreEqual("1024", out_cpint.Value.Value);
            Assert.AreEqual("2", out_cpint.Type.Value);

            Assert.AreEqual("3.14", out_cpfloat.Value.Value);
            Assert.AreEqual("2", out_cpfloat.Type.Value);

            Assert.AreEqual("DATETIME(\"03/31/1979\")", out_cpdatetime.Value.Value);
            Assert.AreEqual("5", out_cpdatetime.Type.Value);

            Assert.AreEqual("TRUE", out_cpbool.Value.Value);
            Assert.AreEqual("3", out_cpbool.Type.Value);

            page1.Delete(0);
        }

    }
}