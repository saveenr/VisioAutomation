using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;
using VA=VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VTest.Core.Extensions
{
    [MUT.TestClass]
    public class CustomPropsTest : Framework.VTest
    {

        [MUT.TestMethod]
        public void CustomProps_SetCustomProps1()
        {
            var page1 = this.GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 2, 2);

            // By default a shape has ZERO custom Properties
            MUT.Assert.AreEqual(0, VA.Shapes.CustomPropertyHelper.GetCount(s1));

            // Add a Custom Property
            var cp = new VA.Shapes.CustomPropertyCells();
            cp.Value = "\"BAR1\"";
            VA.Shapes.CustomPropertyHelper.Set(s1, "FOO1", cp);
            // Asset that now we have ONE CustomProperty
            MUT.Assert.AreEqual(1, VA.Shapes.CustomPropertyHelper.GetCount(s1));
            // Check that it is called FOO1
            MUT.Assert.AreEqual(true, VA.Shapes.CustomPropertyHelper.Contains(s1, "FOO1"));

            // Check that non-existent properties can't be found
            MUT.Assert.AreEqual(false, VA.Shapes.CustomPropertyHelper.Contains(s1, "FOOX"));

            // Delete that custom property
            VA.Shapes.CustomPropertyHelper.Delete(s1, "FOO1");
            // Verify that we have zero Custom Properties
            MUT.Assert.AreEqual(0, VA.Shapes.CustomPropertyHelper.GetCount(s1));

            page1.Delete(0);
        }

        [MUT.TestMethod]
        public void CustomProps_SetSamePropMultipleTimes()
        {
            var page1 = this.GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 2, 2);

            // By default a shape has ZERO custom Properties
            MUT.Assert.AreEqual(0, VA.Shapes.CustomPropertyHelper.GetCount(s1));

            int cp_type = 0; // string type

            // Add the same one multiple times Custom Property
            VA.Shapes.CustomPropertyHelper.Set(s1, "FOO1", "\"BAR1\"", cp_type);
            // Asset that now we have ONE CustomProperty
            MUT.Assert.AreEqual(1, VA.Shapes.CustomPropertyHelper.GetCount(s1));
            // Check that it is called FOO1
            MUT.Assert.AreEqual(true, VA.Shapes.CustomPropertyHelper.Contains(s1, "FOO1"));

            // Try to SET the same property again many times
            VA.Shapes.CustomPropertyHelper.Set(s1, "FOO1", "\"BAR2\"", cp_type);
            VA.Shapes.CustomPropertyHelper.Set(s1, "FOO1", "\"BAR3\"", cp_type);
            VA.Shapes.CustomPropertyHelper.Set(s1, "FOO1", "\"BAR4\"", cp_type);

            // Asset that now we have ONE CustomProperty
            MUT.Assert.AreEqual(1, VA.Shapes.CustomPropertyHelper.GetCount(s1));
            // Check that it is called FOO1
            MUT.Assert.AreEqual(true, VA.Shapes.CustomPropertyHelper.Contains(s1, "FOO1"));

            page1.Delete(0);
        }

        [MUT.TestMethod]
        public void CustomProps_InvalidPropName()
        {
            bool caught = false;
            var page1 = this.GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 2, 2);

            MUT.Assert.AreEqual(0, VA.Shapes.CustomPropertyHelper.GetDictionary(s1, VisioAutomation.Core.CellValueType.Formula).Count);

            int cp_type = 0; // 0 for string

            try
            {
                VA.Shapes.CustomPropertyHelper.Set(s1, "FOO 1", "BAR1", cp_type);
            }
            catch (System.ArgumentException)
            {
                page1.Delete(0);
                caught = true;
            }

            if (!caught)
            {
                MUT.Assert.Fail("Did not catch expected exception");
            }
        }

        [MUT.TestMethod]
        public void CustomProps_VerifyCustomPropAttributes()
        {
            var page1 = this.GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 2, 2);

            var in_cp = new VA.Shapes.CustomPropertyCells();
            in_cp.Label = "\"The Foo property\"";
            in_cp.Value = "\"Some value\"";
            in_cp.Prompt = "\"Some Prompt\"";
            in_cp.LangID = 1034;
            in_cp.Type = 0; // 0 = string. see: http://msdn.microsoft.com/en-us/library/aa200980(v=office.10).aspx
            in_cp.Calendar = (int) IVisio.VisCellVals.visCalWestern;
            in_cp.Invisible = 0;
            VA.Shapes.CustomPropertyHelper.Set(s1, "foo", in_cp);

            var out_cp = VA.Shapes.CustomPropertyHelper.GetDictionary(s1, VisioAutomation.Core.CellValueType.Formula);

            MUT.Assert.AreEqual(1, out_cp.Count);
            page1.Delete(0);
        }

        [MUT.TestMethod]
        public void CustomProps_PropertyNames()
        {
            var page1 = this.GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 2, 2);

            int cp_type = 0; // 0 for string

            MUT.Assert.AreEqual(0, VA.Shapes.CustomPropertyHelper.GetCount(s1));
            VA.Shapes.CustomPropertyHelper.Set(s1, "FOO1", "\"BAR1\"", cp_type);
            MUT.Assert.AreEqual(1, VA.Shapes.CustomPropertyHelper.GetCount(s1));
            VA.Shapes.CustomPropertyHelper.Set(s1, "FOO1", "\"BAR2\"", cp_type);
            MUT.Assert.AreEqual(1, VA.Shapes.CustomPropertyHelper.GetCount(s1));
            VA.Shapes.CustomPropertyHelper.Set(s1, "FOO2", "\"BAR3\"", cp_type);

            var names1 = VA.Shapes.CustomPropertyHelper.GetNames(s1);
            MUT.Assert.AreEqual(2,names1.Count);
            MUT.Assert.IsTrue(names1.Contains("FOO1"));
            MUT.Assert.IsTrue(names1.Contains("FOO2"));

            MUT.Assert.AreEqual(2, VA.Shapes.CustomPropertyHelper.GetCount(s1));
            VA.Shapes.CustomPropertyHelper.Delete(s1, "FOO1");

            var names2 = VA.Shapes.CustomPropertyHelper.GetNames(s1);
            MUT.Assert.AreEqual(1, names2.Count);
            MUT.Assert.IsTrue(names2.Contains("FOO2"));

            VA.Shapes.CustomPropertyHelper.Set(s1, "FOO3", "\"BAR1\"", cp_type);
            var names3 = VA.Shapes.CustomPropertyHelper.GetNames(s1);
            MUT.Assert.AreEqual(2, names3.Count);
            MUT.Assert.IsTrue(names3.Contains("FOO3"));
            MUT.Assert.IsTrue(names3.Contains("FOO2"));

            VA.Shapes.CustomPropertyHelper.Delete(s1, "FOO3");

            MUT.Assert.AreEqual(1, VA.Shapes.CustomPropertyHelper.GetCount(s1));
            VA.Shapes.CustomPropertyHelper.Delete(s1, "FOO2");

            MUT.Assert.AreEqual(0, VA.Shapes.CustomPropertyHelper.GetCount(s1));
            page1.Delete(0);
        }

        [MUT.TestMethod]
        public void CustomProps_GetFromMultipleShapes()
        {
            var page1 = this.GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 2, 2);
            var s2 = page1.DrawRectangle(0, 0, 2, 2);
            var s3 = page1.DrawRectangle(0, 0, 2, 2);
            var s4 = page1.DrawRectangle(0, 0, 2, 2);

            int cp_type = 0; // 0 for string

            VA.Shapes.CustomPropertyHelper.Set(s1, "FOO1", "1", cp_type);
            VA.Shapes.CustomPropertyHelper.Set(s2, "FOO2", "2", cp_type);
            VA.Shapes.CustomPropertyHelper.Set(s2, "FOO3", "3", cp_type);
            VA.Shapes.CustomPropertyHelper.Set(s4, "FOO4", "4", cp_type);
            VA.Shapes.CustomPropertyHelper.Set(s4, "FOO5", "5", cp_type);
            VA.Shapes.CustomPropertyHelper.Set(s4, "FOO6", "6", cp_type);

            var shapes = new[] {s1, s2, s3, s4};
            var shapeidpairs = VisioAutomation.Core.ShapeIDPairs.FromShapes(shapes);
            var allprops = VA.Shapes.CustomPropertyHelper.GetDictionary(page1, shapeidpairs, VisioAutomation.Core.CellValueType.Formula);


            MUT.Assert.AreEqual(4, allprops.Count);
            MUT.Assert.AreEqual(1, allprops[0].Count);
            MUT.Assert.AreEqual(2, allprops[1].Count);
            MUT.Assert.AreEqual(0, allprops[2].Count);
            MUT.Assert.AreEqual(3, allprops[3].Count);

            MUT.Assert.AreEqual("1", allprops[0]["FOO1"].Value.Value);
            MUT.Assert.AreEqual("2", allprops[1]["FOO2"].Value.Value);
            MUT.Assert.AreEqual("3", allprops[1]["FOO3"].Value.Value);
            MUT.Assert.AreEqual("4", allprops[3]["FOO4"].Value.Value);
            MUT.Assert.AreEqual("5", allprops[3]["FOO5"].Value.Value);
            MUT.Assert.AreEqual("6", allprops[3]["FOO6"].Value.Value);

            page1.Delete(0);
        }

        [MUT.TestMethod]
        public void CustomProps_TryAllTypes()
        {
            var page1 = this.GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 2, 2);

            // string
            var cp_string = new VA.Shapes.CustomPropertyCells();
            cp_string.Value = "\"Hello World\"";
            cp_string.Type = VA.Shapes.CustomPropertyCells.CustomPropertyTypeToInt(VA.Shapes.CustomPropertyType.String);

            var cp_int = new VA.Shapes.CustomPropertyCells();
            cp_int.Value = 1024;
            cp_int.Type = VA.Shapes.CustomPropertyCells.CustomPropertyTypeToInt(VA.Shapes.CustomPropertyType.Number);

            var cp_dt = new VA.Shapes.CustomPropertyCells();
            cp_dt.Value = "DATETIME(\"03/31/1979\")";
            cp_dt.Type = VA.Shapes.CustomPropertyCells.CustomPropertyTypeToInt(VA.Shapes.CustomPropertyType.Date);

            var cp_bool = new VA.Shapes.CustomPropertyCells();
            cp_bool.Value = "TRUE";
            cp_bool.Type = VA.Shapes.CustomPropertyCells.CustomPropertyTypeToInt(VA.Shapes.CustomPropertyType.Boolean);

            var cp_float = new VA.Shapes.CustomPropertyCells();
            cp_float.Value = 3.14;
            cp_float.Type = VA.Shapes.CustomPropertyCells.CustomPropertyTypeToInt(VA.Shapes.CustomPropertyType.Number);

            VA.Shapes.CustomPropertyHelper.Set(s1, "PropertyString", cp_string);
            VA.Shapes.CustomPropertyHelper.Set(s1, "PropertyInt", cp_int);
            VA.Shapes.CustomPropertyHelper.Set(s1, "PropertyFloat", cp_float);
            VA.Shapes.CustomPropertyHelper.Set(s1, "PropertyDateTime", cp_dt);
            VA.Shapes.CustomPropertyHelper.Set(s1, "PropertyBool", cp_bool);

            var cpdic = VA.Shapes.CustomPropertyHelper.GetDictionary(s1, VisioAutomation.Core.CellValueType.Formula);

            var out_cpstring = cpdic["PropertyString"];
            var out_cpint = cpdic["PropertyInt"];
            var out_cpfloat = cpdic["PropertyFloat"];
            var out_cpdatetime = cpdic["PropertyDateTime"];
            var out_cpbool = cpdic["PropertyBool"];

            MUT.Assert.AreEqual("\"Hello World\"", out_cpstring.Value.Value);
            MUT.Assert.AreEqual("0", out_cpstring.Type.Value);

            MUT.Assert.AreEqual("1024", out_cpint.Value.Value);
            MUT.Assert.AreEqual("2", out_cpint.Type.Value);

            MUT.Assert.AreEqual("3.14", out_cpfloat.Value.Value);
            MUT.Assert.AreEqual("2", out_cpfloat.Type.Value);

            MUT.Assert.AreEqual("DATETIME(\"03/31/1979\")", out_cpdatetime.Value.Value);
            MUT.Assert.AreEqual("5", out_cpdatetime.Type.Value);

            MUT.Assert.AreEqual("TRUE", out_cpbool.Value.Value);
            MUT.Assert.AreEqual("3", out_cpbool.Type.Value);

            page1.Delete(0);
        }

    }
}