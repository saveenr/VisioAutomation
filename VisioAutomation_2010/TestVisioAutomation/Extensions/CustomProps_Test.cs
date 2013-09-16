using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Shapes.CustomProperties;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class CustomProps_Test : VisioAutomationTest
    {

        [TestMethod]
        public void CustomProps_SetCustomProps1()
        {
            var page1 = GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 2, 2);

            // By default a shape has ZERO custom Properties
            Assert.AreEqual(0, CustomPropertyHelper.Get(s1).Count);

            // Add a Custom Property
            var cp = new CustomPropertyCells();
            cp.Value = "BAR1";
            CustomPropertyHelper.Set(s1, "FOO1", cp);
            // Asset that now we have ONE CustomProperty
            Assert.AreEqual(1, CustomPropertyHelper.Get(s1).Count);
            // Check that it is called FOO1
            Assert.AreEqual(true, CustomPropertyHelper.Contains(s1, "FOO1"));

            // Check that non-existent properties can't be found
            Assert.AreEqual(false, CustomPropertyHelper.Contains(s1, "FOOX"));

            // Delete that custom property
            CustomPropertyHelper.Delete(s1, "FOO1");
            // Verify that we have zero Custom Properties
            Assert.AreEqual(0, CustomPropertyHelper.Get(s1).Count);

            page1.Delete(0);
        }

        [TestMethod]
        public void CustomProps_SetSamePropMultipleTimes()
        {
            var page1 = GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 2, 2);

            // By default a shape has ZERO custom Properties
            Assert.AreEqual(0, CustomPropertyHelper.Get(s1).Count);

            // Add the same one multiple times Custom Property
            CustomPropertyHelper.Set(s1, "FOO1", "BAR1");
            // Asset that now we have ONE CustomProperty
            Assert.AreEqual(1, CustomPropertyHelper.Get(s1).Count);
            // Check that it is called FOO1
            Assert.AreEqual(true, CustomPropertyHelper.Contains(s1, "FOO1"));

            // Try to SET the same property again many times
            CustomPropertyHelper.Set(s1, "FOO1", "BAR2");
            CustomPropertyHelper.Set(s1, "FOO1", "BAR3");
            CustomPropertyHelper.Set(s1, "FOO1", "BAR4");

            // Asset that now we have ONE CustomProperty
            Assert.AreEqual(1, CustomPropertyHelper.Get(s1).Count);
            // Check that it is called FOO1
            Assert.AreEqual(true, CustomPropertyHelper.Contains(s1, "FOO1"));

            page1.Delete(0);
        }

        [TestMethod]
        public void CustomProps_InvalidPropName()
        {
            bool caught = false;
            var page1 = GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 2, 2);
            Assert.AreEqual(0, CustomPropertyHelper.Get(s1).Count);
            try
            {
                CustomPropertyHelper.Set(s1, "FOO 1", "BAR1");
            }
            catch (VA.AutomationException )
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
            var page1 = GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 2, 2);

            var in_cp = new CustomPropertyCells();
            in_cp.Label = "The Foo property";
            in_cp.Value = "Some value";
            in_cp.Prompt = "Some Prompt";
            in_cp.LangId = 1034;
            in_cp.Type = 5; // 5= date-or-time. see: http://msdn.microsoft.com/en-us/library/aa200980(v=office.10).aspx
            in_cp.Calendar = (int) IVisio.VisCellVals.visCalWestern;
            in_cp.Invisible = 0;
            in_cp.Verify = 0;
            CustomPropertyHelper.Set(s1, "foo", in_cp);
            var out_cp = CustomPropertyHelper.Get(s1);
            Assert.AreEqual(1, out_cp.Count);
            page1.Delete(0);
        }

        [TestMethod]
        public void CustomProps_PropertyNames()
        {
            var page1 = GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 2, 2);

            Assert.AreEqual(0, CustomPropertyHelper.GetCount(s1));
            CustomPropertyHelper.Set(s1, "FOO1", "BAR1");
            Assert.AreEqual(1, CustomPropertyHelper.GetCount(s1));
            CustomPropertyHelper.Set(s1, "FOO1", "BAR2");
            Assert.AreEqual(1, CustomPropertyHelper.GetCount(s1));
            CustomPropertyHelper.Set(s1, "FOO2", "BAR3");

            var names1 = CustomPropertyHelper.GetNames(s1);
            Assert.AreEqual(2,names1.Count);
            Assert.IsTrue(names1.Contains("FOO1"));
            Assert.IsTrue(names1.Contains("FOO2"));

            Assert.AreEqual(2, CustomPropertyHelper.GetCount(s1));
            CustomPropertyHelper.Delete(s1, "FOO1");

            var names2 = CustomPropertyHelper.GetNames(s1);
            Assert.AreEqual(1, names2.Count);
            Assert.IsTrue(names2.Contains("FOO2"));

            CustomPropertyHelper.Set(s1, "FOO3", "BAR1");
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
            var page1 = GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 2, 2);
            var s2 = page1.DrawRectangle(0, 0, 2, 2);
            var s3 = page1.DrawRectangle(0, 0, 2, 2);
            var s4 = page1.DrawRectangle(0, 0, 2, 2);

            CustomPropertyHelper.Set(s1, "FOO1", "1");
            CustomPropertyHelper.Set(s2, "FOO2", "2");
            CustomPropertyHelper.Set(s2, "FOO3", "3");
            CustomPropertyHelper.Set(s4, "FOO4", "4");
            CustomPropertyHelper.Set(s4, "FOO5", "5");
            CustomPropertyHelper.Set(s4, "FOO6", "6");

            var shapes = new[] {s1, s2, s3, s4};
            var allprops = CustomPropertyHelper.Get(page1, shapes);

            Assert.AreEqual(4, allprops.Count);
            Assert.AreEqual(1, allprops[0].Count);
            Assert.AreEqual(2, allprops[1].Count);
            Assert.AreEqual(0, allprops[2].Count);
            Assert.AreEqual(3, allprops[3].Count);

            Assert.AreEqual("\"1\"", allprops[0]["FOO1"].Value.Formula);
            Assert.AreEqual("\"2\"", allprops[1]["FOO2"].Value.Formula);
            Assert.AreEqual("\"3\"", allprops[1]["FOO3"].Value.Formula);
            Assert.AreEqual("\"4\"", allprops[3]["FOO4"].Value.Formula);
            Assert.AreEqual("\"5\"", allprops[3]["FOO5"].Value.Formula);
            Assert.AreEqual("\"6\"", allprops[3]["FOO6"].Value.Formula);

            page1.Delete(0);
        }
    }
}