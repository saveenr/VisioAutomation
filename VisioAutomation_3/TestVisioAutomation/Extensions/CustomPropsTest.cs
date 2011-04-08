using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.CustomProperties;
using VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class CustomPropsTest : VisioAutomationTest
    {
        [TestMethod]
        public void SetCustomProps1()
        {
            var page1 = GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 2, 2);

            // By default a shape has ZERO custom Properties
            Assert.AreEqual(0, CustomPropertyHelper.GetCustomProperties(s1).Count);

            // Add a Custom Property
            var cp = new VA.CustomProperties.CustomPropertyCells();
            cp.Value = "BAR1";
            CustomPropertyHelper.SetCustomProperty(s1, "FOO1", cp);
            // Asset that now we have ONE CustomProperty
            Assert.AreEqual(1, CustomPropertyHelper.GetCustomProperties(s1).Count);
            // Check that it is called FOO1
            Assert.AreEqual(true, CustomPropertyHelper.HasCustomProperty(s1, "FOO1"));

            // Check that non-existent properties can't be found
            Assert.AreEqual(false, CustomPropertyHelper.HasCustomProperty(s1, "FOOX"));

            // Delete that custom property
            CustomPropertyHelper.DeleteCustomProperty(s1, "FOO1");
            // Verify that we have zero Custom Properties
            Assert.AreEqual(0, CustomPropertyHelper.GetCustomProperties(s1).Count);

            page1.Delete(0);
        }

        [TestMethod]
        public void SetSamePropMultipleTimes()
        {
            var page1 = GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 2, 2);

            // By default a shape has ZERO custom Properties
            Assert.AreEqual(0, CustomPropertyHelper.GetCustomProperties(s1).Count);

            // Add the same one multiple times Custom Property
            CustomPropertyHelper.SetCustomProperty(s1, "FOO1", "BAR1");
            // Asset that now we have ONE CustomProperty
            Assert.AreEqual(1, CustomPropertyHelper.GetCustomProperties(s1).Count);
            // Check that it is called FOO1
            Assert.AreEqual(true, CustomPropertyHelper.HasCustomProperty(s1, "FOO1"));

            // Try to SET the same property again many times
            CustomPropertyHelper.SetCustomProperty(s1, "FOO1", "BAR2");
            CustomPropertyHelper.SetCustomProperty(s1, "FOO1", "BAR3");
            CustomPropertyHelper.SetCustomProperty(s1, "FOO1", "BAR4");

            // Asset that now we have ONE CustomProperty
            Assert.AreEqual(1, CustomPropertyHelper.GetCustomProperties(s1).Count);
            // Check that it is called FOO1
            Assert.AreEqual(true, CustomPropertyHelper.HasCustomProperty(s1, "FOO1"));

            page1.Delete(0);
        }

        [TestMethod]
        public void InvalidPropName()
        {
            bool caught = false;
            var page1 = GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 2, 2);
            Assert.AreEqual(0, CustomPropertyHelper.GetCustomProperties(s1).Count);
            try
            {
                CustomPropertyHelper.SetCustomProperty(s1, "FOO 1", "BAR1");
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
        public void AdditionalProperties()
        {
            var page1 = GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 2, 2);
            Assert.AreEqual(0, CustomPropertyHelper.GetCustomProperties(s1).Count);

            var cp = new VA.CustomProperties.CustomPropertyCells();
            cp.Label = "The Foo property";
            cp.Value = "Some value";
            cp.Prompt = "Some Prompt";
            cp.LangId = 1034;
            cp.Type = (int) VA.CustomProperties.Format.DateOrTime;
            cp.Calendar = (int)IVisio.VisCellVals.visCalWestern;
            CustomPropertyHelper.SetCustomProperty(s1, "foo", cp);
            var z = CustomPropertyHelper.GetCustomProperties(s1);
            Assert.AreEqual(1, CustomPropertyHelper.GetCustomProperties(s1).Count);
            page1.Delete(0);
        }

        [TestMethod]
        public void PropertyNames()
        {
            var page1 = GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 2, 2);

            Assert.AreEqual(0, CustomPropertyHelper.GetCustomPropertyCount(s1));
            CustomPropertyHelper.SetCustomProperty(s1, "FOO1", "BAR1");
            Assert.AreEqual(1, CustomPropertyHelper.GetCustomPropertyCount(s1));
            CustomPropertyHelper.SetCustomProperty(s1, "FOO1", "BAR2");
            Assert.AreEqual(1, CustomPropertyHelper.GetCustomPropertyCount(s1));
            CustomPropertyHelper.SetCustomProperty(s1, "FOO2", "BAR3");

            var names1 = CustomPropertyHelper.GetCustomPropertyNames(s1);
            Assert.AreEqual("FOO1", names1[0]);
            Assert.AreEqual("FOO2", names1[1]);

            Assert.AreEqual(2, CustomPropertyHelper.GetCustomPropertyCount(s1));
            CustomPropertyHelper.DeleteCustomProperty(s1, "FOO1");

            var names2 = CustomPropertyHelper.GetCustomPropertyNames(s1);
            Assert.AreEqual("FOO2", names2[0]);

            CustomPropertyHelper.SetCustomProperty(s1, "FOO3", "BAR1");
            var names3 = CustomPropertyHelper.GetCustomPropertyNames(s1);
            Assert.AreEqual("FOO3", names3[0]);
            Assert.AreEqual("FOO2", names3[1]);

            CustomPropertyHelper.DeleteCustomProperty(s1, "FOO3");

            Assert.AreEqual(1, CustomPropertyHelper.GetCustomPropertyCount(s1));
            CustomPropertyHelper.DeleteCustomProperty(s1, "FOO2");

            Assert.AreEqual(0, CustomPropertyHelper.GetCustomPropertyCount(s1));
            page1.Delete(0);
        }

        [TestMethod]
        public void GetCustomPropsForMultipleShapes()
        {
            var page1 = GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 2, 2);
            var s2 = page1.DrawRectangle(0, 0, 2, 2);
            var s3 = page1.DrawRectangle(0, 0, 2, 2);
            var s4 = page1.DrawRectangle(0, 0, 2, 2);

            CustomPropertyHelper.SetCustomProperty(s1, "FOO1", "1");
            CustomPropertyHelper.SetCustomProperty(s2, "FOO2", "2");
            CustomPropertyHelper.SetCustomProperty(s2, "FOO3", "3");
            CustomPropertyHelper.SetCustomProperty(s4, "FOO4", "4");
            CustomPropertyHelper.SetCustomProperty(s4, "FOO5", "5");
            CustomPropertyHelper.SetCustomProperty(s4, "FOO6", "6");

            var shapes = new[] {s1, s2, s3, s4};
            var allprops = CustomPropertyHelper.GetCustomProperties(page1, shapes);

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