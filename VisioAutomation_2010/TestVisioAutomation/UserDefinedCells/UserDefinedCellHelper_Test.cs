using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class UserDefinedCellHelper_Test : VisioAutomationTest
    {
        [TestMethod]
        public void BasicScenario()
        {
            var page1 = GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 2, 2);

            // By default a shape has ZERO custom Properties
            Assert.AreEqual(0, VA.UserDefinedCells.UserDefinedCellsHelper.GetUserDefinedCellCount(s1));

            // Add a Custom Property
            var prop = new VA.UserDefinedCells.UserDefinedCell("FOO1", "BAR1");
            VA.UserDefinedCells.UserDefinedCellsHelper.SetUserDefinedCell(s1, prop.Name, prop.Value, prop.Prompt);
            Assert.AreEqual(1, VA.UserDefinedCells.UserDefinedCellsHelper.GetUserDefinedCellCount(s1));
            // Check that it is called FOO1
            Assert.AreEqual(true, VA.UserDefinedCells.UserDefinedCellsHelper.HasUserDefinedCell(s1, "FOO1"));

            // Check that non-existent properties can't be found
            Assert.AreEqual(false, VA.CustomProperties.CustomPropertyHelper.HasCustomProperty(s1, "FOOX"));

            // Delete that custom property
            VA.UserDefinedCells.UserDefinedCellsHelper.DeleteUserDefinedCell(s1, "FOO1");
            // Verify that we have zero Custom Properties
            Assert.AreEqual(0, VA.UserDefinedCells.UserDefinedCellsHelper.GetUserDefinedCellCount(s1));

            page1.Delete(0);
        }

        [TestMethod]
        public void GetPropsForMultipleShapes()
        {
            var page1 = GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 1, 1);
            var s2 = page1.DrawRectangle(1, 1, 2, 2);
            var shapes = new[] { s1, s2 };

            var props = VA.UserDefinedCells.UserDefinedCellsHelper.GetUserDefinedCells(page1, shapes);
            Assert.AreEqual(2, props.Count);
            Assert.AreEqual(0, props[0].Count);
            Assert.AreEqual(0, props[1].Count);

            VA.UserDefinedCells.UserDefinedCellsHelper.SetUserDefinedCell(s1, "foo", "bar", null);
            var props1 = VA.UserDefinedCells.UserDefinedCellsHelper.GetUserDefinedCells(page1, shapes);
            Assert.AreEqual(2, props1.Count);
            Assert.AreEqual(1, props1[0].Count);
            Assert.AreEqual(0, props1[1].Count);

            page1.Delete(0);
        }

        [TestMethod]
        public void SetSamePropMultipleTimes()
        {
            var page1 = GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 2, 2);

            // By default a shape has ZERO custom Properties
            Assert.AreEqual(0, VA.CustomProperties.CustomPropertyHelper.GetCustomProperties(s1).Count);

            // Add the same one multiple times Custom Property
            VA.UserDefinedCells.UserDefinedCellsHelper.SetUserDefinedCell(s1, "FOO1", "BAR1", null);
            // Asset that now we have ONE CustomProperty
            Assert.AreEqual(1, VA.UserDefinedCells.UserDefinedCellsHelper.GetUserDefinedCellCount(s1));
            // Check that it is called FOO1
            Assert.AreEqual(true, VA.UserDefinedCells.UserDefinedCellsHelper.HasUserDefinedCell(s1, "FOO1"));

            // Try to SET the same property again many times
            VA.UserDefinedCells.UserDefinedCellsHelper.SetUserDefinedCell(s1, "FOO1", "BAR2", null);
            VA.UserDefinedCells.UserDefinedCellsHelper.SetUserDefinedCell(s1, "FOO1", "BAR3", null);
            VA.UserDefinedCells.UserDefinedCellsHelper.SetUserDefinedCell(s1, "FOO1", "BAR4", null);

            // Asset that now we have ONE CustomProperty
            Assert.AreEqual(1, VA.UserDefinedCells.UserDefinedCellsHelper.GetUserDefinedCellCount(s1));
            // Check that it is called FOO1
            Assert.AreEqual(true, VA.UserDefinedCells.UserDefinedCellsHelper.HasUserDefinedCell(s1, "FOO1"));
            page1.Delete(0);
        }

        [TestMethod]
        public void InvalidPropName()
        {
            bool caught = false;
            var page1 = GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 2, 2);
            Assert.AreEqual(0, VA.UserDefinedCells.UserDefinedCellsHelper.GetUserDefinedCellCount(s1));
            try
            {
                VA.UserDefinedCells.UserDefinedCellsHelper.SetUserDefinedCell(s1, "FOO 1", "BAR1", null);
            }
            catch (VA.AutomationException )
            {
                // this was expected
                page1.Delete(0);
                caught = true;
            }
            if (caught == false)
            {
                Assert.Fail("Did not catch expected exception");
            }
        }

        [TestMethod]
        public void AdditionalProperties()
        {
            var page1 = GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 2, 2);
            Assert.AreEqual(0, VA.UserDefinedCells.UserDefinedCellsHelper.GetUserDefinedCellCount(s1));

            var prop = new VA.UserDefinedCells.UserDefinedCell("foo");
            prop.Prompt = "Some Prompt";
            VA.UserDefinedCells.UserDefinedCellsHelper.SetUserDefinedCell(s1, prop.Name, prop.Value, prop.Prompt);
            Assert.AreEqual(1, VA.UserDefinedCells.UserDefinedCellsHelper.GetUserDefinedCellCount(s1));
            page1.Delete(0);
        }

        [TestMethod]
        public void GetPropNames()
        {
            var page1 = GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 2, 2);

            Assert.AreEqual(0, VA.UserDefinedCells.UserDefinedCellsHelper.GetUserDefinedCellCount(s1));
            VA.UserDefinedCells.UserDefinedCellsHelper.SetUserDefinedCell(s1, "FOO1", "BAR1", null);
            Assert.AreEqual(1, VA.UserDefinedCells.UserDefinedCellsHelper.GetUserDefinedCellCount(s1));
            VA.UserDefinedCells.UserDefinedCellsHelper.SetUserDefinedCell(s1, "FOO1", "BAR2", null);
            Assert.AreEqual(1, VA.UserDefinedCells.UserDefinedCellsHelper.GetUserDefinedCellCount(s1));
            VA.UserDefinedCells.UserDefinedCellsHelper.SetUserDefinedCell(s1, "FOO2", "BAR3", null);

            var names1 = VA.UserDefinedCells.UserDefinedCellsHelper.GetUserDefinedCellNames(s1);
            Assert.AreEqual("FOO1", names1[0]);
            Assert.AreEqual("FOO2", names1[1]);

            Assert.AreEqual(2, VA.UserDefinedCells.UserDefinedCellsHelper.GetUserDefinedCellCount(s1));
            VA.UserDefinedCells.UserDefinedCellsHelper.DeleteUserDefinedCell(s1, "FOO1");

            var names2 = VA.UserDefinedCells.UserDefinedCellsHelper.GetUserDefinedCellNames(s1);
            Assert.AreEqual("FOO2", names2[0]);

            VA.UserDefinedCells.UserDefinedCellsHelper.SetUserDefinedCell(s1, "FOO3", "BAR1", null);
            var names3 = VA.UserDefinedCells.UserDefinedCellsHelper.GetUserDefinedCellNames(s1);
            Assert.AreEqual("FOO3", names3[0]);
            Assert.AreEqual("FOO2", names3[1]);

            VA.UserDefinedCells.UserDefinedCellsHelper.DeleteUserDefinedCell(s1, "FOO3");

            Assert.AreEqual(1, VA.UserDefinedCells.UserDefinedCellsHelper.GetUserDefinedCellCount(s1));
            VA.UserDefinedCells.UserDefinedCellsHelper.DeleteUserDefinedCell(s1, "FOO2");

            Assert.AreEqual(0, VA.UserDefinedCells.UserDefinedCellsHelper.GetUserDefinedCellCount(s1));

            page1.Delete(0);
        }

        [TestMethod]
        public void PropsForMultipleShapes()
        {
            var page1 = GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 2, 2);
            var s2 = page1.DrawRectangle(0, 0, 2, 2);
            var s3 = page1.DrawRectangle(0, 0, 2, 2);
            var s4 = page1.DrawRectangle(0, 0, 2, 2);

            VA.UserDefinedCells.UserDefinedCellsHelper.SetUserDefinedCell(s1, "FOO1", "1", null);
            VA.UserDefinedCells.UserDefinedCellsHelper.SetUserDefinedCell(s2, "FOO2", "2", null);
            VA.UserDefinedCells.UserDefinedCellsHelper.SetUserDefinedCell(s2, "FOO3", "3", null);
            VA.UserDefinedCells.UserDefinedCellsHelper.SetUserDefinedCell(s4, "FOO4", "4", null);
            VA.UserDefinedCells.UserDefinedCellsHelper.SetUserDefinedCell(s4, "FOO5", "5", null);
            VA.UserDefinedCells.UserDefinedCellsHelper.SetUserDefinedCell(s4, "FOO6", "6", null);

            var shapeids = new[] {s1, s2, s3, s4};
            var allprops = VA.UserDefinedCells.UserDefinedCellsHelper.GetUserDefinedCells(page1, shapeids);

            Assert.AreEqual(4, allprops.Count);
            Assert.AreEqual(1, allprops[0].Count);
            Assert.AreEqual(2, allprops[1].Count);
            Assert.AreEqual(0, allprops[2].Count);
            Assert.AreEqual(3, allprops[3].Count);

            Assert.AreEqual("\"1\"", allprops[0][0].Value);
            Assert.AreEqual("\"2\"", allprops[1][0].Value);
            Assert.AreEqual("\"3\"", allprops[1][1].Value);
            Assert.AreEqual("\"4\"", allprops[3][0].Value);
            Assert.AreEqual("\"5\"", allprops[3][1].Value);
            Assert.AreEqual("\"6\"", allprops[3][2].Value);
            page1.Delete(0);
        }

        [TestMethod]
        public void GetProps1()
        {
            var page1 = GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 2, 2);

            var p1 = VA.UserDefinedCells.UserDefinedCellsHelper.GetUserDefinedCells(s1);
            Assert.AreEqual(0, p1.Count);

            VA.UserDefinedCells.UserDefinedCellsHelper.SetUserDefinedCell(s1, "FOO1", "1", null);
            VA.UserDefinedCells.UserDefinedCellsHelper.SetUserDefinedCell(s1, "FOO2", "2", null);

            var p2 = VA.UserDefinedCells.UserDefinedCellsHelper.GetUserDefinedCells(s1);
            Assert.AreEqual(2, p2.Count);
            Assert.AreEqual("FOO1",p2[0].Name);
            Assert.AreEqual("\"1\"", p2[0].Value);
            Assert.AreEqual("FOO2", p2[1].Name);
            Assert.AreEqual("\"2\"", p2[1].Value);

            page1.Delete(0);
        }

    }
}