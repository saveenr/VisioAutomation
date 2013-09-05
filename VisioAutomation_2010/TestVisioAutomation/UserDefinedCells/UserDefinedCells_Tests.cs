using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using System.Linq;
using VisioAutomation.Shapes.CustomProperties;
using VisioAutomation.Shapes.UserDefinedCells;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class UserDefinedCells_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void UserDefinedCells_Scenario1()
        {
            var page1 = GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 2, 2);

            // By default a shape has ZERO custom Properties
            Assert.AreEqual(0, UserDefinedCellsHelper.GetCount(s1));

            // Add a Custom Property
            var prop = new UserDefinedCell("FOO1", "BAR1");
            UserDefinedCellsHelper.Set(s1, prop.Name, prop.Value, prop.Prompt);
            Assert.AreEqual(1, UserDefinedCellsHelper.GetCount(s1));
            // Check that it is called FOO1
            Assert.AreEqual(true, UserDefinedCellsHelper.Contains(s1, "FOO1"));

            // Check that non-existent properties can't be found
            Assert.AreEqual(false, CustomPropertyHelper.Contains(s1, "FOOX"));

            // Delete that custom property
            UserDefinedCellsHelper.Delete(s1, "FOO1");
            // Verify that we have zero Custom Properties
            Assert.AreEqual(0, UserDefinedCellsHelper.GetCount(s1));

            page1.Delete(0);
        }

        [TestMethod]
        public void UserDefinedCells_GetFromMultipleShapes()
        {
            var page1 = GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 1, 1);
            var s2 = page1.DrawRectangle(1, 1, 2, 2);
            var shapes = new[] { s1, s2 };

            UserDefinedCellsHelper.Set(s1, "foo", "bar", null);
            var props1 = UserDefinedCellsHelper.Get(page1, shapes);
            Assert.AreEqual(2, props1.Count);
            Assert.AreEqual(1, props1[0].Count);
            Assert.AreEqual(0, props1[1].Count);

            page1.Delete(0);
        }

        [TestMethod]
        public void UserDefinedCells_GetFromMultipleShapes_WithAdditionalProps()
        {
            var page1 = GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 1, 1);
            var s2 = page1.DrawRectangle(1, 1, 2, 2);
            var shapes = new[] { s1, s2 };

            UserDefinedCellsHelper.Set(s1, "foo", "bar", null);

            var queryex = new VA.ShapeSheet.Query.CellQuery();
            var sec = queryex.Sections.Add(IVisio.VisSectionIndices.visSectionUser);
            var Value = sec.Columns.Add(VA.ShapeSheet.SRCConstants.User_Value, "Value");
            var Prompt = sec.Columns.Add(VA.ShapeSheet.SRCConstants.User_Prompt, "Prompt");

            var formulas = queryex.GetFormulas(page1, shapes.Select(s => s.ID).ToList());


            page1.Delete(0);
        }

        [TestMethod]
        public void UserDefinedCells_SetMultipleTimes()
        {
            var page1 = GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 2, 2);

            // By default a shape has ZERO custom Properties
            Assert.AreEqual(0, CustomPropertyHelper.Get(s1).Count);

            // Add the same one multiple times Custom Property
            UserDefinedCellsHelper.Set(s1, "FOO1", "BAR1", null);
            // Asset that now we have ONE CustomProperty
            Assert.AreEqual(1, UserDefinedCellsHelper.GetCount(s1));
            // Check that it is called FOO1
            Assert.AreEqual(true, UserDefinedCellsHelper.Contains(s1, "FOO1"));

            // Try to SET the same property again many times
            UserDefinedCellsHelper.Set(s1, "FOO1", "BAR2", null);
            UserDefinedCellsHelper.Set(s1, "FOO1", "BAR3", null);
            UserDefinedCellsHelper.Set(s1, "FOO1", "BAR4", null);

            // Asset that now we have ONE CustomProperty
            Assert.AreEqual(1, UserDefinedCellsHelper.GetCount(s1));
            // Check that it is called FOO1
            Assert.AreEqual(true, UserDefinedCellsHelper.Contains(s1, "FOO1"));
            page1.Delete(0);
        }

        [TestMethod]
        public void UserDefinedCells_InvalidNames()
        {
            if (UserDefinedCellsHelper.IsValidName("A") == false)
            {
                Assert.Fail();
            }

            if (UserDefinedCellsHelper.IsValidName("A.B") == false)
            {
                Assert.Fail();
            }

            if (UserDefinedCellsHelper.IsValidName("A B") == true)
            {
                Assert.Fail();
            }

            if (UserDefinedCellsHelper.IsValidName(" ") == true)
            {
                Assert.Fail();
            }
        }

        [TestMethod]
        public void UserDefinedCells_CheckInvalidNamesNotAllowed()
        {
            bool caught = false;
            var page1 = GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 2, 2);
            Assert.AreEqual(0, UserDefinedCellsHelper.GetCount(s1));
            try
            {
                UserDefinedCellsHelper.Set(s1, "FOO 1", "BAR1", null);
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
        public void UserDefinedCells_SetAdditionalProperties()
        {
            var page1 = GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 2, 2);
            Assert.AreEqual(0, UserDefinedCellsHelper.GetCount(s1));

            var prop = new UserDefinedCell("foo");
            prop.Prompt = "Some Prompt";
            UserDefinedCellsHelper.Set(s1, prop.Name, prop.Value, prop.Prompt);
            Assert.AreEqual(1, UserDefinedCellsHelper.GetCount(s1));
            page1.Delete(0);
        }

        [TestMethod]
        public void UserDefinedCells_GetNames()
        {
            var page1 = GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 2, 2);

            Assert.AreEqual(0, UserDefinedCellsHelper.GetCount(s1));
            UserDefinedCellsHelper.Set(s1, "FOO1", "BAR1", null);
            Assert.AreEqual(1, UserDefinedCellsHelper.GetCount(s1));
            UserDefinedCellsHelper.Set(s1, "FOO1", "BAR2", null);
            Assert.AreEqual(1, UserDefinedCellsHelper.GetCount(s1));
            UserDefinedCellsHelper.Set(s1, "FOO2", "BAR3", null);

            var names1 = UserDefinedCellsHelper.GetNames(s1);
            Assert.AreEqual(2,names1.Count);
            Assert.IsTrue(names1.Contains("FOO1"));
            Assert.IsTrue(names1.Contains("FOO2"));

            Assert.AreEqual(2, UserDefinedCellsHelper.GetCount(s1));
            UserDefinedCellsHelper.Delete(s1, "FOO1");

            var names2 = UserDefinedCellsHelper.GetNames(s1);
            Assert.AreEqual(1, names2.Count);
            Assert.IsTrue(names2.Contains("FOO2"));

            UserDefinedCellsHelper.Set(s1, "FOO3", "BAR1", null);
            var names3 = UserDefinedCellsHelper.GetNames(s1);
            Assert.AreEqual(2, names3.Count);
            Assert.IsTrue(names3.Contains("FOO2"));
            Assert.IsTrue(names3.Contains("FOO3"));

            UserDefinedCellsHelper.Delete(s1, "FOO3");

            Assert.AreEqual(1, UserDefinedCellsHelper.GetCount(s1));
            UserDefinedCellsHelper.Delete(s1, "FOO2");

            Assert.AreEqual(0, UserDefinedCellsHelper.GetCount(s1));

            page1.Delete(0);
        }

        [TestMethod]
        public void UserDefinedCells_SetForMultipleShapes()
        {
            var page1 = GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 2, 2);
            var s2 = page1.DrawRectangle(0, 0, 2, 2);
            var s3 = page1.DrawRectangle(0, 0, 2, 2);
            var s4 = page1.DrawRectangle(0, 0, 2, 2);

            UserDefinedCellsHelper.Set(s1, "FOO1", "1", "p1");
            UserDefinedCellsHelper.Set(s2, "FOO2", "2", "p2");
            UserDefinedCellsHelper.Set(s2, "FOO3", "3", "p3");
            UserDefinedCellsHelper.Set(s4, "FOO4", "4", "p4");
            UserDefinedCellsHelper.Set(s4, "FOO5", "5", "p4");
            UserDefinedCellsHelper.Set(s4, "FOO6", "6", "p6");

            var shapeids = new[] {s1, s2, s3, s4};
            var allprops = UserDefinedCellsHelper.Get(page1, shapeids);

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
        public void UserDefinedCells_ValueQuoting()
        {
            var page1 = GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 2, 2);

            var p1 = UserDefinedCellsHelper.Get(s1);
            Assert.AreEqual(0, p1.Count);

            UserDefinedCellsHelper.Set(s1, "FOO1", "1", null);
            UserDefinedCellsHelper.Set(s1, "FOO2", "2", null);
            UserDefinedCellsHelper.Set(s1, "FOO3", "3\"4", null);

            var p2 = UserDefinedCellsHelper.Get(s1);
            Assert.AreEqual(3, p2.Count);
            
            Assert.AreEqual("FOO1",p2[0].Name);
            Assert.AreEqual("\"1\"", p2[0].Value);

            Assert.AreEqual("FOO2", p2[1].Name);
            Assert.AreEqual("\"2\"", p2[1].Value);

            Assert.AreEqual("FOO3", p2[2].Name);
            Assert.AreEqual("\"3\"\"4\"", p2[2].Value);

            page1.Delete(0);
        }
    }
}