using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class TextHelper_Tests : VisioAutomationTest
    {
        private static readonly VA.Text.Markup.Field field_title = new VA.Text.Markup.Field(IVisio.VisFieldCategories.visFCatDocument, IVisio.VisFieldCodes.visFCodePrintDate, IVisio.VisFieldFormats.visFmtNumGenNoUnits);
        private static readonly VA.Text.Markup.Field field_shapewidth = new VA.Text.Markup.Field(IVisio.VisFieldCategories.visFCatGeometry, IVisio.VisFieldCodes.visFCodeBackgroundName, IVisio.VisFieldFormats.visFmtNumGenNoUnits);
 
        [TestMethod]
        public void Fields_Scenario_1()
        {
            var doc1 = this.GetNewDoc();
            var page1 = this.GetVisioApplication().ActivePage;
            doc1.Title = "Fields_Scenario_1";

            var s0 = page1.DrawRectangle(0, 0, 4, 1);
            VA.Text.TextHelper.SetTextFormatFields(s0, "DOCNAME: {0}", field_title);

            var s1 = page1.DrawRectangle(0, 1, 4, 2);
            VA.Text.TextHelper.SetTextFormatFields(s1, "SHAPE WIDTH: {0}", field_shapewidth);

            var s1_shape_size = VisioAutomationTest.GetSize(s1);

            var shape_area = page1.DrawRectangle(4, 1, 8, 2);
            VA.Text.TextHelper.SetTextFormatFields(shape_area, "SHAPEAREA: {0}", "Width*Height");

            var s0_characters = s0.Characters;
            Assert.AreEqual("DOCNAME: Fields_Scenario_1", s0_characters.Text);
            var s1_characters = s1.Characters;
            Assert.AreEqual(string.Format("SHAPE WIDTH: {0}", s1_shape_size.Width),s1_characters.Text);
            var s3_characters = shape_area.Characters;
            Assert.AreEqual("SHAPEAREA: 4", s3_characters.Text);

            page1.Delete(0);
            doc1.Close(true);
        }

        [TestMethod]
        public void Fields_Scenario_2()
        {
            var page1 = GetNewPage();
            var doc1 = page1.Document;

            var s1 = page1.DrawRectangle(0, 0, 4, 1);
            doc1.Title = "Fields_Scenario_2";
            VA.Text.TextHelper.SetTextFormatFields(s1, "DOCNAME: ", field_title);

            page1.Delete(0);
        }

        [TestMethod]
        public void Fields_Scenario_3()
        {
            bool caught = false;
            var page1 = GetNewPage();
            var doc1 = page1.Document;

            var s1 = page1.DrawRectangle(0, 0, 4, 1);
            doc1.Title = "Fields_Scenario_3";
            try
            {
                VA.Text.TextHelper.SetTextFormatFields(s1, "DOCNAME: {0} {1}", field_title);
            }
            catch (System.ArgumentOutOfRangeException )
            {
                // this is expected
                page1.Delete(0);
                caught = true;
            }

            if (!caught)
            {
                Assert.Fail("Did not catch expected exception");
            }

        }

        [TestMethod]
        public void FitShapeToText_Scenario_1()
        {
            var page1 = GetNewPage();
            var doc1 = page1.Document;

            var s1 = page1.DrawRectangle(0, 0, 0.1, 0.1);

            var font = doc1.Fonts["Arial"];
            var src_charfont = VisioAutomation.ShapeSheet.SRCConstants.Char_Font;
            var cell_charfont = s1.CellsSRC[src_charfont.Section, src_charfont.Row, src_charfont.Cell];
            cell_charfont.FormulaU = font.ID.ToString(System.Globalization.CultureInfo.InvariantCulture);
            s1.Text = TestCommon.Helper.LoremIpsumText;

            VisioAutomation.Text.TextHelper.FitShapeToText(page1, new[] { s1 });

            var new_size = VisioAutomationTest.GetSize(s1);
            Assert.AreEqual(0.138835906982422, new_size.Width, 0.0001);
            Assert.AreEqual(89.6006546020508, new_size.Height, 0.0001);

            page1.Delete(0);
        }

        [TestMethod]
        public void SetTabStops()
        {
            var no_tab_stops = new VA.Text.TabStop[] { };
            var tabstops = new[]
                               {
                                   new VA.Text.TabStop(0.5, VA.Text.TabStopAlignment.Left),
                                   new VA.Text.TabStop(1.5, VA.Text.TabStopAlignment.Right),
                                   new VA.Text.TabStop(2.5, VA.Text.TabStopAlignment.Center),
                                   new VA.Text.TabStop(4.0, VA.Text.TabStopAlignment.Decimal)
                               };

            var page1 = GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 4, 4);

            // shapes shoould not have any tab stops by default
            var m0 = VA.Text.TextFormat.GetFormat(s1);
            Assert.AreEqual(0, m0.TabStops.Count);

            // clearing tab stops shoudl work even if there are no tab stops
            VA.Text.TextFormat.SetTabStops(s1, no_tab_stops);
            var m1 = VA.Text.TextFormat.GetFormat(s1);
            Assert.AreEqual(0, m1.TabStops.Count);

            // set the 3 tab stops
            VA.Text.TextFormat.SetTabStops(s1, tabstops);

            // should have exactly the same number we set
            var m2 = VA.Text.TextFormat.GetFormat(s1);
            Assert.AreEqual(tabstops.Length, m2.TabStops.Count);

            // clear the tab stops
            VA.Text.TextFormat.SetTabStops(s1, no_tab_stops);
            var m3 = VA.Text.TextFormat.GetFormat(s1);
            Assert.AreEqual(0, m3.TabStops.Count);

            page1.Delete(0);
        }
    }
}