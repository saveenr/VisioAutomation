using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using VisioAutomation.Text.Markup;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class TextHelper_Tests : VisioAutomationTest
    {
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
    }
}