using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class CharacterFormatTest : VisioAutomationTest
    {       
        [TestMethod]
        public void SetCharacterSizeMultupleRegions()
        {
            // Create a simple shape with text that has multiple character formatting rows
            // Then format the text and all character rows should be altered

            var page1 = GetNewPage();
            var shape0 = page1.DrawRectangle(1, 1, 3, 3);

            shape0.Text = TestCommon.Helper.LoremIpsumText;

            var fmt0 = new VA.Text.CharacterFormatCells();
            var pts_10 = VA.Convert.PointsToInches(10);
            fmt0.Size = pts_10;

            var fmt1 = new VA.Text.CharacterFormatCells();
            var pts_6 = VA.Convert.PointsToInches(6);
            fmt1.Size = pts_6;

            var fmt2 = new VA.Text.CharacterFormatCells();
            var pts_18 = VA.Convert.PointsToInches(18);
            fmt2.Size = pts_18;

            var fmt3 = new VA.Text.CharacterFormatCells();
            var pts_9 = VA.Convert.PointsToInches(9);
            fmt3.Size = pts_9;

            VisioAutomation.Text.TextFormat.Format(shape0, fmt0);
            VisioAutomation.Text.TextFormat.FormatRange(shape0, fmt1, 10, 20);
            VisioAutomation.Text.TextFormat.FormatRange(shape0, fmt2, 30, 40);

            // retrieve the text size
            var out_formats1 = VA.Text.TextFormat.GetFormat(shape0);


            // veriy all the sizes are present
            Assert.AreEqual(5,out_formats1.CharacterFormats.Count);
            Assert.AreEqual(pts_10, out_formats1.CharacterFormats[0].Size.Result, 0.000000005);
            Assert.AreEqual(pts_6, out_formats1.CharacterFormats[1].Size.Result, 0.000000005);
            Assert.AreEqual(pts_10, out_formats1.CharacterFormats[2].Size.Result, 0.000000005);
            Assert.AreEqual(pts_18, out_formats1.CharacterFormats[3].Size.Result, 0.000000005);
            Assert.AreEqual(pts_10, out_formats1.CharacterFormats[4].Size.Result, 0.000000005);

            // new replaces all the sizes with a single specific sizes
            // all the ranges will still exist but will all have the same size
            VisioAutomation.Text.TextFormat.Format(shape0, fmt3);
            var out_formats2 = VA.Text.TextFormat.GetFormat(shape0);

            Assert.AreEqual(5, out_formats2.CharacterFormats.Count);
            Assert.AreEqual(pts_9, out_formats2.CharacterFormats[0].Size.Result, 0.000000005);
            Assert.AreEqual(pts_9, out_formats2.CharacterFormats[1].Size.Result, 0.000000005);
            Assert.AreEqual(pts_9, out_formats2.CharacterFormats[2].Size.Result, 0.000000005);
            Assert.AreEqual(pts_9, out_formats2.CharacterFormats[3].Size.Result, 0.000000005);
            Assert.AreEqual(pts_9, out_formats2.CharacterFormats[4].Size.Result, 0.000000005);


            // now retrieve with unit codes to verify that
            // our conversion of points and inches matches reality
            var out_formats3 = VA.Text.TextFormat.GetFormat(shape0);
            var inches_for_9pts = VA.Convert.PointsToInches(9.0);
            Assert.AreEqual(5, out_formats3.CharacterFormats.Count);
            Assert.AreEqual(inches_for_9pts, out_formats3.CharacterFormats[0].Size.Result, 0.000000005);
            Assert.AreEqual(inches_for_9pts, out_formats3.CharacterFormats[1].Size.Result, 0.000000005);
            Assert.AreEqual(inches_for_9pts, out_formats3.CharacterFormats[2].Size.Result, 0.000000005);
            Assert.AreEqual(inches_for_9pts, out_formats3.CharacterFormats[3].Size.Result, 0.000000005);
            Assert.AreEqual(inches_for_9pts, out_formats3.CharacterFormats[4].Size.Result, 0.000000005);

            page1.Delete(0);
        }
        
        [TestMethod]
        public void FormatCharactersAndCheckTextRuns()
        {
            var page1 = GetNewPage();

            string original_text = "Lorem ipsum dolor sit amet";
            var s1 = page1.DrawRectangle(0, 0, 5, 5);
            s1.Text = original_text;

            var textruns0 = VA.Text.TextFormat.GetFormat(s1);
            Assert.AreEqual(1, textruns0.CharacterTextRuns.Count);
            Assert.AreEqual(0, textruns0.CharacterTextRuns[0].Begin);
            Assert.AreEqual(original_text.Length + 1, textruns0.CharacterTextRuns[0].End);
            Assert.AreEqual(original_text, textruns0.CharacterTextRuns[0].Text);

            var charfmt1 = new VA.Text.CharacterFormatCells();
            charfmt1.Color = new VA.Drawing.ColorRGB(0xff0000).ToFormula();
            VA.Text.TextFormat.FormatRange(s1, charfmt1, 0, 5);

            var textruns1 = VA.Text.TextFormat.GetFormat(s1);
            Assert.AreEqual(2, textruns1.CharacterTextRuns.Count);
            Assert.AreEqual(0, textruns1.CharacterTextRuns[0].Begin);
            Assert.AreEqual(5, textruns1.CharacterTextRuns[0].End);
            Assert.AreEqual("Lorem", textruns1.CharacterTextRuns[0].Text);
            Assert.AreEqual(5, textruns1.CharacterTextRuns[1].Begin);
            Assert.AreEqual(27, textruns1.CharacterTextRuns[1].End);
            Assert.AreEqual(" ipsum dolor sit amet", textruns1.CharacterTextRuns[1].Text);

            var charfmt2 = new VA.Text.CharacterFormatCells();
            charfmt2.Style = (int)(VA.Text.CharStyle.Italic | VA.Text.CharStyle.UnderLine);
            VA.Text.TextFormat.FormatRange(s1, charfmt2, 5, 7);

            var textruns2 = VA.Text.TextFormat.GetFormat(s1);
            Assert.AreEqual(3, textruns2.CharacterTextRuns.Count);
            Assert.AreEqual(0, textruns2.CharacterTextRuns[0].Begin);
            Assert.AreEqual(5, textruns2.CharacterTextRuns[0].End);
            Assert.AreEqual("Lorem", textruns2.CharacterTextRuns[0].Text);
            Assert.AreEqual(5, textruns2.CharacterTextRuns[1].Begin);
            Assert.AreEqual(7, textruns2.CharacterTextRuns[1].End);
            Assert.AreEqual(" i", textruns2.CharacterTextRuns[1].Text);
            Assert.AreEqual(7, textruns2.CharacterTextRuns[2].Begin);
            Assert.AreEqual(27, textruns2.CharacterTextRuns[2].End);
            Assert.AreEqual("psum dolor sit amet", textruns2.CharacterTextRuns[2].Text);
            page1.Delete(0);
        }
    }
}
