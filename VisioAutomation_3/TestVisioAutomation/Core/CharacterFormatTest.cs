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
        public void SetCharacterSize()
        {
            double input_text_size = VA.Convert.PointsToInches(50);
            
            // Create a simple shape with no text and set the size to a specific value
            var page1 = GetNewPage();
            var shape0 = page1.DrawRectangle(1, 1, 3, 3);

            // set the text size
            var incharformat = new VA.Text.CharacterFormatCells();
            incharformat.Size = input_text_size;
            VisioAutomation.Text.TextHelper.SetFormat(incharformat, shape0);

            // retrieve the text size
            var query = new VA.ShapeSheet.Query.CellQuery();
            var col_charsize = query.AddColumn(VA.ShapeSheet.SRCConstants.Char_Size);
            var results = query.GetResults<double>(shape0);
            
            // before & after sizes should be the same
            Assert.AreEqual(input_text_size, results[0,col_charsize], 0.005);
            page1.Delete(0);
        }
        
        [TestMethod]
        public void SetMultipleCharacterSizes()
        {
            // Create a simple shape with text that has multiple character formatting rows
            // Then format the text and all character rows should be altered

            var page1 = GetNewPage();
            var shape0 = page1.DrawRectangle(1, 1, 3, 3);

            shape0.Text = TestVisioAutomation.TestHelper.LoremIpsumText;

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

            VisioAutomation.Text.TextHelper.SetFormat(shape0,fmt0 );
            VisioAutomation.Text.TextHelper.SetFormat(shape0,fmt1 , 10, 20);
            VisioAutomation.Text.TextHelper.SetFormat(shape0,fmt2 , 30, 40);
            
            // retrieve the text size
            var query = new VA.ShapeSheet.Query.SectionQuery(IVisio.VisSectionIndices.visSectionCharacter);
            query.AddColumn(IVisio.VisCellIndices.visCharacterSize);

            var table1 = query.GetResults<double>(shape0);

            // veriy all the sizes are present
            Assert.AreEqual(5,table1.Rows.Count);
            Assert.AreEqual(pts_10, table1[0, 0], 0.000000005);
            Assert.AreEqual(pts_6,  table1[1, 0], 0.000000005);
            Assert.AreEqual(pts_10, table1[2, 0], 0.000000005);
            Assert.AreEqual(pts_18, table1[3, 0], 0.000000005);
            Assert.AreEqual(pts_10, table1[4, 0], 0.000000005);

            // new replaces all the sizes with a single specific sizes
            // all the ranges will still exist but will all have the same size
            VisioAutomation.Text.TextHelper.SetFormat(fmt3, shape0);
            var table2 = query.GetResults<double>(shape0);

            Assert.AreEqual(5, table2.Rows.Count);
            Assert.AreEqual(pts_9, table2[0, 0], 0.000000005);
            Assert.AreEqual(pts_9, table2[1, 0], 0.000000005);
            Assert.AreEqual(pts_9, table2[2, 0], 0.000000005);
            Assert.AreEqual(pts_9, table2[3, 0], 0.000000005);
            Assert.AreEqual(pts_9, table2[4, 0], 0.000000005);


            // now retrieve with unit codes to verify that
            // our conversion of points and inches matches reality
            query.Columns[0].UnitCode = IVisio.VisUnitCodes.visPoints;
            var table3 = query.GetResults<double>(shape0);

            Assert.AreEqual(5, table3.Rows.Count);
            Assert.AreEqual(9.0, table3[0, 0], 0.000000005);
            Assert.AreEqual(9.0, table3[1, 0], 0.000000005);
            Assert.AreEqual(9.0, table3[2, 0], 0.000000005);
            Assert.AreEqual(9.0, table3[3, 0], 0.000000005);
            Assert.AreEqual(9.0, table3[4, 0], 0.000000005);

            page1.Delete(0);
        }
        
        [TestMethod]
        public void CheckTextRuns()
        {
            var page1 = GetNewPage();

            string original_text = "Lorem ipsum dolor sit amet";
            var s1 = page1.DrawRectangle(0, 0, 5, 5);
            s1.Text = original_text;

            var textruns0 = VA.Text.TextHelper.GetTextRuns(s1, IVisio.VisRunTypes.visCharPropRow, true);
            Assert.AreEqual(1, textruns0.Count);
            Assert.AreEqual(0, textruns0[0].Begin);
            Assert.AreEqual(original_text.Length + 1, textruns0[0].End);
            Assert.AreEqual(original_text, textruns0[0].Text);

            var charfmt1 = new VA.Text.CharacterFormatCells();
            charfmt1.Color = new VA.Drawing.ColorRGB(0xff0000).ToFormula();
            VA.Text.TextHelper.SetFormat(s1,  charfmt1, 0, 5);

            var textruns1 = VA.Text.TextHelper.GetTextRuns(s1, IVisio.VisRunTypes.visCharPropRow, true);
            Assert.AreEqual(2, textruns1.Count);
            Assert.AreEqual(0, textruns1[0].Begin);
            Assert.AreEqual(5, textruns1[0].End);
            Assert.AreEqual("Lorem", textruns1[0].Text);
            Assert.AreEqual(5, textruns1[1].Begin);
            Assert.AreEqual(27, textruns1[1].End);
            Assert.AreEqual(" ipsum dolor sit amet", textruns1[1].Text);

            var charfmt2 = new VA.Text.CharacterFormatCells();
            charfmt2.Style = (int)(VA.Text.CharStyle.Italic | VA.Text.CharStyle.UnderLine);
            VA.Text.TextHelper.SetFormat(s1, charfmt2, 5, 7);

            var textruns2 = VA.Text.TextHelper.GetTextRuns(s1, IVisio.VisRunTypes.visCharPropRow, true);
            Assert.AreEqual(3, textruns2.Count);
            Assert.AreEqual(0, textruns2[0].Begin);
            Assert.AreEqual(5, textruns2[0].End);
            Assert.AreEqual("Lorem", textruns2[0].Text);
            Assert.AreEqual(5, textruns2[1].Begin);
            Assert.AreEqual(7, textruns2[1].End);
            Assert.AreEqual(" i", textruns2[1].Text);
            Assert.AreEqual(7, textruns2[2].Begin);
            Assert.AreEqual(27, textruns2[2].End);
            Assert.AreEqual("psum dolor sit amet", textruns2[2].Text);
            page1.Delete(0);
        }
    }
}
