using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using VisioAutomation.Extensions;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace TestVisioAutomation
{
    [TestClass]
    public class TextMarkupTests : VisioAutomationTest
    {
        [TestMethod]
        public void ValidateFormattingRegions()
        {
            // Check that the formatting regions are correctly
            // mapped given a number of Text structures

            var el1 = new VA.Text.Markup.TextElement();
            var markup1 = el1.GetMarkupInfo();
            var regions1 = markup1.FormatRegions;
            Assert.AreEqual(1, markup1.FormatRegions.Count);
            Assert.AreEqual(0, regions1[0].TextLength);
            Assert.AreEqual(0, regions1[0].TextStartPos);
            Assert.AreEqual(0, regions1[0].TextEndPos);


            var el2 = new VA.Text.Markup.TextElement("HELLO");
            var markup2 = el2.GetMarkupInfo();
            var regions2 = markup2.FormatRegions;
            Assert.AreEqual(1, markup2.FormatRegions.Count);
            Assert.AreEqual(5, regions2[0].TextLength);
            Assert.AreEqual(0, regions2[0].TextStartPos);
            Assert.AreEqual(5, regions2[0].TextEndPos);

            var el3 = new VA.Text.Markup.TextElement("HELLO");
            el3.AppendText(" WORLD");
            var markup3 = el3.GetMarkupInfo();
            var regions3 = markup3.FormatRegions;
            Assert.AreEqual(1, markup3.FormatRegions.Count);
            Assert.AreEqual(11, regions3[0].TextLength);
            Assert.AreEqual(0, regions3[0].TextStartPos);
            Assert.AreEqual(11, regions3[0].TextEndPos);

            var el4 = new VA.Text.Markup.TextElement();
            el4.AppendElement("HELLO");
            el4.AppendElement(" WORLD");
            var markup4 = el4.GetMarkupInfo();
            var regions4 = markup4.FormatRegions;
            Assert.AreEqual(3, markup4.FormatRegions.Count);
            Assert.AreEqual(11, regions4[0].TextLength);
            Assert.AreEqual(0, regions4[0].TextStartPos);
            Assert.AreEqual(11, regions4[0].TextEndPos);
            Assert.AreEqual(5, regions4[1].TextLength);
            Assert.AreEqual(0, regions4[1].TextStartPos);
            Assert.AreEqual(5, regions4[1].TextEndPos);
            Assert.AreEqual(6, regions4[2].TextLength);
            Assert.AreEqual(5, regions4[2].TextStartPos);
            Assert.AreEqual(11, regions4[2].TextEndPos);


            var el5 = new VA.Text.Markup.TextElement();
            var el5_a = el5.AppendElement("HELLO");
            var el5_b = el5_a.AppendElement(" WORLD");

            var markup5 = el5.GetMarkupInfo();
            var regions5 = markup5.FormatRegions;
            Assert.AreEqual(3, markup5.FormatRegions.Count);
            Assert.AreEqual(11, regions5[0].TextLength);
            Assert.AreEqual(0, regions5[0].TextStartPos);
            Assert.AreEqual(11, regions5[0].TextEndPos);
            Assert.AreEqual(11, regions5[1].TextLength);
            Assert.AreEqual(0, regions5[1].TextStartPos);
            Assert.AreEqual(11, regions5[1].TextEndPos);
            Assert.AreEqual(6, regions5[2].TextLength);
            Assert.AreEqual(5, regions5[2].TextStartPos);
            Assert.AreEqual(11, regions5[2].TextEndPos);

        }


        [TestMethod]
        public void TextElement_with_multiple_text_nodes()
        {
            // Validate that multiple text elements in the structure
            // all make it into a real visio shep when the text is render

            var el0 = new VA.Text.Markup.TextElement();
            var el1 = el0.AppendElement("HELLO");
            var el2 = el0.AppendElement(" WORLD");

            var page1 = GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 4, 4);

            el0.SetText(s1);

            Assert.AreEqual("HELLO WORLD", s1.Text);

            page1.Delete(0);
        }

        [TestMethod]
        public void Element_with_bold_and_italic_text()
        {
            // Validate that basic formatting works when rendering

            var el0 = new VA.Text.Markup.TextElement();
            var el1 = el0.AppendElement("HELLO");
            var el2 = el0.AppendElement(" WORLD");

            el1.CharacterFormat.CharStyle = VA.Text.CharStyle.Bold;
            el2.CharacterFormat.CharStyle = VA.Text.CharStyle.Italic;

            var page1 = GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 4, 4);

            el0.SetText(s1);

            var fmts = VA.Text.TextFormat.GetFormat(s1);
            Assert.AreEqual(3, fmts.CharacterFormats.Count);
            Assert.AreEqual((int)VA.Text.CharStyle.Bold, fmts.CharacterFormats[0].Style.Result);
            Assert.AreEqual((int)VA.Text.CharStyle.Italic, fmts.CharacterFormats[1].Style.Result);
            Assert.AreEqual((int)VA.Text.CharStyle.None, fmts.CharacterFormats[2].Style.Result);

            page1.Delete(0);
        }

        [TestMethod]
        public void Style_inheritance()
        {
            // Validate that sub elements inherit the formatting of parent elements

            var el0 = new VA.Text.Markup.TextElement();
            var el1 = el0.AppendElement("HELLO");
            var el2 = el1.AppendElement(" WORLD");

            el0.CharacterFormat.FontSize = 14;
            el0.CharacterFormat.FontSize = 7;
            
            el1.CharacterFormat.Font = "Impact";
            el1.CharacterFormat.CharStyle = VA.Text.CharStyle.Bold;
            
            el2.CharacterFormat.Font = "Courier New";
            el2.CharacterFormat.FontSize = 20;
            el2.CharacterFormat.CharStyle = VA.Text.CharStyle.Italic;

            var markup = el0.GetMarkupInfo();
            var regions = markup.FormatRegions;
            Assert.AreEqual(3, regions.Count);
            Assert.AreEqual(6, regions[2].TextLength);
            Assert.AreEqual(5, regions[2].TextStartPos);
            Assert.AreEqual(11, regions[2].TextEndPos);
            Assert.AreEqual(11, regions[1].TextLength);
            Assert.AreEqual(0, regions[1].TextStartPos);
            Assert.AreEqual(11, regions[1].TextEndPos);
            Assert.AreEqual(11, regions[0].TextLength);
            Assert.AreEqual(0, regions[0].TextStartPos);
            Assert.AreEqual(11, regions[0].TextEndPos);

            Assert.AreEqual(el0, regions[0].Element);
            Assert.AreEqual(el1, regions[1].Element);
            Assert.AreEqual(el2, regions[2].Element);

            var page1 = GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 4, 4);

            el0.SetText(s1);

            var fmts = VA.Text.TextFormat.GetFormat(s1);
            Assert.AreEqual(3, fmts.CharacterFormats.Count);

            page1.Delete(0);
        }

        [TestMethod]
        public void Test_Format_Text_field()
        {
            // Now account for field insertion

            var el0 = new VA.Text.Markup.TextElement();
            el0.AppendText("HELLO ");
            el0.AppendField(VA.Text.Markup.Fields.Height);
            el0.AppendText(" WORLD");

            var page1 = GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 4, 4);

            string it = el0.GetInnerText();
            Assert.AreEqual("HELLO " + VA.Text.Markup.Fields.Height.PlaceholderText + " WORLD", it);
            el0.SetText(s1);

            var shape_size = VisioAutomationTest.GetSize(s1);

            string s = string.Format("HELLO {0} WORLD", shape_size.Height);
            var s1_characters = s1.Characters;
            Assert.AreEqual(s, s1_characters.Text);

            page1.Delete(0);
        }

        [TestMethod]
        public void CharacterFormatCells_Check_SetFormat_1()
        {
            var page1 = GetNewPage(new VA.Drawing.Size(10, 10));
            var s1 = page1.DrawRectangle(0, 0, 10, 10);

            var sb = new System.Text.StringBuilder();
            for (int y = 0; y < 10; y++)
            {
                for (int x = 0; x < 10; x++)
                {
                    int n = (y * 10 + x) % 5;
                    sb.Append(n.ToString());
                }
            }
            s1.Text = sb.ToString();


            var c0 = new VA.DOM.ShapeCells();
            c0.CharSize = 0.6;
            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();
            c0.Apply(update, s1.ID16);
            update.Execute(page1);

            var c1 = new VA.Text.CharacterFormatCells();
            c1.Color = new VA.Drawing.ColorRGB(0xff0000).ToFormula();
            VA.Text.TextFormat.FormatRange(s1, c1, 0, 5);

            var c2 = new VA.Text.CharacterFormatCells();
            c2.Size = 0.5;
            VA.Text.TextFormat.FormatRange(s1, c2, 5, 10);

            var c3 = new VA.Text.CharacterFormatCells();
            c3.Font = page1.Document.Fonts["Impact"].ID;
            VA.Text.TextFormat.FormatRange(s1, c3, 10, 15);

            var c4 = new VA.Text.CharacterFormatCells();
            c4.Style = (int) (VA.Text.CharStyle.Italic | VA.Text.CharStyle.UnderLine);
            VA.Text.TextFormat.FormatRange(s1, c4, 15, 20);

            var c5 = new VA.Text.CharacterFormatCells();
            c5.Transparency = 0.5;
            VA.Text.TextFormat.FormatRange(s1, c5, 20, 25);

            var formatting = VA.Text.TextFormat.GetFormat(s1);
            Assert.AreEqual("RGB(255,0,0)", formatting.CharacterFormats[0].Color.Formula);
            Assert.AreEqual(0.5, formatting.CharacterFormats[1].Size.Result);
            Assert.AreEqual(page1.Document.Fonts["Impact"].ID, formatting.CharacterFormats[2].Font.Result);
            Assert.AreEqual("6", formatting.CharacterFormats[3].Style.Formula);
            Assert.AreEqual("50%", formatting.CharacterFormats[4].Transparency.Formula);
            Assert.AreEqual(0.6, formatting.CharacterFormats[5].Size.Result);

            //page1.Delete(0);
        }

        [TestMethod]
        public void CharacterFormatCells_Check_SetFormat_2()
        {
            var page1 = GetNewPage(new VA.Drawing.Size(10, 10));
            IVisio.Shape s1;

            var fmt = new VA.Text.CharacterFormatCells();

            s1 = page1.DrawRectangle(0,0,1,1);
            s1.Text = "Plain";

            fmt = new VA.Text.CharacterFormatCells();
            s1 = page1.DrawRectangle(1,1,2,2);
            s1.Text = "Bold";
            fmt.Style = (int) VA.Text.CharStyle.Bold;
            VA.Text.TextFormat.Format(s1, fmt);

            fmt = new VA.Text.CharacterFormatCells();
            s1 = page1.DrawRectangle(2,2,3,3);
            s1.Text = "Italic";
            fmt.Style = (int)VA.Text.CharStyle.Italic;
            VA.Text.TextFormat.Format(s1, fmt);

            fmt = new VA.Text.CharacterFormatCells();
            s1 = page1.DrawRectangle(3,3,4,4);
            s1.Text = "Underline";
            fmt.Style = (int)VA.Text.CharStyle.UnderLine;
            VA.Text.TextFormat.Format(s1, fmt);

            fmt = new VA.Text.CharacterFormatCells();
            s1 = page1.DrawRectangle(4,4,5,5);
            s1.Text = "Smallcaps";
            fmt.Style = (int)VA.Text.CharStyle.SmallCaps;
            VA.Text.TextFormat.Format(s1, fmt);

            fmt = new VA.Text.CharacterFormatCells();
            s1 = page1.DrawRectangle(5,5,6,6);
            s1.Text = "Red";
            fmt.Color = new VA.Drawing.ColorRGB(0xff0000).ToFormula();
            VA.Text.TextFormat.Format(s1, fmt);

            fmt = new VA.Text.CharacterFormatCells();
            s1 = page1.DrawRectangle(6,6,7,7);
            s1.Text = "#ec35a7";
            fmt.Color = new VA.Drawing.ColorRGB(0xec35a7).ToFormula();
            VA.Text.TextFormat.Format(s1, fmt);

            fmt = new VA.Text.CharacterFormatCells();
            s1 = page1.DrawRectangle(7,7,8,8);
            s1.Text = "#34f178";
            fmt.Color = new VA.Drawing.ColorRGB(0x34f178).ToFormula();
            VA.Text.TextFormat.Format(s1, fmt);

            fmt = new VA.Text.CharacterFormatCells();
            s1 = page1.DrawRectangle(8,8,9,9);
            s1.Text = "Calibri";
            fmt.Font = page1.Document.Fonts["Calibri"].ID;
            VA.Text.TextFormat.Format(s1, fmt);

            fmt = new VA.Text.CharacterFormatCells();
            s1 = page1.DrawRectangle(9,9,10,10);
            s1.Text = "Impact";
            fmt.Font = page1.Document.Fonts["Impact"].ID;
            VA.Text.TextFormat.Format(s1, fmt);

            fmt = new VA.Text.CharacterFormatCells();
            s1 = page1.DrawRectangle(10,10,11,11);
            s1.Text = "Segoe UI";
            fmt.Font = page1.Document.Fonts["Segoe UI"].ID;
            VA.Text.TextFormat.Format(s1, fmt);

            fmt = new VA.Text.CharacterFormatCells();
            s1 = page1.DrawRectangle(11,11,12,12);
            s1.Text = "6pt";
            fmt.Size = VA.Convert.PointsToInches(6);
            VA.Text.TextFormat.Format(s1, fmt);

            fmt = new VA.Text.CharacterFormatCells();
            s1 = page1.DrawRectangle(12,12,13,13);
            s1.Text = "8pt";
            fmt.Size = VA.Convert.PointsToInches(8);
            VA.Text.TextFormat.Format(s1, fmt);

            fmt = new VA.Text.CharacterFormatCells();
            s1 = page1.DrawRectangle(13,13,14,14);
            s1.Text = "11pt";
            fmt.Size = VA.Convert.PointsToInches(8);
            VA.Text.TextFormat.Format(s1, fmt);

            fmt = new VA.Text.CharacterFormatCells();
            s1 = page1.DrawRectangle(14,14,15,15);
            s1.Text = "15pt";
            fmt.Size = VA.Convert.PointsToInches(15);
            VA.Text.TextFormat.Format(s1, fmt);

            page1.Delete(0);
        }

        [TestMethod]
        public void ParagraphFormatCells_Check_SetFormat_1()
        {
            var page1 = GetNewPage(new VA.Drawing.Size(10, 10));

            var s1 = page1.DrawRectangle(0,0,5,5);
            s1.Text = "Line1\nLine2\nLine3\nLine4\nLine5\nLine6";

            var formats0 = VA.Text.TextFormat.GetFormat(s1);
            Assert.AreEqual(1, formats0.ParagraphFormats.Count);

            var fmt1 = new VA.Text.ParagraphFormatCells();
            fmt1.IndentLeft = 0.25;

            var cfmt1 = new VA.Text.CharacterFormatCells();
            cfmt1.Color = "RGB(255,0,0)";


            VA.Text.TextFormat.FormatRange(s1, cfmt1, 2, 3);
            VA.Text.TextFormat.FormatRange(s1, fmt1, 2, 3);

            var formats1 = VA.Text.TextFormat.GetFormat(s1);
            Assert.AreEqual(2, formats1.ParagraphFormats.Count);

            var fmt2 = new VA.Text.ParagraphFormatCells();
            fmt2.BulletIndex = 2;

            var cfmt2 = new VA.Text.CharacterFormatCells();
            cfmt2.Color = "RGB(0,255,0)";

            VA.Text.TextFormat.FormatRange(s1, cfmt2, 13, 14);
            VA.Text.TextFormat.FormatRange(s1, fmt2, 13, 14);

            var formats2 = VA.Text.TextFormat.GetFormat(s1);
            Assert.AreEqual(4, formats2.ParagraphFormats.Count);

            Assert.AreEqual(0.25, formats2.ParagraphFormats[0].IndentLeft.Result);
            Assert.AreEqual(0, formats2.ParagraphFormats[0].IndentFirst.Result);

            Assert.AreEqual(0, formats2.ParagraphFormats[2].IndentLeft.Result);
            Assert.AreEqual(0, formats2.ParagraphFormats[2].IndentFirst.Result);

            Assert.AreEqual(2, formats2.ParagraphFormats[1].BulletIndex.Result);

            Assert.AreEqual(0, formats2.ParagraphFormats[1].IndentLeft.Result);
            Assert.AreEqual(0, formats2.ParagraphFormats[2].BulletIndex.Result);

            page1.Delete(0);
        }

        [TestMethod]
        public void TextBlockFormatCells_Check_SetFormat_1()
        {
            var page1 = GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 4, 4);
            var s2 = page1.DrawRectangle(5, 5, 7, 7);

            var tf0 = VA.Text.TextFormat.GetFormat(s1);
            Assert.AreEqual("4 pt",tf0.TextBlocks.BottomMargin.Formula);

            var tb1 = new VA.Text.TextBlockFormatCells();
            tb1.BottomMargin = "8 pt";

            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();
            tb1.Apply(update,s1.ID16);
            update.Execute(page1);

            var tf2 = VA.Text.TextFormat.GetFormat(s1);
            Assert.AreEqual("8 pt", tf2.TextBlocks.BottomMargin.Formula);

            var tfs = VA.Text.TextFormat.GetFormat(page1, new[] { s1.ID, s2.ID });
            Assert.AreEqual("8 pt", tfs[0].TextBlocks.BottomMargin.Formula);
            Assert.AreEqual("4 pt", tfs[1].TextBlocks.BottomMargin.Formula);


            page1.Delete(0);
        }

    }
}