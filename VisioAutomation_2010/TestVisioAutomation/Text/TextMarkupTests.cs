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
        public void Markup1()
        {
            // Validate that setting text with no values works
            var el0 = new VA.Text.Markup.TextElement("HELLO");

            var page1 = GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 4, 4);
            el0.SetText(s1);

            var fmts = VA.Text.TextFormat.GetFormat(s1);
            Assert.AreEqual(1, fmts.CharacterFormats.Count);
            Assert.AreEqual(1, fmts.ParagraphFormats.Count);

            page1.Delete(0);
        }

        [TestMethod]
        public void Markup2()
        {
            // Validate that setting text with no values works
            var el0 = new VA.Text.Markup.TextElement("HELLO");
            var color_red = new VA.Drawing.ColorRGB(0xff0000);
            el0.CharacterCells.Color = color_red.ToFormula();

            var page1 = GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 4, 4);
            el0.SetText(s1);

            var fmts = VA.Text.TextFormat.GetFormat(s1);
            Assert.AreEqual(1, fmts.CharacterFormats.Count);
            Assert.AreEqual(1, fmts.ParagraphFormats.Count);

            Assert.AreEqual("RGB(255,0,0)", fmts.CharacterFormats[0].Color.Formula);

            page1.Delete(0);
        }


        [TestMethod]
        public void ValidateFormattingRegions()
        {
            // Check that the formatting regions are correctly
            // mapped given a number of Text structures

            var el1 = new VA.Text.Markup.TextElement();
            var markup1 = el1.GetMarkupInfo();
            var regions1 = markup1.FormatRegions;
            Assert.AreEqual(1, markup1.FormatRegions.Count);
            Assert.AreEqual(0, regions1[0].Length);
            Assert.AreEqual(0, regions1[0].Start);
            Assert.AreEqual(0, regions1[0].End);


            var el2 = new VA.Text.Markup.TextElement("HELLO");
            var markup2 = el2.GetMarkupInfo();
            var regions2 = markup2.FormatRegions;
            Assert.AreEqual(1, markup2.FormatRegions.Count);
            Assert.AreEqual(5, regions2[0].Length);
            Assert.AreEqual(0, regions2[0].Start);
            Assert.AreEqual(5, regions2[0].End);

            var el3 = new VA.Text.Markup.TextElement("HELLO");
            el3.AddText(" WORLD");
            var markup3 = el3.GetMarkupInfo();
            var regions3 = markup3.FormatRegions;
            Assert.AreEqual(1, markup3.FormatRegions.Count);
            Assert.AreEqual(11, regions3[0].Length);
            Assert.AreEqual(0, regions3[0].Start);
            Assert.AreEqual(11, regions3[0].End);

            var el4 = new VA.Text.Markup.TextElement();
            el4.AddElement("HELLO");
            el4.AddElement(" WORLD");
            var markup4 = el4.GetMarkupInfo();
            var regions4 = markup4.FormatRegions;
            Assert.AreEqual(3, markup4.FormatRegions.Count);
            Assert.AreEqual(11, regions4[0].Length);
            Assert.AreEqual(0, regions4[0].Start);
            Assert.AreEqual(11, regions4[0].End);
            Assert.AreEqual(5, regions4[1].Length);
            Assert.AreEqual(0, regions4[1].Start);
            Assert.AreEqual(5, regions4[1].End);
            Assert.AreEqual(6, regions4[2].Length);
            Assert.AreEqual(5, regions4[2].Start);
            Assert.AreEqual(11, regions4[2].End);


            var el5 = new VA.Text.Markup.TextElement();
            var el5_a = el5.AddElement("HELLO");
            var el5_b = el5_a.AddElement(" WORLD");

            var markup5 = el5.GetMarkupInfo();
            var regions5 = markup5.FormatRegions;
            Assert.AreEqual(3, markup5.FormatRegions.Count);
            Assert.AreEqual(11, regions5[0].Length);
            Assert.AreEqual(0, regions5[0].Start);
            Assert.AreEqual(11, regions5[0].End);
            Assert.AreEqual(11, regions5[1].Length);
            Assert.AreEqual(0, regions5[1].Start);
            Assert.AreEqual(11, regions5[1].End);
            Assert.AreEqual(6, regions5[2].Length);
            Assert.AreEqual(5, regions5[2].Start);
            Assert.AreEqual(11, regions5[2].End);

        }


        [TestMethod]
        public void TextElement_with_multiple_text_nodes()
        {
            // Validate that multiple text elements in the structure
            // all make it into the Visio shape when the text is rendered

            var el0 = new VA.Text.Markup.TextElement();
            var el1 = el0.AddElement("HELLO");
            var el2 = el0.AddElement(" WORLD");

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
            var el1 = el0.AddElement("HELLO");
            var el2 = el0.AddElement(" WORLD");

            el1.CharacterCells.Style = (int)VA.Text.CharStyle.Bold;
            el2.CharacterCells.Style = (int)VA.Text.CharStyle.Italic;

            var page1 = GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 4, 4);
            el0.SetText(s1);

            var fmts = VA.Text.TextFormat.GetFormat(s1);
            Assert.AreEqual(3, fmts.CharacterFormats.Count);
            Assert.AreEqual((int)VA.Text.CharStyle.Bold, fmts.CharacterFormats[0].Style.Result);
            Assert.AreEqual((int)VA.Text.CharStyle.Italic, fmts.CharacterFormats[1].Style.Result);
                        
            // The code line below returns a different value in Visio 2007. 
            // Commenting-out that line to keep unit tests consistent
            // Assert.AreEqual((int)VA.Text.CharStyle.Bold, fmts.CharacterFormats[2].Style.Result);

            page1.Delete(0);
        }

        [TestMethod]
        public void Style_inheritance()
        {
            // Validate that sub elements inherit the formatting of parent elements
            var page1 = GetNewPage();
            var courier = page1.Document.Fonts["Courier New"];
            var impact = page1.Document.Fonts["Impact"];
            
            var el0 = new VA.Text.Markup.TextElement();
            var el1 = el0.AddElement("HELLO");
            var el2 = el1.AddElement(" WORLD");

            el0.CharacterCells.Font = "14pt";
            el0.CharacterCells.Size = "7pt";
            
            el1.CharacterCells.Font = impact.ID;
            el1.CharacterCells.Style = (int) VA.Text.CharStyle.Bold;
            
            el2.CharacterCells.Font = courier.ID;
            el2.CharacterCells.Size = "20pt";
            el2.CharacterCells.Style = (int) VA.Text.CharStyle.Italic;

            var markup = el0.GetMarkupInfo();
            var regions = markup.FormatRegions;
            Assert.AreEqual(3, regions.Count);
            Assert.AreEqual(6, regions[2].Length);
            Assert.AreEqual(5, regions[2].Start);
            Assert.AreEqual(11, regions[2].End);
            Assert.AreEqual(11, regions[1].Length);
            Assert.AreEqual(0, regions[1].Start);
            Assert.AreEqual(11, regions[1].End);
            Assert.AreEqual(11, regions[0].Length);
            Assert.AreEqual(0, regions[0].Start);
            Assert.AreEqual(11, regions[0].End);

            Assert.AreEqual(el0, regions[0].Element);
            Assert.AreEqual(el1, regions[1].Element);
            Assert.AreEqual(el2, regions[2].Element);

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
            el0.AddText("HELLO ");
            el0.AddField(VA.Text.Markup.FieldConstants.Height);
            el0.AddText(" WORLD");

            var page1 = GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 4, 4);

            string it = el0.GetInnerText();
            Assert.AreEqual("HELLO " + VA.Text.Markup.FieldConstants.Height.PlaceholderText + " WORLD", it);
            el0.SetText(s1);

            var shape_size = VisioAutomationTest.GetSize(s1);

            string s = string.Format("HELLO {0} WORLD", shape_size.Height);
            var s1_characters = s1.Characters;
            Assert.AreEqual(s, s1_characters.Text);

            page1.Delete(0);
        }


        [TestMethod]
        public void TextBlockFormatCells_Check_SetFormat_1()
        {
            var page1 = GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 4, 4);
            var s2 = page1.DrawRectangle(5, 5, 7, 7);

            var tf0 = VA.Text.TextFormat.GetFormat(s1);
            Assert.AreEqual("4 pt",tf0.TextBlock.BottomMargin.Formula);

            var tb1 = new VA.Text.TextBlockFormatCells();
            tb1.BottomMargin = "8 pt";

            var update = new VA.ShapeSheet.Update();
            tb1.Apply(update,s1.ID16);
            update.Execute(page1);

            var tf2 = VA.Text.TextFormat.GetFormat(s1);
            Assert.AreEqual("8 pt", tf2.TextBlock.BottomMargin.Formula);

            var tfs = VA.Text.TextFormat.GetFormat(page1, new[] { s1.ID, s2.ID });
            Assert.AreEqual("8 pt", tfs[0].TextBlock.BottomMargin.Formula);
            Assert.AreEqual("4 pt", tfs[1].TextBlock.BottomMargin.Formula);

            page1.Delete(0);
        }

        [TestMethod]
        public void TextXformCells_Check_SetFormat_1()
        {
            var page1 = GetNewPage();
            var s1 = page1.DrawRectangle(0, 0, 4, 4);
            s1.Text = TestCommon.Helper.LoremIpsumText;
            
            var tb1 = new VA.Text.TextXFormCells();
            tb1.TxtAngle = "20 deg";
            tb1.TxtPinX = "Width*1.3";
            tb1.TxtPinY = "Height*0.5";
            tb1.TxtLocPinX = "TxtWidth*0.3";
            tb1.TxtLocPinY = "TxtHeight*0.4";
            tb1.TxtHeight = "Height*1.5";
            tb1.TxtWidth = "Width*0.7";

            var update = new VA.ShapeSheet.Update();
            tb1.Apply(update, s1.ID16);
            update.Execute(page1);

            var tb2 = VA.Text.TextXFormCells.GetCells(s1);
            Assert.AreEqual(tb1.TxtAngle.Formula,tb2.TxtAngle.Formula);
            Assert.AreEqual(tb1.TxtPinX.Formula, tb2.TxtPinX.Formula);
            Assert.AreEqual(tb1.TxtPinY.Formula, tb2.TxtPinY.Formula);
            Assert.AreEqual(tb1.TxtHeight.Formula, tb2.TxtHeight.Formula);
            Assert.AreEqual(tb1.TxtWidth.Formula, tb2.TxtWidth.Formula);
            Assert.AreEqual(tb1.TxtLocPinX.Formula, tb2.TxtLocPinX.Formula);
            Assert.AreEqual(tb1.TxtLocPinY.Formula, tb2.TxtLocPinY.Formula);

            page1.Delete(0);
        }


        [TestMethod]
        public void Test_Fields1()
        {
            var page1 = GetNewPage();
            var shape = page1.DrawRectangle(0, 0, 4, 2);

            // case 1 - markup is just a single field element
            var markup_1 = new VA.Text.Markup.TextElement();
            markup_1.AddField(VA.Text.Markup.FieldConstants.Height);
            markup_1.SetText(shape);
            Assert.AreEqual("2", shape.Characters.Text);

            // case 2 - markup contains a single field surrounded by literal text
            var markup2 = new VA.Text.Markup.TextElement();
            markup2.AddText("HELLO ");
            markup2.AddField(VA.Text.Markup.FieldConstants.Height);
            markup2.AddText(" WORLD");
            markup2.SetText(shape);
            Assert.AreEqual("HELLO 2 WORLD", shape.Characters.Text);

            // case 3 - markup contains a single literal surrounded by two fields
            var markup3 = new VA.Text.Markup.TextElement();
            markup3.AddField(VA.Text.Markup.FieldConstants.Height);
            markup3.AddText(" HELLO ");
            markup3.AddField(VA.Text.Markup.FieldConstants.Width);
            markup3.SetText(shape);
            Assert.AreEqual("2 HELLO 4", shape.Characters.Text);

            var markup4 = new VA.Text.Markup.TextElement();
            markup4.AddField(VA.Text.Markup.FieldConstants.Height);
            markup4.AddText(" HELLO ");
            markup4.AddField(VA.Text.Markup.FieldConstants.Width);
            markup4.AddField(VA.Text.Markup.FieldConstants.Width);
            markup4.SetText(shape);
            Assert.AreEqual("2 HELLO 44", shape.Characters.Text);

            page1.Delete(0);
        }
    }
}