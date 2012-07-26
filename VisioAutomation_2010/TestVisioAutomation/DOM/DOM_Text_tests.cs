using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using IVisio=Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace TestVisioAutomation
{
    [TestClass]
    public class DOM_Text_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void MarkupCharacterPlain()
        {
            var m = new VA.Text.Markup.TextElement("{Normal}");
            var page1 = this.GetNewPage(new VA.Drawing.Size(5, 5));
            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            m.SetText(s0);

            var textfmt = VA.Text.TextFormat.GetFormat(s0);
            var charfmt = textfmt.CharacterFormats;
            Assert.AreEqual(1,charfmt.Count);

            page1.Delete(0);
        }

        [TestMethod]
        public void MarkupCharacterBold()
        {
            var m = new VA.Text.Markup.TextElement("{Bold}");
            m.CharacterFormat.Style = VA.Text.CharStyle.Bold;
            var page1 = this.GetNewPage(new VA.Drawing.Size(5, 5));
            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            m.SetText(s0);

            var textfmt = VA.Text.TextFormat.GetFormat(s0);
            var charfmt = textfmt.CharacterFormats;
            Assert.AreEqual(1, charfmt.Count);
            Assert.AreEqual((int)VA.Text.CharStyle.Bold, charfmt[0].Style.Result);

            page1.Delete(0);
        }

        [TestMethod]
        public void MarkupCharacterItalic()
        {
            var m = new VA.Text.Markup.TextElement("{Italic}");
            m.CharacterFormat.Style = VA.Text.CharStyle.Italic;
            var page1 = this.GetNewPage(new VA.Drawing.Size(5, 5));
            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            m.SetText(s0);

            var textfmt = VA.Text.TextFormat.GetFormat(s0);
            var charfmt = textfmt.CharacterFormats;
            Assert.AreEqual(1, charfmt.Count);
            Assert.AreEqual((int)VA.Text.CharStyle.Italic, charfmt[0].Style.Result);

            page1.Delete(0);
        }

        [TestMethod]
        public void MarkupCharacterFont()
        {
            var page1 = this.GetNewPage(new VA.Drawing.Size(5, 5));

            var impact = page1.Document.Fonts["Impact"];
            var m = new VA.Text.Markup.TextElement("Normal Text in Impact Font");
            m.CharacterFormat.FontID = impact.ID;
            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            m.SetText(s0);

            var textfmt = VA.Text.TextFormat.GetFormat(s0);
            var charfmt = textfmt.CharacterFormats;
            Assert.AreEqual(1, charfmt.Count);
            Assert.AreEqual(0, charfmt[0].Style.Result);
            Assert.AreEqual(impact.ID, charfmt[0].Font.Result);

            page1.Delete(0);
        }

        [TestMethod]
        public void MarkupCharacterComplex()
        {
            var page1 = this.GetNewPage(new VA.Drawing.Size(5, 5));
            var doc = page1.Document;
            var fonts = doc.Fonts;

            var segoeui = fonts["Segoe UI"];
            var impact = fonts["Impact"];
            var couriernew = fonts["Courier New"];
            var georgia = fonts["Georgia"];

            var t1 = new VA.Text.Markup.TextElement("{Normal}");
            t1.CharacterFormat.FontID = segoeui.ID;
            
            var t2 = t1.AppendElement("{Italic}");
            t2.CharacterFormat.Style = VA.Text.CharStyle.Italic;
            t2.CharacterFormat.FontID = impact.ID;

            var t3 = t2.AppendElement("{Bold}");
            t3.CharacterFormat.Style = VA.Text.CharStyle.Bold;
            t3.CharacterFormat.FontID = couriernew.ID;

            var t4 = t2.AppendElement("{Bold Italic}");
            t4.CharacterFormat.Style = VA.Text.CharStyle.Bold | VA.Text.CharStyle.Italic;
            t4.CharacterFormat.FontID = georgia.ID;

            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            t1.SetText(s0);

            var textfmt = VA.Text.TextFormat.GetFormat(s0);
            var charfmt = textfmt.CharacterFormats;
            
            // check the number of character regions
            Assert.AreEqual(5, charfmt.Count);

            // check the fonts
            Assert.AreEqual(segoeui.ID, charfmt[0].Font.Result);
            Assert.AreEqual(impact.ID, charfmt[1].Font.Result);
            Assert.AreEqual(couriernew.ID, charfmt[2].Font.Result);
            Assert.AreEqual(georgia.ID, charfmt[3].Font.Result);
            Assert.AreEqual(segoeui.ID, charfmt[4].Font.Result);


            // check the styles
            Assert.AreEqual((int)VA.Text.CharStyle.None, charfmt[0].Style.Result);
            Assert.AreEqual((int)VA.Text.CharStyle.Italic, charfmt[1].Style.Result);
            Assert.AreEqual((int)VA.Text.CharStyle.Bold, charfmt[2].Style.Result);
            Assert.AreEqual((int) (VA.Text.CharStyle.Italic | VA.Text.CharStyle.Bold), charfmt[3].Style.Result);
            Assert.AreEqual((int)(VA.Text.CharStyle.None), charfmt[4].Style.Result);

            // check the text run content
            var charruns= textfmt.CharacterTextRuns;
            Assert.AreEqual(4, charruns.Count);
            Assert.AreEqual("{Normal}", charruns[0].Text);
            Assert.AreEqual("{Italic}", charruns[1].Text);
            Assert.AreEqual("{Bold}", charruns[2].Text);
            Assert.AreEqual("{Bold Italic}", charruns[3].Text);

            // cleanup
            page1.Delete(0);
        }

        [TestMethod]
        public void MarkupParagraphDefault()
        {
            var m = new VA.Text.Markup.TextElement("{DefaultPara}");
            var page1 = this.GetNewPage(new VA.Drawing.Size(5, 5));
            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            m.SetText(s0);

            var textfmt = VA.Text.TextFormat.GetFormat(s0);
            var parafmt = textfmt.ParagraphFormats;
            Assert.AreEqual(1, parafmt.Count);

            page1.Delete(0);
        }

        [TestMethod]
        public void MarkupParagraphLeft()
        {
            var m = new VA.Text.Markup.TextElement("{LeftHAlign}");
            m.ParagraphFormat.HAlign = VA.Drawing.AlignmentHorizontal.Left;
            var page1 = this.GetNewPage(new VA.Drawing.Size(5, 5));
            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            m.SetText(s0);

            var textfmt = VA.Text.TextFormat.GetFormat(s0);
            var parafmt = textfmt.ParagraphFormats;
            Assert.AreEqual(1, parafmt.Count);

            Assert.AreEqual((int)VA.Drawing.AlignmentHorizontal.Left, parafmt[0].HorizontalAlign.Result);

            page1.Delete(0);
        }

        [TestMethod]
        public void MarkupParagraphCenter()
        {
            var m = new VA.Text.Markup.TextElement("{CenterHAlign}");
            m.ParagraphFormat.HAlign = VA.Drawing.AlignmentHorizontal.Center;
            var page1 = this.GetNewPage(new VA.Drawing.Size(5, 5));
            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            m.SetText(s0);

            var textfmt = VA.Text.TextFormat.GetFormat(s0);
            var parafmt = textfmt.ParagraphFormats;
            Assert.AreEqual(1, parafmt.Count);

            Assert.AreEqual((int)VA.Drawing.AlignmentHorizontal.Center, parafmt[0].HorizontalAlign.Result);

            page1.Delete(0);
        }

        [TestMethod]
        public void MarkupParagraphRight()
        {
            var m = new VA.Text.Markup.TextElement("{RightHAlign}");
            m.ParagraphFormat.HAlign = VA.Drawing.AlignmentHorizontal.Right;
            var page1 = this.GetNewPage(new VA.Drawing.Size(5, 5));
            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            m.SetText(s0);

            var textfmt = VA.Text.TextFormat.GetFormat(s0);
            var parafmt = textfmt.ParagraphFormats;
            Assert.AreEqual(1, parafmt.Count);

            Assert.AreEqual((int)VA.Drawing.AlignmentHorizontal.Right, parafmt[0].HorizontalAlign.Result);

            page1.Delete(0);
        }

    }
}