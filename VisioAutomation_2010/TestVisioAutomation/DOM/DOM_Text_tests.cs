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
        public void Markup_Simple_Plain()
        {
            var m = new VA.Text.Markup.TextElement("Normal Text");
            var page1 = this.GetNewPage(new VA.Drawing.Size(5, 5));
            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            m.SetText(s0);

            var textfmt = VA.Text.TextFormat.GetFormat(s0);
            var charfmt = textfmt.CharacterFormats;
            Assert.AreEqual(1,charfmt.Count);

            page1.Delete(0);
        }

        [TestMethod]
        public void Markup_Simple_Bold()
        {
            var m = new VA.Text.Markup.TextElement("Bold Text");
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
        public void Markup_Simple_Italic()
        {
            var m = new VA.Text.Markup.TextElement("Italic Text");
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
        public void Markup_Simple_Font()
        {
            var page1 = this.GetNewPage(new VA.Drawing.Size(5, 5));

            var impact = page1.Document.Fonts["Impact"];
            var m = new VA.Text.Markup.TextElement("Normal Text in Impact Font");
            m.CharacterFormat.Font = impact.ID;
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
        public void Render_Markup_Simple_Font_Multiple()
        {
            var page1 = this.GetNewPage(new VA.Drawing.Size(5, 5));
            var impact = page1.Document.Fonts["Impact"];
            var m = new VA.Text.Markup.TextElement("Normal Text in Impact Font with Red Color");
            m.CharacterFormat.Font = impact.ID;
            m.CharacterFormat.Color = new VA.Drawing.ColorRGB(0xff0000);
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
        public void Markup_Overlap_Multiple()
        {
            var page1 = this.GetNewPage(new VA.Drawing.Size(5, 5));
            var segoeui = page1.Document.Fonts["Segoe UI"];

            var t1 = new VA.Text.Markup.TextElement("{Normal}");
            t1.CharacterFormat.Font = segoeui.ID;
            
            var t2 = t1.AppendElement("{Italic}");
            t2.CharacterFormat.Style = VA.Text.CharStyle.Italic;

            var t3 = t2.AppendElement("{Bold}");
            t3.CharacterFormat.Style = VA.Text.CharStyle.Bold;

            var t4 = t2.AppendElement("{Bold Italic}");
            t4.CharacterFormat.Style = VA.Text.CharStyle.Bold | VA.Text.CharStyle.Italic;

            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            t1.SetText(s0);

            var textfmt = VA.Text.TextFormat.GetFormat(s0);
            var charfmt = textfmt.CharacterFormats;
            Assert.AreEqual(5, charfmt.Count);
            Assert.AreEqual((int)VA.Text.CharStyle.None, charfmt[0].Style.Result);
            Assert.AreEqual((int)VA.Text.CharStyle.Italic, charfmt[1].Style.Result);
            Assert.AreEqual((int)VA.Text.CharStyle.Bold, charfmt[2].Style.Result);
            Assert.AreEqual((int) (VA.Text.CharStyle.Italic | VA.Text.CharStyle.Bold), charfmt[3].Style.Result);
            Assert.AreEqual((int)(VA.Text.CharStyle.None), charfmt[4].Style.Result);

            page1.Delete(0);
        }
    }
}