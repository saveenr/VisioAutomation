using MUT=Microsoft.VisualStudio.TestTools.UnitTesting;
using VABOX = VisioAutomation.Models.Layouts.Box;
using VA=VisioAutomation;

namespace VTest.Dom
{
    [MUT.TestClass]
    public class Dom_Text : VisioAutomationTest
    {
        [MUT.TestMethod]
        public void DomText_CharacterFormatting()
        {
            this.DomText_CharacterBold();
            this.DomText_CharacterComplex();
            this.DomText_CharacterFont();
            this.DomText_CharacterItalic();
            this.DomText_CharacterPlain();
            this.DomText_ParagraphCenter();
            this.DomText_ParagraphDefault();
            this.DomText_ParagraphLeft();
            this.DomText_ParagraphRight();
        }

        public void DomText_CharacterPlain()
        {
            var m = new VisioAutomation.Models.Text.Element("{Normal}");
            var page1 = this.GetNewPage(new VisioAutomation.Core.Size(5, 5));
            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            m.SetText(s0);

            var textfmt = VisioAutomation.Text.TextFormat.GetFormat(s0, VisioAutomation.Core.CellValueType.Formula);
            var charfmt = textfmt.CharacterFormats;
            MUT.Assert.AreEqual(1, charfmt.Count);

            page1.Delete(0);
        }

        public void DomText_CharacterBold()
        {
            var m = new VisioAutomation.Models.Text.Element("{Bold}");
            m.CharacterFormatting.Style = (int)VA.Models.Text.CharStyle.Bold;
            var page1 = this.GetNewPage(new VisioAutomation.Core.Size(5, 5));
            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            m.SetText(s0);

            var textfmt = VisioAutomation.Text.TextFormat.GetFormat(s0, VisioAutomation.Core.CellValueType.Formula);
            var charfmt = textfmt.CharacterFormats;
            MUT.Assert.AreEqual(1, charfmt.Count);
            MUT.Assert.AreEqual(((int)VA.Models.Text.CharStyle.Bold).ToString(), charfmt[0].Style.Value);

            page1.Delete(0);
        }

        public void DomText_CharacterItalic()
        {
            var m = new VisioAutomation.Models.Text.Element("{Italic}");
            m.CharacterFormatting.Style = (int)VA.Models.Text.CharStyle.Italic;
            var page1 = this.GetNewPage(new VisioAutomation.Core.Size(5, 5));
            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            m.SetText(s0);

            var textfmt = VisioAutomation.Text.TextFormat.GetFormat(s0, VisioAutomation.Core.CellValueType.Formula);
            var charfmt = textfmt.CharacterFormats;
            MUT.Assert.AreEqual(1, charfmt.Count);
            MUT.Assert.AreEqual(((int)VA.Models.Text.CharStyle.Italic).ToString(), charfmt[0].Style.Value);

            page1.Delete(0);
        }

        public void DomText_CharacterFont()
        {
            var page1 = this.GetNewPage(new VisioAutomation.Core.Size(5, 5));

            var impact = page1.Document.Fonts["Arial"];
            var m = new VisioAutomation.Models.Text.Element("Normal Text in Impact Font");
            m.CharacterFormatting.Font = impact.ID;
            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            m.SetText(s0);

            var textfmt = VisioAutomation.Text.TextFormat.GetFormat(s0, VisioAutomation.Core.CellValueType.Formula);
            var charfmt = textfmt.CharacterFormats;
            MUT.Assert.AreEqual(1, charfmt.Count);

            AssertUtil.OneOf( new[] { "0", "THEMEVAL()" }, charfmt[0].Style.Value);

            MUT.Assert.AreEqual(impact.ID.ToString(), charfmt[0].Font.Value);

            page1.Delete(0);
        }

        public void DomText_CharacterComplex()
        {
            var page1 = this.GetNewPage(new VisioAutomation.Core.Size(5, 5));
            var doc = page1.Document;
            var fonts = doc.Fonts;

            var segoeui = fonts["Segoe UI"];
            var impact = fonts["Arial"];
            var couriernew = fonts["Courier New"];
            var georgia = fonts["Georgia"];

            var t1 = new VisioAutomation.Models.Text.Element("{Normal}");
            t1.CharacterFormatting.Font = segoeui.ID;

            var t2 = t1.AddElement("{Italic}");
            t2.CharacterFormatting.Style = (int)VA.Models.Text.CharStyle.Italic;
            t2.CharacterFormatting.Font = impact.ID;

            var t3 = t2.AddElement("{Bold}");
            t3.CharacterFormatting.Style = (int)VA.Models.Text.CharStyle.Bold;
            t3.CharacterFormatting.Font = couriernew.ID;

            var t4 = t2.AddElement("{Bold Italic}");
            t4.CharacterFormatting.Style = (int)(VA.Models.Text.CharStyle.Bold | VA.Models.Text.CharStyle.Italic);
            t4.CharacterFormatting.Font = georgia.ID;

            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            t1.SetText(s0);

            var textfmt = VisioAutomation.Text.TextFormat.GetFormat(s0, VisioAutomation.Core.CellValueType.Formula);
            var charfmt = textfmt.CharacterFormats;

            // check the number of character regions
            MUT.Assert.AreEqual(5, charfmt.Count);

            // check the fonts
            MUT.Assert.AreEqual(segoeui.ID.ToString(), charfmt[0].Font.Value);
            MUT.Assert.AreEqual(impact.ID.ToString(), charfmt[1].Font.Value);
            MUT.Assert.AreEqual(couriernew.ID.ToString(), charfmt[2].Font.Value);
            MUT.Assert.AreEqual(georgia.ID.ToString(), charfmt[3].Font.Value);
            MUT.Assert.AreEqual(segoeui.ID.ToString(), charfmt[4].Font.Value);


            // check the styles

            string style_value_0 = charfmt[0].Style.Value;
            string style_value_4 = charfmt[4].Style.Value;
            string style_none = ((int)VA.Models.Text.CharStyle.None).ToString();

            AssertUtil.OneOf(new[] { style_none, "THEMEVAL()" }, style_value_0);

            MUT.Assert.AreEqual(((int)VA.Models.Text.CharStyle.Italic).ToString(), charfmt[1].Style.Value);
            MUT.Assert.AreEqual(((int)VA.Models.Text.CharStyle.Bold).ToString(), charfmt[2].Style.Value);
            MUT.Assert.AreEqual(((int)(VA.Models.Text.CharStyle.Italic | VA.Models.Text.CharStyle.Bold)).ToString(), charfmt[3].Style.Value);

            AssertUtil.OneOf(new[] { style_none, "THEMEVAL()" }, style_value_4);

            // check the text run content
            var charruns = textfmt.CharacterTextRuns;
            MUT.Assert.AreEqual(4, charruns.Count);
            MUT.Assert.AreEqual("{Normal}", charruns[0].Text);
            MUT.Assert.AreEqual("{Italic}", charruns[1].Text);
            MUT.Assert.AreEqual("{Bold}", charruns[2].Text);
            MUT.Assert.AreEqual("{Bold Italic}", charruns[3].Text);

            // cleanup
            page1.Delete(0);
        }

        public void DomText_ParagraphDefault()
        {
            var m = new VisioAutomation.Models.Text.Element("{DefaultPara}");
            var page1 = this.GetNewPage(new VisioAutomation.Core.Size(5, 5));
            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            m.SetText(s0);

            var textfmt = VisioAutomation.Text.TextFormat.GetFormat(s0, VisioAutomation.Core.CellValueType.Formula);
            var parafmt = textfmt.ParagraphFormats;
            MUT.Assert.AreEqual(1, parafmt.Count);

            page1.Delete(0);
        }

        public void DomText_ParagraphLeft()
        {
            var m = new VisioAutomation.Models.Text.Element("{LeftHAlign}");
            m.ParagraphFormatting.HorizontalAlign = (int)VABOX.AlignmentHorizontal.Left;
            var page1 = this.GetNewPage(new VisioAutomation.Core.Size(5, 5));
            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            m.SetText(s0);

            var textfmt = VisioAutomation.Text.TextFormat.GetFormat(s0, VisioAutomation.Core.CellValueType.Formula);
            var parafmt = textfmt.ParagraphFormats;
            MUT.Assert.AreEqual(1, parafmt.Count);

            MUT.Assert.AreEqual(((int)VABOX.AlignmentHorizontal.Left).ToString(), parafmt[0].HorizontalAlign.Value);

            page1.Delete(0);
        }

        public void DomText_ParagraphCenter()
        {
            var m = new VisioAutomation.Models.Text.Element("{CenterHAlign}");
            m.ParagraphFormatting.HorizontalAlign = (int)VABOX.AlignmentHorizontal.Center;
            var page1 = this.GetNewPage(new VisioAutomation.Core.Size(5, 5));
            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            m.SetText(s0);

            var textfmt = VisioAutomation.Text.TextFormat.GetFormat(s0, VisioAutomation.Core.CellValueType.Formula);
            var parafmt = textfmt.ParagraphFormats;
            MUT.Assert.AreEqual(1, parafmt.Count);

            MUT.Assert.AreEqual(((int)VABOX.AlignmentHorizontal.Center).ToString(), parafmt[0].HorizontalAlign.Value);

            page1.Delete(0);
        }

        public void DomText_ParagraphRight()
        {
            var m = new VisioAutomation.Models.Text.Element("{RightHAlign}");
            m.ParagraphFormatting.HorizontalAlign = (int)VABOX.AlignmentHorizontal.Right;
            var page1 = this.GetNewPage(new VisioAutomation.Core.Size(5, 5));
            var s0 = page1.DrawRectangle(0, 0, 4, 4);
            m.SetText(s0);

            var textfmt = VisioAutomation.Text.TextFormat.GetFormat(s0, VisioAutomation.Core.CellValueType.Formula);
            var parafmt = textfmt.ParagraphFormats;
            MUT.Assert.AreEqual(1, parafmt.Count);

            MUT.Assert.AreEqual(((int)VABOX.AlignmentHorizontal.Right).ToString(), parafmt[0].HorizontalAlign.Value);

            page1.Delete(0);
        }
    }
}