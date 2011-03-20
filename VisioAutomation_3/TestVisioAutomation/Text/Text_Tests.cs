using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using VisioAutomation.Extensions;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace TestVisioAutomation
{
    [TestClass]
    public class Text_Tests : VisioAutomationTest
    {
        [TestMethod]
        public void Test_Xml_To_TextElement_Whitespace()
        {
            string text =
                @"
<text>
    <space/>
    <tab/>
    <newline/>
    Normal text
    <br/>
</text> ";

            var root_el = VA.Text.Markup.TextElement.FromXml(text, false);

            var actual = root_el.GetInnerText();
            Assert.AreEqual(" \t\nNormal text\n", actual);
        }

        [TestMethod]
        public void Test_Xml_To_TextElement()
        {
            string text =
                @"
<text>
    Normal text
        <br/>

    <text font=""Courier New"">
        Courier new at default size
        <br/>

        <text size=""20"">
            Now at 20pt 
        <br/>
                <text bold=""1"">and this text is bold</text>
                <text italic=""1"">and this text is italic</text>
        </text>
    </text>

</text> ";

            var root_el = VA.Text.Markup.TextElement.FromXml(text, false);

            Assert.AreEqual(1, root_el.Children.Count);
            var root_elements = root_el.Elements.ToList();

            var n0 = root_el.Children[0];
            Assert.AreEqual(3, n0.Children.Count);
            var n1 = (VA.Text.Markup.Literal)n0.Children[0];
            var n2 = (VA.Text.Markup.Literal)n0.Children[1];
            var n3 = (VA.Text.Markup.TextElement)n0.Children[2];
            Assert.AreEqual(VA.Text.Markup.NodeType.Literal, n1.NodeType);
            Assert.AreEqual(VA.Text.Markup.NodeType.Literal, n2.NodeType);
            Assert.AreEqual(VA.Text.Markup.NodeType.Element, n3.NodeType);

            Assert.AreEqual("Normal text", n1.GetInnerText());
            Assert.AreEqual("\n", n2.GetInnerText());
            Assert.AreEqual(3, n3.Children.Count);

            var n4 = (VA.Text.Markup.Literal)n3.Children[0];
            var n5 = (VA.Text.Markup.Literal)n3.Children[1];
            var n6 = (VA.Text.Markup.TextElement)n3.Children[2];
            Assert.AreEqual(VA.Text.Markup.NodeType.Literal, n4.NodeType);
            Assert.AreEqual(VA.Text.Markup.NodeType.Literal, n5.NodeType);
            Assert.AreEqual(VA.Text.Markup.NodeType.Element, n6.NodeType);

            Assert.AreEqual("Courier new at default size", n4.GetInnerText());
            Assert.AreEqual("\n", n5.GetInnerText());
            Assert.AreEqual("Now at 20pt\nand this text is boldand this text is italic", n6.GetInnerText());

            Assert.AreEqual(4, n6.Children.Count);
            var n7 = (VA.Text.Markup.Literal) n6.Children[0];
            var n8 = (VA.Text.Markup.Literal) n6.Children[1];
            var n9 = (VA.Text.Markup.TextElement) n6.Children[2];
            var n10 = (VA.Text.Markup.TextElement) n6.Children[3];

            Assert.AreEqual(VA.Text.Markup.NodeType.Literal, n7.NodeType);
            Assert.AreEqual(VA.Text.Markup.NodeType.Literal, n8.NodeType);
            Assert.AreEqual(VA.Text.Markup.NodeType.Element, n9.NodeType);
            Assert.AreEqual(VA.Text.Markup.NodeType.Element, n10.NodeType);
            Assert.AreEqual("Now at 20pt", n7.GetInnerText());
            Assert.AreEqual("\n", n8.GetInnerText());
            Assert.AreEqual("and this text is bold", n9.GetInnerText());
            Assert.AreEqual("and this text is italic", n10.GetInnerText());

            Assert.AreEqual("Courier New", n3.TextFormat.Font);
            Assert.AreEqual(20, n6.TextFormat.FontSize);
            Assert.AreEqual(VA.Text.CharStyle.Bold, n9.TextFormat.CharStyle.Value);
            Assert.AreEqual(VA.Text.CharStyle.Italic, n10.TextFormat.CharStyle.Value);
        }

        [TestMethod]
        public void Test_Element_To_Region_1()
        {
            var el = new VA.Text.Markup.TextElement();
            var markup = el.GetMarkupInfo();
            var regions = markup.FormatRegions;
            Assert.AreEqual(1, markup.FormatRegions.Count);
            Assert.AreEqual(0, regions[0].TextLength);
            Assert.AreEqual(0, regions[0].TextStartPos);
            Assert.AreEqual(0, regions[0].TextEndPos);
        }

        [TestMethod]
        public void Test_Element_To_Region_2()
        {
            var el = new VA.Text.Markup.TextElement("HELLO");
            var markup = el.GetMarkupInfo();
            var regions = markup.FormatRegions;
            Assert.AreEqual(1, markup.FormatRegions.Count);
            Assert.AreEqual(5, regions[0].TextLength);
            Assert.AreEqual(0, regions[0].TextStartPos);
            Assert.AreEqual(5, regions[0].TextEndPos);
        }

        [TestMethod]
        public void Test_Element_To_Region_3()
        {
            var el = new VA.Text.Markup.TextElement("HELLO");
            el.AppendText(" WORLD");
            var markup = el.GetMarkupInfo();
            var regions = markup.FormatRegions;
            Assert.AreEqual(1, markup.FormatRegions.Count);
            Assert.AreEqual(11, regions[0].TextLength);
            Assert.AreEqual(0, regions[0].TextStartPos);
            Assert.AreEqual(11, regions[0].TextEndPos);
        }

        [TestMethod]
        public void Test_Element_To_Region_4()
        {
            var el = new VA.Text.Markup.TextElement();
            el.AppendNewElement("HELLO");
            el.AppendNewElement(" WORLD");
            var markup = el.GetMarkupInfo();
            var regions = markup.FormatRegions;
            Assert.AreEqual(3, markup.FormatRegions.Count);
            Assert.AreEqual(11, regions[0].TextLength);
            Assert.AreEqual(0, regions[0].TextStartPos);
            Assert.AreEqual(11, regions[0].TextEndPos);
            Assert.AreEqual(5, regions[1].TextLength);
            Assert.AreEqual(0, regions[1].TextStartPos);
            Assert.AreEqual(5, regions[1].TextEndPos);
            Assert.AreEqual(6, regions[2].TextLength);
            Assert.AreEqual(5, regions[2].TextStartPos);
            Assert.AreEqual(11, regions[2].TextEndPos);
        }

        [TestMethod]
        public void Test_Element_To_Region_5()
        {
            var el0 = new VA.Text.Markup.TextElement();
            var el1 = el0.AppendNewElement("HELLO");
            var el2 = el1.AppendNewElement(" WORLD");

            var markup = el0.GetMarkupInfo();
            var regions = markup.FormatRegions;
            Assert.AreEqual(3, markup.FormatRegions.Count);
            Assert.AreEqual(11, regions[0].TextLength);
            Assert.AreEqual(0, regions[0].TextStartPos);
            Assert.AreEqual(11, regions[0].TextEndPos);
            Assert.AreEqual(11, regions[1].TextLength);
            Assert.AreEqual(0, regions[1].TextStartPos);
            Assert.AreEqual(11, regions[1].TextEndPos);
            Assert.AreEqual(6, regions[2].TextLength);
            Assert.AreEqual(5, regions[2].TextStartPos);
            Assert.AreEqual(11, regions[2].TextEndPos);
        }

        [TestMethod]
        public void TextElement_with_multiple_text_nodes()
        {
            var el0 = new VA.Text.Markup.TextElement();
            var el1 = el0.AppendNewElement("HELLO");
            var el2 = el0.AppendNewElement(" WORLD");

            var page1 = GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 4, 4);

            el0.SetShapeText(s1);

            Assert.AreEqual("HELLO WORLD", s1.Text);

            page1.Delete(0);
        }

        [TestMethod]
        public void Element_with_bold_and_italic_text()
        {
            var el0 = new VA.Text.Markup.TextElement();
            var el1 = el0.AppendNewElement("HELLO");
            var el2 = el0.AppendNewElement(" WORLD");

            el1.TextFormat.CharStyle = VA.Text.CharStyle.Bold;
            el2.TextFormat.CharStyle = VA.Text.CharStyle.Italic;

            var page1 = GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 4, 4);

            el0.SetShapeText(s1);

            var fmts = VA.Text.TextHelper.GetCharacterFormat(s1);
            Assert.AreEqual(3, fmts.Count);
            Assert.AreEqual(VA.Text.CharStyle.Bold, fmts[0].Style.Result);
            Assert.AreEqual(VA.Text.CharStyle.Italic, fmts[1].Style.Result);
            Assert.AreEqual(VA.Text.CharStyle.None, fmts[2].Style.Result);

            page1.Delete(0);
        }

        [TestMethod]
        public void Style_inheritance()
        {
            var el0 = new VA.Text.Markup.TextElement();
            var el1 = el0.AppendNewElement("HELLO");
            var el2 = el1.AppendNewElement(" WORLD");

            el0.TextFormat.FontSize = 7;
            el1.TextFormat.Font = "Impact";
            el0.TextFormat.FontSize = 14;
            el1.TextFormat.CharStyle = VA.Text.CharStyle.Bold;
            el2.TextFormat.Font = "Courier New";
            el2.TextFormat.FontSize = 20;
            el2.TextFormat.CharStyle = VA.Text.CharStyle.Italic;

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

            el0.SetShapeText(s1);

            var fmts = VA.Text.TextHelper.GetCharacterFormat(s1);
            Assert.AreEqual(3, fmts.Count);

            page1.Delete(0);
        }

        [TestMethod]
        public void Test_Format_Text_field()
        {
            var el0 = new VA.Text.Markup.TextElement();
            el0.AppendText("HELLO ");
            el0.AppendField(VA.Text.Markup.Fields.Height);
            el0.AppendText(" WORLD");

            var page1 = GetNewPage();

            var s1 = page1.DrawRectangle(0, 0, 4, 4);

            string it = el0.GetInnerText();
            Assert.AreEqual("HELLO X WORLD", it);
            el0.SetShapeText(s1);

            var shape_size = VisioAutomationTest.GetSize(s1);

            string s = string.Format("HELLO {0} WORLD", shape_size.Height);
            var s1_characters = s1.Characters;
            Assert.AreEqual(s, s1_characters.Text);

            page1.Delete(0);
        }

        [TestMethod]
        public void CharacterFormat1()
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
                //sb.AppendLine();                
            }
            s1.Text = sb.ToString();


            var c0 = new VA.DOM.ShapeCells();
            c0.CharSize = 0.6;
            VisioTestCommon.setformulas(c0, page1, s1);

            var c1 = new VA.Text.CharacterFormatCells();
            c1.Color = new VA.Drawing.ColorRGB(0xff0000).ToFormula();
            VA.Text.TextHelper.SetFormat(s1, c1, 0, 5);

            var c2 = new VA.Text.CharacterFormatCells();
            c2.Size = 0.5;
            VA.Text.TextHelper.SetFormat(s1, c2, 5, 10);

            var c3 = new VA.Text.CharacterFormatCells();
            c3.Font = page1.Document.Fonts["Impact"].ID;
            VA.Text.TextHelper.SetFormat(s1, c3, 10, 15);

            var c4 = new VA.Text.CharacterFormatCells();
            c4.Style = (int) (VA.Text.CharStyle.Italic | VA.Text.CharStyle.UnderLine);
            VA.Text.TextHelper.SetFormat(s1, c4, 15, 20);

            var c5 = new VA.Text.CharacterFormatCells();
            c5.Transparency = 0.5;
            VA.Text.TextHelper.SetFormat(s1, c5, 20, 25);

            var formatting = VA.Text.TextHelper.GetCharacterFormat(s1);
            Assert.AreEqual("RGB(255,0,0)", formatting[0].Color.Formula);
            Assert.AreEqual(0.5, formatting[1].Size.Result);
            Assert.AreEqual(page1.Document.Fonts["Impact"].ID, formatting[2].Font.Result);
            Assert.AreEqual("6", formatting[3].Style.Formula);
            Assert.AreEqual("50%", formatting[4].Transparency.Formula);
            Assert.AreEqual(0.6, formatting[5].Size.Result);

            //page1.Delete(0);
        }

        [TestMethod]
        public void Test_SetCharacterFormat2()
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
            VA.Text.TextHelper.SetFormat(fmt, s1);

            fmt = new VA.Text.CharacterFormatCells();
            s1 = page1.DrawRectangle(2,2,3,3);
            s1.Text = "Italic";
            fmt.Style = (int)VA.Text.CharStyle.Italic;
            VA.Text.TextHelper.SetFormat(fmt, s1);

            fmt = new VA.Text.CharacterFormatCells();
            s1 = page1.DrawRectangle(3,3,4,4);
            s1.Text = "Underline";
            fmt.Style = (int)VA.Text.CharStyle.UnderLine;
            VA.Text.TextHelper.SetFormat(fmt, s1);

            fmt = new VA.Text.CharacterFormatCells();
            s1 = page1.DrawRectangle(4,4,5,5);
            s1.Text = "Smallcaps";
            fmt.Style = (int)VA.Text.CharStyle.SmallCaps;
            VA.Text.TextHelper.SetFormat(fmt, s1);

            fmt = new VA.Text.CharacterFormatCells();
            s1 = page1.DrawRectangle(5,5,6,6);
            s1.Text = "Red";
            fmt.Color = new VA.Drawing.ColorRGB(0xff0000).ToFormula();
            VA.Text.TextHelper.SetFormat(fmt, s1);

            fmt = new VA.Text.CharacterFormatCells();
            s1 = page1.DrawRectangle(6,6,7,7);
            s1.Text = "#ec35a7";
            fmt.Color = new VA.Drawing.ColorRGB(0xec35a7).ToFormula();
            VA.Text.TextHelper.SetFormat(fmt, s1);

            fmt = new VA.Text.CharacterFormatCells();
            s1 = page1.DrawRectangle(7,7,8,8);
            s1.Text = "#34f178";
            fmt.Color = new VA.Drawing.ColorRGB(0x34f178).ToFormula();
            VA.Text.TextHelper.SetFormat(fmt, s1);

            fmt = new VA.Text.CharacterFormatCells();
            s1 = page1.DrawRectangle(8,8,9,9);
            s1.Text = "Calibri";
            fmt.Font = page1.Document.Fonts["Calibri"].ID;
            VA.Text.TextHelper.SetFormat(fmt, s1);

            fmt = new VA.Text.CharacterFormatCells();
            s1 = page1.DrawRectangle(9,9,10,10);
            s1.Text = "Impact";
            fmt.Font = page1.Document.Fonts["Impact"].ID;
            VA.Text.TextHelper.SetFormat(fmt, s1);

            fmt = new VA.Text.CharacterFormatCells();
            s1 = page1.DrawRectangle(10,10,11,11);
            s1.Text = "Segoe UI";
            fmt.Font = page1.Document.Fonts["Segoe UI"].ID;
            VA.Text.TextHelper.SetFormat(fmt, s1);

            fmt = new VA.Text.CharacterFormatCells();
            s1 = page1.DrawRectangle(11,11,12,12);
            s1.Text = "6pt";
            fmt.Size = VA.Convert.PointsToInches(6);
            VA.Text.TextHelper.SetFormat(fmt, s1);

            fmt = new VA.Text.CharacterFormatCells();
            s1 = page1.DrawRectangle(12,12,13,13);
            s1.Text = "8pt";
            fmt.Size = VA.Convert.PointsToInches(8);
            VA.Text.TextHelper.SetFormat(fmt, s1);

            fmt = new VA.Text.CharacterFormatCells();
            s1 = page1.DrawRectangle(13,13,14,14);
            s1.Text = "11pt";
            fmt.Size = VA.Convert.PointsToInches(8);
            VA.Text.TextHelper.SetFormat(fmt, s1);

            fmt = new VA.Text.CharacterFormatCells();
            s1 = page1.DrawRectangle(14,14,15,15);
            s1.Text = "15pt";
            fmt.Size = VA.Convert.PointsToInches(15);
            VA.Text.TextHelper.SetFormat(fmt, s1);

            page1.Delete(0);
        }

        [TestMethod]
        public void Paragraph_formatting()
        {
            var page1 = GetNewPage(new VA.Drawing.Size(10, 10));

            var s1 = page1.DrawRectangle(0,0,5,5);
            s1.Text = "Line1\nLine2\nLine3\nLine4\nLine5\nLine6";

            var formats0 = VA.Text.TextHelper.GetParagraphFormat(s1);
            Assert.AreEqual(1, formats0.Count);

            var fmt1 = new VA.Text.ParagraphFormatCells();
            fmt1.IndentLeft = 0.25;

            var cfmt1 = new VA.Text.CharacterFormatCells();
            cfmt1.Color = "RGB(255,0,0)";


            VA.Text.TextHelper.SetFormat(s1, cfmt1, 2, 3);
            VA.Text.TextHelper.SetFormat(s1, fmt1, 2, 3);

            var formats1 = VA.Text.TextHelper.GetParagraphFormat(s1);
            Assert.AreEqual(2, formats1.Count);

            var fmt2 = new VA.Text.ParagraphFormatCells();
            fmt2.BulletIndex = 2;

            var cfmt2 = new VA.Text.CharacterFormatCells();
            cfmt2.Color = "RGB(0,255,0)";

            VA.Text.TextHelper.SetFormat(s1, cfmt2, 13, 14);
            VA.Text.TextHelper.SetFormat(s1, fmt2, 13, 14);

            var formats2 = VA.Text.TextHelper.GetParagraphFormat(s1);
            Assert.AreEqual(4, formats2.Count);

            Assert.AreEqual(0.25, formats2[0].IndentLeft.Result);
            Assert.AreEqual(0, formats2[0].IndentFirst.Result);

            Assert.AreEqual(0, formats2[2].IndentLeft.Result);
            Assert.AreEqual(0, formats2[2].IndentFirst.Result);

            Assert.AreEqual(2, formats2[1].BulletIndex.Result);

            Assert.AreEqual(0, formats2[1].IndentLeft.Result);
            Assert.AreEqual(0, formats2[2].BulletIndex.Result);

            //page1.Delete(0);
        }

    }
}