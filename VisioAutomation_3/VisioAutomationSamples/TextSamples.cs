using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using System.Linq;
using System.Collections.Generic;

namespace VisioAutomationSamples
{
    public static partial class TextSamples
    {
        private static string text1 =
            @"
<text font=""Calibri"" size=""15"" color=""#ff0000"">

    <text>Hello World [Calibri 15pt red]</text>
    <br/><br/><br/><br/>
    <text font=""Segoe UI"" size=""20"" color=""#0000ff"">
        Hello World [Segoe UI 20 pt blue]
        <text italic=""1"" bold=""1"" halign=""left"" color=""#505050"">
            Hello World [ left italic gray ]
            <text bold=""1"" italic=""1"" halign=""right"">
                Hello World [ bold,italic right]
            </text>
            <text bold=""0"" italic=""0"" halign=""center"">
                Hello World [ nobold,noitalic,center]
            </text>
            Hello World [ left italic gray ]
        </text>
        Hello World [Segoe UI 20 pt blue]
    </text>
    Hello World [Calibri 15pt red]
</text>";

        private static string text2 =
            @"
    <text size=""20"" font=""Segoe UI""> The lines underneath should be bulleted 

<text bullets=""1"" halign=""left""> This a demonstration of <text bold=""1"">bold</text> text
A demonstration of <text italic=""1"">italic</text> text
CellsPackage can be combined to form <text bold=""1"" italic=""1"">bold italic</text>
This word is <text underline=""1"" >under<text underline=""0"" >lined</text></text>
This word is <text smallcaps=""1"" >smallcaps</text>
</text>

The bullets have ended.
</text> ";

        private static string text3 =
            @"
    <text size=""20"" font=""Segoe UI"" halign=""left""> The lines below should be indented. And they should get increasingly more transparent.

<text indent=""25"" transparency=""25"">
indent 25
</text>
<text indent=""50"" transparency=""45"">
indent 50
</text>
<text indent=""75"" transparency=""65"">
indent 75
</text>
<text indent=""100"" transparency=""85"">
indent 100
</text>

The indenting has ended.


</text> ";

        private static string text4 =
            "<text size=\"30\"> This tests special characters <text> Carriage return [\r] </text> <text> Line feed [\n] </text> </text> ";

        public static void NonRotatingText()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();
            var s0 = page.DrawRectangle(1, 1, 4, 4);
            s0.Text = "Hello World";

            s0.GetCell(VisioAutomation.ShapeSheet.SRCConstants.TxtAngle).Formula = "-Angle";
        }

        public static void TextFields()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();
            var s0 = page.DrawRectangle(1, 1, 4, 4);

            VA.Text.TextHelper.SetTextFormatFields(s0, "{0} ({1} of {2})", VisioAutomation.Text.Markup.Fields.NumberOfPages,
                                              VisioAutomation.Text.Markup.Fields.PageNumber,
                                              VisioAutomation.Text.Markup.Fields.PageName);
        }

        public static void TextMarkup1()
        {
            // the backspace char \b is not valid text in XML, that is why it is not included above

            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            //vi.Page.Name = "14orgchart"; 

            var s1 = page.DrawRectangle(0, 0, 8, 8);
            var s2 = page.DrawRectangle(8, 0, 16, 8);
            var s3 = page.DrawRectangle(0, 8, 8, 16);
            var s4 = page.DrawRectangle(8, 8, 16, 16);

            var markup_doms = new[] {text1, text2, text3, text4}
                .Select(s => VA.Text.Markup.TextElement.FromXml(s, true))
                .ToList();
            var shapes = new[] {s1, s2, s3, s4};

            for (int i = 0; i < shapes.Length; i++)
            {
                var shape = shapes[i];
                var markup_dom = markup_doms[i%markup_doms.Count];
                markup_dom.SetShapeText(shape);
            }
        }

        public static void TextMarkup2()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();
            var s0 = page.DrawRectangle(1, 1, 4, 4);

            page.DrawRectangle(1, 1, 4, 4);

            var tokens = new[] {"The ", "Quick ", "Brown ", "Fox"};
            var e1 = new VisioAutomation.Text.Markup.TextElement();
            foreach (var token in tokens)
            {
                e1.AppendText(token);
            }
            //vi.Text.Markup = e1;
        }

        public static void TextSizing()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            var s0 = page.DrawRectangle(0, 0, 4, 4);

            // Alignment Box fits to accomodate text
            s0.Text = "Alignment Box fits to accomodate text";

            s0.GetCell(VisioAutomation.ShapeSheet.SRCConstants.Width).Formula = "2.0";
            s0.GetCell(VisioAutomation.ShapeSheet.SRCConstants.Height).Formula = "GUARD(TxtHeight)";
            s0.GetCell(VisioAutomation.ShapeSheet.SRCConstants.TxtWidth).Formula = "Width*1";
            s0.GetCell(VisioAutomation.ShapeSheet.SRCConstants.TxtHeight).Formula = "TEXTHEIGHT(TheText,TxtWidth)";

            // Text Scales Proportional to Shape Height
            var s1 = page.DrawRectangle(0, 4, 8, 8);
            s1.Text = "Text Scales Proportional to Shape Height";
            s0.GetCell(VisioAutomation.ShapeSheet.SRCConstants.Char_Size).Formula = "Height*0.25";

            // Text scales smaller to fit more text
            var s2 = page.DrawRectangle(4, 0, 8, 4);
            s2.Text = "Text scales smaller to fit more text";
            s2.GetCell(VisioAutomation.ShapeSheet.SRCConstants.Char_Size).Formula =
                "11pt * 10/SQRT(LEN(SHAPETEXT(TheText)))";
        }

        public static void FontChart()
        {
            var stencil = SampleEnvironment.Application.Documents.OpenStencil("basic_u.vss");
            var master = stencil.Masters["Rectangle"];
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            var fonts = new[] {"Segoe UI", "Calibri", "Arial"};
            var sizes = new[] {"28.0pt", "18.0pt", "14.0pt", "12.0pt", "10.0pt"};
            var fontids = fonts.Select(f => page.Document.Fonts[f].ID).ToList();

            var shapeids = VA.Layout.LayoutHelper.DrawGrid(page, master, new VA.Drawing.Size(3.0, 0.5), sizes.Length, fonts.Length);
            var shapes = page.Shapes.GetShapesFromIDs(shapeids);

            var items = from fi in Enumerable.Range(0, fonts.Count())
                        from size in sizes
                        select new {font = fonts[fi], size = size, fontid = fontids[fi]};

            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();

            int i = 0;
            foreach (var item in items)
            {
                var shape = shapes[i];
                shape.Text = item.font + " " + item.size;
                var shapeid = shapeids[i];
                update.SetFormula(shapeid, VisioAutomation.ShapeSheet.SRCConstants.Char_Size, item.size);
                update.SetFormula(shapeid, VisioAutomation.ShapeSheet.SRCConstants.Char_Font, item.fontid);
                i++;
            }
        }
    }
}