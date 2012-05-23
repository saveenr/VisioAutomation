using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomationSamples
{
    public static class TextSamples
    {
        public static void NonRotatingText()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();
            var s0 = page.DrawRectangle(1, 1, 4, 4);
            s0.Text = "Hello World";

            s0.GetCell(VA.ShapeSheet.SRCConstants.TxtAngle).Formula = "-Angle";
        }

        public static void TextFields()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();
            var s0 = page.DrawRectangle(1, 1, 4, 4);

            var markup1 = new VA.Text.Markup.TextElement();
            markup1.AppendField(VA.Text.Markup.FieldConstants.PageName);
            markup1.AppendText(" (");
            markup1.AppendField(VA.Text.Markup.FieldConstants.PageNumber);
            markup1.AppendText(" of ");
            markup1.AppendField(VA.Text.Markup.FieldConstants.NumberOfPages);
            markup1.AppendText(") ");
            markup1.SetText(s0);
        }

        public static void TextSizing()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            var s0 = page.DrawRectangle(0, 0, 4, 4);

            // Alignment Box fits to accomodate text
            s0.Text = "Alignment Box fits to accomodate text";

            s0.GetCell(VA.ShapeSheet.SRCConstants.Width).Formula = "2.0";
            s0.GetCell(VA.ShapeSheet.SRCConstants.Height).Formula = "GUARD(TxtHeight)";
            s0.GetCell(VA.ShapeSheet.SRCConstants.TxtWidth).Formula = "Width*1";
            s0.GetCell(VA.ShapeSheet.SRCConstants.TxtHeight).Formula = "TEXTHEIGHT(TheText,TxtWidth)";

            // Text Scales Proportional to Shape Height
            var s1 = page.DrawRectangle(0, 4, 8, 8);
            s1.Text = "Text Scales Proportional to Shape Height";
            s0.GetCell(VA.ShapeSheet.SRCConstants.Char_Size).Formula = "Height*0.25";

            // Text scales smaller to fit more text
            var s2 = page.DrawRectangle(4, 0, 8, 4);
            s2.Text = "Text scales smaller to fit more text";
            s2.GetCell(VA.ShapeSheet.SRCConstants.Char_Size).Formula =
                "11pt * 10/SQRT(LEN(SHAPETEXT(TheText)))";
        }
    }
}