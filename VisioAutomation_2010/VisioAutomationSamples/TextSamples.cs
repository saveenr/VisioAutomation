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
            markup1.AddField(VA.Text.Markup.FieldConstants.PageName);
            markup1.AddText(" (");
            markup1.AddField(VA.Text.Markup.FieldConstants.PageNumber);
            markup1.AddText(" of ");
            markup1.AddField(VA.Text.Markup.FieldConstants.NumberOfPages);
            markup1.AddText(") ");
            markup1.SetText(s0);
        }
    }
}