using VA = VisioAutomation;

namespace VisioAutomationSamples
{
    public static class TextSamples
    {
        public static void NonRotatingText()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();
            var s0 = page.DrawRectangle(1, 1, 4, 4);
            s0.Text = "Hello World";

            var src = VA.ShapeSheet.SRCConstants.TxtAngle;
            var cell = s0.CellsSRC[src.Section, src.Row, src.Cell];
            cell.Formula = "-Angle";
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