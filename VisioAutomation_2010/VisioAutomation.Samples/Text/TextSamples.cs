
namespace VisioAutomationSamples;

public static class TextSamples
{
    public static void NonRotatingText()
    {
        var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();
        var s0 = page.DrawRectangle(1, 1, 4, 4);
        s0.Text = "Hello World";

        var src = VA.ShapeSheet.SrcConstants.TextXFormAngle;
        var cell = s0.CellsSRC[src.Section, src.Row, src.Cell];
        cell.Formula = "-Angle";
    }

    public static void TextFields()
    {
        var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();
        var s0 = page.DrawRectangle(1, 1, 4, 4);

        var markup1 = new VisioAutomation.Models.Text.Element();
        markup1.AddField(VisioAutomation.Models.Text.FieldConstants.PageName);
        markup1.AddText(" (");
        markup1.AddField(VisioAutomation.Models.Text.FieldConstants.PageNumber);
        markup1.AddText(" of ");
        markup1.AddField(VisioAutomation.Models.Text.FieldConstants.NumberOfPages);
        markup1.AddText(") ");
        markup1.SetText(s0);
    }
}