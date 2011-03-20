using Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    [System.Flags]
    public enum CharStyle
    {
        None = 0,
        Bold = VisCellVals.visBold,
        Italic = VisCellVals.visItalic,
        UnderLine = VisCellVals.visUnderLine,
        SmallCaps = VisCellVals.visSmallCaps
    }
}