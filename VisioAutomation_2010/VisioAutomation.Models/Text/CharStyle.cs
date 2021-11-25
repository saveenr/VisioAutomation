
namespace VisioAutomation.Models.Text;

[System.Flags]
public enum CharStyle
{
    None = 0,
    Bold = IVisio.VisCellVals.visBold,
    Italic = IVisio.VisCellVals.visItalic,
    UnderLine = IVisio.VisCellVals.visUnderLine,
    SmallCaps = IVisio.VisCellVals.visSmallCaps
}