using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    public enum TabStopAlignment
    {
        Left = IVisio.VisCellVals.visTabStopLeft,
        Center = IVisio.VisCellVals.visTabStopCenter,
        Right = IVisio.VisCellVals.visTabStopRight,
        Decimal = IVisio.VisCellVals.visTabStopDecimal,
        Comma = IVisio.VisCellVals.visTabStopComma
    }
}