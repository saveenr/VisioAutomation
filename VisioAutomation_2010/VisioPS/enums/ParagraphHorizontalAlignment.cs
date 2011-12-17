using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPS
{
    public enum ParagraphHorizontalAlignment
    {
        None,
        Left = IVisio.VisCellVals.visHorzLeft,
        Center = IVisio.VisCellVals.visHorzCenter,
        Right = IVisio.VisCellVals.visHorzRight,
        Justify = IVisio.VisCellVals.visHorzJustify,
        Force = IVisio.VisCellVals.visHorzForce
    }
}