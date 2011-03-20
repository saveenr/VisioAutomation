using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPS
{
    public enum TextVerticalAlignment
    {
        None = -1,       
        Top = IVisio.VisCellVals.visVertTop,
        Middle = IVisio.VisCellVals.visVertMiddle,
        Bottom = IVisio.VisCellVals.visVertBottom
    }
}