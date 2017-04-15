using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting.Models
{
    public enum PageOrientation
    {
        SameAsPrinter = IVisio.VisCellVals.visPPOSameAsPrinter,
        Portrait = IVisio.VisCellVals.visPPOPortrait,
        Landscape = IVisio.VisCellVals.visPPOLandscape
    }
}