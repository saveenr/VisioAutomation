using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Scripting.Layout
{
    public enum PrintPageOrientation
    {
        SameAsPrinter = IVisio.VisCellVals.visPPOSameAsPrinter,
        Portrait = IVisio.VisCellVals.visPPOPortrait,
        Landscape = IVisio.VisCellVals.visPPOLandscape
    }
}