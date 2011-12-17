using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages
{
    public enum PrintPageOrientation
    {
        SameAsPrinter = IVisio.VisCellVals.visPPOSameAsPrinter,
        Portrait = IVisio.VisCellVals.visPPOPortrait,
        Landscape = IVisio.VisCellVals.visPPOLandscape
    }
}