using VA = VisioAutomation;

namespace VisioPS
{
    public enum PageOrientation
    {
        None = -1,
        Portrait = VA.Layout.PrintPageOrientation.Portrait,
        Landscape = VA.Layout.PrintPageOrientation.Landscape,
        SameAsPrinter = VA.Layout.PrintPageOrientation.SameAsPrinter
    }
}