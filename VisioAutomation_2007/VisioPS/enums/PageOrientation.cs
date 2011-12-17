using VA = VisioAutomation;

namespace VisioPS
{
    public enum PageOrientation
    {
        None = -1,
        Portrait = VA.Pages.PrintPageOrientation.Portrait,
        Landscape = VA.Pages.PrintPageOrientation.Landscape,
        SameAsPrinter = VA.Pages.PrintPageOrientation.SameAsPrinter
    }
}