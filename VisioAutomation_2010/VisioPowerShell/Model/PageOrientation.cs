using VA = VisioAutomation;

namespace VisioPowerShell
{
    public enum PageOrientation
    {
        None = -1,
        Portrait = VA.Pages.PrintPageOrientation.Portrait,
        Landscape = VA.Pages.PrintPageOrientation.Landscape,
        SameAsPrinter = VA.Pages.PrintPageOrientation.SameAsPrinter
    }
}