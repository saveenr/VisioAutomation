using VisioAutomation.Pages;
using VA = VisioAutomation;

namespace VisioPowerShell
{
    public enum PageOrientation
    {
        None = -1,
        Portrait = PrintPageOrientation.Portrait,
        Landscape = PrintPageOrientation.Landscape,
        SameAsPrinter = PrintPageOrientation.SameAsPrinter
    }
}