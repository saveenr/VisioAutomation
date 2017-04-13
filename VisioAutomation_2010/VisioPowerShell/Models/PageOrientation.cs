using VisioAutomation.Scripting.Models;

namespace VisioPowerShell.Models
{
    public enum PageOrientation
    {
        None = -1,
        Portrait = PrintPageOrientation.Portrait,
        Landscape = PrintPageOrientation.Landscape,
        SameAsPrinter = PrintPageOrientation.SameAsPrinter
    }
}