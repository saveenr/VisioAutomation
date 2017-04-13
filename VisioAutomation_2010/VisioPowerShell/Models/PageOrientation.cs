namespace VisioPowerShell.Models
{
    public enum PageOrientation
    {
        None = -1,
        Portrait = VisioAutomation.Scripting.Layout.PrintPageOrientation.Portrait,
        Landscape = VisioAutomation.Scripting.Layout.PrintPageOrientation.Landscape,
        SameAsPrinter = VisioAutomation.Scripting.Layout.PrintPageOrientation.SameAsPrinter
    }
}