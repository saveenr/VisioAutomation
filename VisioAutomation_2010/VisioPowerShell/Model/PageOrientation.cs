namespace VisioPowerShell.Model
{
    public enum PageOrientation
    {
        None = -1,
        Portrait = VisioAutomation.Pages.PrintPageOrientation.Portrait,
        Landscape = VisioAutomation.Pages.PrintPageOrientation.Landscape,
        SameAsPrinter = VisioAutomation.Pages.PrintPageOrientation.SameAsPrinter
    }
}