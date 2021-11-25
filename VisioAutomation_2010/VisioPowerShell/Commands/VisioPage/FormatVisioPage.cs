

namespace VisioPowerShell.Commands.VisioPage;

[SMA.Cmdlet(SMA.VerbsCommon.Format, Nouns.VisioPage)]
public class FormatVisioPage: VisioCmdlet
{
    [SMA.Parameter(Mandatory = false)]
    public double Width = -1;

    [SMA.Parameter(Mandatory = false)]
    public double Height = -1;

    [SMA.Parameter(Mandatory = false)]
    public SMA.SwitchParameter FitContents;

    [SMA.Parameter(Mandatory = false)]
    public double BorderWidth { get; set; }

    [SMA.Parameter(Mandatory = false)]
    public double BorderHeight { get; set; }

    [SMA.Parameter(Mandatory = false)] 
    public VisioScripting.Models.PageOrientation? Orientation = null;
        
    [SMA.Parameter(Mandatory = false)] 
    public string BackgroundPage = null;
        
    [SMA.Parameter(Mandatory = false)]
    public VisioAutomation.Models.LayoutStyles.LayoutStyleBase LayoutStyle = null;

    // CONTEXT:PAGES
    [SMA.Parameter(Mandatory = false)]
    public IVisio.Page[] Page;

    protected override void ProcessRecord()
    {
        var targetpages = new VisioScripting.TargetPages(this.Page).ResolveToPages(this.Client);

        if (this.FitContents)
        {
            var bordersize = new VisioAutomation.Geometry.Size(this.BorderWidth, this.BorderHeight);
            this.Client.Page.ResizePageToFitContents(targetpages, bordersize);
            this.Client.View.SetZoomToObject(VisioScripting.TargetWindow.Auto, VisioScripting.Models.ZoomToObject.Page);
        }

        if (this.Width >0 || this.Height >0)
        {
            var page_format_cells = new VisioAutomation.Pages.PageFormatCells();

            if (this.Width > 0)
            {
                page_format_cells.Width = this.Width;
            }

            if (this.Height > 0)
            {
                page_format_cells.Height = this.Height;
            }

            this.Client.Page.SetPageFormatCells(targetpages, page_format_cells);
        }


        if (this.Orientation.HasValue)
        {
            this.Client.Page.SetPageOrientation(targetpages,this.Orientation.Value);
        }

        if (this.BackgroundPage != null)
        {
            this.Client.Page.SetPageBackground(targetpages, this.BackgroundPage);
        }

        if (this.LayoutStyle!=null)
        {
            this.Client.Page.LayoutPage(targetpages, this.LayoutStyle);
        }
    }
}