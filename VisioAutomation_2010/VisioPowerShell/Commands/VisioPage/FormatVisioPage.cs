using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioPage
{
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

        [SMA.Parameter(Mandatory = false)] public IVisio.Page[] Page;

        protected override void ProcessRecord()
        {
            var targetpages = new VisioScripting.TargetPages(this.Page).Resolve(this.Client);
            if (this.FitContents || this.Width >0 || this.Height >0)
            {
                if (this.FitContents)
                {
                    var bordersize = new VisioAutomation.Geometry.Size(this.BorderWidth, this.BorderWidth);
                    this.Client.Page.ResizePageToFitContents(targetpages, bordersize);
                    var activewindow = new VisioScripting.TargetWindow();
                    this.Client.View.SetZoomToObject(activewindow, VisioScripting.Models.ZoomToObject.Page);
                }

                if (this.Width > 0 || this.Height > 0)
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
            }


            if (this.Orientation.HasValue)
            {
                this.Client.Page.SetPageOrientation(targetpages,this.Orientation.Value);
            }

            if (this.BackgroundPage != null)
            {
                // TODO: SetActivePageBackground should handle targetpages
                this.Client.Page.SetPageBackground(targetpages, this.BackgroundPage);
            }

            if (this.LayoutStyle!=null)
            {
                this.Client.Page.LayoutPage(targetpages, this.LayoutStyle);
            }
        }
    }
}