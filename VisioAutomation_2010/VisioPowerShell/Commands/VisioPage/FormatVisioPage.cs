using SMA = System.Management.Automation;

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

        protected override void ProcessRecord()
        {
            if (this.FitContents || this.Width >0 || this.Height >0)
            {
                var cmdtarget = this.Client.GetCommandTargetPage();
                var targetpages = new VisioScripting.Models.TargetPages(cmdtarget.ActivePage);

                if (this.FitContents)
                {
                    var bordersize = new VisioAutomation.Geometry.Size(this.BorderWidth, this.BorderWidth);
                    this.Client.Page.ResizePageToFitContents(targetpages, bordersize);
                    this.Client.View.SetActiveWindowZoomToObject(VisioScripting.Models.ZoomToObject.Page);
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
                var cmdtarget = this.Client.GetCommandTargetPage();
                var targetpages = new VisioScripting.Models.TargetPages(cmdtarget.ActivePage);
                this.Client.Page.SetPageOrientation(targetpages,this.Orientation.Value);
            }

            if (this.BackgroundPage != null)
            {
                this.Client.Page.SetActivePageBackground(this.BackgroundPage);
            }

            if (this.LayoutStyle!=null)
            {
                var cmdtarget = this.Client.GetCommandTargetPage();
                var targetpage = new VisioScripting.Models.TargetPage(cmdtarget.ActivePage);
                this.Client.Page.LayoutPage(targetpage, this.LayoutStyle);
            }
        }
    }
}