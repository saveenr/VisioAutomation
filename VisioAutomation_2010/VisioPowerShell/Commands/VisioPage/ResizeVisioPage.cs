using SMA = System.Management.Automation;
using VisioAutomation.ShapeSheet;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Resize, VisioPowerShell.Commands.Nouns.VisioPage)]
    public class ResizeVisioPage : VisioCmdlet
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

        protected override void ProcessRecord()
        {
            var cmdtarget = this.Client.GetCommandTargetPage();
            var tp = new VisioScripting.Models.TargetPages(cmdtarget.ActivePage);

            if (this.FitContents)
            {
                var bordersize = new VisioAutomation.Geometry.Size(this.BorderWidth, this.BorderWidth);
                this.Client.Page.ResizePageToFitContents(tp, bordersize);
                this.Client.View.ZoomActiveWindowToObject(VisioScripting.Models.Zoom.ToPage);
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

                this.Client.Page.SetPageFormatCells(tp, page_format_cells);
            }
        }
    }
}