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
            if (this.FitContents)
            {
                var bordersize = new VisioAutomation.Geometry.Size(this.BorderWidth, this.BorderWidth);
                this.Client.Page.ResizeActivePageToFitContents(bordersize, true);                
            }

            if (this.Width > 0 || this.Height > 0)
            {
                var app = this.Client.Application.GetActiveApplication();
                var page = app.ActivePage;

                var page_format_cells = new VisioAutomation.Pages.PageFormatCells();
                
                if (this.Width > 0)
                {
                    page_format_cells.Width = this.Width;
                }

                if (this.Height > 0)
                {
                    page_format_cells.Height = this.Height;
                }

                var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();
                page_format_cells.SetFormulas(writer);
                writer.BlastGuards = true;

                writer.Commit(page);
            }
        }
    }
}