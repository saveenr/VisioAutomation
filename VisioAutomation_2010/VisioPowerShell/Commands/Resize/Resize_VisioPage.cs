using System.Management.Automation;

namespace VisioPowerShell.Commands.Resize
{
    [Cmdlet(VerbsCommon.Resize, "VisioPage")]
    public class Resize_VisioPage : VisioCmdlet
    {
        [Parameter(Mandatory = false)] public double Width = -1;

        [Parameter(Mandatory = false)] public double Height = -1;

        [Parameter(Mandatory = false)]
        public SwitchParameter FitContents;

        [Parameter(Mandatory = false)]
        public double BorderWidth { get; set; }

        [Parameter(Mandatory = false)]
        public double BorderHeight { get; set; }

        protected override void ProcessRecord()
        {
            if (this.FitContents)
            {
                var bordersize = new VisioAutomation.Drawing.Size(this.BorderWidth, this.BorderWidth);
                this.client.Page.ResizeToFitContents(bordersize, true);                
            }

            if (this.Width > 0 || this.Height > 0)
            {
                var page = this.client.Application.Get().ActivePage;
                var pagecells = VisioAutomation.Pages.PageCells.GetCells(page.PageSheet);

                var newpagecells = new VisioAutomation.Pages.PageCells();
                
                if (this.Width > 0)
                {
                    newpagecells.PageWidth = this.Width;
                }

                if (this.Height > 0)
                {
                    newpagecells.PageHeight = this.Height;
                }

                var update = new VisioAutomation.ShapeSheet.Update();
                update.SetFormulas(newpagecells);
                update.BlastGuards = true;
                update.Execute(page);
            }
        }
    }
}