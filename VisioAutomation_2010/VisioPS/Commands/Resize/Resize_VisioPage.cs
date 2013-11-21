using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Resize, "VisioPage")]
    public class Invoke_VisioResizePageToFitContents : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Mandatory = false)] public double Width = -1;

        [SMA.Parameter(Mandatory = false)] public double Height = -1;

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter FitContents;

        [SMA.Parameter(Mandatory = false)]
        public double BorderWidth { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public double BorderHeight { get; set; }

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            if (FitContents)
            {
                var bordersize = new VA.Drawing.Size(BorderWidth, BorderWidth);
                scriptingsession.Page.ResizeToFitContents(bordersize, true);                
            }

            if (Width > 0 || Height > 0)
            {
                var page = scriptingsession.VisioApplication.ActivePage;
                var pagecells = VA.Pages.PageCells.GetCells(page.PageSheet);

                var newpagecells = new VA.Pages.PageCells();
                
                if (Width > 0)
                {
                    newpagecells.PageWidth = this.Width;
                }

                if (Height > 0)
                {
                    newpagecells.PageHeight = this.Height;
                }

                var update = new VA.ShapeSheet.Update();
                update.SetFormulas(newpagecells);
                update.BlastGuards = true;
                update.Execute(page);
            }
        }
    }
}