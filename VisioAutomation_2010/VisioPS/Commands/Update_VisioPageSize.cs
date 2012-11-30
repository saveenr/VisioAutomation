using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsData.Update, "VisioPageSize")]
    public class Update_VisioPageSize : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public PageSizeOperations PageSizeOperations { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public double BorderWidth { get; set; }

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var bordersize = new VA.Drawing.Size(BorderWidth, BorderWidth);
            if (this.PageSizeOperations == VisioPS.PageSizeOperations.FitContents)
            {
                scriptingsession.Page.ResizeToFitContents(bordersize, true);
            }
        }
    }
}