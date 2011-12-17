using VAS=VisioAutomation.Scripting;
using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("MoveTo", "Page")]
    public class MoveTo_Page : VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = false)]
        public VA.Pages.PageNavigation Flag { get; set; }

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            scriptingsession.Page.GoTo(this.Flag);
        }
    }
}