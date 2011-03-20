using VAS=VisioAutomation.Scripting;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Set", "Selection")]
    public class Set_Selection : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = false)]
        public SelectionOperation Operation { get; set; }

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;

            if (this.Operation == VisioPS.SelectionOperation.All)
            {
                scriptingsession.Selection.SelectAll();
            }
            else if (this.Operation == VisioPS.SelectionOperation.None)
            {
                scriptingsession.Selection.SelectNone();
            }
            else if (this.Operation == VisioPS.SelectionOperation.Invert)
            {
                scriptingsession.Selection.SelectInvert();
            }
        }
    }
}