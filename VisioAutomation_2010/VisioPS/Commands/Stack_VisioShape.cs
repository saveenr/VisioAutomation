using VAS=VisioAutomation.Scripting;
using VA=VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Stack", "VisioShape")]
    public class Stack_VisioShape : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public VA.Drawing.Axis Axis { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true)] public double Distance = 0.0;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            scriptingsession.Layout.Stack(this.Axis, this.Distance);
        }
    }
}