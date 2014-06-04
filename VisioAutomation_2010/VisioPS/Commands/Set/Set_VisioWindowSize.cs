using VAS = VisioAutomation.Scripting;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, "VisioWindowSize")]
    public class Set_VisioWindowSize : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public int Width { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public int Height { get; set; }
        
        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var w = scriptingsession.Application.Window;
            w.SetSize(Width, Height);
        }
    }
}