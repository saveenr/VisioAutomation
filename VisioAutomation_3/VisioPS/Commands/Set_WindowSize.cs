using VAS=VisioAutomation.Scripting;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Set", "WindowSize")]
    public class Set_WindowSize : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public int Width { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public int Height { get; set; }
        
        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            scriptingsession.SetApplicationWindowSize(Width, Height);
        }
    }
}