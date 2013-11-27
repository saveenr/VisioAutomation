using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Close, "VisioApplication")]
    public class Close_VisioApplication : VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = false)]
        public SMA.SwitchParameter Force { get; set; }
        
        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            scriptingsession.Application.Close(this.Force);
        }
    }
}