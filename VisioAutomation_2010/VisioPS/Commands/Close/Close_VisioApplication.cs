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
            var app = scriptingsession.VisioApplication;

            if (app==null)
            {
                this.WriteWarning("There is no Visio Application to stop");
                return;
            }

            // TODO: Add proper quit method to ScriptinSession

            if (this.Force)
            {
                this.ScriptingSession.Application.ForceClose();
            }
            else
            {
                app.Quit();
                scriptingsession.VisioApplication = null;                
            }
        }
    }
}