using VAS=VisioAutomation.Scripting;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "VisioApplication")]
    public class New_VisioApplication : VisioPSCmdlet
    {
        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var app = scriptingsession.Application.NewApplication();
            Globals.Application = app;
            //this.WriteObject(app); // TODO: investigate why calling write-object and returning app can cause the visio application to have an error when it shuts down
        }
    }
}