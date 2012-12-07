using VAS=VisioAutomation.Scripting;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsDiagnostic.Test, "VisioApplication")]
    public class Test_VisioApplication: VisioPS.VisioPSCmdlet
    {
        // checks to see if we hae an active drawing open
        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var app = scriptingsession.VisioApplication;

            if (app == null)
            {
                // there is no application object 
                this.WriteObject(false);
            }
            else
            {
                // there is an application object associated with this session

                try
                {
                    // try to do something simple, read-only, and fast with the application object
                    var app_version = app.Version;
                    this.WriteObject(true);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    // If a COMException is thrown, this indicates that the
                    // application object is invalid
                    this.WriteObject(false);
                }
            }
        }
    }
}