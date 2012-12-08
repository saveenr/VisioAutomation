using VAS=VisioAutomation.Scripting;
using SMA = System.Management.Automation;
using VisioPS.Extensions;

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
                this.WriteVerbose("Session's Application object is null");
                this.WriteObject(false);
            }
            else
            {
                this.WriteVerbose("Session's Application object is not null");
                try
                {
                    this.WriteVerbose("Attempting to read Visio Application's Version property");
                    // try to do something simple, read-only, and fast with the application object
                    var app_version = app.Version;
                    this.WriteVerbose(
                        "No COMException was thrown when reading Version property. This application instance seems valid");
                    this.WriteObject(true);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    this.WriteVerbose("COMException thrown");
                    this.WriteVerbose("This application instance is invalid");
                    // If a COMException is thrown, this indicates that the
                    // application object is invalid
                    this.WriteObject(false);
                }
                catch (System.Exception exc)
                {
                    this.WriteVerbose("An exception besides COMException was thrown");
                    // just re-raise it.
                    throw exc;
                }
            }
        }
    }
}