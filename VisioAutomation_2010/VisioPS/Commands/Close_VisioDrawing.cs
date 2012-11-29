using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Close", "VisioDocument")]
    public class Close_VisioDocument: VisioPS.VisioPSCmdlet
    {
        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            scriptingsession.Document.Close(true);

            // TODO: This cmdlet should accept an optional parameter identifying which document to close
        }
    }
}