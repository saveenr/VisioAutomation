using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands.VisioScriptingClient
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioScriptingClient)]
    public class GetVisioScriptingClient : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            this.WriteObject(this.Client);
        }
    }
}