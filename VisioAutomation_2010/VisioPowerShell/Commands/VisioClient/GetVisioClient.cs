

namespace VisioPowerShell.Commands.VisioScriptingClient
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioClient)]
    public class GetVisioClient : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            this.WriteObject(this.Client);
        }
    }
}