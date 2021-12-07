using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands.VisioClient
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