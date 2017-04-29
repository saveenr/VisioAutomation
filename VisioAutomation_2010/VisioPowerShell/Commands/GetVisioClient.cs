using System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.Get, VisioPowerShell.Commands.Nouns.VisioClient)]
    public class GetVisioClient : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            this.WriteObject(this.Client);
        }
    }
}