using System.Management.Automation;

namespace VisioPowerShell.Commands.Get
{

    [Cmdlet(VerbsCommon.Get, "VisioClient")]
    public class Get_VisioClient : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            this.WriteObject(this.client);
        }
    }
}