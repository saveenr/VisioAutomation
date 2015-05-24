using System.Management.Automation;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands.Get
{

    [Cmdlet(SMA.VerbsCommon.Get, "VisioClient")]
    public class Get_VisioClient : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            this.WriteObject(this.client);
        }
    }
}