using SMA = System.Management.Automation;
using VA=VisioAutomation;

namespace VisioPowerShell.Commands
{

    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioClient")]
    public class Get_VisioClient : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            this.WriteObject(this.client);
        }
    }
}