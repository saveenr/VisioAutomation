using System.Management.Automation;
using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands.Get
{
    [Cmdlet(VerbsCommon.Get, "VisioApplication")]
    public class Get_VisioApplication : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var app = this.client.Application.Get();
            if (app == null)
            {
                this.WriteVerbose("A Visio Application Instance is NOT Attached");
                this.WriteObject(null);
            }
            else
            {
                this.WriteVerbose("A Visio Application Instance is Attached");
                this.WriteObject(app);
            }
        }
    }
}