using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.Get, "VisioApplication")]
    public class Get_VisioApplication : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            if (this.client.Application.Get() == null)
            {
                this.WriteVerbose("A Visio Application Instance is NOT Attached");
                this.WriteObject(null);
            }
            else
            {
                this.WriteVerbose("A Visio Application Instance is Attached");
                this.WriteObject(this.client.Application.Get());
            }
        }
    }
}