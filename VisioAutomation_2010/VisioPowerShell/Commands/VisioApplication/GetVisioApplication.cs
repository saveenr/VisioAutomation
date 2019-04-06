using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands.VisioApplication
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioApplication)]
    public class GetVisioApplication : VisioCmdlet
    {
        protected override void ProcessRecord()
        {
            var app = this.Client.Application.GetAttachedApplication();
            if (app == null)
            {
                this.WriteVerbose("A Visio Application Instance is not Attached");
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