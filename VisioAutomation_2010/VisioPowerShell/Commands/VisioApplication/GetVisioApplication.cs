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
                this.WriteVerbose("VisioPS not attached to a Visio application instance");
            }
            else
            {
                this.WriteVerbose("VisioPS attached to a Visio application instance");
            }

            this.WriteObject(app);
        }
    }
}