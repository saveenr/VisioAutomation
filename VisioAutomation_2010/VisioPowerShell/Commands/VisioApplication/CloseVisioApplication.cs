using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands.VisioApplication
{
    [SMA.Cmdlet(SMA.VerbsCommon.Close, Nouns.VisioApplication)]
    public class CloseVisioApplication : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter Force { get; set; }
        
        protected override void ProcessRecord()
        {
            this.Client.Application.CloseActiveApplication(this.Force);
        }
    }
}
