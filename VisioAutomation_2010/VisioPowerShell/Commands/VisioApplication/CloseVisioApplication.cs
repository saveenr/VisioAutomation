using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Close, VisioPowerShell.Commands.Nouns.VisioApplication)]
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
