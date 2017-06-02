using System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.Close, VisioPowerShell.Commands.Nouns.VisioApplication)]
    public class CloseVisioApplication : VisioCmdlet
    {
        [Parameter(Position = 0, Mandatory = false)]
        public SwitchParameter Force { get; set; }
        
        protected override void ProcessRecord()
        {
            this.Client.Application.Close(this.Force);
        }
    }
}
