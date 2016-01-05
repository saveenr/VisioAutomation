using System.Management.Automation;

namespace VisioPowerShell.Commands.Close
{
    [Cmdlet(VerbsCommon.Close, VisioPowerShell.Nouns.VisioApplication)]
    public class Close_VisioApplication : VisioCmdlet
    {
        [Parameter(Position = 0, Mandatory = false)]
        public SwitchParameter Force { get; set; }
        
        protected override void ProcessRecord()
        {
            this.Client.Application.Close(this.Force);
        }
    }
}
