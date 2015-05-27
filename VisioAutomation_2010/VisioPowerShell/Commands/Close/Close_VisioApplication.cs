using System.Management.Automation;

namespace VisioPowerShell.Commands.Close
{
    [Cmdlet(VerbsCommon.Close, "VisioApplication")]
    public class Close_VisioApplication : VisioCmdlet
    {
        [Parameter(Position = 0, Mandatory = false)]
        public SwitchParameter Force { get; set; }
        
        protected override void ProcessRecord()
        {
            this.client.Application.Close(this.Force);
        }
    }
}