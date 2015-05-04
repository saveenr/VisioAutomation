using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.Close, "VisioApplication")]
    public class Close_VisioApplication : VisioCmdlet
    {
        [SMA.ParameterAttribute(Position = 0, Mandatory = false)]
        public SMA.SwitchParameter Force { get; set; }
        
        protected override void ProcessRecord()
        {
            this.client.Application.Close(this.Force);
        }
    }
}