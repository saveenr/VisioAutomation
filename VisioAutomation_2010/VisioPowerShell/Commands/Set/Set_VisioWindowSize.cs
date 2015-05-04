using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.Set, "VisioWindowSize")]
    public class Set_VisioWindowSize : VisioCmdlet
    {
        [SMA.ParameterAttribute(Position = 0, Mandatory = true)]
        public int Width { get; set; }

        [SMA.ParameterAttribute(Position = 1, Mandatory = true)]
        public int Height { get; set; }
        
        protected override void ProcessRecord()
        {
            var w = this.client.Application.Window;
            w.SetSize(this.Width, this.Height);
        }
    }
}