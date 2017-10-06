using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, VisioPowerShell.Commands.Nouns.VisioWindowSize)]
    public class SetVisioWindowSize : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public int Width { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public int Height { get; set; }
        
        protected override void ProcessRecord()
        {
            var w = this.Client.Application.Window;
            w.SetSize(this.Width, this.Height);
        }
    }
}