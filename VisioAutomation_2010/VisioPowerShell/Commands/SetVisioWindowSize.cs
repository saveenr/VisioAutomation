using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, VisioPowerShell.Commands.Nouns.VisioWindow)]
    public class SetVisioWindow : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public int Width { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public int Height { get; set; }
        
        protected override void ProcessRecord()
        {
            if (this.Width > 0 || this.Height > 0)
            {
                var w = this.Client.Window;
                w.SetSize(this.Width, this.Height);
            }
        }
    }
}