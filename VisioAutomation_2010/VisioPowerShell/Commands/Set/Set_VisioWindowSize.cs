using System.Management.Automation;

namespace VisioPowerShell.Commands.Set
{
    [Cmdlet(VerbsCommon.Set, "VisioWindowSize")]
    public class Set_VisioWindowSize : VisioCmdlet
    {
        [Parameter(Position = 0, Mandatory = true)]
        public int Width { get; set; }

        [Parameter(Position = 1, Mandatory = true)]
        public int Height { get; set; }
        
        protected override void ProcessRecord()
        {
            var w = this.client.Application.Window;
            w.SetSize(this.Width, this.Height);
        }
    }
}