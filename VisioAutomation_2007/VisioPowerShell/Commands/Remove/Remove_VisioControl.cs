using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Remove, "VisioControl")]
    public class Remove_VisioControl : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public int ControlIndex { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            this.client.Control.Delete(this.Shapes,this.ControlIndex);
        }
    }
}