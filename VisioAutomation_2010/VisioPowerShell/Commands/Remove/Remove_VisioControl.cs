using System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.Remove
{
    [Cmdlet(VerbsCommon.Remove, "VisioControl")]
    public class Remove_VisioControl : VisioCmdlet
    {
        [Parameter(Position = 0, Mandatory = true)]
        public int ControlIndex { get; set; }

        [Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            this.client.Control.Delete(this.Shapes,this.ControlIndex);
        }
    }
}