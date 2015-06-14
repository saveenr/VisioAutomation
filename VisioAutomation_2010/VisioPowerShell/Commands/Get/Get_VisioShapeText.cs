using System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.Get
{
    [Cmdlet(VerbsCommon.Get, "VisioShapeText")]
    public class Get_VisioShapeText : VisioCmdlet
    {
        [Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            var t = this.client.Text.Get(this.Shapes);
            this.WriteObject(t);
        }
    }
}