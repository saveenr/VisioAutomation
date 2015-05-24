using System.Management.Automation;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.Get
{
    [Cmdlet(SMA.VerbsCommon.Get, "VisioText")]
    public class Get_VisioText : VisioCmdlet
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