using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.Get, "VisioText")]
    public class Get_VisioText : VisioCmdlet
    {
        [SMA.ParameterAttribute(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            var t = this.client.Text.Get(this.Shapes);
            this.WriteObject(t);
        }
    }
}