using System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.Copy
{
    [Cmdlet(VerbsCommon.Copy, "VisioShape")]
    public class Copy_VisioShape : VisioCmdlet
    {
        [Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            this.client.Selection.Duplicate(this.Shapes);
        }
    }
}