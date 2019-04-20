using VisioAutomation.ShapeSheet;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioHyperlink
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioHyperlink)]
    public class GetVisioHyperlink : VisioCmdlet
    {
        // CONTEXT:SHAPE
        // [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            var targetshapes = new VisioScripting.TargetShapes(this.Shapes);
            var dicof_shape_to_hyperlinks = this.Client.Hyperlink.GetHyperlinks(targetshapes, CellValueType.Formula);
            this.WriteObject(dicof_shape_to_hyperlinks);

        }
    }
}
 