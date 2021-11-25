using VisioAutomation.ShapeSheet;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioHyperlink
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioHyperlink)]
    public class GetVisioHyperlink : VisioCmdlet
    {
        // CONTEXT:SHAPES
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shape;

        protected override void ProcessRecord()
        {
            var targetshapes = new VisioScripting.TargetShapes(this.Shape);
            var dicof_shape_to_hyperlinks = this.Client.Hyperlink.GetHyperlinks(targetshapes, VisioAutomation.Core.CellValueType.Formula);
            this.WriteObject(dicof_shape_to_hyperlinks);

        }
    }
}
 