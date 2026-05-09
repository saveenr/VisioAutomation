using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioShape
{
    [SMA.Cmdlet(SMA.VerbsDiagnostic.Measure, Nouns.VisioShape)]
    public class MeasureVisioShape: VisioCmdlet
    {
        // CONTEXT:SHAPES
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shape;

        protected override void ProcessRecord()
        {
            var targetshapes = new VisioScripting.TargetShapes(this.Shape);
            var list_shapedim = this.Client.Selection.GetShapeDimensions(targetshapes);
            this.WriteObject(list_shapedim,true);
        }
    }
}