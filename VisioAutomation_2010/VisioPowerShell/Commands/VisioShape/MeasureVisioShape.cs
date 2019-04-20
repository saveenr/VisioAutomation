using System.Linq;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioPage
{
    [SMA.Cmdlet(SMA.VerbsDiagnostic.Measure, Nouns.VisioShape)]
    public class MeasureVisioShape: VisioCmdlet
    {
        // CONTEXT:SHAPE 
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shape;

        protected override void ProcessRecord()
        {

            var targetshapes = new VisioScripting.TargetShapes(this.Shape).Resolve(this.Client);

            if (targetshapes.Shapes.Count < 1)
            {
                return;
            }


            var shapeids = VisioAutomation.ShapeIDPairs.FromShapes(targetshapes.Shapes).Select(i => i.ShapeID).ToList();
            var page = targetshapes.Shapes[0].ContainingPage;
            var list_shapedim = VisioScripting.Models.ShapeDimensions.Get_ShapeDimensions(page, shapeids);

            this.WriteObject(list_shapedim,true);

        }


    }
}