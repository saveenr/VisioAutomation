using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioConnectionPoint
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioConnectionPoint)]
    public class GetVisioConnectionPoint : VisioCmdlet
    {

        // CONTEXT:SHAPES
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shape;

        protected override void ProcessRecord()
        {
            var targetshapes = new VisioScripting.TargetShapes(this.Shape);

            var dic = this.Client.ConnectionPoint.GetConnectionPoints(targetshapes);

            this.WriteObject(dic);
        }
    }
}