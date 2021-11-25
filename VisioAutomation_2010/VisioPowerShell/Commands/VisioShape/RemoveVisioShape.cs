

namespace VisioPowerShell.Commands.VisioShape
{
    [SMA.Cmdlet(SMA.VerbsCommon.Remove, Nouns.VisioShape)]
    public class RemoveVisioShape : VisioCmdlet
    {
        // CONTEXT:SHAPES
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shape;

        protected override void ProcessRecord()
        {
            var targetshapes = new VisioScripting.TargetShapes(this.Shape);
            this.Client.Application.DeleteShapes(targetshapes);
        }
    }
}