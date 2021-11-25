

namespace VisioPowerShell.Commands.VisioText
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioText)]
    public class GetVisioText : VisioCmdlet
    {
        // CONTEXT:SHAPES 

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shape;
        protected override void ProcessRecord()
        {
            var targetshapes = new VisioScripting.TargetShapes(this.Shape);
            var listof_string = this.Client.Text.GetShapeText(targetshapes);
            this.WriteObject(listof_string);
        }
    }
}