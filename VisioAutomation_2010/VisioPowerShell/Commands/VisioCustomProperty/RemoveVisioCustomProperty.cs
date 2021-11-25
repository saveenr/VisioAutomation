

namespace VisioPowerShell.Commands.VisioCustomProperty
{
    [SMA.Cmdlet(SMA.VerbsCommon.Remove, Nouns.VisioCustomProperty)]
    public class RemoveVisioCustomProperty : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public string Name { get; set; }

        // CONTEXT:SHAPES
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shape;

        protected override void ProcessRecord()
        {
            var targetshapes = new VisioScripting.TargetShapes(this.Shape);
            this.Client.CustomProperty.DeleteCustomPropertyWithName(targetshapes, this.Name);
        }
    }
}

