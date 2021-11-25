using VisioAutomation.ShapeSheet;


namespace VisioPowerShell.Commands.VisioLockCells
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioLockCells)]
    public class GetVisioLockCells : VisioCmdlet
    {
        // CONTEXT:SHAPES
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shape;

        protected override void ProcessRecord()
        {
            var targetshapes = new VisioScripting.TargetShapes(this.Shape);
            var dic = this.Client.Lock.GetLockCells(targetshapes, CellValueType.Formula);
            this.WriteObject(dic, true);
        }
    }
}
