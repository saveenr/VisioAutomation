using VisioAutomation.ShapeSheet;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioLockCells
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioLockCells)]
    public class GetVisioLockCells : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            var targetshapes = new VisioScripting.TargetShapes(this.Shapes);
            var dic = this.Client.Lock.GetLockCells(targetshapes, CellValueType.Formula);
            this.WriteObject(dic, true);
        }
    }
}
