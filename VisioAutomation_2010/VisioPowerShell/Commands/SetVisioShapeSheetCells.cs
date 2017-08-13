using VisioAutomation.ShapeSheet.Writers;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, VisioPowerShell.Commands.Nouns.VisioShapeSheetCells)]
    public class SetVisioShapeSheetCells : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false, Position = 0)]
        public VisioPowerShell.Models.BaseCells Cells { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter BlastGuards { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TestCircular { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes { get; set; }

        protected override void ProcessRecord()
        {
            var target_shapes = this.Shapes ?? this.Client.Selection.GetShapes();
            var targets = new VisioScripting.Models.TargetShapes(target_shapes);

            this.Client.ShapeSheet.SetShapeCells(targets, this.ApplyCells, this.BlastGuards, this.TestCircular);
        }

        public void ApplyCells(SidSrcWriter writer, short id)
        {
            this.Cells.Apply(writer, id);
        }
    }
}