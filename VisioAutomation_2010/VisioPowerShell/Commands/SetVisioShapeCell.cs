using System.Collections;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, VisioPowerShell.Commands.Nouns.VisioShapeCell)]
    public class SetVisioShapeCell : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false, Position = 0)]
        public VisioScripting.Models.ShapeCells Cells { get; set; }

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

            this.Client.ShapeSheet.SetShapeCells(targets, this.Cells, this.BlastGuards, this.TestCircular);
        }
    }
}