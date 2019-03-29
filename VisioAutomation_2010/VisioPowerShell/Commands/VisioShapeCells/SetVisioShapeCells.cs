using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, Nouns.VisioShapeCells)]
    public class SetVisioShapeCells : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = true, Position = 0)]
        public VisioPowerShell.Models.ShapeCells[] Cells { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter BlastGuards { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TestCircular { get; set; }

        protected override void ProcessRecord()
        {
            if (this.Cells == null)
            {
                return;
            }

            if (this.Cells.Length < 1)
            {
                return;
            }

            var target_shapes = this.Shapes ?? this.Client.Selection.GetShapesInSelection();

            if (target_shapes.Count < 1)
            {
                return;
            }

            var targets = new VisioScripting.Models.TargetShapes(target_shapes).ResolveShapes(this.Client);
            var target_shapeids = targets.ToShapeIDs();

            var writer = new VisioAutomation.ShapeSheet.Writers.SidSrcWriter();
            writer.BlastGuards = this.BlastGuards;
            writer.TestCircular = this.TestCircular;

            for (int i = 0; i < target_shapeids.ShapeIDs.Count; i++)
            {
                var shapeid = target_shapeids.ShapeIDs[i];
                var shape_cells = this.Cells[i % this.Cells.Length];

                shape_cells.Apply(writer, (short)shapeid);
            }

            var surface = this.Client.ShapeSheet.GetShapeSheetSurface();

            this.Client.Output.WriteVerbose("BlastGuards: {0}", this.BlastGuards);
            this.Client.Output.WriteVerbose("TestCircular: {0}", this.TestCircular);
            this.Client.Output.WriteVerbose("Number of Shapes : {0}", target_shapeids.ShapeIDs.Count);

            using (var undoscope = this.Client.Undo.NewUndoScope(nameof(SetVisioShapeCells)))
            {
                this.Client.Output.WriteVerbose("Start Update");
                writer.CommitFormulas(surface);
                this.Client.Output.WriteVerbose("End Update");
            }
        }
    }
}