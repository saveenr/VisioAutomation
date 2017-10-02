using VisioAutomation.ShapeSheet.Writers;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, VisioPowerShell.Commands.Nouns.VisioShapeSheetCells)]
    public class SetVisioShapeSheetCells : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = true, Position = 0)]
        public VisioPowerShell.Models.BaseCells[] Cells { get; set; }

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

            if (this.Cells == null)
            {
                return;
            }

            if (this.Cells.Length < 1)
            {
                return;
            }

            targets = targets.ResolveShapes(this.Client);

            if (targets.Shapes.Count < 1)
            {
                return;
            }

            var target_ids = targets.ToShapeIDs();

            var writer = new SidSrcWriter();
            writer.BlastGuards = this.BlastGuards;
            writer.TestCircular = this.TestCircular;

            for (int i = 0; i < target_ids.ShapeIDs.Count; i++)
            {
                var shape_id = target_ids.ShapeIDs[i];
                var cells = this.Cells[i % this.Cells.Length];

                cells.Apply(writer, (short)shape_id);
            }

            var surface = this.Client.ShapeSheet.GetShapeSheetSurface();

            this.Client.WriteVerbose("BlastGuards: {0}", this.BlastGuards);
            this.Client.WriteVerbose("TestCircular: {0}", this.TestCircular);
            this.Client.WriteVerbose("Number of Shapes : {0}", target_ids.ShapeIDs.Count);

            using (var undoscope = this.Client.Application.NewUndoScope("Set Shape Cells"))
            {
                this.Client.WriteVerbose("Start Update");
                writer.Commit(surface);
                this.Client.WriteVerbose("End Update");
            }
        }
    }
}