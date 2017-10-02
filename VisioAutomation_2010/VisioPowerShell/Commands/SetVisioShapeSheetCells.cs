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

            this.SetCells(this.Client, targets, this.ApplyCells, this.BlastGuards, this.TestCircular);
        }

        public void ApplyCells(SidSrcWriter writer, short id)
        {
            this.Cells.Apply(writer, id);
        }

        public void SetCells(VisioScripting.Client _client, VisioScripting.Models.TargetShapes targets, System.Action<SidSrcWriter, short> apply_cells, bool blast_guards, bool test_circular)
        {
            targets = targets.ResolveShapes(_client);
            var target_ids = targets.ToShapeIDs();
            this.SetCells(_client, target_ids, apply_cells, blast_guards, test_circular);
        }

        public void SetCells(VisioScripting.Client _client, VisioScripting.Models.TargetShapeIDs targets, System.Action<SidSrcWriter, short> apply_cells, bool blast_guards, bool test_circular)
        {
            var writer = new SidSrcWriter();
            writer.BlastGuards = blast_guards;
            writer.TestCircular = test_circular;

            foreach (var shape_id in targets.ShapeIDs)
            {
                apply_cells(writer, (short)shape_id);
            }

            var surface = _client.ShapeSheet.GetShapeSheetSurface();

            _client.WriteVerbose("BlastGuards: {0}", blast_guards);
            _client.WriteVerbose("TestCircular: {0}", test_circular);
            _client.WriteVerbose("Number of Shapes : {0}", targets.ShapeIDs.Count);

            using (var undoscope = _client.Application.NewUndoScope("Set Shape Cells"))
            {
                _client.WriteVerbose("Start Update");
                writer.Commit(surface);
                _client.WriteVerbose("End Update");
            }
        }

    }
}