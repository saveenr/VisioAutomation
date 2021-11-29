using System.Linq;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioShapeCells
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, Nouns.VisioShapeCells)]
    public class SetVisioShapeCells : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = true, Position = 0)]
        public VisioPowerShell.Models.ShapeCells[] Cells { get; set; }
        
        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter BlastGuards { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TestCircular { get; set; }

        // CONTEXT:SHAPES 
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shape { get; set; }

        protected override void ProcessRecord()
        {
           
            var targetshapes = new VisioScripting.TargetShapes(this.Shape).ResolveToShapes(this.Client);

            if (targetshapes.Shapes.Count < 1)
            {
                return;
            }

            if (this.Cells == null || this.Cells.Length < 1)
            {
                return;
            }

            var targetshapeids = targetshapes.ToShapeIDs();

            var writer = new VisioAutomation.ShapeSheet.Writers.SidSrcWriter();
            writer.BlastGuards = this.BlastGuards;
            writer.TestCircular = this.TestCircular;

            foreach (int i in Enumerable.Range(0, targetshapeids.Count))
            {
                int shapeid_index = i;
                int cells_index = i % this.Cells.Length;

                var shapeid = targetshapeids[shapeid_index];
                var shape_cells = this.Cells[cells_index];

                shape_cells.Apply(writer, (short)shapeid);
            }

            var page = targetshapes.Shapes[0].ContainingPage;
            var visobjtarget = new VisioAutomation.Core.VisioObjectTarget(page);

            this.Client.Output.WriteVerbose("BlastGuards: {0}", this.BlastGuards);
            this.Client.Output.WriteVerbose("TestCircular: {0}", this.TestCircular);
            this.Client.Output.WriteVerbose("Number of Shapes : {0}", targetshapeids.Count);

            using (var undoscope = this.Client.Undo.NewUndoScope(nameof(SetVisioShapeCells)))
            {
                this.Client.Output.WriteVerbose("Start Update");
                writer.CommitFormulas(visobjtarget);
                this.Client.Output.WriteVerbose("End Update");
            }
        }
    }
}