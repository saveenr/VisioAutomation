using VisioAutomation.ShapeSheet.Writers;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, VisioPowerShell.Commands.Nouns.VisioPageCells)]
    public class SetVisioPageCells : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = true, Position = 0)]
        public VisioPowerShell.Models.PageCells Cells { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Page Page { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter BlastGuards { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TestCircular { get; set; }

        protected override void ProcessRecord()
        {
            var target_page = this.Page ?? this.Client.Page.Get();
            var target_shapes = new[] {target_page.PageSheet};
            var targets = new VisioScripting.Models.TargetShapes(target_shapes);

            if (this.Cells == null)
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

            this.Cells.Apply(writer, (short)target_page.PageSheet.ID);


            this.Client.WriteVerbose("BlastGuards: {0}", this.BlastGuards);
            this.Client.WriteVerbose("TestCircular: {0}", this.TestCircular);
            this.Client.WriteVerbose("Number of Shapes : {0}", target_ids.ShapeIDs.Count);

            using (var undoscope = this.Client.Application.NewUndoScope("Set Shape Cells"))
            {
                this.Client.WriteVerbose("Start Update");
                writer.Commit(target_page);
                this.Client.WriteVerbose("End Update");
            }
        }
    }

}