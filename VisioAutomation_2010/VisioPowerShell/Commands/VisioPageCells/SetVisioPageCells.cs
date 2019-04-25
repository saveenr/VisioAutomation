using System.Linq;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioPageCells
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, Nouns.VisioPageCells)]
    public class SetVisioPageCells : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = true, Position = 0)]
        public VisioPowerShell.Models.PageCells[] Cells { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter BlastGuards { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TestCircular { get; set; }

        // CONTEXT:PAGES
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Page[] Page { get; set; }

        protected override void ProcessRecord()
        {
            var targetpages = new VisioScripting.TargetPages(this.Page).Resolve(this.Client);

            if (targetpages.Pages.Count < 1)
            {
                return;
            }

            if (this.Cells == null || this.Cells.Length < 1)
            {
                return;
            }

            this.Client.Output.WriteVerbose("BlastGuards: {0}", this.BlastGuards);
            this.Client.Output.WriteVerbose("TestCircular: {0}", this.TestCircular);

            using (var undoscope = this.Client.Undo.NewUndoScope(nameof(SetVisioPageCells)))
            {
                foreach (int i in Enumerable.Range(0,targetpages.Pages.Count))
                {
                    int page_index = i;
                    int cells_index = i % this.Cells.Length;

                    var page = targetpages.Pages[page_index];
                    var cells = this.Cells[cells_index];

                    this.Client.Output.WriteVerbose("Start Update Page Name={0}", page.NameU);

                    var shapesheet = page.PageSheet;
                    int shapeid = shapesheet.ID;

                    var writer = new VisioAutomation.ShapeSheet.Writers.SidSrcWriter();
                    writer.BlastGuards = this.BlastGuards;
                    writer.TestCircular = this.TestCircular;
                    cells.Apply(writer, (short)shapeid);
                    writer.Commit(page, VisioAutomation.ShapeSheet.CellValueType.Formula);

                    this.Client.Output.WriteVerbose("End Update Page Name={0}", page.NameU);
                }
            }
        }
    }
}