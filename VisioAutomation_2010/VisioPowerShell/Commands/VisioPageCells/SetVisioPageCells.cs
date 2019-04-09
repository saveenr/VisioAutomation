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
        public IVisio.Page[] Pages { get; set; }

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

            var targetpages = new VisioScripting.TargetPages(this.Pages);

            this.Client.Output.WriteVerbose("BlastGuards: {0}", this.BlastGuards);
            this.Client.Output.WriteVerbose("TestCircular: {0}", this.TestCircular);

            using (var undoscope = this.Client.Undo.NewUndoScope(new VisioScripting.TargetActiveApplication(), nameof(SetVisioPageCells)))
            {
                for (int i = 0; i < targetpages.Pages.Count; i++)
                {
                    var targetpage = targetpages.Pages[i];
                    this.Client.Output.WriteVerbose("Start Update Page Name={0}", targetpage.NameU);

                    var targetpage_shapesheet = targetpage.PageSheet;
                    int targetpage_shapesheetid = targetpage_shapesheet.ID;
                    var target_cells = this.Cells[i % this.Cells.Length];
                    var writer = new VisioAutomation.ShapeSheet.Writers.SidSrcWriter();
                    writer.BlastGuards = this.BlastGuards;
                    writer.TestCircular = this.TestCircular;
                    target_cells.Apply(writer, (short)targetpage_shapesheetid);
                    writer.Commit(targetpage, VisioAutomation.ShapeSheet.CellValueType.Formula);

                    this.Client.Output.WriteVerbose("End Update Page Name={0}", targetpage.NameU);
                }
            }
        }
    }
}