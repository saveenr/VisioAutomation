using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, VisioPowerShell.Commands.Nouns.VisioPageCells)]
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

            var target_pages = this.Pages ?? new []{ this.Client.Page.GetActivePage() };

            if (target_pages.Length < 1)
            {
                return;
            }


            using (var undoscope = this.Client.Application.NewUndoScope("Set Page Cells"))
            {
                for (int i = 0; i < target_pages.Length; i++)
                {
                    var target_page = target_pages[i];
                    var target_cells = this.Cells[i % this.Cells.Length];

                    this.Client.Output.WriteVerbose("Start Update Page Name={0}", target_page.NameU);

                    var target_pagesheet = target_page.PageSheet;
                    int target_pagesheet_id = target_pagesheet.ID;

                    var writer = new VisioAutomation.ShapeSheet.Writers.SidSrcWriter();
                    writer.BlastGuards = this.BlastGuards;
                    writer.TestCircular = this.TestCircular;
                    target_cells.Apply(writer, (short)target_pagesheet_id);
                    this.Client.Output.WriteVerbose("BlastGuards: {0}", this.BlastGuards);
                    this.Client.Output.WriteVerbose("TestCircular: {0}", this.TestCircular);

                    writer.Commit(target_page);
                    this.Client.Output.WriteVerbose("End Update Page Name={0}", target_page.NameU);
                }
            }
        }
    }
}