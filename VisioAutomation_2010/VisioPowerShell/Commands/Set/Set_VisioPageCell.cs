using System.Collections;
using System.Management.Automation;
using VisioAutomation.ShapeSheet.Writers;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.Set
{
    [Cmdlet(VerbsCommon.Set, VisioPowerShell.Nouns.VisioPageCell)]
    public class Set_VisioPageCell: VisioCmdlet
    {
        [Parameter(Mandatory = true,Position=0)] 
        public Hashtable Hashtable  { get; set; }

        [Parameter(Mandatory = false)]
        public SwitchParameter BlastGuards { get; set; }

        [Parameter(Mandatory = false)]
        public SwitchParameter TestCircular { get; set; }

        [Parameter(Mandatory = false)]
        public IVisio.Page[] Pages { get; set; }

        protected override void ProcessRecord()
        {
            var writer = new FormulaWriterSRC();
            writer.BlastGuards = this.BlastGuards;
            writer.TestCircular= this.TestCircular;

            var target_pages = this.Pages ?? new[] { this.Client.Page.Get() };

            var cellmap = VisioAutomation.Scripting.ShapeSheet.CellSRCDictionary.GetCellMapForPages();
            var valuemap = new VisioAutomation.Scripting.ShapeSheet.CellValueDictionary(cellmap, this.Hashtable);

            this.DumpValues(valuemap);

            foreach (var page in target_pages)
            {
                var pagesheet = page.PageSheet;

                foreach (var cellname in valuemap.CellNames)
                {
                    string cell_value = valuemap[cellname];
                    var cell_src = valuemap.GetSRC(cellname);
                    writer.SetFormula( cell_src , cell_value);
                }
                this.WriteVerbose("BlastGuards: {0}", this.BlastGuards);
                this.WriteVerbose("TestCircular: {0}", this.TestCircular);
                this.WriteVerbose("Number of Shapes : {0}", 1);
                this.WriteVerbose("Number of Total Updates: {0}", writer.Count);

                var application = this.Client.Application.Get();
                using (var undoscope = this.Client.Application.NewUndoScope("SetPageCells"))
                {
                    this.WriteVerbose("Start Update");
                    writer.Commit(pagesheet);
                    this.WriteVerbose("End Update");
                }
            }
        }
    }
}