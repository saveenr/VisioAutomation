using System.Collections;
using System.Linq;
using System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.Set
{
    [Cmdlet(VerbsCommon.Set, "VisioPageCell")]
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
            var update = new VisioAutomation.ShapeSheet.Update();
            update.BlastGuards = this.BlastGuards;
            update.TestCircular= this.TestCircular;

            var target_pages = this.Pages ?? new[] { this.client.Page.Get() };

            var cellmap = CellSRCDictionary.GetCellMapForPages();
            var valuemap = new CellValueDictionary(cellmap, this.Hashtable);

            this.DumpValues(valuemap);

            foreach (var page in target_pages)
            {
                var pagesheet = page.PageSheet;

                foreach (var cellname in valuemap.CellNames)
                {
                    string cell_value = valuemap[cellname];
                    var cell_src = valuemap.GetSRC(cellname);
                    update.SetFormulaIgnoreNull( cell_src , cell_value);
                }
                this.WriteVerbose("BlastGuards: {0}", this.BlastGuards);
                this.WriteVerbose("TestCircular: {0}", this.TestCircular);
                this.WriteVerbose("Number of Shapes : {0}", 1);
                this.WriteVerbose("Number of Total Updates: {0}", update.Count());

                var application = this.client.Application.Get();
                using (var undoscope = this.client.Application.NewUndoScope("SetPageCells"))
                {
                    this.WriteVerbose("Start Update");
                    update.Execute(pagesheet);
                    this.WriteVerbose("End Update");
                }
            }
        }
    }
}