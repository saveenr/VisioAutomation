using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;
using System.Linq;
using VA = VisioAutomation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, "VisioPageCell")]
    public class Set_VisioPageCell: VisioCmdlet
    {
        [SMA.Parameter(Mandatory = true,Position=0)] 
        public System.Collections.Hashtable Hashtable  { get; set; }
 
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Page[] Pages { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter BlastGuards { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter TestCircular { get; set; }
        
        protected override void ProcessRecord()
        {
            var update = new VisioAutomation.ShapeSheet.Update();
            update.BlastGuards = this.BlastGuards;
            update.TestCircular= this.TestCircular;

            var target_pages = this.Pages ?? new[] { this.client.Page.Get() };

            var dic = CellMap.GetCellMapForPages();
            var valuemap = new CellValueMap(dic);

            valuemap.UpdateValueMap(this.Hashtable);

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
                this.WriteVerbose("Number of Updates per Shape: {0}", update.Count() / 1);

                using (var undoscope = new VA.Application.UndoScope(this.client.VisioApplication, "SetPageCells"))
                {
                    this.WriteVerbose("Start Update");
                    update.Execute(pagesheet);
                    this.WriteVerbose("End Update");
                }
            }

        }

    }
}