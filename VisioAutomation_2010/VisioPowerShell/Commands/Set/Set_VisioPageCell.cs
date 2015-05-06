using System.Collections;
using System.Linq;
using VisioAutomation.Application;
using VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;
using VA = VisioAutomation;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.Set, "VisioPageCell")]
    public class Set_VisioPageCell: VisioCmdlet
    {
        [SMA.ParameterAttribute(Mandatory = true,Position=0)] 
        public Hashtable Hashtable  { get; set; }
 
        [SMA.ParameterAttribute(Mandatory = false)]
        public IVisio.Page[] Pages { get; set; }

        [SMA.ParameterAttribute(Mandatory = false)]
        public SMA.SwitchParameter BlastGuards { get; set; }

        [SMA.ParameterAttribute(Mandatory = false)]
        public SMA.SwitchParameter TestCircular { get; set; }
        
        protected override void ProcessRecord()
        {
            var update = new Update();
            update.BlastGuards = this.BlastGuards;
            update.TestCircular= this.TestCircular;

            var target_pages = this.Pages ?? new[] { this.client.Page.Get() };

            var cellmap = CellSRCDictionary.GetCellMapForPages();
            var valuemap = new CellValueDictionary(cellmap);
            valuemap.UpdateValueMap(this.Hashtable,cellmap);

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

                using (var undoscope = new UndoScope(this.client.VisioApplication, "SetPageCells"))
                {
                    this.WriteVerbose("Start Update");
                    update.Execute(pagesheet);
                    this.WriteVerbose("End Update");
                }
            }

        }

    }
}