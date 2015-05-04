using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;
using VA = VisioAutomation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioPageCell")]
    public class Get_VisioPageCell : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false, Position = 0)]
        public string[] Cells { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Page Page { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public SMA.SwitchParameter GetResults;

        [SMA.Parameter(Mandatory = false)]
        public ResultType ResultType = ResultType.String;

        protected override void ProcessRecord()
        {
            EnsureEnoughCellNames(this.Cells);
            var target_page = this.Page ?? this.client.Page.Get();
            var cellmap = CellSRCDictionary.GetCellMapForPages();
            this.WriteVerbose("Valid Names: " + string.Join(",", cellmap.GetNames()));
            var query = cellmap.CreateQueryFromCellNames(this.Cells);
            var surface = new ShapeSheetSurface(target_page);
            var target_shapeids = new[] { surface.Target.Page.ID };
            var dt = Helpers.QueryToDataTable(query, this.GetResults, this.ResultType, target_shapeids, surface);
            this.WriteObject(dt);
        }

        public static void EnsureEnoughCellNames(IList<string> Cells)
        {
            if (Cells == null)
            {
                throw new System.ArgumentNullException("Cells");
            }

            if (Cells.Count< 1)
            {
                string msg = "Must provide at least one cell name";
                throw new System.ArgumentException(msg,"Cells");
            }
        }
    }
}