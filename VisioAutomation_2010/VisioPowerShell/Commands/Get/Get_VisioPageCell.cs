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
        public SMA.SwitchParameter GetResults;

        [SMA.Parameter(Mandatory = false)]
        public ResultType ResultType = ResultType.String;

        protected override void ProcessRecord()
        {
            if (this.Cells == null)
            {
                throw new System.ArgumentException("Cells");
            }

            if (this.Cells.Length < 1)
            {
                string msg = "Must provide at least one cell name";
                throw new System.ArgumentException(msg);
            }
            var target_page = this.client.Page.Get();

            var cellmap = CellMap.GetPageCellDictionary();
            // CheckForInvalidNames(cellmap);
            
            var query = new VA.ShapeSheet.Query.CellQuery();
            Get_VisioPageCell.SetFromCellNames(query, this.Cells, cellmap);

            // Perform Query
            var surface = new ShapeSheetSurface(target_page);
            var target_shapeids = new[] { surface.Target.Page.ID };
            var dt = Helpers.QueryToDataTable(query, this.GetResults, this.ResultType, target_shapeids, surface);
            this.WriteObject(dt);
        }

        public static void SetFromCellNames(VA.ShapeSheet.Query.CellQuery query, string[] Cells, CellMap dic)
        {
            if (Cells == null)
            {
                return;
            }

            foreach (string resolved_cellname in dic.ResolveNames(Cells))
            {
                if (!query.CellColumns.Contains(resolved_cellname))
                {
                    query.AddCell(dic[resolved_cellname], resolved_cellname);
                }
            }
        }
    }
}