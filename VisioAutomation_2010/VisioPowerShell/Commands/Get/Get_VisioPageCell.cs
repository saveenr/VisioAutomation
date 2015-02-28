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
            var query = new VA.ShapeSheet.Query.CellQuery();

            var dic = CellMap.GetPageCellDictionary();
            SetFromCellNames(query, this.Cells, dic);

            var surface = new ShapeSheetSurface(this.client.Page.Get());

            var target_shapeids = new[] { surface.Target.Page.ID };

            this.WriteVerbose("Number of Cells: {0}", query.CellColumns.Count);

            this.WriteVerbose("Start Query");

            var dt = Helpers.QueryToDataTable(query, this.GetResults, this.ResultType, target_shapeids, surface);

            this.WriteObject(dt);
            this.WriteVerbose("End Query");
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

        private void addcell(VA.ShapeSheet.Query.CellQuery query, bool switchpar, string cellname)
        {
            var dic = CellMap.GetPageCellDictionary();
            if (switchpar)
            {
                query.AddCell(dic[cellname], cellname);
            }
        }


    }
}