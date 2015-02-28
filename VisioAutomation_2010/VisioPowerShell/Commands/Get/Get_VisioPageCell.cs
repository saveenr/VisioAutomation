using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;
using VA = VisioAutomation;
using System.Linq;

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
            EnsureEnoughCellNames(this.Cells);
            var target_page = this.client.Page.Get();
            var cellmap = CellMap.GetPageCellDictionary();
            this.WriteVerbose("Valid Names: " + string.Join(",", cellmap.GetNames()));
            CheckForInvalidNames(cellmap,this.Cells);
            var query = Get_VisioPageCell.CreateQueryFromCellNames(this.Cells, cellmap);
            var surface = new ShapeSheetSurface(target_page);
            var target_shapeids = new[] { surface.Target.Page.ID };
            var dt = Helpers.QueryToDataTable(query, this.GetResults, this.ResultType, target_shapeids, surface);
            this.WriteObject(dt);
        }

        public static void EnsureEnoughCellNames(IList<string> Cells)
        {
            if (Cells == null)
            {
                throw new System.ArgumentException("Cells");
            }

            if (Cells.Count< 1)
            {
                string msg = "Must provide at least one cell name";
                throw new System.ArgumentException(msg);
            }
        }

        public static VisioAutomation.ShapeSheet.Query.CellQuery CreateQueryFromCellNames(string[] Cells, CellMap dic)
        {
            var query = new VisioAutomation.ShapeSheet.Query.CellQuery();

            foreach (string resolved_cellname in dic.ResolveNames(Cells))
            {
                if (!query.CellColumns.Contains(resolved_cellname))
                {
                    query.AddCell(dic[resolved_cellname], resolved_cellname);
                }
            }
            return query;
        }

        public static void CheckForInvalidNames(CellMap cellmap, IList<string> Cells)
        {
            var invalid_names = Cells.Where(cellname => !cellmap.ContainsCell(cellname)).ToList();
            if (invalid_names.Count > 0)
            {
                string msg = "Invalid cell names: " + string.Join(",", invalid_names);
                throw new System.ArgumentException(msg);
            }
        }

    }
}