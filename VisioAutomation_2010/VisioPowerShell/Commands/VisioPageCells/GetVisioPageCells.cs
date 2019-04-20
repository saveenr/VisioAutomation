using System;
using System.Collections.Generic;
using System.Linq;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioPageCells
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioPageCells)]
    public class GetVisioPageCells: VisioCmdlet
    {

        [SMA.Parameter(Mandatory = false)]
        public VisioPowerShell.Models.CellOutputType OutputType = VisioPowerShell.Models.CellOutputType.Formula;

        // CONTEXT:PAGES
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Page[] Pages { get; set; }

        protected override void ProcessRecord()
        {
            var targetpages = new VisioScripting.TargetPages(this.Pages);

            var template = new VisioPowerShell.Models.PageCells();
            var celldic = VisioPowerShell.Models.NamedSrcDictionary.FromCells(template);
            var cellnames = celldic.Keys.ToArray();
            var query = _CreateQuery(celldic, cellnames);
            var surface = this.Client.ShapeSheet.GetShapeSheetSurface();
            
            var result_dt = new System.Data.DataTable();

            foreach (var targetpage in targetpages.Pages)
            {
                var targetpage_shapesheet = targetpage.PageSheet;
                var targetpage_shapeshapeids = new List<int> { targetpage_shapesheet.ID };
                var dt = VisioPowerShell.Models.DataTableHelpers.QueryToDataTable(query, this.OutputType, targetpage_shapeshapeids, surface);
                result_dt.Merge(dt);
            }

            // Annotate the returned datatable to disambiguate rows
            var pageindex_col = result_dt.Columns.Add("PageIndex", typeof(System.Int32));
            pageindex_col.SetOrdinal(0);
            for (int row_index = 0; row_index < targetpages.Pages.Count; row_index++)
            {
                result_dt.Rows[row_index][pageindex_col.ColumnName] = targetpages.Pages[row_index].Index;
            }

            this.WriteObject(result_dt);
        }

        private VisioAutomation.ShapeSheet.Query.CellQuery _CreateQuery(
            VisioPowerShell.Models.NamedSrcDictionary celldic,
            IList<string> cellnames)
        {
            var invalid_names = cellnames.Where(cellname => !celldic.ContainsKey(cellname)).ToList();

            if (invalid_names.Count > 0)
            {
                var quoted_names = invalid_names.Select( s=> string.Format("\"{0}\"",s));
                string msg = "Invalid cell names: " + string.Join(",", quoted_names);
                throw new ArgumentException(nameof(cellnames),msg);
            }

            var query = new VisioAutomation.ShapeSheet.Query.CellQuery();

            foreach (string cellname in cellnames)
            {
                // resolve any wildcards to the actual cell names
                var resolved_cellnames = celldic.ExpandKeyWildcard(cellname);

                foreach (var resolved_cellname in resolved_cellnames)
                {
                    if (!query.Columns.Contains(resolved_cellname))
                    {
                        var resolved_src = celldic[resolved_cellname];
                        query.Columns.Add(resolved_src, resolved_cellname);
                    }
                }
            }

            return query;
        }
    }
}