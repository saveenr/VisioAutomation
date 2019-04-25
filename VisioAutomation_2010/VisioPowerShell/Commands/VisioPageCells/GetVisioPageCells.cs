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
        public string[] Cell { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public VisioPowerShell.Models.CellOutputType OutputType = VisioPowerShell.Models.CellOutputType.Formula;

        // CONTEXT:PAGES
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Page[] Page { get; set; }

        protected override void ProcessRecord()
        {
            var targetpages = new VisioScripting.TargetPages(this.Page).Resolve(this.Client);

            if (targetpages.Pages.Count < 1)
            {
                return;
            }

            var template = new VisioPowerShell.Models.PageCells();
            var dicof_name_to_cell = VisioPowerShell.Models.NamedSrcDictionary.FromCells(template);
            var desired_cells = this.Cell ?? dicof_name_to_cell.Keys.ToArray();
            var query = _create_query(dicof_name_to_cell, desired_cells);
            var surface = this.Client.ShapeSheet.GetShapeSheetSurface();
            
            var datatable = new System.Data.DataTable();

            foreach (var targetpage in targetpages.Pages)
            {
                var shapesheet = targetpage.PageSheet;
                var shapeids = new List<int> { shapesheet.ID };
                var temp_datatable = VisioPowerShell.Models.DataTableHelpers.QueryToDataTable(query, this.OutputType, shapeids, surface);
                datatable.Merge(temp_datatable);
            }

            // Annotate the returned datatable to disambiguate rows
            var pageindex_col = datatable.Columns.Add("PageID", typeof(int));
            pageindex_col.SetOrdinal(0);
            foreach (int row_index in Enumerable.Range(0,targetpages.Pages.Count))
            {
                datatable.Rows[row_index][pageindex_col.ColumnName] = targetpages.Pages[row_index].ID;
            }

            this.WriteObject(datatable);
        }

        private VisioAutomation.ShapeSheet.Query.CellQuery _create_query(
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