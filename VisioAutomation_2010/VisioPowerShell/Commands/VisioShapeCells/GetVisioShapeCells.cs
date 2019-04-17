using System;
using System.Collections.Generic;
using System.Linq;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioShapeCells
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioShapeCells)]
    public class GetVisioShapeCells : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes { get; set; }

        [SMA.Parameter(Mandatory = false)] 
        public VisioPowerShell.Models.CellOutputType OutputType = VisioPowerShell.Models.CellOutputType.Formula;

        protected override void ProcessRecord()
        {
            var target_shapes = this.Shapes ?? this.Client.Selection.GetSelectedShapes(VisioScripting.TargetWindow.Active);

            var template = new VisioPowerShell.Models.ShapeCells();
            var dicof_name_to_cell = VisioPowerShell.Models.NamedSrcDictionary.FromCells(template);
            var arrayof_cellnames = dicof_name_to_cell.Keys.ToArray();
            var query = _create_query(dicof_name_to_cell, arrayof_cellnames);
            var surface = this.Client.ShapeSheet.GetShapeSheetSurface();
            var target_shapeids = target_shapes.Select(s => s.ID).ToList();
            var dt = VisioPowerShell.Models.DataTableHelpers.QueryToDataTable(query, this.OutputType, target_shapeids, surface);

            // Annotate the returned datatable to disambiguate rows
            var shapeid_col = dt.Columns.Add("ShapeID", typeof(System.Int32));
            shapeid_col.SetOrdinal(0);

            foreach (int row_index in Enumerable.Range(0,target_shapeids.Count))
            {
                dt.Rows[row_index][shapeid_col.ColumnName] = target_shapeids[row_index];
            }

            this.WriteObject(dt);
        }

        private VisioAutomation.ShapeSheet.Query.CellQuery _create_query(
            VisioPowerShell.Models.NamedSrcDictionary dicof_named_to_cell, 
            IList<string> cellnames)
        {
            var invalid_names = cellnames.Where(cellname => !dicof_named_to_cell.ContainsKey(cellname)).ToList();

            if (invalid_names.Count > 0)
            {
                string msg = "Invalid cell names: " + string.Join(",", invalid_names);
                throw new ArgumentException(msg);
            }

            var query = new VisioAutomation.ShapeSheet.Query.CellQuery();

            foreach (string cellname in cellnames)
            {
                foreach (var resolved_cellname in dicof_named_to_cell.ExpandKeyWildcard(cellname))
                {
                    if (!query.Columns.Contains(resolved_cellname))
                    {
                        var resolved_src = dicof_named_to_cell[resolved_cellname];
                        query.Columns.Add(resolved_src, resolved_cellname);
                    }
                }
            }

            return query;
        }
    }
}