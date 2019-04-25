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
        public string[] Cell { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public VisioPowerShell.Models.CellOutputType OutputType = VisioPowerShell.Models.CellOutputType.Formula;

        // CONTEXT:SHAPES 
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shape { get; set; }

        protected override void ProcessRecord()
        {
            var target_shapes = new VisioScripting.TargetShapes(this.Shape).Resolve(this.Client);

            if (target_shapes.Shapes.Count < 1)
            {
                return;
            }

            var template = new VisioPowerShell.Models.ShapeCells();

            var dicof_name_to_cell = VisioPowerShell.Models.NamedSrcDictionary.FromCells(template);

            var desired_cells = this.Cell ?? dicof_name_to_cell.Keys.ToArray();

            var query = _create_query(dicof_name_to_cell, desired_cells);
            var surface = this.Client.ShapeSheet.GetShapeSheetSurface();
            var shapeids = target_shapes.Shapes.Select(s => s.ID).ToList();
            var datatable = VisioPowerShell.Models.DataTableHelpers.QueryToDataTable(query, this.OutputType, shapeids, surface);

            // Annotate the returned datatable to disambiguate rows
            var shapeid_col = datatable.Columns.Add("ShapeID", typeof(int));
            shapeid_col.SetOrdinal(0);

            foreach (int row_index in Enumerable.Range(0,shapeids.Count))
            {
                datatable.Rows[row_index][shapeid_col.ColumnName] = shapeids[row_index];
            }

            this.WriteObject(datatable);
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