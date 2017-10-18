using System;
using System.Collections.Generic;
using System.Linq;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, VisioPowerShell.Commands.Nouns.VisioShapeCells)]
    public class GetVisioShapeCells : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes { get; set; }

        [SMA.Parameter(Mandatory = false)] 
        public VisioPowerShell.Models.CellOutputType OutputType = VisioPowerShell.Models.CellOutputType.Formula;

        protected override void ProcessRecord()
        {
            var target_shapes = this.Shapes ?? this.Client.Selection.GetShapesInSelection();

            var template = new VisioPowerShell.Models.ShapeCells();
            var celldic = VisioPowerShell.Models.NamedCellDictionary.FromCells(template);
            var cellnames = celldic.Keys.ToArray();
            var query = _CreateQuery(celldic, cellnames);
            var surface = this.Client.ShapeSheet.GetShapeSheetSurface();
            var target_shapeids = target_shapes.Select(s => s.ID).ToList();
            var dt = VisioPowerShell.Models.DataTableHelpers.QueryToDataTable(query, this.OutputType, target_shapeids, surface);

            // Annotate the returned datatable to disambiguate rows
            var shapeid_col = dt.Columns.Add("ShapeID", typeof(System.Int32));
            shapeid_col.SetOrdinal(0);
            for (int row_index = 0; row_index < target_shapeids.Count; row_index++)
            {
                dt.Rows[row_index][shapeid_col.ColumnName] = target_shapeids[row_index];
            }

            this.WriteObject(dt);
        }

        private VisioAutomation.ShapeSheet.Query.CellQuery _CreateQuery(
            VisioPowerShell.Models.NamedCellDictionary celldic, 
            IList<string> cellnames)
        {
            var invalid_names = cellnames.Where(cellname => !celldic.ContainsKey(cellname)).ToList();

            if (invalid_names.Count > 0)
            {
                string msg = "Invalid cell names: " + string.Join(",", invalid_names);
                throw new ArgumentException(msg);
            }

            var query = new VisioAutomation.ShapeSheet.Query.CellQuery();

            foreach (string cell in cellnames)
            {
                foreach (var resolved_cellname in celldic.ExpandKeyWildcard(cell))
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