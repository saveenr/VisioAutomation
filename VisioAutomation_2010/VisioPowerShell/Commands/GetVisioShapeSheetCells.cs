using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.Get, VisioPowerShell.Commands.Nouns.VisioShapeSheetCells)]
    public class GetVisioShapeSheetCells : VisioCmdlet
    {
        [Parameter(Mandatory = true, Position = 0)]
        public VisioPowerShell.Models.CellType Type { get; set; }

        [Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes { get; set; }

        [Parameter(Mandatory = false)] 
        public SwitchParameter GetResults;

        [Parameter(Mandatory = false)] 
        public VisioPowerShell.Models.ResultType ResultType = VisioPowerShell.Models.ResultType.String;

        protected override void ProcessRecord()
        {
            var target_shapes = this.Shapes ?? this.Client.Selection.GetShapes();
            var celldic = VisioPowerShell.Models.BaseCells.GetDictionary(this.Type);
            var cells = celldic.Keys.ToArray();
            var query = _CreateQuery(celldic, cells);
            var surface = this.Client.ShapeSheet.GetShapeSheetSurface();
            var target_shapeids = target_shapes.Select(s => s.ID).ToList();
            var dt = VisioPowerShell.Models.DataTableHelpers.QueryToDataTable(query, this.GetResults, this.ResultType, target_shapeids, surface);
            this.WriteObject(dt);
        }

        private VisioAutomation.ShapeSheet.Query.CellQuery _CreateQuery(
            VisioPowerShell.Models.NamedCellDictionary celldic, 
            IList<string> cells)
        {
            var invalid_names = cells.Where(cellname => !celldic.ContainsKey(cellname)).ToList();

            if (invalid_names.Count > 0)
            {
                string msg = "Invalid cell names: " + string.Join(",", invalid_names);
                throw new ArgumentException(msg);
            }

            var query = new VisioAutomation.ShapeSheet.Query.CellQuery();

            foreach (string cell in cells)
            {
                foreach (var resolved_cellname in celldic.ExpandKeyWildcard(cell))
                {
                    if (!query.Columns.Contains(resolved_cellname))
                    {
                        var resolved_src = celldic[resolved_cellname];
                        query.AddColumn(resolved_src, resolved_cellname);
                    }
                }
            }

            return query;
        }
    }
}