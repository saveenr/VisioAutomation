using System.Linq;
using System.Management.Automation;
using VisioPowerShell.Models;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [Cmdlet(VerbsCommon.Get, VisioPowerShell.Commands.Nouns.VisioShapeSheetCells)]
    public class GetVisioShapeSheetCells : VisioCmdlet
    {
        [Parameter(Mandatory = true, Position = 0)]
        public CellsType Type { get; set; }

        [Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes { get; set; }

        [Parameter(Mandatory = false)] 
        public SwitchParameter GetResults;

        [Parameter(Mandatory = false)] 
        public ResultType ResultType = ResultType.String;

        protected override void ProcessRecord()
        {

            var target_shapes = this.Shapes ?? this.Client.Selection.GetShapes();

            NamedSrcDictionary cellmap;

            if (this.Type == CellsType.Page)
            {
                cellmap = VisioPowerShell.Models.PageCells.GetCellDictionary();
            }
            else if (this.Type == CellsType.ShapeFormat)
            {
                cellmap = VisioPowerShell.Models.ShapeCells.GetCellDictionary();
            }
            else if (this.Type == CellsType.TextFormat)
            {
                cellmap = VisioPowerShell.Models.TextCells.GetCellDictionary();
            }
            else
            {
                throw new System.ArgumentException();
            }

            var cells = cellmap.ExpandCellNames(null);

            var query = cellmap.ToQuery(cells);
            var surface = this.Client.ShapeSheet.GetShapeSheetSurface();
            var target_shapeids = target_shapes.Select(s => s.ID).ToList();
            var dt = DataTableHelpers.QueryToDataTable(query, this.GetResults, this.ResultType, target_shapeids, surface);
            this.WriteObject(dt);
        }
    }
}