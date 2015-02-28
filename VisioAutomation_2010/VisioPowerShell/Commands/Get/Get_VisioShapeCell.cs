using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;
using System.Linq;
using VA = VisioAutomation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, "VisioShapeCell")]
    public class Get_VisioShapeCell : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = true, Position = 0)]
        public string[] Cells { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes { get; set; }

        [SMA.Parameter(Mandatory = false)] 
        public SMA.SwitchParameter GetResults;

        [SMA.Parameter(Mandatory = false)] 
        public ResultType ResultType = ResultType.String;

        protected override void ProcessRecord()
        {
            Get_VisioPageCell.EnsureEnoughCellNames(this.Cells);
            var target_shapes = this.Shapes ?? this.client.Selection.GetShapes();
            var cellmap = CellMap.GetShapeCellDictionary();
            this.WriteVerbose("Valid Names: " + string.Join(",", cellmap.GetNames()));
            Get_VisioPageCell.CheckForInvalidNames(cellmap, this.Cells);
            var query = Get_VisioPageCell.CreateQueryFromCellNames(this.Cells, cellmap);
            var surface = this.client.ShapeSheet.GetShapeSheetSurface();
            var target_shapeids = target_shapes.Select(s => s.ID).ToList();
            var dt = Helpers.QueryToDataTable(query, this.GetResults, this.ResultType, target_shapeids, surface);
            this.WriteObject(dt);
        }
    }
}