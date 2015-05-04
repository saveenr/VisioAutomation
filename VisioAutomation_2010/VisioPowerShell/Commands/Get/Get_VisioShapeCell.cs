using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;
using SMA = System.Management.Automation;
using VA = VisioAutomation;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.Get, "VisioShapeCell")]
    public class Get_VisioShapeCell : VisioCmdlet
    {
        [SMA.ParameterAttribute(Mandatory = true, Position = 0)]
        public string[] Cells { get; set; }

        [SMA.ParameterAttribute(Mandatory = false)]
        public IVisio.Shape[] Shapes { get; set; }

        [SMA.ParameterAttribute(Mandatory = false)] 
        public SMA.SwitchParameter GetResults;

        [SMA.ParameterAttribute(Mandatory = false)] 
        public ResultType ResultType = ResultType.String;

        protected override void ProcessRecord()
        {
            Get_VisioPageCell.EnsureEnoughCellNames(this.Cells);
            var target_shapes = this.Shapes ?? this.client.Selection.GetShapes();
            var cellmap = CellSRCDictionary.GetCellMapForShapes();
            this.WriteVerbose("Valid Names: " + string.Join(",", cellmap.GetNames()));
            var query = cellmap.CreateQueryFromCellNames(this.Cells);
            var surface = this.client.ShapeSheet.GetShapeSheetSurface();
            var target_shapeids = target_shapes.Select(s => s.ID).ToList();
            var dt = Helpers.QueryToDataTable(query, this.GetResults, this.ResultType, target_shapeids, surface);
            this.WriteObject(dt);
        }
    }
}