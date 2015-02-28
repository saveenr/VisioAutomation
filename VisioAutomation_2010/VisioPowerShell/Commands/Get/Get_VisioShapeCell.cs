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

            if (this.Cells == null)
            {
                throw new System.ArgumentException("Cells");
            }

            if (this.Cells.Length < 1)
            {
                string msg = "Must provide at least one cell name";
                throw new System.ArgumentException(msg);
            }

            var cellmap = CellMap.GetShapeCellDictionary();
            var invalid_names = this.Cells.Where(cellname => !cellmap.ContainsCell(cellname)).ToList();
            if (invalid_names.Count > 0)
            {
                var names = cellmap.GetNames();
                string valid_names = string.Join(",", names);
                this.WriteVerbose( "Valid Names: " + valid_names);
                string msg = "Invalid cell names: " + string.Join(",",invalid_names);
                throw new System.ArgumentException(msg);
            }

            var query = new VisioAutomation.ShapeSheet.Query.CellQuery();

            var target_shapes = this.Shapes ?? this.client.Selection.GetShapes();
            var target_shapeids = target_shapes.Select(s => s.ID).ToList();

            Get_VisioPageCell.SetFromCellNames(query, this.Cells, cellmap);

            var surface = this.client.ShapeSheet.GetShapeSheetSurface();

            this.WriteVerbose("Number of Shapes : {0}", target_shapes.Count);
            this.WriteVerbose("Number of Cells: {0}", query.CellColumns.Count);

            this.WriteVerbose("Start Query");

            var dt = Helpers.QueryToDataTable(query, this.GetResults, this.ResultType, target_shapeids, surface);
            this.WriteObject(dt);

            this.WriteVerbose("End Query");
        }
    }
}