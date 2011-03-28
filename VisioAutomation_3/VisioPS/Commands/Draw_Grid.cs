using IVisio=Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("Draw", "Grid")]
    public class Draw_Grid : VisioPS.VisioPSCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public IVisio.Master Master { get; set; }

        [SMA.Parameter(Position=1, Mandatory = true)]
        public int Columns { get; set; }

        [SMA.Parameter(Position=2, Mandatory = true)]
        public int Rows { get; set; }

        [SMA.Parameter(Position = 3, Mandatory = true)]
        public double CellWidth = 0.5;
        
        [SMA.Parameter(Position = 4, Mandatory = true)]
        public double CellHeight = 0.5;

        [SMA.Parameter(Position = 3, Mandatory = true)]
        public double X = 0.0;

        [SMA.Parameter(Position = 4, Mandatory = true)]
        public double Y = 0.0;

        protected override void ProcessRecord()
        {
            var scriptingsession = this.ScriptingSession;
            var cellsize = new VA.Drawing.Size(CellWidth, CellHeight);
            var grid_origin = new VA.Drawing.Point(this.X, this.Y);
            var shapes = scriptingsession.Draw.DrawGrid(Master, cellsize, Columns, Rows, grid_origin);
            this.WriteObject(shapes,false);
        }
    }
}