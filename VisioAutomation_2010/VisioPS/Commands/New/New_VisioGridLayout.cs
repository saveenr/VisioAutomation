using IVisio=Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using SMA = System.Management.Automation;
using GRID = VisioAutomation.Layout.Models.Grid;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "VisioGridLayout")]
    public class New_VisioGridLayout : VisioPS.VisioPSCmdlet
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

        [SMA.Parameter(Position = 5, Mandatory = false)]
        public GRID.RowDirection RowDirection = GRID.RowDirection.BottomToTop;

        [SMA.Parameter(Position = 5, Mandatory = false)]
        public GRID.ColumnDirection ColumnDirection = GRID.ColumnDirection.LeftToRight;

        protected override void ProcessRecord()
        {
            var cellsize = new VA.Drawing.Size(CellWidth, CellHeight);
            var layout = new GRID.GridLayout(this.Columns, this.Rows, cellsize, this.Master);
            layout.RowDirection = this.RowDirection;
            layout.ColumnDirection = this.ColumnDirection;
            this.WriteObject(layout);
        }
    }
}