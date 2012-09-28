using IVisio=Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using SMA = System.Management.Automation;

namespace VisioPS.Commands
{
    [SMA.Cmdlet("New", "GridLayout")]
    public class New_GridLayout : VisioPS.VisioPSCmdlet
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
        public VA.Layout.Models.Grid.RowDirection RowDirection = VA.Layout.Models.Grid.RowDirection.BottomToTop;

        [SMA.Parameter(Position = 5, Mandatory = false)]
        public VA.Layout.Models.Grid.ColumnDirection ColumnDirection = VA.Layout.Models.Grid.ColumnDirection.LeftToRight;

        protected override void ProcessRecord()
        {
            var cellsize = new VA.Drawing.Size(CellWidth, CellHeight);
            var layout = new VA.Layout.Models.Grid.GridLayout(this.Columns, this.Rows, cellsize, this.Master);
            layout.RowDirection = this.RowDirection;
            layout.ColumnDirection = this.ColumnDirection;
            this.WriteObject(layout);
        }
    }
}