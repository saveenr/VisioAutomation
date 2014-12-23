using IVisio=Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using SMA = System.Management.Automation;
using GRID = VisioAutomation.Models.Grid;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, "VisioGridLayout")]
    public class New_VisioGridLayout : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public IVisio.Master Master { get; set; }

        [SMA.Parameter(Mandatory = true)]
        public int Columns { get; set; }

        [SMA.Parameter(Mandatory = true)]
        public int Rows { get; set; }

        [SMA.Parameter(Mandatory = true)]
        public double CellWidth = 0.5;
        
        [SMA.Parameter(Mandatory = true)]
        public double CellHeight = 0.5;

        [SMA.Parameter(Mandatory = false)]
        public double CellHorizontalSpacing = 0.25;

        [SMA.Parameter(Mandatory = false)]
        public double CellVerticalSpacing = 0.25;

        [SMA.Parameter(Mandatory = false)]
        public GRID.RowDirection RowDirection = GRID.RowDirection.BottomToTop;

        [SMA.Parameter(Mandatory = false)]
        public GRID.ColumnDirection ColumnDirection = GRID.ColumnDirection.LeftToRight;

        protected override void ProcessRecord()
        {
            var cellsize = new VA.Drawing.Size(CellWidth, CellHeight);
            var layout = new GRID.GridLayout(this.Columns, this.Rows, cellsize, this.Master);
            layout.CellSpacing = new VA.Drawing.Size(this.CellHorizontalSpacing, this.CellVerticalSpacing);
            layout.RowDirection = this.RowDirection;
            layout.ColumnDirection = this.ColumnDirection;
            this.WriteObject(layout);
        }
    }
}