using System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using GRID = VisioAutomation.Models.Grid;

namespace VisioPowerShell.Commands.New
{
    [Cmdlet(VerbsCommon.New, "VisioGridLayout")]
    public class New_VisioGridLayout : VisioCmdlet
    {
        [Parameter(Position = 0, Mandatory = true)]
        public IVisio.Master Master { get; set; }

        [Parameter(Mandatory = true)]
        public int Columns { get; set; }

        [Parameter(Mandatory = true)]
        public int Rows { get; set; }

        [Parameter(Mandatory = true)]
        public double CellWidth = 0.5;
        
        [Parameter(Mandatory = true)]
        public double CellHeight = 0.5;

        [Parameter(Mandatory = false)]
        public double CellHorizontalSpacing = 0.25;

        [Parameter(Mandatory = false)]
        public double CellVerticalSpacing = 0.25;

        [Parameter(Mandatory = false)]
        public GRID.RowDirection RowDirection = GRID.RowDirection.BottomToTop;

        [Parameter(Mandatory = false)]
        public GRID.ColumnDirection ColumnDirection = GRID.ColumnDirection.LeftToRight;

        protected override void ProcessRecord()
        {
            var cellsize = new VA.Drawing.Size(this.CellWidth, this.CellHeight);
            var layout = new GRID.GridLayout(this.Columns, this.Rows, cellsize, this.Master);
            layout.CellSpacing = new VA.Drawing.Size(this.CellHorizontalSpacing, this.CellVerticalSpacing);
            layout.RowDirection = this.RowDirection;
            layout.ColumnDirection = this.ColumnDirection;
            this.WriteObject(layout);
        }
    }
}