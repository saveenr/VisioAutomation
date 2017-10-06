using SMA = System.Management.Automation;
using VisioAutomation.Models.Layouts.Grid;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, VisioPowerShell.Commands.Nouns.VisioModelGrid)]
    public class NewVisioModelGrid : VisioCmdlet
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
        public RowDirection RowDirection = RowDirection.BottomToTop;

        [SMA.Parameter(Mandatory = false)]
        public ColumnDirection ColumnDirection = ColumnDirection.LeftToRight;

        protected override void ProcessRecord()
        {
            var cellsize = new VA.Geometry.Size(this.CellWidth, this.CellHeight);
            var layout = new GridLayout(this.Columns, this.Rows, cellsize, this.Master);
            layout.CellSpacing = new VA.Geometry.Size(this.CellHorizontalSpacing, this.CellVerticalSpacing);
            layout.RowDirection = this.RowDirection;
            layout.ColumnDirection = this.ColumnDirection;
            this.WriteObject(layout);
        }
    }
}