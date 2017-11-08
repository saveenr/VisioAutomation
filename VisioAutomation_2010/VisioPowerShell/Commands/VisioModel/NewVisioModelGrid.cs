using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, Nouns.VisioModelGrid)]
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
        public VisioAutomation.Models.Layouts.Grid.RowDirection RowDirection = VisioAutomation.Models.Layouts.Grid.RowDirection.BottomToTop;

        [SMA.Parameter(Mandatory = false)]
        public VisioAutomation.Models.Layouts.Grid.ColumnDirection ColumnDirection = VisioAutomation.Models.Layouts.Grid.ColumnDirection.LeftToRight;

        protected override void ProcessRecord()
        {
            var cellsize = new VisioAutomation.Geometry.Size(this.CellWidth, this.CellHeight);
            var layout = new VisioAutomation.Models.Layouts.Grid.GridLayout(this.Columns, this.Rows, cellsize, this.Master);
            layout.CellSpacing = new VisioAutomation.Geometry.Size(this.CellHorizontalSpacing, this.CellVerticalSpacing);
            layout.RowDirection = this.RowDirection;
            layout.ColumnDirection = this.ColumnDirection;
            this.WriteObject(layout);
        }
    }
}