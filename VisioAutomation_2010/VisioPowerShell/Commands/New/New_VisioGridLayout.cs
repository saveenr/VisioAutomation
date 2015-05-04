using VisioAutomation.Drawing;
using IVisio=Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using SMA = System.Management.Automation;
using GRID = VisioAutomation.Models.Grid;

namespace VisioPowerShell.Commands
{
    [SMA.CmdletAttribute(SMA.VerbsCommon.New, "VisioGridLayout")]
    public class New_VisioGridLayout : VisioCmdlet
    {
        [SMA.ParameterAttribute(Position = 0, Mandatory = true)]
        public IVisio.Master Master { get; set; }

        [SMA.ParameterAttribute(Mandatory = true)]
        public int Columns { get; set; }

        [SMA.ParameterAttribute(Mandatory = true)]
        public int Rows { get; set; }

        [SMA.ParameterAttribute(Mandatory = true)]
        public double CellWidth = 0.5;
        
        [SMA.ParameterAttribute(Mandatory = true)]
        public double CellHeight = 0.5;

        [SMA.ParameterAttribute(Mandatory = false)]
        public double CellHorizontalSpacing = 0.25;

        [SMA.ParameterAttribute(Mandatory = false)]
        public double CellVerticalSpacing = 0.25;

        [SMA.ParameterAttribute(Mandatory = false)]
        public GRID.RowDirection RowDirection = GRID.RowDirection.BottomToTop;

        [SMA.ParameterAttribute(Mandatory = false)]
        public GRID.ColumnDirection ColumnDirection = GRID.ColumnDirection.LeftToRight;

        protected override void ProcessRecord()
        {
            var cellsize = new Size(this.CellWidth, this.CellHeight);
            var layout = new GRID.GridLayout(this.Columns, this.Rows, cellsize, this.Master);
            layout.CellSpacing = new Size(this.CellHorizontalSpacing, this.CellVerticalSpacing);
            layout.RowDirection = this.RowDirection;
            layout.ColumnDirection = this.ColumnDirection;
            this.WriteObject(layout);
        }
    }
}