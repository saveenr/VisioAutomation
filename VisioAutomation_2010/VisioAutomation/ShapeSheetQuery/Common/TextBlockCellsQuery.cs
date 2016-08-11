using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;

namespace VisioAutomation.ShapeSheetQuery.Common
{
    class TextBlockCellsQuery : CellQuery
    {
        public CellColumn BottomMargin { get; set; }
        public CellColumn LeftMargin { get; set; }
        public CellColumn RightMargin { get; set; }
        public CellColumn TopMargin { get; set; }
        public CellColumn DefaultTabStop { get; set; }
        public CellColumn TextBkgnd { get; set; }
        public CellColumn TextBkgndTrans { get; set; }
        public CellColumn TextDirection { get; set; }
        public CellColumn VerticalAlign { get; set; }
        public CellColumn TxtWidth { get; set; }
        public CellColumn TxtHeight { get; set; }
        public CellColumn TxtPinX { get; set; }
        public CellColumn TxtPinY { get; set; }
        public CellColumn TxtLocPinX { get; set; }
        public CellColumn TxtLocPinY { get; set; }
        public CellColumn TxtAngle { get; set; }

        public TextBlockCellsQuery() :
            base()
        {
            this.BottomMargin = this.AddCell(SRCCON.BottomMargin, nameof(SRCCON.BottomMargin));
            this.LeftMargin = this.AddCell(SRCCON.LeftMargin, nameof(SRCCON.LeftMargin));
            this.RightMargin = this.AddCell(SRCCON.RightMargin, nameof(SRCCON.RightMargin));
            this.TopMargin = this.AddCell(SRCCON.TopMargin, nameof(SRCCON.TopMargin));
            this.DefaultTabStop = this.AddCell(SRCCON.DefaultTabStop, nameof(SRCCON.DefaultTabStop));
            this.TextBkgnd = this.AddCell(SRCCON.TextBkgnd, nameof(SRCCON.TextBkgnd));
            this.TextBkgndTrans = this.AddCell(SRCCON.TextBkgndTrans, nameof(SRCCON.TextBkgndTrans));
            this.TextDirection = this.AddCell(SRCCON.TextDirection, nameof(SRCCON.TextDirection));
            this.VerticalAlign = this.AddCell(SRCCON.VerticalAlign, nameof(SRCCON.VerticalAlign));
            this.TxtPinX = this.AddCell(SRCCON.TxtPinX, nameof(SRCCON.TxtPinX));
            this.TxtPinY = this.AddCell(SRCCON.TxtPinY, nameof(SRCCON.TxtPinY));
            this.TxtLocPinX = this.AddCell(SRCCON.TxtLocPinX, nameof(SRCCON.TxtLocPinX));
            this.TxtLocPinY = this.AddCell(SRCCON.TxtLocPinY, nameof(SRCCON.TxtLocPinY));
            this.TxtWidth = this.AddCell(SRCCON.TxtWidth, nameof(SRCCON.TxtWidth));
            this.TxtHeight = this.AddCell(SRCCON.TxtHeight, nameof(SRCCON.TxtHeight));
            this.TxtAngle = this.AddCell(SRCCON.TxtAngle, nameof(SRCCON.TxtAngle));

        }

        public Text.TextBlockCells GetCells(SectionResultRow<ShapeSheet.CellData<double>> row)
        {
            var cells = new Text.TextBlockCells();
            cells.BottomMargin = row.Cells[this.BottomMargin];
            cells.LeftMargin = row.Cells[this.LeftMargin];
            cells.RightMargin = row.Cells[this.RightMargin];
            cells.TopMargin = row.Cells[this.TopMargin];
            cells.DefaultTabStop = row.Cells[this.DefaultTabStop];
            cells.TextBkgnd = Extensions.CellDataMethods.ToInt(row.Cells[this.TextBkgnd]);
            cells.TextBkgndTrans = row.Cells[this.TextBkgndTrans];
            cells.TextDirection = Extensions.CellDataMethods.ToInt(row.Cells[this.TextDirection]);
            cells.VerticalAlign = Extensions.CellDataMethods.ToInt(row.Cells[this.VerticalAlign]);
            cells.TxtPinX = row.Cells[this.TxtPinX];
            cells.TxtPinY = row.Cells[this.TxtPinY];
            cells.TxtLocPinX = row.Cells[this.TxtLocPinX];
            cells.TxtLocPinY = row.Cells[this.TxtLocPinY];
            cells.TxtWidth = row.Cells[this.TxtWidth];
            cells.TxtHeight = row.Cells[this.TxtHeight];
            cells.TxtAngle = row.Cells[this.TxtAngle];
            return cells;
        }
    }
}