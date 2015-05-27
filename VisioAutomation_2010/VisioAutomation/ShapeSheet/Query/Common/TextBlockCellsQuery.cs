using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;

namespace VisioAutomation.ShapeSheet.Query.Common
{
    class TextBlockCellsQuery : CellQuery
    {
        public Query.CellColumn BottomMargin { get; set; }
        public Query.CellColumn LeftMargin { get; set; }
        public Query.CellColumn RightMargin { get; set; }
        public Query.CellColumn TopMargin { get; set; }
        public Query.CellColumn DefaultTabStop { get; set; }
        public Query.CellColumn TextBkgnd { get; set; }
        public Query.CellColumn TextBkgndTrans { get; set; }
        public Query.CellColumn TextDirection { get; set; }
        public Query.CellColumn VerticalAlign { get; set; }
        public Query.CellColumn TxtWidth { get; set; }
        public Query.CellColumn TxtHeight { get; set; }
        public Query.CellColumn TxtPinX { get; set; }
        public Query.CellColumn TxtPinY { get; set; }
        public Query.CellColumn TxtLocPinX { get; set; }
        public Query.CellColumn TxtLocPinY { get; set; }
        public Query.CellColumn TxtAngle { get; set; }

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

        public VisioAutomation.Text.TextBlockCells GetCells(System.Collections.Generic.IList<ShapeSheet.CellData<double>> row)
        {
            var cells = new VisioAutomation.Text.TextBlockCells();
            cells.BottomMargin = row[this.BottomMargin];
            cells.LeftMargin = row[this.LeftMargin];
            cells.RightMargin = row[this.RightMargin];
            cells.TopMargin = row[this.TopMargin];
            cells.DefaultTabStop = row[this.DefaultTabStop];
            cells.TextBkgnd = Extensions.CellDataMethods.ToInt(row[this.TextBkgnd]);
            cells.TextBkgndTrans = row[this.TextBkgndTrans];
            cells.TextDirection = Extensions.CellDataMethods.ToInt(row[this.TextDirection]);
            cells.VerticalAlign = Extensions.CellDataMethods.ToInt(row[this.VerticalAlign]);
            cells.TxtPinX = row[this.TxtPinX];
            cells.TxtPinY = row[this.TxtPinY];
            cells.TxtLocPinX = row[this.TxtLocPinX];
            cells.TxtLocPinY = row[this.TxtLocPinY];
            cells.TxtWidth = row[this.TxtWidth];
            cells.TxtHeight = row[this.TxtHeight];
            cells.TxtAngle = row[this.TxtAngle];
            return cells;
        }
    }
}