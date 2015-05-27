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
            this.BottomMargin = this.AddCell(ShapeSheet.SRCConstants.BottomMargin, nameof(ShapeSheet.SRCConstants.BottomMargin));
            this.LeftMargin = this.AddCell(ShapeSheet.SRCConstants.LeftMargin, nameof(ShapeSheet.SRCConstants.LeftMargin));
            this.RightMargin = this.AddCell(ShapeSheet.SRCConstants.RightMargin, nameof(ShapeSheet.SRCConstants.RightMargin));
            this.TopMargin = this.AddCell(ShapeSheet.SRCConstants.TopMargin, nameof(ShapeSheet.SRCConstants.TopMargin));
            this.DefaultTabStop = this.AddCell(ShapeSheet.SRCConstants.DefaultTabStop, nameof(ShapeSheet.SRCConstants.DefaultTabStop));
            this.TextBkgnd = this.AddCell(ShapeSheet.SRCConstants.TextBkgnd, nameof(ShapeSheet.SRCConstants.TextBkgnd));
            this.TextBkgndTrans = this.AddCell(ShapeSheet.SRCConstants.TextBkgndTrans, nameof(ShapeSheet.SRCConstants.TextBkgndTrans));
            this.TextDirection = this.AddCell(ShapeSheet.SRCConstants.TextDirection, nameof(ShapeSheet.SRCConstants.TextDirection));
            this.VerticalAlign = this.AddCell(ShapeSheet.SRCConstants.VerticalAlign, nameof(ShapeSheet.SRCConstants.VerticalAlign));
            this.TxtPinX = this.AddCell(ShapeSheet.SRCConstants.TxtPinX, nameof(ShapeSheet.SRCConstants.TxtPinX));
            this.TxtPinY = this.AddCell(ShapeSheet.SRCConstants.TxtPinY, nameof(ShapeSheet.SRCConstants.TxtPinY));
            this.TxtLocPinX = this.AddCell(ShapeSheet.SRCConstants.TxtLocPinX, nameof(ShapeSheet.SRCConstants.TxtLocPinX));
            this.TxtLocPinY = this.AddCell(ShapeSheet.SRCConstants.TxtLocPinY, nameof(ShapeSheet.SRCConstants.TxtLocPinY));
            this.TxtWidth = this.AddCell(ShapeSheet.SRCConstants.TxtWidth, nameof(ShapeSheet.SRCConstants.TxtWidth));
            this.TxtHeight = this.AddCell(ShapeSheet.SRCConstants.TxtHeight, nameof(ShapeSheet.SRCConstants.TxtHeight));
            this.TxtAngle = this.AddCell(ShapeSheet.SRCConstants.TxtAngle, nameof(ShapeSheet.SRCConstants.TxtAngle));

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