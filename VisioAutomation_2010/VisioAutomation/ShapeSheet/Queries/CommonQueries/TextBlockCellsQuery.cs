using VisioAutomation.ShapeSheet.Queries.Columns;
using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;

namespace VisioAutomation.ShapeSheet.Queries.CommonQueries
{
    class TextBlockCellsQuery : CellGroupSingleRowQuery<Text.TextBlockCells, double>
    {
        public ColumnQuery BottomMargin { get; set; }
        public ColumnQuery LeftMargin { get; set; }
        public ColumnQuery RightMargin { get; set; }
        public ColumnQuery TopMargin { get; set; }
        public ColumnQuery DefaultTabStop { get; set; }
        public ColumnQuery TextBkgnd { get; set; }
        public ColumnQuery TextBkgndTrans { get; set; }
        public ColumnQuery TextDirection { get; set; }
        public ColumnQuery VerticalAlign { get; set; }
        public ColumnQuery TxtWidth { get; set; }
        public ColumnQuery TxtHeight { get; set; }
        public ColumnQuery TxtPinX { get; set; }
        public ColumnQuery TxtPinY { get; set; }
        public ColumnQuery TxtLocPinX { get; set; }
        public ColumnQuery TxtLocPinY { get; set; }
        public ColumnQuery TxtAngle { get; set; }

        public TextBlockCellsQuery()
        {
            this.BottomMargin = this.query.AddCell(SRCCON.BottomMargin, nameof(SRCCON.BottomMargin));
            this.LeftMargin = this.query.AddCell(SRCCON.LeftMargin, nameof(SRCCON.LeftMargin));
            this.RightMargin = this.query.AddCell(SRCCON.RightMargin, nameof(SRCCON.RightMargin));
            this.TopMargin = this.query.AddCell(SRCCON.TopMargin, nameof(SRCCON.TopMargin));
            this.DefaultTabStop = this.query.AddCell(SRCCON.DefaultTabStop, nameof(SRCCON.DefaultTabStop));
            this.TextBkgnd = this.query.AddCell(SRCCON.TextBkgnd, nameof(SRCCON.TextBkgnd));
            this.TextBkgndTrans = this.query.AddCell(SRCCON.TextBkgndTrans, nameof(SRCCON.TextBkgndTrans));
            this.TextDirection = this.query.AddCell(SRCCON.TextDirection, nameof(SRCCON.TextDirection));
            this.VerticalAlign = this.query.AddCell(SRCCON.VerticalAlign, nameof(SRCCON.VerticalAlign));
            this.TxtPinX = this.query.AddCell(SRCCON.TxtPinX, nameof(SRCCON.TxtPinX));
            this.TxtPinY = this.query.AddCell(SRCCON.TxtPinY, nameof(SRCCON.TxtPinY));
            this.TxtLocPinX = this.query.AddCell(SRCCON.TxtLocPinX, nameof(SRCCON.TxtLocPinX));
            this.TxtLocPinY = this.query.AddCell(SRCCON.TxtLocPinY, nameof(SRCCON.TxtLocPinY));
            this.TxtWidth = this.query.AddCell(SRCCON.TxtWidth, nameof(SRCCON.TxtWidth));
            this.TxtHeight = this.query.AddCell(SRCCON.TxtHeight, nameof(SRCCON.TxtHeight));
            this.TxtAngle = this.query.AddCell(SRCCON.TxtAngle, nameof(SRCCON.TxtAngle));

        }

        public override Text.TextBlockCells CellDataToCellGroup(ShapeSheet.CellData<double>[] row)
        {
            var cells = new Text.TextBlockCells();
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