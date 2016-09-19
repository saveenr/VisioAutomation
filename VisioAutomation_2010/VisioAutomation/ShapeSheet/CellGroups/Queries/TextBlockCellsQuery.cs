using VisioAutomation.ShapeSheet.Queries.Columns;
using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;
using VisioAutomation.Extensions;

namespace VisioAutomation.ShapeSheet.CellGroups.Queries
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

        }

        public override Text.TextBlockCells CellDataToCellGroup(ShapeSheet.CellData<double>[] row)
        {
            var cells = new Text.TextBlockCells();
            cells.BottomMargin = row[this.BottomMargin];
            cells.LeftMargin = row[this.LeftMargin];
            cells.RightMargin = row[this.RightMargin];
            cells.TopMargin = row[this.TopMargin];
            cells.DefaultTabStop = row[this.DefaultTabStop];
            cells.TextBkgnd = row[this.TextBkgnd].ToInt();
            cells.TextBkgndTrans = row[this.TextBkgndTrans];
            cells.TextDirection = row[this.TextDirection].ToInt();
            cells.VerticalAlign = row[this.VerticalAlign].ToInt();
            return cells;
        }
    }

    class TextXFormCellsQuery : CellGroupSingleRowQuery<Text.TextXFormCells, double>
    {
        public ColumnQuery TxtWidth { get; set; }
        public ColumnQuery TxtHeight { get; set; }
        public ColumnQuery TxtPinX { get; set; }
        public ColumnQuery TxtPinY { get; set; }
        public ColumnQuery TxtLocPinX { get; set; }
        public ColumnQuery TxtLocPinY { get; set; }
        public ColumnQuery TxtAngle { get; set; }

        public TextXFormCellsQuery()
        {
            this.TxtPinX = this.query.AddCell(SRCCON.TxtPinX, nameof(SRCCON.TxtPinX));
            this.TxtPinY = this.query.AddCell(SRCCON.TxtPinY, nameof(SRCCON.TxtPinY));
            this.TxtLocPinX = this.query.AddCell(SRCCON.TxtLocPinX, nameof(SRCCON.TxtLocPinX));
            this.TxtLocPinY = this.query.AddCell(SRCCON.TxtLocPinY, nameof(SRCCON.TxtLocPinY));
            this.TxtWidth = this.query.AddCell(SRCCON.TxtWidth, nameof(SRCCON.TxtWidth));
            this.TxtHeight = this.query.AddCell(SRCCON.TxtHeight, nameof(SRCCON.TxtHeight));
            this.TxtAngle = this.query.AddCell(SRCCON.TxtAngle, nameof(SRCCON.TxtAngle));

        }

        public override Text.TextXFormCells CellDataToCellGroup(ShapeSheet.CellData<double>[] row)
        {
            var cells = new Text.TextXFormCells();
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