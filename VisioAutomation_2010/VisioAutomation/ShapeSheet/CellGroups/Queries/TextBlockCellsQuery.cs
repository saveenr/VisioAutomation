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

        public override Text.TextBlockCells CellDataToCellGroup(ShapeSheet.CellData[] row)
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
}