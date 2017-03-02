using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Text
{
    class TextBlockCellsReader : SingleRowReader<Text.TextBlockCells>
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

        public TextBlockCellsReader()
        {
            this.BottomMargin = this.query.AddCell(SrcConstants.BottomMargin, nameof(SrcConstants.BottomMargin));
            this.LeftMargin = this.query.AddCell(SrcConstants.LeftMargin, nameof(SrcConstants.LeftMargin));
            this.RightMargin = this.query.AddCell(SrcConstants.RightMargin, nameof(SrcConstants.RightMargin));
            this.TopMargin = this.query.AddCell(SrcConstants.TopMargin, nameof(SrcConstants.TopMargin));
            this.DefaultTabStop = this.query.AddCell(SrcConstants.DefaultTabStop, nameof(SrcConstants.DefaultTabStop));
            this.TextBkgnd = this.query.AddCell(SrcConstants.TextBkgnd, nameof(SrcConstants.TextBkgnd));
            this.TextBkgndTrans = this.query.AddCell(SrcConstants.TextBkgndTrans, nameof(SrcConstants.TextBkgndTrans));
            this.TextDirection = this.query.AddCell(SrcConstants.TextDirection, nameof(SrcConstants.TextDirection));
            this.VerticalAlign = this.query.AddCell(SrcConstants.VerticalAlign, nameof(SrcConstants.VerticalAlign));

        }

        public override Text.TextBlockCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
        {
            var cells = new Text.TextBlockCells();
            cells.BottomMargin = row[this.BottomMargin];
            cells.LeftMargin = row[this.LeftMargin];
            cells.RightMargin = row[this.RightMargin];
            cells.TopMargin = row[this.TopMargin];
            cells.DefaultTabStop = row[this.DefaultTabStop];
            cells.TextBkgnd = row[this.TextBkgnd];
            cells.TextBkgndTrans = row[this.TextBkgndTrans];
            cells.TextDirection = row[this.TextDirection];
            cells.VerticalAlign = row[this.VerticalAlign];
            return cells;
        }
    }
}