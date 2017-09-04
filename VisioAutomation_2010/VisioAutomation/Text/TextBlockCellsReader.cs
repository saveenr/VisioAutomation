using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Text
{
    class TextBlockCellsReader : ReaderSingleRow<Text.TextBlockCells>
    {
        public CellColumn BottomMargin { get; set; }
        public CellColumn LeftMargin { get; set; }
        public CellColumn RightMargin { get; set; }
        public CellColumn TopMargin { get; set; }
        public CellColumn DefaultTabStop { get; set; }
        public CellColumn Background { get; set; }
        public CellColumn BackgroundTransparency { get; set; }
        public CellColumn Direction { get; set; }
        public CellColumn VerticalAlign { get; set; }

        public TextBlockCellsReader()
        {
            this.BottomMargin = this.query.AddColumn(SrcConstants.TextBlockBottomMargin, nameof(SrcConstants.TextBlockBottomMargin));
            this.LeftMargin = this.query.AddColumn(SrcConstants.TextBlockLeftMargin, nameof(SrcConstants.TextBlockLeftMargin));
            this.RightMargin = this.query.AddColumn(SrcConstants.TextBlockRightMargin, nameof(SrcConstants.TextBlockRightMargin));
            this.TopMargin = this.query.AddColumn(SrcConstants.TextBlockTopMargin, nameof(SrcConstants.TextBlockTopMargin));
            this.DefaultTabStop = this.query.AddColumn(SrcConstants.TextBlockDefaultTabStop, nameof(SrcConstants.TextBlockDefaultTabStop));
            this.Background = this.query.AddColumn(SrcConstants.TextBlockBackground, nameof(SrcConstants.TextBlockBackground));
            this.BackgroundTransparency = this.query.AddColumn(SrcConstants.TextBlockBackgroundTransparency, nameof(SrcConstants.TextBlockBackgroundTransparency));
            this.Direction = this.query.AddColumn(SrcConstants.TextBlockDirection, nameof(SrcConstants.TextBlockDirection));
            this.VerticalAlign = this.query.AddColumn(SrcConstants.TextBlockVerticalAlign, nameof(SrcConstants.TextBlockVerticalAlign));

        }

        public override Text.TextBlockCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
        {
            var cells = new Text.TextBlockCells();
            cells.BottomMargin = row[this.BottomMargin];
            cells.LeftMargin = row[this.LeftMargin];
            cells.RightMargin = row[this.RightMargin];
            cells.TopMargin = row[this.TopMargin];
            cells.DefaultTabStop = row[this.DefaultTabStop];
            cells.TextBackground = row[this.Background];
            cells.TextBackgroundTransparency = row[this.BackgroundTransparency];
            cells.TextDirection = row[this.Direction];
            cells.VerticalAlign = row[this.VerticalAlign];
            return cells;
        }
    }
}