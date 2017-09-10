using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Text
{
    public class TextBlockCells : CellGroupSingleRow
    {
        public CellValueLiteral BottomMargin { get; set; }
        public CellValueLiteral LeftMargin { get; set; }
        public CellValueLiteral RightMargin { get; set; }
        public CellValueLiteral TopMargin { get; set; }
        public CellValueLiteral DefaultTabStop { get; set; }
        public CellValueLiteral TextBackground { get; set; }
        public CellValueLiteral TextBackgroundTransparency { get; set; }
        public CellValueLiteral TextDirection { get; set; }
        public CellValueLiteral VerticalAlign { get; set; }

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(SrcConstants.TextBlockBottomMargin, this.BottomMargin);
                yield return SrcValuePair.Create(SrcConstants.TextBlockLeftMargin, this.LeftMargin);
                yield return SrcValuePair.Create(SrcConstants.TextBlockRightMargin, this.RightMargin);
                yield return SrcValuePair.Create(SrcConstants.TextBlockTopMargin, this.TopMargin);
                yield return SrcValuePair.Create(SrcConstants.TextBlockDefaultTabStop, this.DefaultTabStop);
                yield return SrcValuePair.Create(SrcConstants.TextBlockBackground, this.TextBackground);
                yield return SrcValuePair.Create(SrcConstants.TextBlockBackgroundTransparency, this.TextBackgroundTransparency);
                yield return SrcValuePair.Create(SrcConstants.TextBlockDirection, this.TextDirection);
                yield return SrcValuePair.Create(SrcConstants.TextBlockVerticalAlign, this.VerticalAlign);
            }
        }

        public static IList<TextBlockCells> GetValues(IVisio.Page page, IList<int> shapeids, CellValueType cvt)
        {
            var query = lazy_query.Value;
            return query.GetValues(page, shapeids, cvt);
        }

        public static TextBlockCells GetValues(IVisio.Shape shape, CellValueType cvt)
        {
            var query = lazy_query.Value;
            return query.GetValues(shape, cvt);
        }

        private static readonly System.Lazy<TextBlockCellsReader> lazy_query = new System.Lazy<TextBlockCellsReader>();

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
                this.BottomMargin = this.query.Columns.Add(SrcConstants.TextBlockBottomMargin, nameof(this.BottomMargin));
                this.LeftMargin = this.query.Columns.Add(SrcConstants.TextBlockLeftMargin, nameof(this.LeftMargin));
                this.RightMargin = this.query.Columns.Add(SrcConstants.TextBlockRightMargin, nameof(this.RightMargin));
                this.TopMargin = this.query.Columns.Add(SrcConstants.TextBlockTopMargin, nameof(this.TopMargin));
                this.DefaultTabStop = this.query.Columns.Add(SrcConstants.TextBlockDefaultTabStop, nameof(this.DefaultTabStop));
                this.Background = this.query.Columns.Add(SrcConstants.TextBlockBackground, nameof(this.Background));
                this.BackgroundTransparency = this.query.Columns.Add(SrcConstants.TextBlockBackgroundTransparency, nameof(this.BackgroundTransparency));
                this.Direction = this.query.Columns.Add(SrcConstants.TextBlockDirection, nameof(this.Direction));
                this.VerticalAlign = this.query.Columns.Add(SrcConstants.TextBlockVerticalAlign, nameof(this.VerticalAlign));

            }

            public override Text.TextBlockCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<string> row)
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
}