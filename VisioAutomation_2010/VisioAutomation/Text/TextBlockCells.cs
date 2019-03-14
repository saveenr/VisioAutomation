using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Text
{
    public class TextBlockCells : CellGroupBase
    {
        public CellValueLiteral BottomMargin { get; set; }
        public CellValueLiteral LeftMargin { get; set; }
        public CellValueLiteral RightMargin { get; set; }
        public CellValueLiteral TopMargin { get; set; }
        public CellValueLiteral DefaultTabStop { get; set; }
        public CellValueLiteral Background { get; set; }
        public CellValueLiteral BackgroundTransparency { get; set; }
        public CellValueLiteral Direction { get; set; }
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
                yield return SrcValuePair.Create(SrcConstants.TextBlockBackground, this.Background);
                yield return SrcValuePair.Create(SrcConstants.TextBlockBackgroundTransparency, this.BackgroundTransparency);
                yield return SrcValuePair.Create(SrcConstants.TextBlockDirection, this.Direction);
                yield return SrcValuePair.Create(SrcConstants.TextBlockVerticalAlign, this.VerticalAlign);
            }
        }

        public static IList<TextBlockCells> GetCells(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var reader = lazy_reader.Value;
            return reader.GetCellsSingleRow(page, shapeids, type);
        }

        public static TextBlockCells GetCells(IVisio.Shape shape, CellValueType type)
        {
            var reader = lazy_reader.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<TextBlockCellsReader> lazy_reader = new System.Lazy<TextBlockCellsReader>();

        class TextBlockCellsReader : CellGroupReader<Text.TextBlockCells>
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
                this.BottomMargin = this.query_singlerow.Columns.Add(SrcConstants.TextBlockBottomMargin, nameof(this.BottomMargin));
                this.LeftMargin = this.query_singlerow.Columns.Add(SrcConstants.TextBlockLeftMargin, nameof(this.LeftMargin));
                this.RightMargin = this.query_singlerow.Columns.Add(SrcConstants.TextBlockRightMargin, nameof(this.RightMargin));
                this.TopMargin = this.query_singlerow.Columns.Add(SrcConstants.TextBlockTopMargin, nameof(this.TopMargin));
                this.DefaultTabStop = this.query_singlerow.Columns.Add(SrcConstants.TextBlockDefaultTabStop, nameof(this.DefaultTabStop));
                this.Background = this.query_singlerow.Columns.Add(SrcConstants.TextBlockBackground, nameof(this.Background));
                this.BackgroundTransparency = this.query_singlerow.Columns.Add(SrcConstants.TextBlockBackgroundTransparency, nameof(this.BackgroundTransparency));
                this.Direction = this.query_singlerow.Columns.Add(SrcConstants.TextBlockDirection, nameof(this.Direction));
                this.VerticalAlign = this.query_singlerow.Columns.Add(SrcConstants.TextBlockVerticalAlign, nameof(this.VerticalAlign));

            }

            public override Text.TextBlockCells ToCellGroup(VisioAutomation.ShapeSheet.Internal.ArraySegment<string> row)
            {
                var cells = new Text.TextBlockCells();
                cells.BottomMargin = row[this.BottomMargin];
                cells.LeftMargin = row[this.LeftMargin];
                cells.RightMargin = row[this.RightMargin];
                cells.TopMargin = row[this.TopMargin];
                cells.DefaultTabStop = row[this.DefaultTabStop];
                cells.Background = row[this.Background];
                cells.BackgroundTransparency = row[this.BackgroundTransparency];
                cells.Direction = row[this.Direction];
                cells.VerticalAlign = row[this.VerticalAlign];
                return cells;
            }
        }
    }
}