using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;

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

        public static IList<TextBlockCells> GetFormulas(IVisio.Page page, IList<int> shapeids)
        {
            var query = lazy_query.Value;
            return query.GetValues(page, shapeids, CellValueType.Formula);
        }

        public static IList<TextBlockCells> GetResults(IVisio.Page page, IList<int> shapeids)
        {
            var query = lazy_query.Value;
            return query.GetValues(page, shapeids, CellValueType.Result);
        }


        public static TextBlockCells GetFormulas(IVisio.Shape shape)
        {
            var query = lazy_query.Value;
            return query.GetValues(shape, CellValueType.Formula);
        }

        public static TextBlockCells GetResults(IVisio.Shape shape)
        {
            var query = lazy_query.Value;
            return query.GetValues(shape, CellValueType.Result);
        }

        private static readonly System.Lazy<TextBlockCellsReader> lazy_query = new System.Lazy<TextBlockCellsReader>();
    }
}