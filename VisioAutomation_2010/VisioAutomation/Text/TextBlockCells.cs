using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    public class TextBlockCells : ShapeSheet.CellGroups.CellGroupSingleRow
    {
        public VisioAutomation.ShapeSheet.CellValueLiteral BottomMargin { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LeftMargin { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral RightMargin { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral TopMargin { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral DefaultTabStop { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral TextBackground { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral TextBackgroundTransparency { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral TextDirection { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral VerticalAlign { get; set; }

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.TextBlockBottomMargin, this.BottomMargin);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.TextBlockLeftMargin, this.LeftMargin);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.TextBlockRightMargin, this.RightMargin);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.TextBlockTopMargin, this.TopMargin);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.TextBlockDefaultTabStop, this.DefaultTabStop);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.TextBlockBackground, this.TextBackground);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.TextBlockBackgroundTransparency, this.TextBackgroundTransparency);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.TextBlockDirection, this.TextDirection);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.TextBlockVerticalAlign, this.VerticalAlign);
            }
        }

        public static IList<TextBlockCells> GetFormulas(IVisio.Page page, IList<int> shapeids)
        {
            var query = TextBlockCells.lazy_query.Value;
            return query.GetValues(page, shapeids, CellValueType.Formula);
        }

        public static IList<TextBlockCells> GetResults(IVisio.Page page, IList<int> shapeids)
        {
            var query = TextBlockCells.lazy_query.Value;
            return query.GetValues(page, shapeids, CellValueType.Result);
        }


        public static TextBlockCells GetFormulas(IVisio.Shape shape)
        {
            var query = TextBlockCells.lazy_query.Value;
            return query.GetValues(shape, CellValueType.Formula);
        }

        public static TextBlockCells GetResults(IVisio.Shape shape)
        {
            var query = TextBlockCells.lazy_query.Value;
            return query.GetValues(shape, CellValueType.Result);
        }

        private static readonly System.Lazy<TextBlockCellsReader> lazy_query = new System.Lazy<TextBlockCellsReader>();
    }
}