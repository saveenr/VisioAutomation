using System.Collections.Generic;
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
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.TextBlockBottomMargin, this.BottomMargin.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.TextBlockLeftMargin, this.LeftMargin.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.TextBlockRightMargin, this.RightMargin.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.TextBlockTopMargin, this.TopMargin.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.TextBlockDefaultTabStop, this.DefaultTabStop.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.TextBlockBackground, this.TextBackground.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.TextBlockBackgroundTransparency, this.TextBackgroundTransparency.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.TextBlockDirection, this.TextDirection.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.TextBlockVerticalAlign, this.VerticalAlign.Value);
            }
        }

        public static IList<TextBlockCells> GetFormulas(IVisio.Page page, IList<int> shapeids)
        {
            var query = TextBlockCells.lazy_query.Value;
            return query.GetFormulas(page, shapeids);
        }

        public static IList<TextBlockCells> GetResults(IVisio.Page page, IList<int> shapeids)
        {
            var query = TextBlockCells.lazy_query.Value;
            return query.GetResults(page, shapeids);
        }


        public static TextBlockCells GetFormulas(IVisio.Shape shape)
        {
            var query = TextBlockCells.lazy_query.Value;
            return query.GetFormulas(shape);
        }

        public static TextBlockCells GetResults(IVisio.Shape shape)
        {
            var query = TextBlockCells.lazy_query.Value;
            return query.GetResults(shape);
        }

        private static readonly System.Lazy<TextBlockCellsReader> lazy_query = new System.Lazy<TextBlockCellsReader>();
    }
}