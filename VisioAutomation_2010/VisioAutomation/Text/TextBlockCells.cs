using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    public class TextBlockCells : ShapeSheet.CellGroups.CellGroupSingleRow
    {
        public ShapeSheet.CellData BottomMargin { get; set; }
        public ShapeSheet.CellData LeftMargin { get; set; }
        public ShapeSheet.CellData RightMargin { get; set; }
        public ShapeSheet.CellData TopMargin { get; set; }
        public ShapeSheet.CellData DefaultTabStop { get; set; }
        public ShapeSheet.CellData TextBackground { get; set; }
        public ShapeSheet.CellData TextBackgroundTransparency { get; set; }
        public ShapeSheet.CellData TextDirection { get; set; }
        public ShapeSheet.CellData VerticalAlign { get; set; }

        public override IEnumerable<SrcFormulaPair> SrcFormulaPairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SrcConstants.TextBlockBottomMargin, this.BottomMargin.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.TextBlockLeftMargin, this.LeftMargin.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.TextBlockRightMargin, this.RightMargin.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.TextBlockTopMargin, this.TopMargin.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.TextBlockDefaultTabStop, this.DefaultTabStop.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.TextBlockBackground, this.TextBackground.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.TextBlockBackgroundTransparency, this.TextBackgroundTransparency.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.TextBlockDirection, this.TextDirection.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.TextBlockVerticalAlign, this.VerticalAlign.Value);
            }
        }

        public static IList<TextBlockCells> GetCells(IVisio.Page page, IList<int> shapeids, VisioAutomation.ShapeSheet.CellValueType cvt)
        {
            var query = TextBlockCells.lazy_query.Value;
            return query.GetCellGroups(page, shapeids, cvt);
        }

        public static TextBlockCells GetCells(IVisio.Shape shape, VisioAutomation.ShapeSheet.CellValueType cvt)
        {
            var query = TextBlockCells.lazy_query.Value;
            return query.GetCellGroup(shape, cvt);
        }

        private static readonly System.Lazy<TextBlockCellsReader> lazy_query = new System.Lazy<TextBlockCellsReader>();
    }
}