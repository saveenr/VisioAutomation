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
        public ShapeSheet.CellData TextBkgnd { get; set; }
        public ShapeSheet.CellData TextBkgndTrans { get; set; }
        public ShapeSheet.CellData TextDirection { get; set; }
        public ShapeSheet.CellData VerticalAlign { get; set; }

        public override IEnumerable<SrcFormulaPair> SrcFormulaPairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SrcConstants.TextBlockBottomMargin, this.BottomMargin.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.TextBlockLeftMargin, this.LeftMargin.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.TextBlockRightMargin, this.RightMargin.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.TextBlockTopMargin, this.TopMargin.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.TextBlockDefaultTabStop, this.DefaultTabStop.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.TextBlockBackground, this.TextBkgnd.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.TextBlockBackgroundTransparency, this.TextBkgndTrans.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.TextBlockDirection, this.TextDirection.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.TextBlockVerticalAlign, this.VerticalAlign.Formula);
            }
        }

        public static IList<TextBlockCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = TextBlockCells.lazy_query.Value;
            return query.GetCellGroups(page, shapeids);
        }

        public static TextBlockCells GetCells(IVisio.Shape shape)
        {
            var query = TextBlockCells.lazy_query.Value;
            return query.GetCellGroup(shape);
        }

        private static readonly System.Lazy<TextBlockCellsReader> lazy_query = new System.Lazy<TextBlockCellsReader>();
    }
}