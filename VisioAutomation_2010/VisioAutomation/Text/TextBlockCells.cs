using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.CellGroups.Queries;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    public class TextBlockCells : ShapeSheet.CellGroups.CellGroupSingleRow
    {
        public ShapeSheet.CellData<double> BottomMargin { get; set; }
        public ShapeSheet.CellData<double> LeftMargin { get; set; }
        public ShapeSheet.CellData<double> RightMargin { get; set; }
        public ShapeSheet.CellData<double> TopMargin { get; set; }
        public ShapeSheet.CellData<double> DefaultTabStop { get; set; }
        public ShapeSheet.CellData<int> TextBkgnd { get; set; }
        public ShapeSheet.CellData<double> TextBkgndTrans { get; set; }
        public ShapeSheet.CellData<int> TextDirection { get; set; }
        public ShapeSheet.CellData<int> VerticalAlign { get; set; }

        public override IEnumerable<SRCFormulaPair> Pairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SRCConstants.BottomMargin, this.BottomMargin.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LeftMargin, this.LeftMargin.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.RightMargin, this.RightMargin.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.TopMargin, this.TopMargin.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.DefaultTabStop, this.DefaultTabStop.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.TextBkgnd, this.TextBkgnd.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.TextBkgndTrans, this.TextBkgndTrans.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.TextDirection, this.TextDirection.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.VerticalAlign, this.VerticalAlign.Formula);
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

        private static System.Lazy<TextBlockCellsQuery> lazy_query = new System.Lazy<TextBlockCellsQuery>();
    }
}