using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text
{
    public class TextBlockCells : ShapeSheet.CellGroups.CellGroup
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
        public ShapeSheet.CellData<double> TxtAngle { get; set; }
        public ShapeSheet.CellData<double> TxtWidth { get; set; }
        public ShapeSheet.CellData<double> TxtHeight { get; set; }
        public ShapeSheet.CellData<double> TxtPinX { get; set; }
        public ShapeSheet.CellData<double> TxtPinY { get; set; }
        public ShapeSheet.CellData<double> TxtLocPinX { get; set; }
        public ShapeSheet.CellData<double> TxtLocPinY { get; set; }

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
                yield return this.newpair(ShapeSheet.SRCConstants.TxtPinX, this.TxtPinX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.TxtPinY, this.TxtPinY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.TxtLocPinX, this.TxtLocPinX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.TxtLocPinY, this.TxtLocPinY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.TxtWidth, this.TxtWidth.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.TxtHeight, this.TxtHeight.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.TxtAngle, this.TxtAngle.Formula);
            }
        }

        public static IList<TextBlockCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = TextBlockCells.lazy_query.Value;
            return ShapeSheet.CellGroups.CellGroup._GetCells<TextBlockCells, double>(page, shapeids, query, query.GetCells);
        }

        public static TextBlockCells GetCells(IVisio.Shape shape)
        {
            var query = TextBlockCells.lazy_query.Value;
            return ShapeSheet.CellGroups.CellGroup._GetCells<TextBlockCells, double>(shape, query, query.GetCells);
        }

        private static System.Lazy<ShapeSheet.Query.Common.TextBlockCellsQuery> lazy_query = new System.Lazy<ShapeSheet.Query.Common.TextBlockCellsQuery>();
    }
}