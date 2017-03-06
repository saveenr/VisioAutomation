using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;

namespace VisioAutomation.Text
{
    public class TextXFormCells : ShapeSheet.CellGroups.CellGroupSingleRow
    {
        public ShapeSheet.CellData TxtAngle { get; set; }
        public ShapeSheet.CellData TxtWidth { get; set; }
        public ShapeSheet.CellData TxtHeight { get; set; }
        public ShapeSheet.CellData TxtPinX { get; set; }
        public ShapeSheet.CellData TxtPinY { get; set; }
        public ShapeSheet.CellData TxtLocPinX { get; set; }
        public ShapeSheet.CellData TxtLocPinY { get; set; }

        public override IEnumerable<SrcFormulaPair> SrcFormulaPairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SrcConstants.TextXFormPinX, this.TxtPinX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.TextXFormPinY, this.TxtPinY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.TextXFormLocPinX, this.TxtLocPinX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.TextXFormLocPinY, this.TxtLocPinY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.TextXFormWidth, this.TxtWidth.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.TextXFormHeight, this.TxtHeight.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.TextXFormAngle, this.TxtAngle.Formula);
            }
        }

        public static List<TextXFormCells> GetCells(Microsoft.Office.Interop.Visio.Page page, IList<int> shapeids)
        {
            var query = TextXFormCells.lazy_query.Value;
            return query.GetCellGroups(page, shapeids);
        }

        public static TextXFormCells GetCells(Microsoft.Office.Interop.Visio.Shape shape)
        {
            var query = TextXFormCells.lazy_query.Value;
            return query.GetCellGroup(shape);
        }

        private static readonly System.Lazy<TextXFormCellsReader> lazy_query = new System.Lazy<TextXFormCellsReader>();
    }
}