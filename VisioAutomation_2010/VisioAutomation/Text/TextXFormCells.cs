using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;

namespace VisioAutomation.Text
{
    public class TextXFormCells : ShapeSheet.CellGroups.CellGroupSingleRow
    {
        public VisioAutomation.ShapeSheet.CellValueLiteral Angle { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Width { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Height { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral PinX { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral PinY { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LocPinX { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LocPinY { get; set; }

        public override IEnumerable<SrcFormulaPair> SrcFormulaPairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SrcConstants.TextXFormPinX, this.PinX.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.TextXFormPinY, this.PinY.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.TextXFormLocPinX, this.LocPinX.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.TextXFormLocPinY, this.LocPinY.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.TextXFormWidth, this.Width.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.TextXFormHeight, this.Height.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.TextXFormAngle, this.Angle.Value);
            }
        }

        public static List<TextXFormCells> GetCells(Microsoft.Office.Interop.Visio.Page page, IList<int> shapeids, VisioAutomation.ShapeSheet.CellValueType cvt)
        {
            var query = TextXFormCells.lazy_query.Value;
            return query.GetCellGroups(page, shapeids, cvt);
        }

        public static TextXFormCells GetCells(Microsoft.Office.Interop.Visio.Shape shape, VisioAutomation.ShapeSheet.CellValueType cvt)
        {
            var query = TextXFormCells.lazy_query.Value;
            return query.GetCellGroup(shape, cvt);
        }

        private static readonly System.Lazy<TextXFormCellsReader> lazy_query = new System.Lazy<TextXFormCellsReader>();
    }
}