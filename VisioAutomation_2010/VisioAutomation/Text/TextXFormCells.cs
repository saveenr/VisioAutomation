using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;

namespace VisioAutomation.Text
{
    public class TextXFormCells : ShapeSheet.CellGroups.CellGroupSingleRow
    {
        public ShapeSheet.CellData Angle { get; set; }
        public ShapeSheet.CellData Width { get; set; }
        public ShapeSheet.CellData Height { get; set; }
        public ShapeSheet.CellData PinX { get; set; }
        public ShapeSheet.CellData PinY { get; set; }
        public ShapeSheet.CellData LocPinX { get; set; }
        public ShapeSheet.CellData LocPinY { get; set; }

        public override IEnumerable<SrcFormulaPair> SrcFormulaPairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SrcConstants.TextXFormPinX, this.PinX.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.TextXFormPinY, this.PinY.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.TextXFormLocPinX, this.LocPinX.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.TextXFormLocPinY, this.LocPinY.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.TextXFormWidth, this.Width.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.TextXFormHeight, this.Height.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.TextXFormAngle, this.Angle.ValueF);
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