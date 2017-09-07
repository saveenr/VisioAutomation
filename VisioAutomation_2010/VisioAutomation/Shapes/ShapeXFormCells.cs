using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class ShapeXFormCells : ShapeSheet.CellGroups.CellGroupSingleRow
    {
        public VisioAutomation.ShapeSheet.CellValueLiteral PinX { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral PinY { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LocPinX { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LocPinY { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Width { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Height { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Angle { get; set; }

        public override IEnumerable<SrcFormulaPair> SrcFormulaPairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SrcConstants.XFormPinX, this.PinX.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.XFormPinY, this.PinY.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.XFormLocPinX, this.LocPinX.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.XFormLocPinY, this.LocPinY.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.XFormWidth, this.Width.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.XFormHeight, this.Height.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.XFormAngle, this.Angle.Value);
            }
        }

        public static List<ShapeXFormCells> GetFormulas(IVisio.Page page, IList<int> shapeids)
        {
            var query = ShapeXFormCells.lazy_query.Value;
            return query.GetFormulas(page, shapeids);
        }

        public static List<ShapeXFormCells> GetResults(IVisio.Page page, IList<int> shapeids)
        {
            var query = ShapeXFormCells.lazy_query.Value;
            return query.GetResults(page, shapeids);
        }

        public static ShapeXFormCells GetFormulas(IVisio.Shape shape)
        {
            var query = ShapeXFormCells.lazy_query.Value;
            return query.GetFormulas(shape);
        }

        public static ShapeXFormCells GetResults(IVisio.Shape shape)
        {
            var query = ShapeXFormCells.lazy_query.Value;
            return query.GetResults(shape);
        }

        private static readonly System.Lazy<ShapeXFormCellsReader> lazy_query = new System.Lazy<ShapeXFormCellsReader>();
    }
}