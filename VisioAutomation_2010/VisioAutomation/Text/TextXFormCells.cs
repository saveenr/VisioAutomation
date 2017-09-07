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

        public static List<TextXFormCells> GetFormulas(Microsoft.Office.Interop.Visio.Page page, IList<int> shapeids)
        {
            var query = TextXFormCells.lazy_query.Value;
            return query.GetFormulas(page, shapeids);
        }

        public static List<TextXFormCells> GetResults(Microsoft.Office.Interop.Visio.Page page, IList<int> shapeids)
        {
            var query = TextXFormCells.lazy_query.Value;
            return query.GetResults(page, shapeids);
        }


        public static TextXFormCells GetFormulas(Microsoft.Office.Interop.Visio.Shape shape)
        {
            var query = TextXFormCells.lazy_query.Value;
            return query.GetFormulas(shape);
        }

        public static TextXFormCells GetResults(Microsoft.Office.Interop.Visio.Shape shape)
        {
            var query = TextXFormCells.lazy_query.Value;
            return query.GetResults(shape);
        }

        private static readonly System.Lazy<TextXFormCellsReader> lazy_query = new System.Lazy<TextXFormCellsReader>();
    }
}