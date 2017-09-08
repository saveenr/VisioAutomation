using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
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

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.TextXFormPinX, this.PinX.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.TextXFormPinY, this.PinY.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.TextXFormLocPinX, this.LocPinX.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.TextXFormLocPinY, this.LocPinY.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.TextXFormWidth, this.Width.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.TextXFormHeight, this.Height.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.TextXFormAngle, this.Angle.Value);
            }
        }

        public static List<TextXFormCells> GetFormulas(Microsoft.Office.Interop.Visio.Page page, IList<int> shapeids)
        {
            var query = TextXFormCells.lazy_query.Value;
            return query.GetValues(page, shapeids, CellValueType.Formula);
        }

        public static List<TextXFormCells> GetResults(Microsoft.Office.Interop.Visio.Page page, IList<int> shapeids)
        {
            var query = TextXFormCells.lazy_query.Value;
            return query.GetValues(page, shapeids, CellValueType.Result);
        }


        public static TextXFormCells GetFormulas(Microsoft.Office.Interop.Visio.Shape shape)
        {
            var query = TextXFormCells.lazy_query.Value;
            return query.GetValues(shape, CellValueType.Formula);
        }

        public static TextXFormCells GetResults(Microsoft.Office.Interop.Visio.Shape shape)
        {
            var query = TextXFormCells.lazy_query.Value;
            return query.GetValues(shape, CellValueType.Result);
        }

        private static readonly System.Lazy<TextXFormCellsReader> lazy_query = new System.Lazy<TextXFormCellsReader>();
    }
}