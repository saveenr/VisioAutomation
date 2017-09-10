using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;

namespace VisioAutomation.Text
{
    public class TextXFormCells : CellGroupSingleRow
    {
        public CellValueLiteral Angle { get; set; }
        public CellValueLiteral Width { get; set; }
        public CellValueLiteral Height { get; set; }
        public CellValueLiteral PinX { get; set; }
        public CellValueLiteral PinY { get; set; }
        public CellValueLiteral LocPinX { get; set; }
        public CellValueLiteral LocPinY { get; set; }

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(SrcConstants.TextXFormPinX, this.PinX);
                yield return SrcValuePair.Create(SrcConstants.TextXFormPinY, this.PinY);
                yield return SrcValuePair.Create(SrcConstants.TextXFormLocPinX, this.LocPinX);
                yield return SrcValuePair.Create(SrcConstants.TextXFormLocPinY, this.LocPinY);
                yield return SrcValuePair.Create(SrcConstants.TextXFormWidth, this.Width);
                yield return SrcValuePair.Create(SrcConstants.TextXFormHeight, this.Height);
                yield return SrcValuePair.Create(SrcConstants.TextXFormAngle, this.Angle);
            }
        }

        public static List<TextXFormCells> GetFormulas(Microsoft.Office.Interop.Visio.Page page, IList<int> shapeids)
        {
            var query = lazy_query.Value;
            return query.GetValues(page, shapeids, CellValueType.Formula);
        }

        public static List<TextXFormCells> GetResults(Microsoft.Office.Interop.Visio.Page page, IList<int> shapeids)
        {
            var query = lazy_query.Value;
            return query.GetValues(page, shapeids, CellValueType.Result);
        }


        public static TextXFormCells GetFormulas(Microsoft.Office.Interop.Visio.Shape shape)
        {
            var query = lazy_query.Value;
            return query.GetValues(shape, CellValueType.Formula);
        }

        public static TextXFormCells GetResults(Microsoft.Office.Interop.Visio.Shape shape)
        {
            var query = lazy_query.Value;
            return query.GetValues(shape, CellValueType.Result);
        }

        private static readonly System.Lazy<TextXFormCellsReader> lazy_query = new System.Lazy<TextXFormCellsReader>();
    }
}