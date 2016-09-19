using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.CellGroups.Queries;

namespace VisioAutomation.Text
{
    public class TextXFormCells : ShapeSheet.CellGroups.CellGroupSingleRow
    {
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
                yield return this.newpair(ShapeSheet.SRCConstants.TxtPinX, this.TxtPinX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.TxtPinY, this.TxtPinY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.TxtLocPinX, this.TxtLocPinX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.TxtLocPinY, this.TxtLocPinY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.TxtWidth, this.TxtWidth.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.TxtHeight, this.TxtHeight.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.TxtAngle, this.TxtAngle.Formula);
            }
        }

        public static IList<TextXFormCells> GetCells(Microsoft.Office.Interop.Visio.Page page, IList<int> shapeids)
        {
            var query = TextXFormCells.lazy_query.Value;
            return query.GetCellGroups(page, shapeids);
        }

        public static TextXFormCells GetCells(Microsoft.Office.Interop.Visio.Shape shape)
        {
            var query = TextXFormCells.lazy_query.Value;
            return query.GetCellGroup(shape);
        }

        private static System.Lazy<TextXFormCellsQuery> lazy_query = new System.Lazy<TextXFormCellsQuery>();
    }
}