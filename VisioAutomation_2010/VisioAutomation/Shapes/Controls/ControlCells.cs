using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes.Controls
{
    public class ControlCells : ShapeSheetQuery.QueryGroups.QueryGroupMultiRow
    {
        public ShapeSheet.CellData<int> CanGlue { get; set; }
        public ShapeSheet.CellData<int> Tip { get; set; }
        public ShapeSheet.CellData<double> X { get; set; }
        public ShapeSheet.CellData<double> Y { get; set; }
        public ShapeSheet.CellData<int> YBehavior { get; set; }
        public ShapeSheet.CellData<int> XBehavior { get; set; }
        public ShapeSheet.CellData<int> XDynamics { get; set; }
        public ShapeSheet.CellData<int> YDynamics { get; set; }


        public override IEnumerable<SRCFormulaPair> Pairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SRCConstants.Controls_CanGlue, this.CanGlue.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Controls_Tip, this.Tip.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Controls_X, this.X.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Controls_Y, this.Y.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Controls_YCon, this.YBehavior.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Controls_XCon, this.XBehavior.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Controls_XDyn, this.XDynamics.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Controls_YDyn, this.YDynamics.Formula);
            }
        }

        public static IList<List<ControlCells>> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = ControlCells.lazy_query.Value;
            return ShapeSheetQuery.QueryGroups.QueryGroupMultiRow._GetCells<ControlCells, double>(page, shapeids, query, query.GetCells);
        }

        public static IList<ControlCells> GetCells(IVisio.Shape shape)
        {
            var query = ControlCells.lazy_query.Value;
            return ShapeSheetQuery.QueryGroups.QueryGroupMultiRow._GetCells<ControlCells, double>(shape, query, query.GetCells);
        }

        private static System.Lazy<ShapeSheetQuery.CommonQueries.ControlCellsQuery> lazy_query = new System.Lazy<ShapeSheetQuery.CommonQueries.ControlCellsQuery>();
    }
}