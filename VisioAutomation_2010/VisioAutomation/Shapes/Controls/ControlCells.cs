using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.CellGroups.Queries;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes.Controls
{
    public class ControlCells : ShapeSheet.CellGroups.CellGroupMultiRow
    {
        public ShapeSheet.CellData CanGlue { get; set; }
        public ShapeSheet.CellData Tip { get; set; }
        public ShapeSheet.CellData X { get; set; }
        public ShapeSheet.CellData Y { get; set; }
        public ShapeSheet.CellData YBehavior { get; set; }
        public ShapeSheet.CellData XBehavior { get; set; }
        public ShapeSheet.CellData XDynamics { get; set; }
        public ShapeSheet.CellData YDynamics { get; set; }


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
            return query.GetCellGroups(page, shapeids);
        }

        public static IList<ControlCells> GetCells(IVisio.Shape shape)
        {
            var query = ControlCells.lazy_query.Value;
            return query.GetCellGroups(shape);
        }

        private static System.Lazy<ControlCellsQuery> lazy_query = new System.Lazy<ControlCellsQuery>();
    }
}