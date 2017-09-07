using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
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

        public override IEnumerable<SrcFormulaPair> SrcFormulaPairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SrcConstants.ControlCanGlue, this.CanGlue.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ControlTip, this.Tip.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ControlX, this.X.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ControlY, this.Y.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ControlYBehavior, this.YBehavior.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ControlXBehavior, this.XBehavior.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ControlXDynamics, this.XDynamics.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.ControlYDynamics, this.YDynamics.Value);
            }
        }

        public static List<List<ControlCells>> GetCells(IVisio.Page page, IList<int> shapeids, VisioAutomation.ShapeSheet.CellValueType cvt)
        {
            var query = ControlCells.lazy_query.Value;
            return query.GetCellGroups(page, shapeids, cvt);
        }

        public static List<ControlCells> GetCells(IVisio.Shape shape, VisioAutomation.ShapeSheet.CellValueType cvt)
        {
            var query = ControlCells.lazy_query.Value;
            return query.GetCellGroups(shape, cvt);
        }

        private static readonly System.Lazy<ControlCellsReader> lazy_query = new System.Lazy<ControlCellsReader>();
    }
}