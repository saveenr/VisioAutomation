using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class ConnectionPointCells : ShapeSheet.CellGroups.CellGroupMultiRow
    {
        public ShapeSheet.CellData X { get; set; }
        public ShapeSheet.CellData Y { get; set; }
        public ShapeSheet.CellData DirX { get; set; }
        public ShapeSheet.CellData DirY { get; set; }
        public ShapeSheet.CellData Type { get; set; }

        public override IEnumerable<SrcFormulaPair> SrcFormulaPairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SrcConstants.ConnectionPointX, this.X.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ConnectionPointY, this.Y.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ConnectionPointDirX, this.DirX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ConnectionPointDirY, this.DirY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.ConnectionPointType, this.Type.Formula);
            }
        }

        public static List<List<ConnectionPointCells>> GetCells(IVisio.Page page, IList<int> shapeids, VisioAutomation.ShapeSheet.CellValueType cvt)
        {
            var query = ConnectionPointCells.lazy_query.Value;
            return query.GetCellGroups(page, shapeids, cvt);
        }

        public static List<ConnectionPointCells> GetCells(IVisio.Shape shape, VisioAutomation.ShapeSheet.CellValueType cvt)
        {
            var query = ConnectionPointCells.lazy_query.Value;
            return query.GetCellGroups(shape, cvt);
        }

        private static readonly System.Lazy<ConnectionPointCellsReader> lazy_query = new System.Lazy<ConnectionPointCellsReader>();
    }
}