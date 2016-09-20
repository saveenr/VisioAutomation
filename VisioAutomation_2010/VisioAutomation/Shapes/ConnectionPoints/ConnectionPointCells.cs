using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.CellGroups.Queries;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes.ConnectionPoints
{
    public class ConnectionPointCells : ShapeSheet.CellGroups.CellGroupMultiRow
    {
        public ShapeSheet.CellData X { get; set; }
        public ShapeSheet.CellData Y { get; set; }
        public ShapeSheet.CellData DirX { get; set; }
        public ShapeSheet.CellData DirY { get; set; }
        public ShapeSheet.CellData Type { get; set; }

        public override IEnumerable<SRCFormulaPair> Pairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SRCConstants.Connections_X, this.X.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Connections_Y, this.Y.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Connections_DirX, this.DirX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Connections_DirY, this.DirY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.Connections_Type, this.Type.Formula);
            }
        }

        public static List<List<ConnectionPointCells>> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = ConnectionPointCells.lazy_query.Value;
            return query.GetCellGroups(page, shapeids);
        }

        public static List<ConnectionPointCells> GetCells(IVisio.Shape shape)
        {
            var query = ConnectionPointCells.lazy_query.Value;
            return query.GetCellGroups(shape);
        }

        private static System.Lazy<ConnectionPointCellsQuery> lazy_query = new System.Lazy<ConnectionPointCellsQuery>();
    }
}